using System;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;
using MGCPCB;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.VisualBasic;
using System.Windows.Forms;

namespace xPCB_BankPainter
{
    class Program
    {
        static MGCPCB.Document pcbDoc;
        static MGCPCB.UserLayer FindOrAddLayer( string LayerName, MGCPCB.Document _doc )
        {
            UserLayer _lay = _doc.FindUserLayer(LayerName);
            if ( _lay == null )
            {
                _lay = _doc.SetupParameter.PutUserLayer(LayerName);
            }
            return _lay;
        }

        [STAThread]
        static void Main( string[] args )
        {
            #region Instance Connection Code
            try
            {
                MGCPCBReleaseEnvironmentLib.IMGCPCBReleaseEnvServer _server =
                    (MGCPCBReleaseEnvironmentLib.IMGCPCBReleaseEnvServer)Activator.CreateInstance(
                        Marshal.GetTypeFromCLSID(
                            new Guid("44983CB8-19B0-4695-937A-6FF0B74ECFC5")
                        )
                    );


                _server.SetEnvironment("");
                string VxVersion = _server.sddVersion;
                string strSDD_HOME = _server.sddHome;
                int length = strSDD_HOME.IndexOf("SDD_HOME");
                strSDD_HOME = strSDD_HOME.Substring(0, length).Replace("\\", "\\\\") + "SDD_HOME";
                _server.SetEnvironment(strSDD_HOME);
                string progID = _server.ProgIDVersion;

                object[,] _releases = (object[,])_server.GetInstalledReleases();
                dynamic pcbApp = null;

                for (int i = 1; i < _releases.Length / 4; i++)
                {
                    string _com_version = Convert.ToString(_releases[i, 0]);
                    try
                    {
                        pcbApp = Interaction.GetObject(null, "MGCPCB.Application." + _com_version);

                        pcbDoc = pcbApp.ActiveDocument;
                        dynamic licApp = Interaction.CreateObject("MGCPCBAutomationLicensing.Application." + _com_version);
                        int _token = licApp.GetToken(pcbDoc.Validate(0));
                        pcbDoc.Validate(_token);

                        break;
                    }
                    catch (Exception m)
                    {
                    }
                }


                if (pcbApp == null)
                {
                    System.Windows.Forms.MessageBox.Show("Could not found active Xpedition or PADSPro Application");
                    System.Environment.Exit(1);
                }

                

            }
            catch (Exception m)
            {
                MessageBox.Show(m.Message + "\r\n" + m.Source + "\r\n" + m.StackTrace);
            }
            #endregion

            #region Work Code
            MGCSDDOUTPUTWINDOWLib.MGCSDDOutputLogControl msgWnd = null;
            MGCSDDOUTPUTWINDOWLib.HtmlCtrl _tabCtrl = null;

            foreach (dynamic addin in (dynamic)pcbDoc.Application.Addins)
            {
                if (addin.Name == "Message Window")
                {
                    Console.WriteLine(addin.Control);
                    addin.Visible = true;
                    msgWnd = addin.Control;
                }
            }

            if (msgWnd != null)
            {
                _tabCtrl = msgWnd.AddTab("Paint Gates");
                _tabCtrl.Clear();
                _tabCtrl.Activate();
            }

            var addText = new Action<string>(text =>
            {
                if(_tabCtrl != null)
                {
                    _tabCtrl.AppendText(text + "\r\n");
                }
            });

            var addHtml = new Action<string>(html =>
            {
                if (_tabCtrl != null)
                {
                    _tabCtrl.AppendHTML(html);
                }
            });



            addText("*THIS CODE IS A OPEN-SOURCE SOFTWARE UNDER MIT LICENSE");
            addText("*PROVIDED \"AS IS\" WITHOUT WARRANTY OF ANY KIND,");
            addText("*EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED");
            addText("*WARRANTIES OF MERCHANTABILITY AND / OR FITNESS FOR A PARTICULAR PURPOSE.");
            addText("");
            addHtml("<p style=\"{color: red; font-weight: bold;}\">Copyright 2008-2020; Milbitt Engineering. All rights reserved.<br />" +
                    "<a href=\"https://www.milbitt.com\">www.milbitt.com</a><br />" +
                    "<a href=\"mailto:info@milbitt.com\">info@milbitt.com</a></p>");
            // Everything is OK, then....
            System.Threading.Thread.Sleep(2000);
            
            MGCPCB.Components _comps = pcbDoc.get_Components(EPcbSelectionType.epcbSelectSelected);
            if ( _comps.Count != 1 )
            {
                pcbDoc.Application.Gui.StatusBarText("Select only one component", EPcbStatusField.epcbStatusFieldError);
                addText("Error: Pick only one component");
                return;
            }

            switch (pcbDoc.Application.Gui.DisplayMessage(
                "After the process starts, you will not able to abort in half nor undo it. " +
                "Saving the current state of design is recommended.\r\n\r\n" +
                "Do you want to save the design before process start?", 
                "", 
                EPcbMsgBoxType.epcbMsgBoxYesNoCancel, 
                EPcbMsgBoxIcon.epcbMsgBoxIconExclamation))
            {
                case EPcbMsgBoxAnswer.epcbMsgAnsYes: pcbDoc.Save(); break;
                case EPcbMsgBoxAnswer.epcbMsgAnsNo: break;
                case EPcbMsgBoxAnswer.epcbMsgAnsCancel: pcbDoc.Application.Gui.ProgressBarInitialize(false); System.Environment.Exit(0); break;
            }
            pcbDoc.Application.Gui.ProgressBarInitialize(false, "", 0, 0);
            pcbDoc.Application.Gui.ProgressBarInitialize(true, "Starting Gate Painter", 100, 0);
            addText("Starting Gate Painter");

            try
            {
                MGCPCB.Component _comp = _comps[1];
                //Console.WriteLine("Found part: " + _comp.PartNumber);
                addText("Found part: " + _comp.PartNumber);
                string _celName = _comp.CellName;
                pcbDoc.Application.Gui.ProgressBarInitialize(true, "Gate Painter: Opening PDB Editor", 100, 0);
                addText("Opening PDB Editor");
                pcbDoc.Application.Gui.ProgressBar(10);

                Dictionary<string, string> PinBankList = new Dictionary<string, string>();
                MGCPCBPartsEditor.PartsEditorDlg _peDlg = (MGCPCBPartsEditor.PartsEditorDlg)pcbDoc.PartEditor;
                List<string> symbolNames = new List<string>();
                foreach (MGCPCBPartsEditor.Part _part in _peDlg.ActiveDatabaseEx.ActivePartition.get_Parts(MGCPCBPartsEditor.EPDBPartType.epdbPartAll, _comp.PartNumber))
                {
                    double _percent = 20.0 / _part.PinMapping.Gates.Count;
                    double _stat = 0;
                    foreach (MGCPCBPartsEditor.Gate _gate in _part.PinMapping.Gates)
                    {

                        pcbDoc.Application.Gui.ProgressBarInitialize(true, "Gate Painter: Fetching Gate Informations from PDB", 100, 0);
                        string _symbolName = _gate.SymbolReferences[1].Name.Split(':').LastOrDefault();
                        symbolNames.Add(_symbolName);
                        addText("Fetching Gate Informations from PDB:" + _symbolName);
                        pcbDoc.Application.Gui.ProgressBar(15 + Convert.ToInt16(_stat));
                        
                        foreach (MGCPCBPartsEditor.Slot _slot in _gate.Slots)
                        {
                            foreach (MGCPCBPartsEditor.PinInstance _pin in _slot.Pins)
                            {
                                PinBankList.Add(_pin.Number, _symbolName);
                            }
                        }

                        _stat += _percent;
                    }
                    break;
                }
                _peDlg.Quit();

                if (PinBankList.Count == 0)
                {
                    pcbDoc.Application.Gui.StatusBarText("There's no gate definition found in PDB. Operation halted.", EPcbStatusField.epcbStatusFieldError);
                    addText("Error: There's no gate definition found in PDB. Operation halted.");
                    return;
                }

                addText("Creating Missing User Layers");
                try
                {
                    pcbDoc.TransactionStart(EPcbDRCMode.epcbDRCModeNone);
                    foreach (string symName in symbolNames)
                    {
                        if(pcbDoc.FindUserLayer(symName) == null)
                            pcbDoc.SetupParameter.PutUserLayer(symName);
                    }                                                
                }
                catch
                {

                }
                finally
                {
                    pcbDoc.TransactionEnd();
                }

                CellEditorAddinLib.CellEditorDlg _dlg = (CellEditorAddinLib.CellEditorDlg)pcbDoc.CellEditor;
                pcbDoc.Application.Gui.ProgressBarInitialize(true, "Gate Painter: Opening Cell Editor", 100, 0);
                addText("Opening Cell Editor");
                pcbDoc.Application.Gui.ProgressBar(40);
                CellEditorAddinLib.Cells _designCells = _dlg.ActiveDatabase.ActivePartition.get_Cells();
                foreach (CellEditorAddinLib.Cell _cell in _designCells)
                {
                    if (_cell.Name == _celName)
                    {

                        pcbDoc.Application.Gui.ProgressBarInitialize(true, "Gate Painter: Cell Found, Opening", 100, 0);
                        addText("Cell Found, Opening");
                        pcbDoc.Application.Gui.ProgressBar(45);
                        MGCPCB.Document _cellDoc = (MGCPCB.Document)_cell.Edit();
                        _cellDoc.Application.Visible = false;

                        try
                        {
                            _cellDoc.TransactionStart(EPcbDRCMode.epcbDRCModeNone);
                            pcbDoc.Application.Gui.ProgressBarInitialize(true, "Gate Painter: Indexing Pins", 100, 0);
                            addText("Indexing Pins");
                            pcbDoc.Application.Gui.ProgressBar(50);
                            Dictionary<string, string> _pinIds = new Dictionary<string, string>();
                            foreach (MGCPCB.Pin _pin in _cellDoc.get_Pins(EPcbSelectionType.epcbSelectAll))
                            {
                                _pinIds.Add(_pin.Name, Convert.ToString(_pin.Net.UniqueId));
                            }

                            PinBankList.OrderBy(x => x.Value);

                            UserLayer _lastLayer = null;
                            double _percent = 40.0 / PinBankList.Count;
                            double _stat = 0;
                            pcbDoc.Application.Gui.ProgressBarInitialize(true, "Gate Painter: Generating Shapes", 100, 0);
                            addText("Generating Shapes");
                            pcbDoc.Application.Gui.ProgressBar(55);
                            int progressCounter = 0;
                            foreach (KeyValuePair<string, string> _pinSym in PinBankList)
                            {
                                progressCounter++;
                                if (progressCounter % 50 == 0)
                                {
                                    pcbDoc.Application.Gui.ProgressBar(Convert.ToInt16(55 + _stat));
                                    addText("Generating Shapes, Progress: " + Convert.ToInt16(55 + _stat) + "%");
                                }
                                if (_lastLayer != null)
                                {
                                    if (_lastLayer.Name != _pinSym.Value)
                                        _lastLayer = FindOrAddLayer(_pinSym.Value, _cellDoc);
                                }
                                else
                                    _lastLayer = FindOrAddLayer(_pinSym.Value, _cellDoc);


                                string _pinId = ((KeyValuePair<string, string>)_pinIds.Single(x => x.Key == _pinSym.Key)).Value;

                                MGCPCB.Pin _cellPin = ((MGCPCB.Net)_cellDoc.FindNetByID(_pinId)).get_Pins(EPcbSelectionType.epcbSelectAll)[1];

                                MGCPCB.Geometry _geom = _cellPin.FabricationPads[1].Geometries[1];
                                _cellDoc.PutUserLayerGfx(
                                    _lastLayer, 
                                    0, 
                                    ((object[,])_geom.get_PointsArray(EPcbUnit.epcbUnitCurrent)).Length/3, 
                                    ((object[,])_geom.get_PointsArray(EPcbUnit.epcbUnitCurrent)), 
                                    true,
                                    null, 
                                    EPcbUnit.epcbUnitCurrent
                                );

                                _stat += _percent;
                            }
                            pcbDoc.Application.Gui.ProgressBarInitialize(true, "Gate Painter: Saving Cell", 100, 0);
                            
                            
                        }
                        catch (Exception m)
                        {
                            addHtml("<p style=\"{color: red; font-weight: bold;}\">Error! " + m.Message + "<br/>" + m.Source + "<br/>" + m.StackTrace + "</p>");
                        }
                        finally
                        {
                            _cellDoc.TransactionEnd();
                            addText("Saving Cell");
                            _cellDoc.Save();
                            addText("Closing Cell Editor");
                            pcbDoc.Application.Gui.ProgressBar(95);
                            _cellDoc.Application.Quit();
                            _dlg.SaveActiveDatabase();
                            _dlg.Quit();
                            pcbDoc.Application.Gui.ProgressBar(100);

                            
                            pcbDoc.Application.Gui.ProgressBarInitialize(true, "Gate Painter: Finished", 100, 0);
                            addHtml("<p style=\"{color: #006600; font-weight: bold;}\">Operation Competed</p>");
                            addText("Don't forget making user layers visible!");
                            addText("");
                            addHtml("<p style=\"{color: red; font-weight: bold;}\">Copyright 2008-2020; Milbitt Engineering. All rights reserved.<br />" +
                                "<a href=\"www.milbitt.com\">www.milbitt.com</a></p>");
                            addText("");
                            
                        }

                        break;
                        

                    }

                }
            }
            catch (Exception m)
            {
                addHtml("<p style=\"{color: red; font-weight: bold;}\">Error! " + m.Message + "<br/>" + m.Source + "<br/>" + m.StackTrace + "</p>");
            }
            finally
            {
                pcbDoc.Application.Gui.ProgressBarInitialize(false);
            }
            #endregion 
        }
    }
}

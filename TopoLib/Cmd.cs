//
// Copyright (c) 2020 - 2021 by Bart Duijndam. See: https://www.duijndam.dev 
//
// Licensed under the Apache License, Version 2.0 (the "License"); 
// You may not use this file except in compliance with the License.
// You may obtain a License copy at http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software distributed under the License is
// distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. 
//
// See the License for the specific language governing permissions and limitations under the License.
//
using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using ExcelDna.Documentation;
using ExcelDna.XlDialogBox;
//using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

// Added Bart
using SharpProj;
using SharpProj.Proj;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace TopoLib
{
    public class R1C1
    {
        private readonly string _r1c1;

        public R1C1(string rc) { _r1c1 = rc; }

        public int Top()
        {
            if (_r1c1.Length == 0)
                return 0;

            string[] parts = _r1c1.Split(':');
            if (parts.Length == 0) return -1;

            int[] rows = new int[parts.Length];
            for (int i = 0; i < parts.Length; i++)
            {
                string[] parts2 = parts[i].Split('r','c','R','C');

                string row = parts2[1];
                rows[i] = Convert.ToInt32(row);
            }
            return rows[0];
        }

        public int Rows()
        {
            if (_r1c1.Length == 0)
                return 0;

            string[] parts = _r1c1.Split(':');
            if (parts.Length == 1) return 1;

            int[] rows = new int[2];
            for (int i = 0; i < 2; i++)
            {
                string[] parts2 = parts[i].Split('r','c','R','C');

                string row = parts2[1];
                rows[i] = Convert.ToInt32(row);
            }
            return rows[1] - rows[0] + 1;
        }

        public int Left()
        {
            if (_r1c1.Length == 0)
                return 0;

            string[] parts = _r1c1.Split(':');
            if (parts.Length == 0) return -1;

            int[] cols = new int[parts.Length];
            for (int i = 0; i < parts.Length ; i++)
            {
                string[] parts2 = parts[i].Split('r','c','R','C');

                string col = parts2[2];
                cols[i] = Convert.ToInt32(col);
            }
            return cols[0];
        }

        public int Cols()
        {
            if (_r1c1.Length == 0)
                return 0;

            string[] parts = _r1c1.Split(':');
            if (parts.Length == 1) return 1;

            int[] cols = new int[2];
            for (int i = 0; i < 2; i++)
            {
                string[] parts2 = parts[i].Split('r','c','R','C');

                string col = parts2[2];
                cols[i] = Convert.ToInt32(col);
            }
            return cols[1] - cols[0] + 1;
        }
    }

    public static class Cmd
    {
        // used in the dialogs to export gpx & kml data
        static string gpxRange = "";
        static string kmlRange = "";

        static string SourceCrsRange = "";
        static string TargetCrsRange = "";

        static string SourceAuthority = "EPSG";
        static string TargetAuthority = "EPSG";

        static string FileName = "";
        static string FolderName = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        static int FilterIndex = 1;

        static int SourceAuthIndex = 0; // "EPSG";
        static int TargetAuthIndex = 0; // "EPSG";

        static int SourceCode = 4326;
        static int TargetCode = 4326;

        static int ExportDataType = 0;
        static int DataInputFrom = 0;

        /// <summary>
        /// This is a dummy validation routine
        /// Validation routines only matter if you use a trigger on a control within an XlDialogBox
        /// </summary>
        /// <param name="index">the row index of the control that caused a trigger</param>
        /// <param name="dialogResult">the object array, that the Dialog worked with</param>
        /// <param name="Controls">the collection of controls, that can be edited in the callback function</param>
        /// <returns>
        /// return true, to show the dialog (again) with the updated control settings
        /// return false, if no more changes need to be made
        /// return false will have the same effect as pressing the OK button
        /// </returns>
        static bool Validate(int index, object[,] dialogResult, XlDialogBox.XlDialogControlCollection Controls)
        {
            // just some code to set a break point
            // int i = index;

            return true; // return to dialog
        }

        // app.config can be updated to define dpiAware and/or dpiAwareness, dealing with display scaling.
        // https://docs.microsoft.com/en-us/windows/win32/hidpi/setting-the-default-dpi-awareness-for-a-process
        // https://docs.microsoft.com/en-us/windows/win32/direct2d/how-to--size-a-window-properly-for-high-dpi-displays
        // https://docs.microsoft.com/en-us/windows/win32/sbscs/application-manifests
        // https://docs.microsoft.com/en-us/windows/win32/hidpi/declaring-managed-apps-dpi-aware
        // https://docs.microsoft.com/en-us/windows/win32/hidpi/declaring-managed-apps-dpi-aware#updating-an-existing-wpf-application-to-be-per-monitor-dpi-aware-using-helper-project-in-the-wpf-sample

        // https://stackoverflow.com/questions/13228185/how-to-configure-an-app-to-run-correctly-on-a-machine-with-a-high-dpi-setting-e 
        // this pages provides as an example:
        // <application xmlns="urn:schemas-microsoft-com:asm.v3">
        //   <windowsSettings>
        //     <dpiAware xmlns="http://schemas.microsoft.com/SMI/2005/WindowsSettings">true</dpiAware>
        //   </windowsSettings>
        // </application>

        // use the following link to get info around dpi scaling:
        // https://dzimchuk.net/best-way-to-get-dpi-value-in-wpf/
        // implement this into scaling of the dialog window

        // For showing Message Boxes with Excel-DNA see; https://andysprague.com/2017/07/03/show-message-boxes-with-excel-dna/

        /// <summary>
        /// This is a validation routine for the About TopoLib routine
        /// </summary>
        /// <param name="index">the row index of the control that caused a trigger</param>
        /// <param name="dialogResult">the object array, that the Dialog worked with</param>
        /// <param name="Controls">the collection of controls, that can be edited in the callback function</param>
        /// <returns>
        /// return true, to show the dialog (again) with the updated control settings
        /// return false, if no more changes need to be made
        /// return false will have the same effect as pressing the OK button
        /// </returns>
        static bool ValidateAboutDialog(int index, object[,] dialogResult, XlDialogBox.XlDialogControlCollection Controls)
        {
            System.Diagnostics.Process.Start("https://www.github.com/duijndam-dev/");
            return true; // return to dialog
        }

        /// <summary>
        /// Validation routines only matter if you use a trigger on a control within an XlDialogBox
        /// </summary>
        /// <param name="index">the row index of the control that caused a trigger</param>
        /// <param name="dialogResult">the object array, that the Dialog worked with</param>
        /// <param name="Controls">the collection of controls, that can be edited in the callback function</param>
        /// <returns>
        /// return true, to show the dialog (again) with the updated control settings
        /// return false, if no more changes need to be made
        /// return false will have the same effect as pressing the OK button
        /// </returns>
        static bool ValidateCacheDialog(int index, object[,] dialogResult, XlDialogBox.XlDialogControlCollection Controls)
        {
            // just some code to set a break point
            // int i = index;

            if (index == 2)
            {
                OpenFileDialog dialog = new OpenFileDialog
                {
                    // dialog.CheckFileExists = true;  
                    // dialog.CheckPathExists = true;  
                    InitialDirectory = Path.GetDirectoryName(dialogResult[3, 6].ToString()),
                    FileName = Path.GetFileName(dialogResult[3, 6].ToString()),
                    Title = "Select PROJ Cache File",
                    RestoreDirectory = true,
                    Filter = "Database files (*.db)|*.db|All files (*.*)|*.*"
                };

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    if (dialog.FileName != dialogResult[3,6].ToString())
                    {
                        CctOptions.CachePath = dialog.FileName; // Get the complete path
                        Controls[3].IO = CctOptions.CachePath;  // Stick it back in the control list
                    }
                }
            }
            else if (index == 6)
            {
                DialogResult result = MessageBox.Show
                    ("Do you really want do delete all grid-file information stored in the PROJ cache file ?", 
                    "Please confirm",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);

                if (result == DialogResult.Yes)
                {
                    CctOptions.ProjContext.ClearGridCache();
                }
            }

            return true; // return to dialog
        }

        // See discussion here https://stackoverflow.com/questions/11624298/how-to-use-openfiledialog-to-select-a-folder
        /// <summary>
        /// Validation routines only matter if you use a trigger on a control within an XlDialogBox
        /// </summary>
        /// <param name="index">the row index of the control that caused a trigger</param>
        /// <param name="dialogResult">the object array, that the Dialog worked with</param>
        /// <param name="Controls">the collection of controls, that can be edited in the callback function</param>
        /// <returns>
        /// return true, to show the dialog (again) with the updated control settings
        /// return false, if no more changes need to be made
        /// return false will have the same effect as pressing the OK button
        /// </returns>
        static bool ValidateResourceDialog(int index, object[,] dialogResult, XlDialogBox.XlDialogControlCollection Controls)
        {
            // read the configuration settings to chose between file pickers being used.
            // At this point, all functions have been registered, and it should be safe to do so...
            string sFolderPicker = (string)Cfg.GetKeyValue("FolderPicker", "1");
            int.TryParse(sFolderPicker, out int nFolderPicker);

            string sSelectedPath = dialogResult[2, 6].ToString();

            if (index == 3) // update button
            {
                switch (nFolderPicker)
                {
                    default:
                    case 0:
                        {
                            // can't use 'using' as FolderPicker hasn't been derived from IDisposable
                            var dlg = new FolderPicker
                            {
                                InputPath = sSelectedPath,
                                Title = "Select PROJ resources folder"
                            };
                            if (dlg.ShowDialog(Process.GetCurrentProcess().MainWindowHandle) == true)
                            {
                                sSelectedPath = dlg.ResultPath;
                            }

                        }
                        break;
                    case 1:
                        using (var dialog = new OpenFileDialog())
                        {
                            dialog.Title = "Select any resource file to confirm PROJ_LIB folder";
                            dialog.InitialDirectory = sSelectedPath;
                            dialog.ValidateNames = false;
                            dialog.CheckFileExists = false;
                            dialog.CheckPathExists = true;
                            dialog.Filter = "GeoTiff files (*.tif)|*.tif|All files (*.*)|*.*";

                            // Always default to Folder Selection.
                            if (dialog.ShowDialog() == DialogResult.OK)
                            {
                                sSelectedPath = Path.GetDirectoryName(dialog.FileName);
                            }
                        }
                        break;
                    case 2:
                        using (var dialog = new FolderBrowserDialog())
                        {
                            dialog.SelectedPath = sSelectedPath;
                            dialog.Description = "Select PROJ resources folder";
                            dialog.ShowNewFolderButton = false;

                            DialogResult result = dialog.ShowDialog();

                            if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.SelectedPath))
                            {
                                sSelectedPath = dialog.SelectedPath;
                            }
                        }
                        break;
                } // end switch

                Controls[2].IO = sSelectedPath;

            } // end index == 3
            else if (index == 6)
            {
                Controls[5].IO = "https://cdn.proj.org";
            }

            return true; // return to dialog
        }

        /// <summary>
        /// Validation routines only matter if you use a trigger on a control within an XlDialogBox
        /// </summary>
        /// <param name="index">the row index of the control that caused a trigger</param>
        /// <param name="dialogResult">the object array, that the Dialog worked with</param>
        /// <param name="Controls">the collection of controls, that can be edited in the callback function</param>
        /// <returns>
        /// return true, to show the dialog (again) with the updated control settings
        /// return false, if no more changes need to be made
        /// return false will have the same effect as pressing the OK button
        /// </returns>
        static bool ValidateTransformDialog(int index, object[,] dialogResult, XlDialogBox.XlDialogControlCollection Controls)
        {
            // We need to reset all values in the Controls...
            Controls[ 2].IO = 1;
            Controls[ 8].IO = "EPSG";
            Controls[10].IO = "-1";
            Controls[13].IO = -1000;
            Controls[15].IO = -1000;
            Controls[17].IO = -1000;
            Controls[19].IO = -1000;

            Controls[21].IO = false;
            Controls[22].IO = false;
            Controls[23].IO = false;
            Controls[24].IO = false;
            Controls[25].IO = false;
            Controls[26].IO = false;
            Controls[27].IO = 1;

            return true; // return to dialog
        } // ValidateTransformDialog

        // See also https://stackoverflow.com/questions/40574084/fastest-method-to-remove-empty-rows-and-columns-from-excel-files-using-interop/40726309#40726309
        // And: https://groups.google.com/g/exceldna/c/cu4mRb1UolY/m/ux0y0JnjDwAJ
        [ExcelCommand(
            Name = "Command_Recalculate_Transforms",
            Description = "Recalculates the TopoLib transform functions (only TL.cct.xxx functions)",
            HelpTopic = "TopoLib-AddIn.chm!1201")]
        public static void Command_RecalculateTransforms()
        {
            // See: https://groups.google.com/g/exceldna/c/cu4mRb1UolY/m/ux0y0JnjDwAJ

            // Get the correct application instance
            Microsoft.Office.Interop.Excel.Application xlapp = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;

            // Get active workbook
            Excel.Workbook WorkBook = xlapp.ActiveWorkbook;
            dynamic ActiveSheet = WorkBook.ActiveSheet;

            dynamic formulaCells = null;
            Excel.Range usedRange = null;

            try
            {
                usedRange = ActiveSheet.UsedRange;
                formulaCells = usedRange.SpecialCells(XlCellType.xlCellTypeFormulas, Type.Missing);
            }
            catch (COMException ex)
            {
                if (ex.HResult != -2146827284)
                    throw ex;
            }
/*
            bool bArrayFormulas = false;
            int i = 0;
            if (formulaCells != null)
            {
                foreach (var range in formulaCells)
                {
                    if (i % 10 == 0)
                    {
                        XlCall.Excel(XlCall.xlcMessage, true, $"Refreshed {i} TopoLib routines");
                        
                        bool abort = (bool)XlCall.Excel(XlCall.xlAbort, true);
                        if (abort)
                        {
                            // XlCall.Excel(XlCall.xlcMessage, false);
                            XlCall.Excel(XlCall.xlcMessage, true, $"Interupted after {i} TopoLib routines");
                            return;
                        }
                    }

                    range.Replace("=TL.", "=TL.", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, true, false, false, false);
                    i++;
                }

                // XlCall.Excel(XlCall.xlcMessage, false);
                XlCall.Excel(XlCall.xlcMessage, true, $"Refreshed {i} TopoLib routines");
            }

*/


            try
            {
                // usedRange = ActiveSheet.UsedRange;
                
                // do something with myRange to split usedRange in segments...
                // Excel.Range myRange = ActiveSheet.Range.FromLTRB(1, 1, 11, 11);

                // async functions don't block the UI.
                // maybe use : https://bettersolutions.com/csharp/excel-interop/excel-dna-excelasyncutil-run.htm
                // Some examples here : https://dotnetfiddle.net/kIg2zd https://easysavecode.com/m6JUuoan
                // Se also: https://stackoverflow.com/questions/45199622/using-excel-region-get-range-within-a-range-gives-unexpected-behavior
                
                // Only update the transform functions; the crs functions aren't affected by the optional parameters, that could change in the background.
                ExcelAsyncUtil.QueueAsMacro(() => { bool v = usedRange.Replace("=TL.cct", "=TL.cct", XlLookAt.xlPart, XlSearchOrder.xlByRows, true, false, false, false); });

            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                if (ex.HResult != -2146827284)
                    throw ex;
            }

        } // Recalculate_TopoLib_Transforms


        [ExcelCommand(
            Name = "Command_Show_HelpFile",
            Description = "Shows the Compiled Help file",
            HelpTopic = "TopoLib-AddIn.chm!1202")]
        public static void Command_ShowHelpFile()
        {
            // get the Path of xll file;
            string xllPath = ExcelDnaUtil.XllPath;
            string xllDir  = System.IO.Path.GetDirectoryName(xllPath);

            var CallingMethod = System.Reflection.MethodBase.GetCurrentMethod();
            if (CallingMethod != null)
            {   // is there an ExcelCommandAttribute attribute decorating the method where ShowDialog has been called from ?
                ExcelCommandAttribute attr = (ExcelCommandAttribute)CallingMethod.GetCustomAttributes(typeof(ExcelCommandAttribute), true)[0];
                if (attr != null)
                {
                    // get the HelpTopic string and split it in two parts ([a] file name and [b] helptopic)
                    string[] parts = attr.HelpTopic.Split('!');

                    // the complete helpfile path consists of the xll directory + first part of HelpTopic attribute string 
                    string chmPath = System.IO.Path.Combine(xllDir, parts[0]);

                    // don't bother to start at a particular help topic
                    System.Diagnostics.Process.Start(chmPath);
                }
            }
        } // ShowHelpFile

        [ExcelCommand(
            Name = "Dialog_About_TopoLib",
            Description = "Shows a dialog with a copy right statement and a list of referenced NuGet packages",
            HelpTopic = "TopoLib-AddIn.chm!1203")]
        public static void Dialog_About_TopoLib()
        {
            var dialog  = new XlDialogBox()                  {                   W = 333, H = 240, Text = "About TopoLib",  };
            var ctrl_01 = new XlDialogBox.GroupBox()         { X = 013, Y = 013, W = 307, H = 130, Text = "This library uses the following NuGet packages",  };
            var ctrl_02 = new XlDialogBox.ListBox()          { X = 031, Y = 038, W = 270,          Text = "List_02" };
            var ctrl_03 = new XlDialogBox.OkButton()         { X = 031, Y = 160, W = 270,          Text = "Duijndam.Dev   |   Copyright © 2020 - 2022", IO = 1, };
            var ctrl_04 = new XlDialogBox.OkButton()         { X = 031, Y = 200, W = 100,          Text = "&OK", Default = true, };
            var ctrl_05 = new XlDialogBox.HelpButton2()      { X = 201, Y = 200, W = 100,          Text = "&Help",  };

            ctrl_02.Items.AddRange(new string[]
            {
                "'ExcelDna.AddIn' version='1.5.1' developmentDependency='true'",
                "'ExcelDna.Integration' version='1.5.1'",
                "'ExcelDna.IntelliSense' version='1.5.1'",
                "'ExcelDna.Registration' version='1.5.1'",
                "'ExcelDna.XmlSchemas' version='1.0.0'",
                "'ExcelDnaDoc' version='1.5.1'",
                "'Serilog' version='2.10.0'",
                "'Serilog.Sinks.ExcelDnaLogDisplay' version='1.5.0'",
                "'SharpProj' version='8.2001.106'",
                "'SharpProj.Core' version='8.2001.106'"
            });

            dialog.Controls.Add(ctrl_01);
            dialog.Controls.Add(ctrl_02);
            dialog.Controls.Add(ctrl_03);
            dialog.Controls.Add(ctrl_04);
            dialog.Controls.Add(ctrl_05);

            dialog.CallingMethod = System.Reflection.MethodBase.GetCurrentMethod(); 
            bool bOK = dialog.ShowDialog(ValidateAboutDialog);
            if (bOK == false) return;

        } // Dialog_About_TopoLib

        [ExcelCommand(
            Name = "Dialog_Cache_Settings",
            Description = "Sets global transform parameters for the TopoLib AddIn",
            HelpTopic = "TopoLib-AddIn.chm!1204")]
        public static void Dialog_Cache_Settings()
        {
            var dialog  = new XlDialogBox()                  {	                   W = 360, H = 230, Text = "Proj Library Cache Settings", };
            var ctrl_01 = new XlDialogBox.Label()            {	 X = 020, Y = 010,                   Text = "&Cache Path && File Name",  };
            var ctrl_02 = new XlDialogBox.ApplyButton()      {	 X = 294, Y = 026, W = 025,          Text = "ᐅ",  };
            var ctrl_03 = new XlDialogBox.TextEdit()         {	 X = 020, Y = 027, W = 272,          IO   = "c:\\dummyFolder", };
            var ctrl_04 = new XlDialogBox.GroupBox()         {	 X = 020, Y = 055, W = 300, H = 050, Text = "Cache Useage",  };
            var ctrl_05 = new XlDialogBox.CheckBox()         {	 X = 030, Y = 073, W = 090,          Text = "&Enable Cache",  IO = true, };
            var ctrl_06 = new XlDialogBox.ApplyButton()      {	 X = 150, Y = 070, W = 150,          Text = "&Clear Cache",  IO = 3, };
            var ctrl_07 = new XlDialogBox.GroupBox()         {	 X = 020, Y = 110, W = 300, H = 065, Text = "Cache Parameters",  };
            var ctrl_08 = new XlDialogBox.Label()            {	 X = 030, Y = 128,                   Text = "&File Size [MB]",  };
            var ctrl_09 = new XlDialogBox.IntegerEdit()      {	 X = 030, Y = 144, W = 100,          IO_int = -1, };
            var ctrl_10 = new XlDialogBox.Label()            {	 X = 150, Y = 128,                   Text = "&Expiration Time [s]",  IO = 1, };
            var ctrl_11 = new XlDialogBox.DoubleEdit()       {	 X = 150, Y = 144, W = 100,          IO_double = -1, };
            var ctrl_12 = new XlDialogBox.OkButton()         {	 X = 050, Y = 190, W = 075,          Text = "&OK", Default = true, };
            var ctrl_13 = new XlDialogBox.CancelButton()     {	 X = 137, Y = 190, W = 075,          Text = "&Cancel",  };
            var ctrl_14 = new XlDialogBox.HelpButton2()      {	 X = 224, Y = 190, W = 075,          Text = "&Help",  };

            dialog.Controls.Add(ctrl_01);
            dialog.Controls.Add(ctrl_02);
            dialog.Controls.Add(ctrl_03);
            dialog.Controls.Add(ctrl_04);
            dialog.Controls.Add(ctrl_05);
            dialog.Controls.Add(ctrl_06);
            dialog.Controls.Add(ctrl_07);
            dialog.Controls.Add(ctrl_08);
            dialog.Controls.Add(ctrl_09);
            dialog.Controls.Add(ctrl_10);
            dialog.Controls.Add(ctrl_11);
            dialog.Controls.Add(ctrl_12);
            dialog.Controls.Add(ctrl_13);
            dialog.Controls.Add(ctrl_14);

            dialog.CallingMethod = System.Reflection.MethodBase.GetCurrentMethod(); 
            dialog.DialogScaling = 125.0;  // Use this if the dialog was designed using a display with 120 DPI

            ctrl_03.IO_string  = CctOptions.CachePath;
            ctrl_05.IO_checked = CctOptions.EnableCache;
            ctrl_09.IO_int     = CctOptions.CacheSize;
            ctrl_11.IO_double  = CctOptions.CacheExpiry;
            
            bool bOK = dialog.ShowDialog(ValidateCacheDialog);
            if (bOK == false) return;

            CctOptions.CachePath   = ctrl_03.IO_string;
            CctOptions.EnableCache = ctrl_05.IO_checked;
            CctOptions.CacheSize   = ctrl_09.IO_int;
            CctOptions.CacheExpiry = ctrl_11.IO_double;

            CctOptions.ProjContext.SetGridCache(CctOptions.EnableCache, CctOptions.CachePath, CctOptions.CacheSize, (int) CctOptions.CacheExpiry);

        } // Dialog_Cache_Settings


        [ExcelCommand(
            Name = "Dialog_Export_GPX_data",
            Description = "Dialog to export GPX-data to a file on disk",
            HelpTopic = "TopoLib-AddIn.chm!1205")]
        public static void Dialog_Export_GPX_data()
        {
            var dialog  = new XlDialogBox()                  {	                   W = 345, H = 180, Text = "GPX File Export",  IO = 5, };
            var ctrl_01 = new XlDialogBox.Label()            {	 X = 020, Y = 010,                   Text = "Select the required range of (Long, Lat, [ele, [time]]) values.",  };
            var ctrl_02 = new XlDialogBox.Label()            {	 X = 020, Y = 025,                   Text = "Note: The list of data points must contain 2 - 4 columns.",  };
            var ctrl_03 = new XlDialogBox.Label()            {	 X = 020, Y = 040,                   Text = "Note: Please include header rows if these are available.",  };
            var ctrl_04 = new XlDialogBox.GroupBox()         {	 X = 020, Y = 075, W = 190, H = 055, Text = "EPSG:4326 (long lat [ele]) order",  };
            var ctrl_05 = new XlDialogBox.RefEdit()          {	 X = 030, Y = 095, W = 160,          };
            var ctrl_06 = new XlDialogBox.OkButton()         {	 X = 020, Y = 145, W = 080,          Text = "&Save", Default = true, };
            var ctrl_07 = new XlDialogBox.CancelButton()     {	 X = 130, Y = 145, W = 080,          Text = "&Cancel",  };
            var ctrl_08 = new XlDialogBox.HelpButton2()      {	 X = 240, Y = 145, W = 080,          Text = "&Help",  };

            dialog.Controls.Add(ctrl_01);
            dialog.Controls.Add(ctrl_02);
            dialog.Controls.Add(ctrl_03);
            dialog.Controls.Add(ctrl_04);
            dialog.Controls.Add(ctrl_05);
            dialog.Controls.Add(ctrl_06);
            dialog.Controls.Add(ctrl_07);
            dialog.Controls.Add(ctrl_08);

            dialog.CallingMethod = System.Reflection.MethodBase.GetCurrentMethod(); 
            dialog.DialogScaling = 125.0;  // Use this if the dialog was designed using a display with 120 DPI

            ctrl_05.IO_string = gpxRange;

Show:       bool bOK = dialog.ShowDialog(Validate);
            if (bOK == false) return;

            gpxRange = ctrl_05.IO_string;
            R1C1 gpxR1C1 = new R1C1(gpxRange);

            int nInputRows = gpxR1C1.Rows();
            int nInputCols = gpxR1C1.Cols();

            if (nInputRows < 1)
            {
                System.Windows.Forms.MessageBox.Show("The input data should contain one or more rows", "Range Error");
                goto Show;
            }

            if ((nInputCols < 2) || (nInputCols > 4))
            {
                System.Windows.Forms.MessageBox.Show("The input data should consist of 2 to 4 columns", "Range Error");
                goto Show;
            }

            try
            {
                // now it is time to use dialog results to get the data
                Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
                Workbook xlWb = xlApp.ActiveWorkbook;
                if (xlWb == null) return;

                Worksheet xlWs = xlApp.ActiveWorkbook.ActiveSheet;
                if (xlWs == null) return;

                Range cell1 = xlWs.Cells[gpxR1C1.Top(), gpxR1C1.Left()];
                Range cell2 = xlWs.Cells[gpxR1C1.Top() + nInputRows - 1, gpxR1C1.Left() + nInputCols - 1];
                Range range = xlWs.get_Range(cell1, cell2);

                object[,] oPoints = (object[,]) range.Value2;

                string gpxOut = Gps.GetGpsTracks(oPoints, true);

                // Now we are ready to write the string to a file.
                // public static void WriteAllText (string path, string? contents);

                SaveFileDialog saveFileDialog1 = new SaveFileDialog();

                saveFileDialog1.AddExtension = true;
                saveFileDialog1.ValidateNames = true;
                saveFileDialog1.RestoreDirectory = true;

                saveFileDialog1.DefaultExt = "gpx";  
                saveFileDialog1.Filter = "GPX files (*.gpx)|*.gpx|All files (*.*)|*.*";
                saveFileDialog1.Title = "Save a GPX File";
                saveFileDialog1.ShowDialog();

                if(saveFileDialog1.FileName != "")
                {
                    File.WriteAllText(saveFileDialog1.FileName, gpxOut);
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message, "Unexpected Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                AddIn.ProcessException(ex);
            }
        } // Dialog_Export_GPX_data

        [ExcelCommand(
            Name = "Dialog_Export_KML_data",
            Description = "Dialog to export a Coordinate Reference System or a Coordinate Transform to a file on disk in text format",
            HelpTopic = "TopoLib-AddIn.chm!1222")]
        public static void Dialog_Export_KML_data()
        {
            var dialog  = new XlDialogBox()                  {	                   W = 345, H = 180, Text = "KML File Export",  IO = 5, };
            var ctrl_01 = new XlDialogBox.Label()            {	 X = 020, Y = 010,                   Text = "Select the required range of (Long, Lat, [ele]) values.",  };
            var ctrl_02 = new XlDialogBox.Label()            {	 X = 020, Y = 025,                   Text = "Note: The list of data points must contain 2 - 3 columns.",  };
            var ctrl_03 = new XlDialogBox.Label()            {	 X = 020, Y = 040,                   Text = "Note: Please include header rows if these are available.",  };
            var ctrl_04 = new XlDialogBox.GroupBox()         {	 X = 020, Y = 075, W = 190, H = 055, Text = "EPSG:4326 (long lat [ele]) order",  };
            var ctrl_05 = new XlDialogBox.RefEdit()          {	 X = 030, Y = 095, W = 160,          };
            var ctrl_06 = new XlDialogBox.OkButton()         {	 X = 020, Y = 145, W = 080,          Text = "&Save", Default = true, };
            var ctrl_07 = new XlDialogBox.CancelButton()     {	 X = 130, Y = 145, W = 080,          Text = "&Cancel",  };
            var ctrl_08 = new XlDialogBox.HelpButton2()      {	 X = 240, Y = 145, W = 080,          Text = "&Help",  };

            dialog.Controls.Add(ctrl_01);
            dialog.Controls.Add(ctrl_02);
            dialog.Controls.Add(ctrl_03);
            dialog.Controls.Add(ctrl_04);
            dialog.Controls.Add(ctrl_05);
            dialog.Controls.Add(ctrl_06);
            dialog.Controls.Add(ctrl_07);
            dialog.Controls.Add(ctrl_08);

            dialog.CallingMethod = System.Reflection.MethodBase.GetCurrentMethod(); 
            dialog.DialogScaling = 125.0;  // Use this if the dialog was designed using a display with 120 DPI

            ctrl_05.IO_string = kmlRange;

Show:       bool bOK = dialog.ShowDialog(Validate);
            if (bOK == false) return;

            kmlRange = ctrl_05.IO_string;
            R1C1 kmlR1C1 = new R1C1(kmlRange);

            int nInputRows = kmlR1C1.Rows();
            int nInputCols = kmlR1C1.Cols();

            if (nInputRows < 1)
            {
                System.Windows.Forms.MessageBox.Show("The input data should contain one or more rows", "Range Error");
                goto Show;
            }

            if ((nInputCols < 2) || (nInputCols > 3))
            {
                System.Windows.Forms.MessageBox.Show("The input data should consist of 2 to 3 columns", "Range Error");
                goto Show;
            }

            try
            {
                // now it is time to use dialog results to get the data
                Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
                Workbook xlWb = xlApp.ActiveWorkbook;
                if (xlWb == null) return;

                Worksheet xlWs = xlApp.ActiveWorkbook.ActiveSheet;
                if (xlWs == null) return;

                Range cell1 = xlWs.Cells[kmlR1C1.Top(), kmlR1C1.Left()];
                Range cell2 = xlWs.Cells[kmlR1C1.Top() + nInputRows - 1, kmlR1C1.Left() + nInputCols - 1];
                Range range = xlWs.get_Range(cell1, cell2);

                object[,] oPoints = (object[,]) range.Value2;

                string kmlOut = Gps.GetKmlTracks(oPoints, true);

                // Now we are ready to write the string to a file.
                // public static void WriteAllText (string path, string? contents);

                SaveFileDialog saveFileDialog1 = new SaveFileDialog();

                saveFileDialog1.AddExtension = true;
                saveFileDialog1.ValidateNames = true;
                saveFileDialog1.RestoreDirectory = true;

                saveFileDialog1.DefaultExt = "kml";  
                saveFileDialog1.Filter = "KML files (*.kml)|*.kml|All files (*.*)|*.*";
                saveFileDialog1.Title = "Save a KML File";
                saveFileDialog1.ShowDialog();

                if(saveFileDialog1.FileName != "")
                {
                    File.WriteAllText(saveFileDialog1.FileName, kmlOut);
                }

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message, "Unexpected Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                AddIn.ProcessException(ex);
            }

        } // Dialog_Export_Wizard

        [ExcelCommand(
            Name = "Dialog_Export_Wizard",
            Description = "Dialog to export a CRS- or CCT definition to a file on disk",
            HelpTopic = "TopoLib-AddIn.chm!1206")]
        public static void Dialog_Export_Wizard()
        {
            var dialog1  = new XlDialogBox()                  {	                   W = 360, H = 230, Text = "CRS & CCT String export [1]",  };
            var ctrl_01 = new XlDialogBox.Label()            {	 X = 020, Y = 016,                   Text = "Select Data Type and Data Source",  };
            var ctrl_02 = new XlDialogBox.GroupBox()         {	 X = 020, Y = 040, W = 300, H = 065, Text = "Export Data Type",  };
            var ctrl_03 = new XlDialogBox.RadioButtonGroup() {	                                     IO = 1, };
            var ctrl_04 = new XlDialogBox.RadioButton()      {	          Y = 057, W = 280,          Text = "CRS - Coordinate &Reference System",  };
            var ctrl_05 = new XlDialogBox.RadioButton()      {	          Y = 077, W = 280,          Text = "CCT - Coordinate Conversion and &Transform", };
            var ctrl_06 = new XlDialogBox.GroupBox()         {	 X = 020, Y = 110, W = 300, H = 065, Text = "Data input from",  };
            var ctrl_07 = new XlDialogBox.RadioButtonGroup() {	                                     IO = 1, };
            var ctrl_08 = new XlDialogBox.RadioButton()      {	          Y = 127,                   Text = "&PROJ Database Query",  };
            var ctrl_09 = new XlDialogBox.RadioButton()      {	          Y = 147,                   Text = "Spreadsheet &Cell reference(s)",  };
            var ctrl_10 = new XlDialogBox.OkButton()         {	 X = 030, Y = 190, W = 075,          Text = "< &Back",  Enable = false, };
            var ctrl_11 = new XlDialogBox.OkButton()         {	 X = 125, Y = 190, W = 075,          Text = "&Next >",  };
            var ctrl_12 = new XlDialogBox.CancelButton()     {	 X = 245, Y = 190, W = 075,          Text = "&Cancel",  };

            dialog1.Controls.Add(ctrl_01);
            dialog1.Controls.Add(ctrl_02);
            dialog1.Controls.Add(ctrl_03);
            dialog1.Controls.Add(ctrl_04);
            dialog1.Controls.Add(ctrl_05);
            dialog1.Controls.Add(ctrl_06);
            dialog1.Controls.Add(ctrl_07);
            dialog1.Controls.Add(ctrl_08);
            dialog1.Controls.Add(ctrl_09);
            dialog1.Controls.Add(ctrl_10);
            dialog1.Controls.Add(ctrl_11);
            dialog1.Controls.Add(ctrl_12);

            dialog1.CallingMethod = System.Reflection.MethodBase.GetCurrentMethod(); 
            dialog1.DialogScaling = 125.0;  // Use this if the dialog was designed using a display with 120 DPI

Page1:      ctrl_03.IO_index = ExportDataType;
            ctrl_07.IO_index = DataInputFrom;

            bool bOK = dialog1.ShowDialog(Validate);
            if (bOK == false) return;

            ExportDataType = ctrl_03.IO_index;
            DataInputFrom  = ctrl_07.IO_index;


            var dialog2 = new XlDialogBox()                  {	                   W = 360, H = 230, Text = "CRS & CCT String export [2]",  };

            var ctrl_13 = new XlDialogBox.Label()            {	 X = 020, Y = 016,                   Text = "Select Source and Target CRS",  };
            var ctrl_14 = new XlDialogBox.GroupBox()         {	 X = 020, Y = 040, W = 300, H = 055, Text = "Source CRS",  };

            var ctrl_15 = new XlDialogBox.RefEdit()          {	 X = 030, Y = 060, W = 170,          };

            var ctrl_16 = new XlDialogBox.DropdownList()     {	 X = 030, Y = 060, W = 080,          Text = "List_16", IO = 1, };
            var ctrl_17 = new XlDialogBox.Label()            {	 X = 115, Y = 060, W = 020,          Text = " : ",  };
            var ctrl_18 = new XlDialogBox.DoubleEdit()      {	 X = 130, Y = 060, W = 070,          IO = 4326};

            var ctrl_19 = new XlDialogBox.GroupBox()         {	 X = 020, Y = 110, W = 300, H = 055, Text = "Target CRS",  };

            var ctrl_20 = new XlDialogBox.RefEdit()          {	 X = 030, Y = 130, W = 170,          };

            var ctrl_21 = new XlDialogBox.DropdownList()     {	 X = 030, Y = 130, W = 080,          Text = "List_21", IO = 1,};
            var ctrl_22 = new XlDialogBox.Label()            {	 X = 115, Y = 130, W = 020,          Text = " :",  };
            var ctrl_23 = new XlDialogBox.DoubleEdit()      {	 X = 130, Y = 130, W = 070,          IO = 4326, };

            var ctrl_24 = new XlDialogBox.OkButton()         {	 X = 030, Y = 190, W = 075,          Text = "< &Back",  };
            var ctrl_25 = new XlDialogBox.OkButton()         {	 X = 125, Y = 190, W = 075,          Text = "Save &As",  };
            var ctrl_26 = new XlDialogBox.CancelButton()     {	 X = 245, Y = 190, W = 075,          Text = "&Cancel",  };

            string[] Authorities = new string[]
            {
                "EPSG",
                "ESRI",
                "IAU_2015",
                "IGNF",
                "NKG",
                "OGC",
                "PROJ"
            };

            ctrl_16.Items.AddRange(Authorities);
            ctrl_21.Items.AddRange(Authorities);

            dialog2.Controls.Add(ctrl_13);
            dialog2.Controls.Add(ctrl_14);
            dialog2.Controls.Add(ctrl_15);
            dialog2.Controls.Add(ctrl_16);
            dialog2.Controls.Add(ctrl_17);
            dialog2.Controls.Add(ctrl_18);
            dialog2.Controls.Add(ctrl_19);
            dialog2.Controls.Add(ctrl_20);
            dialog2.Controls.Add(ctrl_21);
            dialog2.Controls.Add(ctrl_22);
            dialog2.Controls.Add(ctrl_23);
            dialog2.Controls.Add(ctrl_24);
            dialog2.Controls.Add(ctrl_25);
            dialog2.Controls.Add(ctrl_26);

            dialog2.CallingMethod = System.Reflection.MethodBase.GetCurrentMethod(); 
            dialog2.DialogScaling = 125.0;  // Use this if the dialog was designed using a display with 120 DPI

Page2:      if (ExportDataType == 0)
            {
                // CRS; adjust text & disable second part of dialog
                ctrl_13.Text = "Select Coordinate Reference System";

                ctrl_19.Visible = false;
                ctrl_20.Visible = false;
                ctrl_21.Visible = false;
                ctrl_22.Visible = false;
                ctrl_23.Visible = false;
            }

            if (DataInputFrom == 0)
            {
                // Database query; no references used
                ctrl_15.Visible   = false;
                ctrl_20.Visible   = false;

                ctrl_16.IO_index  = SourceAuthIndex;
                ctrl_18.IO_double = SourceCode;

                ctrl_21.IO_index  = TargetAuthIndex;
                ctrl_23.IO_double = TargetCode;
            }
            else
            {
                // Reference query; database not used
                ctrl_15.IO_string = SourceCrsRange;
                ctrl_20.IO_string = TargetCrsRange;

                ctrl_16.Visible = false;
                ctrl_17.Visible = false;
                ctrl_18.Visible = false;

                ctrl_21.Visible = false;
                ctrl_22.Visible = false;
                ctrl_23.Visible = false;
            }
            
            bOK = dialog2.ShowDialog(Validate);
            if (bOK == false) return;

            if (dialog2.IO_index == 12) goto Page1; // Back button was pressed

            if (DataInputFrom == 0)  // DB
            {
                // Database query; no references used

                SourceAuthIndex = ctrl_16.IO_index;
                SourceAuthority = Authorities[SourceAuthIndex];

                SourceCode      = (int) ctrl_18.IO_double;

                FileName = SourceAuthority + "(" + SourceCode.ToString() + ")";

                if (ExportDataType == 1) // CCT
                {
                    // CCT; use second part of dialog
                    TargetAuthIndex = ctrl_21.IO_index;
                    TargetAuthority = Authorities[TargetAuthIndex ];

                    TargetCode      = (int) ctrl_23.IO_double;

                    FileName += "_to_" + TargetAuthority + "(" + TargetCode.ToString() + ")";
                }
            }
            else
            {
                // Reference query; database not used
                SourceCrsRange = ctrl_15.IO_string;
                TargetCrsRange = ctrl_20.IO_string;

                R1C1 SourceCrsR1C1 = new R1C1(SourceCrsRange);
                int nSourceCrsRows = SourceCrsR1C1.Rows();
                int nSourceCrsCols = SourceCrsR1C1.Cols();

                if (nSourceCrsRows != 1)
                {
                    System.Windows.Forms.MessageBox.Show("The source-CRS reference should contain one row", "Range Error");
                    goto Page2;
                }

                if ((nSourceCrsCols < 1) || (nSourceCrsCols > 2))
                {
                    System.Windows.Forms.MessageBox.Show("The source-CRS reference should consist of 1 or 2 columns", "Range Error");
                    goto Page2;
                }

                if (ExportDataType == 1)
                {
                    // CCT; use second part of dialog
                    R1C1 TargetCrsR1C1 = new R1C1(TargetCrsRange);
                    int nTargetCrsRows = TargetCrsR1C1.Rows();
                    int nTargetCrsCols = TargetCrsR1C1.Cols();

                    if (nTargetCrsRows != 1)
                    {
                        System.Windows.Forms.MessageBox.Show("The target-CRS reference should contain one row", "Range Error");
                        goto Page2;
                    }

                    if ((nTargetCrsCols < 1) || (nTargetCrsCols > 2))
                    {
                        System.Windows.Forms.MessageBox.Show("The source-CRS reference should consist of 1 or 2 columns", "Range Error");
                        goto Page2;
                    }
                }
            }

            object[,] SourceCells = null;
            object[,] TargetCells = null;
    
            try
            {
                // check the input

                if (DataInputFrom == 0)
                {
                    // Database query; no references used
                }
                else
                {
                    // Reference query; database not used; need to get hold of Excel
                    Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
                    Workbook xlWb = xlApp.ActiveWorkbook;
                    if (xlWb == null) return;

                    Worksheet xlWs = xlApp.ActiveWorkbook.ActiveSheet;
                    if (xlWs == null) return;

                    // Bit of a repeat, to keep things simple
                    R1C1 SourceCrsR1C1 = new R1C1(SourceCrsRange);
                    int nSourceCrsRows = SourceCrsR1C1.Rows();
                    int nSourceCrsCols = SourceCrsR1C1.Cols();

                    // now get the contents of the first reference from Excel
                    Range cell1 = xlWs.Cells[SourceCrsR1C1.Top(), SourceCrsR1C1.Left()];
                    Range cell2 = xlWs.Cells[SourceCrsR1C1.Top() + nSourceCrsRows - 1, SourceCrsR1C1.Left() + nSourceCrsCols - 1];

                    var v1 = cell1.Value2;  // string ?
                    var v2 = cell2.Value2;  // double ?

                    SourceCells = new object[nSourceCrsRows, nSourceCrsCols];
                    SourceCells[0, 0] = cell1.Value2;
                    if (nSourceCrsCols > 1)
                        SourceCells[0, 1] = cell2.Value2;

                    SourceAuthority = nSourceCrsCols == 2 ? Convert.ChangeType(v1, TypeCode.String) : "EPSG";
                    FileName = SourceAuthority + "(" + Convert.ChangeType(v2, TypeCode.String) + ")";

                    if (ExportDataType == 1) // CCT
                    {
                        // CCT; use second part of dialog
                        R1C1 TargetCrsR1C1 = new R1C1(TargetCrsRange);
                        int nTargetCrsRows = TargetCrsR1C1.Rows();
                        int nTargetCrsCols = TargetCrsR1C1.Cols();

                        // now get the contents of the second reference from Excel
                        Range cell3 = xlWs.Cells[TargetCrsR1C1.Top(), TargetCrsR1C1.Left()];
                        Range cell4 = xlWs.Cells[TargetCrsR1C1.Top() + nTargetCrsRows - 1, TargetCrsR1C1.Left() + nTargetCrsCols - 1];

                        var v3 = cell3.Value2;  // string ?
                        var v4 = cell4.Value2;  // double ?

                        TargetCells = new object[nTargetCrsRows, nTargetCrsCols];
                        TargetCells[0, 0] = cell3.Value2;
                        if (nTargetCrsCols > 1)
                            TargetCells[0, 1] = cell4.Value2;

                        TargetAuthority = nTargetCrsCols == 2 ? Convert.ChangeType(v3, TypeCode.String) : "EPSG";
                        FileName += "_to_" + TargetAuthority + "(" + Convert.ChangeType(v4, TypeCode.String) + ")";
                    }
                }

                SaveFileDialog saveFileDialog1 = new SaveFileDialog();

                saveFileDialog1.AddExtension = true;
                saveFileDialog1.ValidateNames = true;
                saveFileDialog1.RestoreDirectory = true;
                saveFileDialog1.FileName = FileName;
                saveFileDialog1.InitialDirectory = FolderName;
                saveFileDialog1.FilterIndex = FilterIndex;
                saveFileDialog1.DefaultExt = "wkt.txt";
                saveFileDialog1.Filter =
                    "PROJ Files (*.proj)|*.proj|" +
                    "WKT files  (*.wkt)|*.wkt|" +
                    "JSON files (*.json)|*.json";
                saveFileDialog1.Title = "Select File Name and (\"Save As\") File Type";
                saveFileDialog1.ShowDialog();

                if(saveFileDialog1.FileName != "")
                {
                    FolderName  = saveFileDialog1.InitialDirectory;
                    FilterIndex = saveFileDialog1.FilterIndex;

                    // now it is time to export the data

                    string FileContents = "";

                    if(ExportDataType == 0)
                    {
                        // we have to export a CRS
                        if (DataInputFrom == 0)
                        {
                            // Database query; no references used
                            // Grab it from SourceAuthority & SourceCode
                            using (ProjContext pjContext = Crs.CreateContext())
                            {
                                using (var crs = CoordinateReferenceSystem.CreateFromDatabase(SourceAuthority, SourceCode, pjContext).WithAxisNormalized())
                                {
                                    switch(saveFileDialog1.FilterIndex)
                                    {
                                    case 1 :
                                        FileContents = crs.AsProjString(new ProjStringOptions { MultiLine = true, WriteApproxFlag = true });
                                        break;

                                    case 2 :
                                        FileContents = crs.AsWellKnownText();
                                        break;

                                    case 3 :
                                        FileContents = crs.AsProjJson();
                                        break;
                                    }
                                }
                            }
                        }
                        else
                        {
                            // Grab it from first reference
                            using (ProjContext pjContext = Crs.CreateContext())
                            {
                                using (var crs = Crs.CreateCrs(SourceCells, pjContext).WithAxisNormalized())

                                {
                                    switch(saveFileDialog1.FilterIndex)
                                    {
                                    case 1 :
                                        FileContents = crs.AsProjString(new ProjStringOptions { MultiLine = true, WriteApproxFlag = true });
                                        break;

                                    case 2 :
                                        FileContents = crs.AsWellKnownText();
                                        break;

                                    case 3 :
                                        FileContents = crs.AsProjJson();
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        // we have to export a CCT
                        if (DataInputFrom == 0)
                        {
                            // Database query; no references used
                            // Grab it from SourceAuthority & SourceCode
                            using (ProjContext pjContext = Crs.CreateContext())
                            {
                                using (var crsSource = CoordinateReferenceSystem.CreateFromDatabase(SourceAuthority, SourceCode, pjContext))
                                using (var crsTarget = CoordinateReferenceSystem.CreateFromDatabase(TargetAuthority, TargetCode, pjContext))
                                {
                                    // Deal with optional parameters
                                    bool bAllowDeprecatedCRS = true;

                                    // setting mode-flag to 128 preserves allowing to use a deprecated CRS; but it can still be overruled in global settings.
                                    var options = Cct.GetCoordinateTransformOptions(128, -1, -1000, -1000, -1000, -1000, ref bAllowDeprecatedCRS);

                                    using(var transform = Cct.CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS))
                                    {
                                        // check if the transform is a choosecoordinatetransform!
                                        ChooseCoordinateTransform transforms = transform as ChooseCoordinateTransform;
                                        if (transforms is null)
                                        {
                                            switch(saveFileDialog1.FilterIndex)
                                            {
                                            case 1 :
                                                FileContents = transform.AsProjString();
                                                break;

                                            case 2 :
                                                FileContents = transform.AsWellKnownText();
                                                break;

                                            case 3 :
                                                FileContents = transform.AsProjJson();
                                                break;
                                            }
                                        }
                                        else
                                        {
                                            switch(saveFileDialog1.FilterIndex)
                                            {
                                            case 1 :
                                                FileContents = transforms[0].AsProjString();
                                                break;

                                            case 2 :
                                                FileContents = transforms[0].AsWellKnownText();
                                                break;

                                            case 3 :
                                                FileContents = transforms[0].AsProjJson();
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            // Grab it from first reference
                            using (ProjContext pjContext = Crs.CreateContext())
                            {
                                using (var crsSource = Crs.CreateCrs(SourceCells, pjContext, false))
                                using (var crsTarget = Crs.CreateCrs(TargetCells, pjContext, false))
                                {
                                    // Deal with optional parameters
                                    bool bAllowDeprecatedCRS = true;

                                    // setting mode-flag to 128 preserves allowing to use a deprecated CRS; but it can still be overruled in global settings.
                                    var options = Cct.GetCoordinateTransformOptions(128, -1, -1000, -1000, -1000, -1000, ref bAllowDeprecatedCRS);

                                    using(var transform = Cct.CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS))
                                    {
                                        // check if the transform is a choosecoordinatetransform!
                                        ChooseCoordinateTransform transforms = transform as ChooseCoordinateTransform;
                                        if (transforms is null)
                                        {
                                            switch(saveFileDialog1.FilterIndex)
                                            {
                                            case 1 :
                                                FileContents = transform.AsProjString();
                                                break;

                                            case 2 :
                                                FileContents = transform.AsWellKnownText();
                                                break;

                                            case 3 :
                                                FileContents = transform.AsProjJson();
                                                break;
                                            }
                                        }
                                        else
                                        {
                                            switch(saveFileDialog1.FilterIndex)
                                            {
                                            case 1 :
                                                FileContents = transforms[0].AsProjString();
                                                break;

                                            case 2 :
                                                FileContents = transforms[0].AsWellKnownText();
                                                break;

                                            case 3 :
                                                FileContents = transforms[0].AsProjJson();
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    
                    if (String.IsNullOrWhiteSpace(FileContents))
                    {
                        throw new Exception("file contents empty");
                    }

                    File.WriteAllText(saveFileDialog1.FileName, FileContents);
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message, "Unexpected Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                AddIn.ProcessException(ex);
            }

        } // Dialog_Export_Wizard

        [ExcelCommand(
            Name = "Dialog_Log_Settings",
            Description = "Sets the logging level for the TopoLib AddIn",
            HelpTopic = "TopoLib-AddIn.chm!1207")]
        public static void Dialog_Log_Settings()
        {
            var dialog  = new XlDialogBox()                  {	                   W = 320, H = 150, Text = "Logging settings",  IO = 2, };
            var ctrl_01 = new XlDialogBox.Label()            {	 X = 020, Y = 010,                   Text = "Select the required logging level in the option list below",  };
            var ctrl_02 = new XlDialogBox.GroupBox()         {	 X = 020, Y = 030, W = 165, H = 100, Text = "Error logging ",  };
            var ctrl_03 = new XlDialogBox.RadioButtonGroup() {	                                     IO = 3, };
            var ctrl_04 = new XlDialogBox.RadioButton()      {	          Y = 045, W = 140,          Text = "&None      ➔  (few messages)",  };
            var ctrl_05 = new XlDialogBox.RadioButton()      {	          Y = 065, W = 140,          Text = "&Errors",  };
            var ctrl_06 = new XlDialogBox.RadioButton()      {	          Y = 085, W = 140,          Text = "&Debug",  };
            var ctrl_07 = new XlDialogBox.RadioButton()      {	          Y = 105, W = 140,          Text = "&Verbose  ➔  (all messages)",  IO = true, };
            var ctrl_08 = new XlDialogBox.OkButton()         {	 X = 210, Y = 060, W = 080,          Text = "&OK", Default = true, };
            var ctrl_09 = new XlDialogBox.CancelButton()     {	 X = 210, Y = 085, W = 080,          Text = "&Cancel",  };
            var ctrl_10 = new XlDialogBox.HelpButton2()      {	 X = 210, Y = 110, W = 080,          Text = "&Help",  };

            dialog.Controls.Add(ctrl_01);
            dialog.Controls.Add(ctrl_02);
            dialog.Controls.Add(ctrl_03);
            dialog.Controls.Add(ctrl_04);
            dialog.Controls.Add(ctrl_05);
            dialog.Controls.Add(ctrl_06);
            dialog.Controls.Add(ctrl_07);
            dialog.Controls.Add(ctrl_08);
            dialog.Controls.Add(ctrl_09);
            dialog.Controls.Add(ctrl_10);
            dialog.CallingMethod = System.Reflection.MethodBase.GetCurrentMethod(); 
            dialog.DialogScaling = 125.0;  // Use this if the dialog was designed using a display with 120 DPI

            ctrl_03.IO_index = CctOptions.LogLevel; // get the LogLevel for use in the dialog

            bool bOK = dialog.ShowDialog(Validate);
            if (bOK == false) return;

            CctOptions.LogLevel = ctrl_03.IO_index; // set the LogLevel; this also takes care of saving the LogLevel in the config file

        } // LogSettingsDialog

        [ExcelCommand(
            Name = "Dialog_Resource_Settings",
            Description = "Sets the acces method for PROJ Resources",
            HelpTopic = "TopoLib-AddIn.chm!1208")]
        public static void Dialog_Resource_Settings()
        {
            var dialog  = new XlDialogBox()                  {	                   W = 360, H = 230, Text = "Resource Settings",  };
            var ctrl_01 = new XlDialogBox.Label()            {	 X = 020, Y = 010,                   Text = "&PROJ_LIB Environment Variable",  };
            var ctrl_02 = new XlDialogBox.TextEdit()         {	 X = 020, Y = 027, W = 272,          IO   = "c:\\program files\\qgis 3.20.0\\share\\proj", };
            var ctrl_03 = new XlDialogBox.OkButton()         {	 X = 294, Y = 026, W = 025,          Text = "ᐅ", IO= 3 };
            var ctrl_04 = new XlDialogBox.GroupBox()         {	 X = 020, Y = 055, W = 300, H = 050, Text = "&WWW Resources from",  };
            var ctrl_05 = new XlDialogBox.TextEdit()         {	          Y = 073, W = 170,          Text = "&Enable Cache",  IO = "https://cdn.proj.org", };
            var ctrl_06 = new XlDialogBox.OkButton()         {	 X = 224, Y = 072, W = 075,          Text = "&Reset",  IO = 3, };
            var ctrl_07 = new XlDialogBox.GroupBox()         {	 X = 020, Y = 110, W = 300, H = 065, Text = "Resource &Useage",  };
            var ctrl_08 = new XlDialogBox.RadioButtonGroup() {	                                     IO = 1, };
            var ctrl_09 = new XlDialogBox.RadioButton()      {	          Y = 127,                   Text = "Use &Local Resources",  };
            var ctrl_10 = new XlDialogBox.RadioButton()      {	          Y = 147,                   Text = "Use &Network Resources",  };
            var ctrl_11 = new XlDialogBox.Label()            {	 X = 030, Y = 016, W = 100,          Visible = false, };
            var ctrl_12 = new XlDialogBox.OkButton()         {	 X = 050, Y = 190, W = 075,          Text = "&OK", Default = true, };
            var ctrl_13 = new XlDialogBox.CancelButton()     {	 X = 137, Y = 190, W = 075,          Text = "&Cancel",  };
            var ctrl_14 = new XlDialogBox.HelpButton2()      {	 X = 224, Y = 190, W = 075,          Text = "&Help",  };

            dialog.Controls.Add(ctrl_01);
            dialog.Controls.Add(ctrl_02);
            dialog.Controls.Add(ctrl_03);
            dialog.Controls.Add(ctrl_04);
            dialog.Controls.Add(ctrl_05);
            dialog.Controls.Add(ctrl_06);
            dialog.Controls.Add(ctrl_07);
            dialog.Controls.Add(ctrl_08);
            dialog.Controls.Add(ctrl_09);
            dialog.Controls.Add(ctrl_10);
            dialog.Controls.Add(ctrl_11);
            dialog.Controls.Add(ctrl_12);
            dialog.Controls.Add(ctrl_13);
            dialog.Controls.Add(ctrl_14);

            dialog.CallingMethod = System.Reflection.MethodBase.GetCurrentMethod(); 
            dialog.DialogScaling = 125.0;  // Use this if the dialog was designed using a display with 120 DPI

            string key = "PROJ_LIB";
            string sProjLib = Environment.GetEnvironmentVariable(key, EnvironmentVariableTarget.User);
            string sProjOld = sProjLib; // keep a copy for dealing with update button 

            ctrl_02.IO_string = sProjLib;
            ctrl_05.IO_string = CctOptions.EndpointUrl;
            ctrl_08.IO_index  = CctOptions.UseNetwork ? 1 : 0;

            bool bOK = dialog.ShowDialog(ValidateResourceDialog);
            if (bOK == false) return;

            sProjLib               = ctrl_02.IO_string;
            CctOptions.EndpointUrl = ctrl_05.IO_string;
            CctOptions.UseNetwork  = ctrl_08.IO_index > 0;

            if (sProjLib != sProjOld)
            {
                // we need to update the environment variable; is it really a directory ?
                if (Directory.Exists(sProjLib))
                {
                    // we can update the environment variable PROJ_LIB
                    Environment.SetEnvironmentVariable(key,  sProjLib, EnvironmentVariableTarget.User);
                    MessageBox.Show(
                        $"The PROJ_LIB environment variable has been updated to [{sProjLib}]. You need to close and re-open all instances of Excel for this change to have effect", 
                        "Please note", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(
                        $"The PROJ_LIB environment variable cannot be updated to [{sProjLib}]. The folder does not exist", 
                        "Please note", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }

            // Get the correct application instance
            Microsoft.Office.Interop.Excel.Application xlapp = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
            xlapp.Calculate();

        } // Dialog_Resource_Settings

        [ExcelCommand(
            Name = "Dialog_Transform_Settings",
            Description = "Sets global transform parameters for the TopoLib AddIn",
            HelpTopic = "TopoLib-AddIn.chm!1209")]
        public static void Dialog_Transform_Settings()
        {
            var dialog  = new XlDialogBox()                  {	                   W = 500, H = 280, Text = "Global Transform Settings",  };
            var ctrl_01 = new XlDialogBox.GroupBox()         {	 X = 020, Y = 020, W = 195, H = 070, Text = "Use Transform Parameters from",  };
            var ctrl_02 = new XlDialogBox.RadioButtonGroup() {	                                     IO = 1, };
            var ctrl_03 = new XlDialogBox.RadioButton()      {	          Y = 040,                   Text = "&Mode Flags",  };
            var ctrl_04 = new XlDialogBox.RadioButton()      {	          Y = 060,                   Text = "&Global Settings",  };
            var ctrl_05 = new XlDialogBox.OkButton()         {	 X = 125, Y = 040, W = 075,          Text = "&Reset",  IO = 3, };
            var ctrl_06 = new XlDialogBox.GroupBox()         {	 X = 020, Y = 095, W = 195, H = 065, Text = "&Miscellaneous",  IO = 2, };
            var ctrl_07 = new XlDialogBox.Label()            {	 X = 030, Y = 110, W = 075,          Text = "Authority",  };
            var ctrl_08 = new XlDialogBox.TextEdit()         {	 X = 030, Y = 126, W = 075,          IO   = "EPSG", };
            var ctrl_09 = new XlDialogBox.Label()            {	 X = 125, Y = 110, W = 075,          Text = "Accurac&y [m]",  };
            var ctrl_10 = new XlDialogBox.DoubleEdit()       {	 X = 125, Y = 126, W = 075,          IO = -1000, };
            var ctrl_11 = new XlDialogBox.GroupBox()         {	 X = 020, Y = 165, W = 195, H = 100, Text = "Area of Interest" };
            var ctrl_12 = new XlDialogBox.Label()            {	 X = 030, Y = 180,                   Text = "Min Lon (W) [°]",  };
            var ctrl_13 = new XlDialogBox.DoubleEdit()       {	 X = 030, Y = 196, W = 075,          IO = -1000, };
            var ctrl_14 = new XlDialogBox.Label()            {	 X = 125, Y = 180,                   Text = "Max Lon (E) [°]",  };
            var ctrl_15 = new XlDialogBox.DoubleEdit()       {	 X = 125, Y = 196, W = 075,          IO = -1000, };
            var ctrl_16 = new XlDialogBox.Label()            {	 X = 030, Y = 220,                   Text = "Min Lat (S) [°]",  };
            var ctrl_17 = new XlDialogBox.DoubleEdit()       {	 X = 030, Y = 236, W = 075,          IO = -1000, };
            var ctrl_18 = new XlDialogBox.Label()            {	 X = 125, Y = 220,                   Text = "Max Lat (N) [°]",  };
            var ctrl_19 = new XlDialogBox.DoubleEdit()       {	 X = 125, Y = 236, W = 075,          IO = -1000, };
            var ctrl_20 = new XlDialogBox.GroupBox()         {	 X = 230, Y = 020, W = 250, H = 205, Text = "&Optional Parameters",  };
            var ctrl_21 = new XlDialogBox.CheckBox()         {	          Y = 040, W = 230,          Text = "Disallow &Ballpark Conversions",  IO = false, };
            var ctrl_22 = new XlDialogBox.CheckBox()         {	          Y = 060, W = 230,          Text = "Allow if &Grid is Missing",  IO = false, };
            var ctrl_23 = new XlDialogBox.CheckBox()         {	          Y = 080, W = 230,          Text = "Use &Primary Grid Names",  IO = false, };
            var ctrl_24 = new XlDialogBox.CheckBox()         {	          Y = 100, W = 230,          Text = "Use &Superseded Transforms",  IO = false, };
            var ctrl_25 = new XlDialogBox.CheckBox()         {	          Y = 120, W = 230,          Text = "Allow &Deprecated CRSs",  IO = false, };
            var ctrl_26 = new XlDialogBox.CheckBox()         {	          Y = 140, W = 230,          Text = "Strictly &Contains Area",  IO = false, };
            var ctrl_27 = new XlDialogBox.RadioButtonGroup() {	          Y = 140, W = 230,          IO = 1, };
            var ctrl_28 = new XlDialogBox.RadioButton()      {	          Y = 160, W = 230,          Text = "&Automatic Intermediate CRS",  };
            var ctrl_29 = new XlDialogBox.RadioButton()      {	          Y = 180, W = 230,          Text = "&Always Allow Intermediate CRS",  };
            var ctrl_30 = new XlDialogBox.RadioButton()      {	          Y = 200, W = 230,          Text = "&Never Allow Intermediate CRS",  };
            var ctrl_31 = new XlDialogBox.Label()            {	 X = 425, Y = 020, W = 033,          Text = " (flag)",  };
            var ctrl_32 = new XlDialogBox.Label()            {	 X = 430, Y = 042, W = 040,          Text = "(8)",  };
            var ctrl_33 = new XlDialogBox.Label()            {	 X = 430, Y = 062, W = 040,          Text = "(16)",  };
            var ctrl_34 = new XlDialogBox.Label()            {	 X = 430, Y = 082, W = 040,          Text = "(32)",  };
            var ctrl_35 = new XlDialogBox.Label()            {	 X = 430, Y = 102, W = 040,          Text = "(64)",  };
            var ctrl_36 = new XlDialogBox.Label()            {	 X = 430, Y = 122, W = 040,          Text = "(128)",  };
            var ctrl_37 = new XlDialogBox.Label()            {	 X = 430, Y = 142, W = 040,          Text = "(256)",  };
            var ctrl_38 = new XlDialogBox.Label()            {	 X = 430, Y = 162, W = 040,          Text = "- - -",  };
            var ctrl_39 = new XlDialogBox.Label()            {	 X = 430, Y = 182, W = 040,          Text = "(512)",  };
            var ctrl_40 = new XlDialogBox.Label()            {	 X = 430, Y = 202, W = 040,          Text = "(1024)",  };
            var ctrl_41 = new XlDialogBox.OkButton()         {	 X = 230, Y = 245, W = 075,          Text = "&OK", Default = true, };
            var ctrl_42 = new XlDialogBox.CancelButton()     {	 X = 317, Y = 245, W = 075,          Text = "&Cancel",  };
            var ctrl_43 = new XlDialogBox.HelpButton2()      {	 X = 404, Y = 245, W = 075,          Text = "&Help",  };

            dialog.Controls.Add(ctrl_01);
            dialog.Controls.Add(ctrl_02);
            dialog.Controls.Add(ctrl_03);
            dialog.Controls.Add(ctrl_04);
            dialog.Controls.Add(ctrl_05);
            dialog.Controls.Add(ctrl_06);
            dialog.Controls.Add(ctrl_07);
            dialog.Controls.Add(ctrl_08);
            dialog.Controls.Add(ctrl_09);
            dialog.Controls.Add(ctrl_10);
            dialog.Controls.Add(ctrl_11);
            dialog.Controls.Add(ctrl_12);
            dialog.Controls.Add(ctrl_13);
            dialog.Controls.Add(ctrl_14);
            dialog.Controls.Add(ctrl_15);
            dialog.Controls.Add(ctrl_16);
            dialog.Controls.Add(ctrl_17);
            dialog.Controls.Add(ctrl_18);
            dialog.Controls.Add(ctrl_19);
            dialog.Controls.Add(ctrl_20);
            dialog.Controls.Add(ctrl_21);
            dialog.Controls.Add(ctrl_22);
            dialog.Controls.Add(ctrl_23);
            dialog.Controls.Add(ctrl_24);
            dialog.Controls.Add(ctrl_25);
            dialog.Controls.Add(ctrl_26);
            dialog.Controls.Add(ctrl_27);
            dialog.Controls.Add(ctrl_28);
            dialog.Controls.Add(ctrl_29);
            dialog.Controls.Add(ctrl_30);
            dialog.Controls.Add(ctrl_31);
            dialog.Controls.Add(ctrl_32);
            dialog.Controls.Add(ctrl_33);
            dialog.Controls.Add(ctrl_34);
            dialog.Controls.Add(ctrl_35);
            dialog.Controls.Add(ctrl_36);
            dialog.Controls.Add(ctrl_37);
            dialog.Controls.Add(ctrl_38);
            dialog.Controls.Add(ctrl_39);
            dialog.Controls.Add(ctrl_40);
            dialog.Controls.Add(ctrl_41);
            dialog.Controls.Add(ctrl_42);
            dialog.Controls.Add(ctrl_43);

            dialog.CallingMethod = System.Reflection.MethodBase.GetCurrentMethod(); 
            dialog.DialogScaling = 125.0;  // Use this if the dialog was designed using a display with 120 DPI

            ctrl_02.IO_index  = CctOptions.UseGlobalSettings ? 1 : 0;

            ctrl_08.IO_string = CctOptions.GlobalAuthority;
            ctrl_10.IO_double = CctOptions.GlobalAccuracy;

            ctrl_13.IO_double = CctOptions.GlobalWestLongitude;
            ctrl_15.IO_double = CctOptions.GlobalEastLongitude;
            ctrl_17.IO_double = CctOptions.GlobalSouthLatitude;
            ctrl_19.IO_double = CctOptions.GlobalNorthLatitude;

            ctrl_21.IO_checked = CctOptions.TransformOptions.NoBallparkConversions;
            ctrl_22.IO_checked = CctOptions.TransformOptions.NoDiscardIfMissing;
            ctrl_23.IO_checked = CctOptions.TransformOptions.UsePrimaryGridNames;
            ctrl_24.IO_checked = CctOptions.TransformOptions.UseSuperseded;
            ctrl_25.IO_checked = CctOptions.AllowDeprecatedCRS;
            ctrl_26.IO_checked = CctOptions.TransformOptions.StrictContains;
            ctrl_27.IO_index   = (int) CctOptions.TransformOptions.IntermediateCrsUsage;

            bool bOK = dialog.ShowDialog(ValidateTransformDialog);
            if (bOK == false) return;

            CctOptions.UseGlobalSettings = ctrl_02.IO_index > 0;

            if (CctOptions.UseGlobalSettings)
                MessageBox.Show(
                    $"These global parameters will be applied to all open spreadsheets. Please be aware this may affect the outcome of earlier transforms.", 
                    "Please note", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            
            CctOptions.GlobalAuthority = ctrl_08.IO_string;
            CctOptions.GlobalAccuracy  = ctrl_10.IO_double;

            // Not the quikest way to set these 4 parameters, but it ensures they area saved in the config file...
            CctOptions.GlobalWestLongitude = ctrl_13.IO_double;
            CctOptions.GlobalEastLongitude = ctrl_15.IO_double;
            CctOptions.GlobalSouthLatitude = ctrl_17.IO_double;
            CctOptions.GlobalNorthLatitude = ctrl_19.IO_double;
            CctOptions.VerifyTransformArea();

            // just to debug
            if (CctOptions.GlobalWestLongitude == -1000) CctOptions.TransformOptions.Area = null;

            CctOptions.TransformOptions.NoBallparkConversions = ctrl_21.IO_checked;
            CctOptions.TransformOptions.NoDiscardIfMissing    = ctrl_22.IO_checked;
            CctOptions.TransformOptions.UsePrimaryGridNames   = ctrl_23.IO_checked;
            CctOptions.TransformOptions.UseSuperseded         = ctrl_24.IO_checked;
            CctOptions.AllowDeprecatedCRS                     = ctrl_25.IO_checked;
            CctOptions.TransformOptions.StrictContains        = ctrl_26.IO_checked;
            CctOptions.TransformOptions.IntermediateCrsUsage  = (IntermediateCrsUsage) ctrl_27.IO_index;

            int optionsFlag = 0;
            if (CctOptions.TransformOptions.NoBallparkConversions) optionsFlag += 8;
            if (CctOptions.TransformOptions.NoDiscardIfMissing   ) optionsFlag += 16;
            if (CctOptions.TransformOptions.UsePrimaryGridNames  ) optionsFlag += 32;
            if (CctOptions.TransformOptions.UseSuperseded        ) optionsFlag += 64;
            if (CctOptions.AllowDeprecatedCRS                    ) optionsFlag += 128;
            if (CctOptions.TransformOptions.StrictContains       ) optionsFlag += 256;
            if (CctOptions.TransformOptions.IntermediateCrsUsage == IntermediateCrsUsage.Always) optionsFlag += 512;
            if (CctOptions.TransformOptions.IntermediateCrsUsage == IntermediateCrsUsage.Never)  optionsFlag += 1024;

            CctOptions.GlobalTransformParameter = optionsFlag;

            // Get the correct application instance
            Microsoft.Office.Interop.Excel.Application xlapp = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
            xlapp.Calculate();

        } // TransformSettingsDialog

        [ExcelCommand(
            Name = "Dialog_TopoLib_Version",
            Description = "Shows a dialog with information on library version and compilation date & time",
            HelpTopic = "TopoLib-AddIn.chm!1210")]
        public static void Dialog_TopoLibVersion()
        {
            var dialog  = new XlDialogBox()             {                    W = 313, H = 200, Text = "Version Info"};
            var ctrl_01 = new XlDialogBox.GroupBox()    {  X = 013, Y = 013, W = 287, H = 130, Text = "Topographical function library",  };
            var ctrl_02 = new XlDialogBox.Label()       {  X = 031, Y = 039,                   Text = "Library version",  };
            var ctrl_03 = new XlDialogBox.TextEdit()    {  X = 031, Y = 058, W = 250,          };
            var ctrl_04 = new XlDialogBox.Label()       {  X = 031, Y = 091,                   Text = "Library compile date",  };
            var ctrl_05 = new XlDialogBox.TextEdit()    {  X = 031, Y = 110, W = 250,          };
            var ctrl_06 = new XlDialogBox.OkButton()    {  X = 031, Y = 160, W = 100,          Text = "&OK", Default = true, };
            var ctrl_07 = new XlDialogBox.HelpButton2() {  X = 181, Y = 160, W = 100,          Text = "&Help",  };

            dialog.Controls.Add(ctrl_01);
            dialog.Controls.Add(ctrl_02);
            dialog.Controls.Add(ctrl_03);
            dialog.Controls.Add(ctrl_04);
            dialog.Controls.Add(ctrl_05);
            dialog.Controls.Add(ctrl_06);
            dialog.Controls.Add(ctrl_07);

            Version v = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            string version = v.ToString();

            System.DateTime date_time = System.IO.File.GetLastWriteTime(ExcelDnaUtil.XllPath);
            string compileDate = date_time.ToString();

            ctrl_03.IO_string = version;
            ctrl_05.IO_string = compileDate;
            dialog.CallingMethod = System.Reflection.MethodBase.GetCurrentMethod(); 
            
            bool bOK = dialog.ShowDialog(Validate);
            if (bOK == false) return;

        } // VersionDialog
    }
}


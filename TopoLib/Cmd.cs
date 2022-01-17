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
using Microsoft.Office.Interop.Excel;

namespace TopoLib
{
    public class R1C1
    {
        private string _r1c1;

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
            int i = index;

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
        static bool ValidateAbout(int index, object[,] dialogResult, XlDialogBox.XlDialogControlCollection Controls)
        {
            System.Diagnostics.Process.Start("https://www.github.com/duijndam-dev/");
            return true; // return to dialog
        }

        [ExcelCommand(
            Name = "Show_Help",
            Description = "Shows the Compiled Help file",
            HelpTopic = "TopoLib-AddIn.chm!1200")]
        public static void ShowHelp()
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
        } // ShowHelp

        [ExcelCommand(
            Name = "About_TopoLib",
            Description = "Shows a dialog with a copy right statement and a list of referenced NuGet packages",
            HelpTopic = "TopoLib-AddIn.chm!1201")]
        public static void ShowPackages()
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
            bool bOK = dialog.ShowDialog(ValidateAbout);
            if (bOK == false) return;
        }

        [ExcelCommand(
            Name = "Version_Info",
            Description = "Shows a dialog with information on library version and compilation date & time",
            HelpTopic = "TopoLib-AddIn.chm!1202")]
        public static void ShowVersion()
        {
            var dialog  = new XlDialogBox()             {                    W = 313, H = 200, Text = "Version Info"};
            var ctrl_01 = new XlDialogBox.GroupBox()    {  X = 013, Y = 013, W = 287, H = 130, Text = "Geophysical and Geomatics function library",  };
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
        }

        [ExcelCommand(
            Name = "Logging_Dialog",
            Description = "Sets the logging level for the TopoLib AddIn",
            HelpTopic = "TopoLib-AddIn.chm!1203")]
        public static void LoggingDialog()
        {
            var dialog  = new XlDialogBox()                  {	                   W = 270, H = 150, Text = "TopoLib logging level",  IO = 2, };
            var ctrl_01 = new XlDialogBox.Label()            {	 X = 020, Y = 010,                   Text = "Select the required logging level in the option list below",  };
            var ctrl_02 = new XlDialogBox.GroupBox()         {	 X = 020, Y = 030, W = 120, H = 100, Text = "Error logging ",  };
            var ctrl_03 = new XlDialogBox.RadioButtonGroup() {	                                     IO = 3, };
            var ctrl_04 = new XlDialogBox.RadioButton()      {	          Y = 045,                   Text = "&None",  };
            var ctrl_05 = new XlDialogBox.RadioButton()      {	          Y = 065,                   Text = "&Errors",  };
            var ctrl_06 = new XlDialogBox.RadioButton()      {	          Y = 085,                   Text = "&Debug",  };
            var ctrl_07 = new XlDialogBox.RadioButton()      {	          Y = 105,                   Text = "&Verbose (trace)",  IO = true, };
            var ctrl_08 = new XlDialogBox.OkButton()         {	 X = 170, Y = 060, W = 080,          Text = "&OK", Default = true, };
            var ctrl_09 = new XlDialogBox.CancelButton()     {	 X = 170, Y = 085, W = 080,          Text = "&Cancel",  };
            var ctrl_10 = new XlDialogBox.HelpButton2()      {	 X = 170, Y = 110, W = 080,          Text = "&Help",  };

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

            ctrl_03.IO_index = Lib.LogLevel;

            bool bOK = dialog.ShowDialog(Validate);
            if (bOK == false) return;

            Lib.LogLevel = ctrl_03.IO_index;

            Cfg.AddOrUpdateKey("LogLevel", ((int)ctrl_03.IO_index).ToString());

        } // LoggingDialog

        [ExcelCommand(
            Name = "PROJ_LIB_Dialog",
            Description = "Starts a File Selector Dialog to set the PROJ_LIB environment variable",
            HelpTopic = "TopoLib-AddIn.chm!1204"
        )]
        public static void ProjLibDialog()
        {
            var dialog  = new XlDialogBox()                  {	                   W = 420, H = 240, Text = "Define location of PROJ grid (*.tif) files ",  IO =  7, };
            var ctrl_01 = new XlDialogBox.GroupBox()         {	 X = 013, Y = 010, W = 394, H = 040, Text = "PROJ_LIB folder at launch of dialog. Use ⭮ button to refresh ",  };
            var ctrl_02 = new XlDialogBox.OkButton()         {	 X = 018, Y = 026, W = 025,          Text = "⭮",  IO = 3, };
            var ctrl_03 = new XlDialogBox.DirectoryLabel()   {	 X = 048, Y = 030, W = 357,          };
            var ctrl_04 = new XlDialogBox.GroupBox()         {	 X = 013, Y = 055, W = 394, H = 140, Text = "File selector. Use *.* to search for all files in a folder ",  };
            var ctrl_05 = new XlDialogBox.TextEdit()         {	 X = 031, Y = 073, W = 170,          IO = "*.*", };
            var ctrl_06 = new XlDialogBox.LinkedFilesList()  {	 X = 220, Y = 073, W = 170, H = 110, IO = 2, };
            var ctrl_07 = new XlDialogBox.LinkedDriveList()  {	 X = 031, Y = 096,          H = 090, };
            var ctrl_08 = new XlDialogBox.OkButton()         {	 X = 151, Y = 205, W = 075,          Text = "&OK",  };
            var ctrl_09 = new XlDialogBox.CancelButton()     {	 X = 238, Y = 205, W = 075,          Text = "&Cancel",  };
            var ctrl_10 = new XlDialogBox.HelpButton2()      {	 X = 330, Y = 205, W = 075,          Text = "&Help",  IO = -1, };

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

            string key = "PROJ_LIB";
            string sProjLib = Environment.GetEnvironmentVariable(key, EnvironmentVariableTarget.User);
            string sOldDir = sProjLib; 

            // The controls in this dialog function by modifying the Current Directory.
            // No information from the controls themselves is being used.

            if( ! string.IsNullOrEmpty(sProjLib))
            {
                Directory.SetCurrentDirectory(sProjLib);
            }

            dialog.CallingMethod = System.Reflection.MethodBase.GetCurrentMethod(); 
            dialog.DialogScaling = 125.0;  // Use this if the dialog was designed using a display with 120 DPI
            bool bOK = dialog.ShowDialog(Validate);

            if (bOK == false) return;

            sProjLib = Directory.GetCurrentDirectory();

            if (sOldDir != sProjLib)
            {
                // changes have been made; we need to do something
                Environment.SetEnvironmentVariable(key,  sProjLib, EnvironmentVariableTarget.User);
                MessageBox.Show(
                    "The PROJ_LIB environment variable has been updated. You need to close and re-open Excel for this change to have effect", 
                    "Please note", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        } // ProjLibDialog

        [ExcelCommand(
            Name = "GlobalSettings_Dialog",
            Description = "Sets global transform parameters for the TopoLib AddIn",
            HelpTopic = "TopoLib-AddIn.chm!1205")]
        public static void GlobalOptionsDialog()
        {
            var dialog  = new XlDialogBox()                  {	                   W = 500, H = 280, Text = "Transform Optional Parameters",  };
            var ctrl_01 = new XlDialogBox.GroupBox()         {	 X = 020, Y = 020, W = 190, H = 065, Text = "&Select Optional Parameters",  };
            var ctrl_02 = new XlDialogBox.RadioButtonGroup() {	                                     IO = 1, };
            var ctrl_03 = new XlDialogBox.RadioButton()      {	          Y = 040,                   Text = "From the &Mode Flags",  };
            var ctrl_04 = new XlDialogBox.RadioButton()      {	          Y = 060,                   Text = "From &Global Settings",  IO = 1, };
            var ctrl_05 = new XlDialogBox.GroupBox()         {	 X = 020, Y = 095, W = 190, H = 060, Text = "&Miscellaneous",  IO = 2, };
            var ctrl_06 = new XlDialogBox.Label()            {	 X = 030, Y = 110, W = 075,          Text = "Authority",  };
            var ctrl_07 = new XlDialogBox.TextEdit()         {	 X = 030, Y = 126, W = 075,          IO = "EPSG", };
            var ctrl_08 = new XlDialogBox.Label()            {	 X = 120, Y = 110, W = 075,          Text = "Accurac&y [m]",  };
            var ctrl_09 = new XlDialogBox.DoubleEdit()       {	 X = 120, Y = 126, W = 075,          IO = -1000, };
            var ctrl_10 = new XlDialogBox.GroupBox()         {	 X = 020, Y = 165, W = 190, H = 100, Text = "Useage Area",  IO = 2, };
            var ctrl_11 = new XlDialogBox.Label()            {	 X = 030, Y = 180,                   Text = "Min Lat. [°]",  };
            var ctrl_12 = new XlDialogBox.DoubleEdit()       {	 X = 030, Y = 196, W = 075,          IO = -1000, };
            var ctrl_13 = new XlDialogBox.Label()            {	 X = 120, Y = 180,                   Text = "Max Lat. [°]",  };
            var ctrl_14 = new XlDialogBox.DoubleEdit()       {	 X = 120, Y = 196, W = 075,          IO = -1000, };
            var ctrl_15 = new XlDialogBox.Label()            {	 X = 030, Y = 220,                   Text = "Min Long. [°]",  };
            var ctrl_16 = new XlDialogBox.DoubleEdit()       {	 X = 030, Y = 236, W = 075,          IO = -1000, };
            var ctrl_17 = new XlDialogBox.Label()            {	 X = 120, Y = 220,                   Text = "Max Long. [°]",  };
            var ctrl_18 = new XlDialogBox.DoubleEdit()       {	 X = 120, Y = 236, W = 075,          IO = -1000, };
            var ctrl_19 = new XlDialogBox.GroupBox()         {	 X = 230, Y = 020, W = 250, H = 210, Text = "&Optional Parameters",  };
            var ctrl_20 = new XlDialogBox.CheckBox()         {	          Y = 040, W = 230,          Text = "Disallow &Ballpark Conversions",  IO = false, };
            var ctrl_21 = new XlDialogBox.CheckBox()         {	          Y = 060, W = 230,          Text = "Allow if &Grid is Missing",  IO = false, };
            var ctrl_22 = new XlDialogBox.CheckBox()         {	          Y = 080, W = 230,          Text = "Use &Primary Grid Names",  IO = false, };
            var ctrl_23 = new XlDialogBox.CheckBox()         {	          Y = 100, W = 230,          Text = "Use &Superseded Transforms",  IO = false, };
            var ctrl_24 = new XlDialogBox.CheckBox()         {	          Y = 120, W = 230,          Text = "Allow &Deprecated CRSs",  IO = false, };
            var ctrl_25 = new XlDialogBox.CheckBox()         {	          Y = 140, W = 230,          Text = "Strictly &Contains Area",  IO = false, };
            var ctrl_26 = new XlDialogBox.CheckBox()         {	          Y = 160, W = 230,          Text = "&Always Allow Intermediate CRS",  IO = false, };
            var ctrl_27 = new XlDialogBox.CheckBox()         {	          Y = 180, W = 230,          Text = "&Never Allow Intermediate CRS",  IO = false, };
            var ctrl_28 = new XlDialogBox.CheckBox()         {	          Y = 200, W = 230,          Text = "&Use Network Connection",  IO = false, };
            var ctrl_29 = new XlDialogBox.Label()            {	 X = 430, Y = 020, W = 030,          Text = "(flag)",  };
            var ctrl_30 = new XlDialogBox.Label()            {	 X = 430, Y = 042, W = 040,          Text = "(8)",  };
            var ctrl_31 = new XlDialogBox.Label()            {	 X = 430, Y = 062, W = 040,          Text = "(16)",  };
            var ctrl_32 = new XlDialogBox.Label()            {	 X = 430, Y = 082, W = 040,          Text = "(32)",  };
            var ctrl_33 = new XlDialogBox.Label()            {	 X = 430, Y = 102, W = 040,          Text = "(64)",  };
            var ctrl_34 = new XlDialogBox.Label()            {	 X = 430, Y = 122, W = 040,          Text = "(128)",  };
            var ctrl_35 = new XlDialogBox.Label()            {	 X = 430, Y = 142, W = 040,          Text = "(256)",  };
            var ctrl_36 = new XlDialogBox.Label()            {	 X = 430, Y = 162, W = 040,          Text = "(512)",  };
            var ctrl_37 = new XlDialogBox.Label()            {	 X = 430, Y = 182, W = 040,          Text = "(1024)",  };
            var ctrl_38 = new XlDialogBox.Label()            {	 X = 430, Y = 202, W = 040,          Text = "(2048)",  };
            var ctrl_39 = new XlDialogBox.OkButton()         {	 X = 230, Y = 245, W = 075,          Text = "&OK", Default = true, };
            var ctrl_40 = new XlDialogBox.CancelButton()     {	 X = 317, Y = 245, W = 075,          Text = "&Cancel",  };
            var ctrl_41 = new XlDialogBox.HelpButton2()      {	 X = 404, Y = 245, W = 075,          Text = "&Help",  };

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
            dialog.CallingMethod = System.Reflection.MethodBase.GetCurrentMethod(); 
            dialog.DialogScaling = 125.0;  // Use this if the dialog was designed using a display with 120 DPI
            bool bOK = dialog.ShowDialog(Validate);
            if (bOK == false) return;
        } // GlobalOptionsDialog

    }
}


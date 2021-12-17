using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using ExcelDna.Integration;
using ExcelDna.Documentation;

// Added Bart
using SharpProj;
using SharpProj.Proj;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Excel = Microsoft.Office.Interop.Excel; // Bart, from https://stackoverflow.com/questions/7916711/get-the-current-workbook-object-in-c-sharp
using System.IO;

namespace TopoLib
{
    public static class Lib
    {
        static int _verbose = 0;

        [ExcelFunction(IsHidden = true)]
        public static int GetVerboseLevel()
        {
            return _verbose;
        }

        [ExcelFunctionDoc(
            IsVolatile = true,
            Name = "TL.lib.Verbose",
            Description = "Returns or sets error handling method",
            Category = "LIB - Library",
            HelpTopic = "TopoLib-AddIn.chm!1200",

            Returns = "the current verbose setting",
            Remarks = "This function takes 1 input parameter 'ErrorHandling' to set the Verbose Level, to manage error handling: " +
                        "<ul>    <li><p>0 = Excel compliant error handling; return #N/A! in case of an error</p></li>" +
                                "<li><p>1 = Where appropriate, return a numerical value, (e.g. -1 or FALSE). Otherwise return error string</p></li>" +
                                "<li><p>2 = Return error string in format: #:[exception message]</p></li></ul>" +
                        "<p>Use the function without an input parameter to get the current verbose setting</p></li>",
            Example = "xxx")]
        public static object LibVerbose(
             [ExcelArgument("Use no arguments to get the current setting; use one argument [0, 1, 2] (0) to set the verbose setting", Name = "ErrorHandling")] object oVerbose)
        {
            // if (ExcelDnaUtil.IsInFunctionWizard()) return ExcelError.ExcelErrorRef;

            int tmp = (int)Optional.Check(oVerbose, (double)_verbose);
             if (tmp < 0 || tmp > 2 ) return ExcelError.ExcelErrorValue;

            _verbose = tmp;

            return (double)_verbose;
        }

        // See: https://stackoverflow.com/questions/136035/catch-multiple-exceptions-at-once
        internal static object ExceptionHandler(Exception ex, object hint = null)
        {
            if (_verbose == 0) 
                return ExcelError.ExcelErrorNA;

            if (_verbose == 1 && hint != null)
                return hint;

            //string errorText = "#:[unknown exception]";
            string errorText;
            
            switch (ex)
            {
                case ArithmeticException _:
                case InvalidOperationException _:
                case ProjException _:
                    errorText =  $"#:[{ex.GetBaseException().Message}]";
                    break;

                case ArgumentNullException _:
                    errorText =  $"#:[{ex.GetBaseException().Message} argument is null]";
                    break;

                case ArgumentException _:
                    errorText =  $"#:[{ex.GetType().Name}: {ex.Message}]";
                    break;
                   
                default:
                    // you can check here [F9] if certain exception types haven't been handled yet
                    errorText =  $"#:[{ex.GetBaseException().Message}]";
                    break;
            }
            return errorText;
        }

        static bool? _supportsDynamicArrays;

        [ExcelFunction(IsHidden = true)]
        private static bool SupportsDynamicArrays()
        {
            if (!_supportsDynamicArrays.HasValue)
            {
                try
                {
                    var result = XlCall.Excel(614, new object[] { 1 }, new object[] { true });
                    _supportsDynamicArrays = true;
                }
                catch
                {
                    _supportsDynamicArrays = false;
                }
            }
            return _supportsDynamicArrays.Value;
        }

        [ExcelFunctionDoc(
            Name = "TL.lib.DynamicArraysSupported",
            Description = "Indicates if your version of Excel supports 'dynamic arrays'",
            Category = "LIB - Library",
            HelpTopic = "TopoLib-AddIn.chm!1201",
            IsHidden = false,
            Returns = "A string showing if your version of Excel supports dynamic arrays",
            Remarks = "This function takes no input parameters",
            Example = "TL.lib.DynamicArraysSupported() returns 'Dynamic arrays are supported'")]
        public static string LibDynamicArrays()
        {
            //            if (ExcelDnaUtil.IsInFunctionWizard())
            //                return ExcelError.ExcelErrorRef;
            //            else
            {
                string strYes = "Dynamic arrays are supported";
                string strNo  = "Dynamic arrays not supported";

                if (SupportsDynamicArrays())
                    return strYes;
                else
                    return strNo;
            }
        }

        [ExcelFunctionDoc(
            Name = "TL.lib.OperatingSystem",
            Description = "Returns version of Operating System and Excel",
            Category = "LIB - Library",
            HelpTopic = "TopoLib-AddIn.chm!1202",

            Returns = "Name and version number of the current operating system",
            Remarks = "This function takes no input parameters",
            Example = "\"Windows (64-bit) NT :.00\" with Win10(64 - bit) and Excel 2016(16.0.6326.1010, 64 - bit)")]
        public static string LibOS()
        {
            //            if (ExcelDnaUtil.IsInFunctionWizard())
            //                return ExcelError.ExcelErrorRef;
            //            else
            {
                Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
                string strOS = xlApp.OperatingSystem;
                return strOS;
            }
        }

        [ExcelFunctionDoc(
            Name = "TL.lib.TopoLibVersion",
            Description = "Returns the TopoLib library version",
            Category = "LIB - Library",
            HelpTopic = "TopoLib-AddIn.chm!1203",

            Returns = "the current version of the installed library ",
            Remarks = "This function takes no input parameters",
            Example = "TL.lib.Version() returns 2.1.1234.5678")]
        public static object TopoLibVersion()
        {
            if (ExcelDnaUtil.IsInFunctionWizard()) return ExcelError.ExcelErrorRef;

            Version v = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            string version = v.ToString();
            return version;
        }

        [ExcelFunctionDoc(
            Name = "TL.lib.CompileTime",
            Description = "Returns the time the library was compiled",
            Category = "LIB - Library",
            HelpTopic = "TopoLib-AddIn.chm!1204",

            Returns = "the compilation date / time of the installed library ",
            Remarks = "This function takes no input parameters",
            Example = "TL.lib.CompileTime() returns 15/04/2021 15:16:38")]
        public static string LibCompileDateTime()
        {
//            if (ExcelDnaUtil.IsInFunctionWizard())
//                return ExcelError.ExcelErrorRef;
//            else
            {
                System.DateTime date_time = System.IO.File.GetLastWriteTime(ExcelDnaUtil.XllPath);
                string date = date_time .ToString();
                return date;
            }
        }

        [ExcelFunctionDoc(
            Name = "TL.lib.InstallationPath",
            Description = "Returns the path where library is installed",
            Category = "LIB - Library",
            HelpTopic = "TopoLib-AddIn.chm!1205",

            Returns = "the path where the AddInis installed",
            Remarks = "This function takes no input parameters")]
        public static string LibInstallationPath()
        {
//            if (ExcelDnaUtil.IsInFunctionWizard())
//                return ExcelError.ExcelErrorRef;
//            else
            {
                // get the Path of xll file;
                string xllPath = ExcelDnaUtil.XllPath;
                return xllPath;
            }
        }

        [ExcelFunctionDoc(
            Name = "TL.lib.ProjVersionInfo",
            Category = "LIB - Library",
            Description = "Version of Proj library and its database(s)",
            HelpTopic = "TopoLib-AddIn.chm!1206",

            Returns = "A string with version information",
            Summary = "Function that returns version of Proj-library or its database(s)",
            Example = "Topo.prj.ProjVersionInfo() returns: 8.1.1"
        )]
        public static string ProjVersionInfo(
             [ExcelArgument("Version Info (0); 0 = LibVersion, 1 = EpsGVersion, 2 = EsriVersion, 3 = IgnfVersion", Name = "Mode")] object mode)
        {
            int nMode = (int)Optional.Check(mode, 0.0);
            
            string info;

            using (var pc = new ProjContext())
            {
                switch (nMode)
                {
                    case 0:
                    default:
                        info = pc.Version.ToString();
                        break;
                    case 1:
                        info = pc.EpsgVersion.ToString();
                        break;
                    case 2:
                        info = pc.EsriVersion.ToString();
                        break;
                    case 3:
                        info = pc.IgnfVersion.ToString();
                        break;
                }
            }

            return info;

        } // Version

        static string _GridCachePath = "";

        [ExcelFunction(IsHidden = true)]
        public static string GetGridCachePath()
        {
            return _GridCachePath ;
        }

        [ExcelFunctionDoc(
            Name = "TL.lib.GridCacheLocation",
            Description = "Gets or sets the path for the cached grid files",
            Category = "LIB - Library",
            HelpTopic = "TopoLib-AddIn.chm!1207",
            IsHidden = false,
            Returns = "A string showing the path for the cached grid files",
            Remarks = "This function takes no input parameters",
            Example = "<p>This function takes 1 input parameter set the path for the cached grid files</p>" +
                      "<p>Use the function without an input parameter to get the current path for the cached grid files</p>")]
        public static string GridCacheLocation(
             [ExcelArgument("Use no arguments to get the current path; use one argument to set the path", Name = "GridCache")] object oGridCache)
        {
            // if (ExcelDnaUtil.IsInFunctionWizard()) return ExcelError.ExcelErrorRef;

            string tmp = Optional.Check(oGridCache, _GridCachePath);

            if (tmp != _GridCachePath)
            {
                // something has changed
                _GridCachePath = tmp;
            }

            if (String.IsNullOrWhiteSpace(_GridCachePath))
            {
                // setup default cache location
                string APPDATA_PATH = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData); // local AppData folder
                string CFGFOLDER_PATH = Path.Combine(APPDATA_PATH, "Topo");     // Path for config folder
                // string CFGFILE_PATH = Path.Combine(CFGFOLDER_PATH, "config.txt");   // Path for config.txt file

                _GridCachePath = CFGFOLDER_PATH;
            }
            return _GridCachePath;

        } // GridCacheLocation

    }
}

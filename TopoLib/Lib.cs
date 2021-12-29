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

// to refresh my memory on access modifiers in C# :
// *internal* is for assembly scope (i.e. only accessible from code in the same .exe or .dll)
// *private* is for class scope (i.e. accessible only from code in the same class).

namespace TopoLib
{
    public static class Lib
    {
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
            Returns = "A string showing if your version of Excel supports dynamic arrays",
            Remarks = "This function takes no input parameters",
            Example = "TL.lib.DynamicArraysSupported() returns 'Dynamic arrays are supported'")]
        public static string DynamicArrays()
        {
            string strYes = "Dynamic arrays are supported";
            string strNo  = "Dynamic arrays not supported";

            if (SupportsDynamicArrays())
                return strYes;
            else
                return strNo;

        } // DynamicArrays

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

        } // ExceptionHandler

        static string _gridCachePath = "";

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

            string tmp = Optional.Check(oGridCache, _gridCachePath);

            if (tmp != _gridCachePath)
            {
                // something has changed
                _gridCachePath = tmp;
            }

            if (String.IsNullOrWhiteSpace(_gridCachePath))
            {
                // setup default cache location
                string APPDATA_PATH = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData); // local AppData folder
                string CFGFOLDER_PATH = Path.Combine(APPDATA_PATH, "TopoLib");                                      // Path for config folder

                _gridCachePath = CFGFOLDER_PATH;
            }
            return _gridCachePath;

        } // GridCacheLocation

        [ExcelFunctionDoc(
            Name = "TL.lib.InstallationPath",
            Description = "Returns the path where library is installed",
            Category = "LIB - Library",
            HelpTopic = "TopoLib-AddIn.chm!1205",

            Returns = "the path where the AddIn is installed",
            Remarks = "This function takes no input parameters")]
        public static string InstallationPath()
        {
            // get the Path of xll file;
            string xllPath = ExcelDnaUtil.XllPath;
            return xllPath;

        } // InstallationPath

        [ExcelFunctionDoc(
            Name = "TL.lib.OperatingSystem",
            Description = "Returns version of Operating System and Excel",
            Category = "LIB - Library",
            HelpTopic = "TopoLib-AddIn.chm!1202",

            Returns = "Name and version number of the current operating system",
            Remarks = "This function takes no input parameters",
            Example = "\"Windows (64-bit) NT :.00\" with Win10(64 - bit) and Excel 2016(16.0.6326.1010, 64 - bit)")]
        public static string OperatingSystem()
        {
            //            if (ExcelDnaUtil.IsInFunctionWizard())
            //                return ExcelError.ExcelErrorRef;
            //            else
            {
                Excel.Application xlApp = (Excel.Application)ExcelDnaUtil.Application;
                string strOS = xlApp.OperatingSystem;
                return strOS;
            }
        } // OperatingSystem


        private static int _verbose = 0;

        [ExcelFunctionDoc(
            IsVolatile = true,
            Name = "TL.lib.Verbose",
            Description = "Returns or sets error handling method",
            Category = "LIB - Library",
            HelpTopic = "TopoLib-AddIn.chm!1200",

            Returns = "the current verbose setting",
            Remarks = "This function takes 1 input parameter 'ErrorHandling' to set the Verbose Level, to manage error handling: " +
                        "<ul>    <li><p>0 = Excel compliant error handling; return #N/A! in case of an error</p></li>" +
                                "<li><p>1 = Where appropriate, return a numerical value, (e.g. -1 or FALSE). Otherwise return #N/A! error</p></li>" +
                                "<li><p>2 = Return error string in format: #:[exception message]</p></li></ul>" +
                        "<p>Use the function without an input parameter to get the current verbose setting</p>" +
                        "<p>This functionality will be updated when a 'decent' logger has been added to the library</p>.",
            Example = "xxx")]
        public static object Verbose(
             [ExcelArgument("Use no arguments to get the current setting; use one argument [0, 1, 2] (0) to set the verbose setting", Name = "ErrorHandling")] object oVerbose)
        {
            // if (ExcelDnaUtil.IsInFunctionWizard()) return ExcelError.ExcelErrorRef;

            int tmp = (int)Optional.Check(oVerbose, (double)_verbose);
             if (tmp < 0 || tmp > 2 ) return ExcelError.ExcelErrorValue;

            _verbose = tmp;

            return (double)_verbose;
        
        } // Verbose

        [ExcelFunctionDoc(
            Name = "TL.lib.VersionInfo",
            Category = "LIB - Library",
            Description = "Version of Proj library and its database(s)",
            HelpTopic = "TopoLib-AddIn.chm!1206",

            Returns = "A string with version information",
            Summary = "Function that returns version of Proj-library or its database(s)",
            Example = "Topo.prj.ProjVersionInfo() returns: 8.1.1"
        )]
        public static string VersionInfo(
             [ExcelArgument("Version Info (0); 0 = Proj lib version, 1 = EPSG version, 2 = ESRI version, 3 = IGNF vesrion , 4 = TopoLib version, 5 = Compile time", Name = "Mode")] object oMode)
        {
            int nMode = (int)Optional.Check(oMode, 0.0);
            
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
                    case 4:
                        Version v = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                        info = v.ToString();
                        break;
                    case 5:
                        System.DateTime date_time = System.IO.File.GetLastWriteTime(ExcelDnaUtil.XllPath);
                        info = date_time .ToString();
                        break;
                }
            }
            return info;

        } // VersionInfo

    }
}

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
// using Microsoft.VisualStudio.TestTools.UnitTesting;
using Excel = Microsoft.Office.Interop.Excel; // Bart, from https://stackoverflow.com/questions/7916711/get-the-current-workbook-object-in-c-sharp
using System.IO;

namespace TopoLib
{
    public static class Lib
    {
        static bool? _supportsDynamicArrays;

        [ExcelFunction(IsHidden = true)]
        internal static bool SupportsDynamicArrays()
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
            HelpTopic = "TopoLib-AddIn.chm!1500",
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

        static string _gridCachePath = "";

        [ExcelFunctionDoc(
            Name = "TL.lib.GridCacheLocation",
            Description = "Gets or sets the path for the cached grid files",
            Category = "LIB - Library",
            HelpTopic = "TopoLib-AddIn.chm!1501",
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
            HelpTopic = "TopoLib-AddIn.chm!1502",

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
            HelpTopic = "TopoLib-AddIn.chm!1503",

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

        [ExcelFunctionDoc(
            IsVolatile = true,
            Name = "TL.lib.LoggingLevel",
            Description = "Returns or sets the logging level",
            Category = "LIB - Library",
            HelpTopic = "TopoLib-AddIn.chm!1504",

            Returns = "the current logging level",
            Remarks = "This function takes 1 input parameter to set the logging Level, to QC error handling: " +
                        "<ul>    <li><p>0 = Verbose - all messages being logged (default)</p></li>" +
                                "<li><p>1 = Debug messages being logged </p></li>" +
                                "<li><p>2 = Warnings being logged </p></li>" +
                                "<li><p>3 = Errors (only) being logged</p></li>" +
                        "</ul>" +
                        "<p>Use the function without an input parameter to get the current verbose setting</p>" +
                        "<p>Using this function to set the logging level is not recommended, as it may interfere with setting the logging level from the TopoLib Ribbon!</p>.",
            Example = "xxx")]
        public static object SetLoggingLevel(
             [ExcelArgument("Use no arguments to get the current setting; use one argument [0, 1, 2, 3] (0) to set the logging level", Name = "Logging")] object oLogging)
        {
            // if (ExcelDnaUtil.IsInFunctionWizard()) return ExcelError.ExcelErrorRef;

            int tmp = (int)Optional.Check(oLogging, (double)CctOptions.ProjContext.LogLevel); // use the current logging level as a default value
             if (tmp < 0 || tmp > 3 ) return ExcelError.ExcelErrorValue;

            CctOptions.ProjContext.LogLevel = (ProjLogLevel) tmp;

            return (double)tmp;
        
        } // SetLoggingLevel

        // For Version Info, see also:
        // https://proj.org/development/reference/functions.html?highlight=proj_operation_factory_context#various

        [ExcelFunctionDoc(
            Name = "TL.lib.VersionInfo",
            Category = "LIB - Library",
            Description = "Version of Proj library and its database(s)",
            HelpTopic = "TopoLib-AddIn.chm!1505",

            Returns = "A string with version information",
            Summary = "Function that returns version of Proj-library or its database(s)",
            Remarks = "This function takes 1 input parameter to get information from : " +
            "<ul>    <li><p>0 = Version of the PROJ-database (0 = default)</p></li>" +
                    "<li><p>1 = Version of the PROJ-data package with which this database is most compatible</p></li>" +
                    "<li><p>2 = Version of the EPSG-database</p></li>" +
                    "<li><p>3 = Version of the ESRI-database</p></li>" +
                    "<li><p>4 = Version of the IGNF-database</p></li>" +
                    "<li><p>5 = Version of the TopoLib AddIn</p></li>" +
                    "<li><p>6 = Compile time of the TopoLib AddIn</p></li>" +
            "</ul>",
            Example = "Topo.prj.ProjVersionInfo() returns: 8.2.1"
        )]
        public static string VersionInfo(
             [ExcelArgument("Version Info (0); 0 = PROJ database version, 1 = PROJ data version, 2 = EPSG databaseversion, 3 = ESRI databaseversion, 4 = IGNF database version ,  5 = TopoLib version, 6 = TopoLib compile time", Name = "Mode")] object oMode)
        {
            int nMode = (int)Optional.Check(oMode, 0.0);
            
            string info;

            using (ProjContext pjContext = Crs.CreateContext())
            {
                switch (nMode)
                {
                    case 0:
                    default:
                        info = pjContext.Version.ToString();
                        break;
                    case 1:
                        info = pjContext.ProjDataVersion.ToString();
                        break;
                    case 2:
                        info = pjContext.EpsgVersion.ToString();
                        break;
                    case 3:
                        info = pjContext.EsriVersion.ToString();
                        break;
                    case 4:
                        info = pjContext.IgnfVersion.ToString();
                        break;
                    case 5:
                        Version v = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                        info = v.ToString();
                        break;
                    case 6:
                        System.DateTime date_time = System.IO.File.GetLastWriteTime(ExcelDnaUtil.XllPath);
                        info = date_time .ToString();
                        break;
                }
            }
            return info;

        } // VersionInfo

    } // class Lib

} // namespace TopoLib


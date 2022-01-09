using ExcelDna.Integration;
using ExcelDna.IntelliSense;

// Next line is needed to set ComVisible to false. Not sure if it is neded. See:
// https://docs.microsoft.com/en-us/dotnet/api/system.runtime.interopservices.comvisibleattribute?view=net-6.0
// using System.Runtime.InteropServices;

// for more information, please consult the following web page: 
// https://github.com/Excel-DNA/IntelliSense/wiki/Usage-Instructions


namespace TopoLib
{
//    [ComVisible(false)]
    internal class ExcelAddin : IExcelAddIn
    {
        public void AutoOpen()
        {
            IntelliSenseServer.Install();
        }
        public void AutoClose()
        {
            IntelliSenseServer.Uninstall();
        }
    }
}

// See the following on the debugger kicking in with a breakpoint:
// https://github.com/Excel-DNA/IntelliSense/wiki/Usage-Instructions#note-about-the-loaderlock-managed-debugging-assistant
// Text copied below for convenience

// Note about the 'LoaderLock' Managed Debugging Assistant
// When debugging an add-in that includes the Integrated display server, you might see the following warning:

// LoaderLock occured:
// Managed Debugging Assistant 'LoaderLock' has detected a problem in 'C:\Program Files (x86)\Microsoft Office\Office14\EXCEL.EXE'.
// Additional information: Attempting managed execution inside OS Loader lock. Do not attempt to run managed code inside a DllMain or image initialization function since doing so can cause the application to hang.

// This warning relates to code in the IntelliSense server that monitors the Excel process for libraries loaded or unloaded. 
// I believe the callback we run here is safe to execute under the OS Loader Lock, so this warning can be switched off in the debugger.


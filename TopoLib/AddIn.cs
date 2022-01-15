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
using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using ExcelDna.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Text;
using System.Threading.Tasks;
using Serilog;

namespace TopoLib
{
    [ComVisible(true)]
    public class AddIn : IExcelAddIn
    {
        private static ILogger _log = Serilog.Log.Logger;

        public static ILogger Logger   // property
		{
			get { return _log; }   // get method
			set { _log =  value; }  // set method
		}

        public void AutoOpen()
        {
            _log = Serilog.Log.Logger = ConfigureLogging();

            Version v = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            string version = v.ToString();

            System.DateTime date_time = System.IO.File.GetLastWriteTime(ExcelDnaUtil.XllPath);
            string compileDate = date_time.ToString();

            _log.Information($"Starting TopoLib version {version}, compiled on {compileDate}.");

            IntelliSenseServer.Install();
        }

        public void AutoClose()
        {
            IntelliSenseServer.Uninstall();

            Serilog.Log.CloseAndFlush();
        }

        public static void ProcessUnhandledException(Exception ex, string message = null, [CallerMemberName] string caller = null)
        {
            try
            {
                _log.Error(ex, message ?? $"Unhandled exception on {caller}");
            }
            catch (Exception lex)
            {
                try
                {
                    Serilog.Debugging.SelfLog.WriteLine(lex.ToString());
                }
                catch
                {
                    // Do nothing...
                }
            }

            if (ex is TargetInvocationException && !(ex.InnerException is null))
            {
                ProcessUnhandledException(ex.InnerException, message, caller);
                return;
            }

#if DEBUG
            MessageBox.Show(ex.ToString(), "Unexpected Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
#else
            const string errorMessage = "An unexpected error ocurred. Please try again in a few minutes, and if the error persists, contact support";
            MessageBox.Show(errorMessage, "Unexpected Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
#endif
        }

        private static ILogger ConfigureLogging()
        {
            return new LoggerConfiguration()
                .MinimumLevel.Verbose()
                .WriteTo.ExcelDnaLogDisplay(displayOrder: DisplayOrder.NewestFirst)
                .CreateLogger();
        }
    }
}

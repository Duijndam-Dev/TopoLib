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
using System.Resources;
using System.Reflection;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using ExcelDna.Logging;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Serilog;


// for Ribbon fundamentals, look here :
// https://github.com/Excel-DNA/Tutorials/tree/master/Fundamentals/RibbonBasics

namespace TopoLib
{
    [ComVisible(true)]
    public class CustomRibbon : ExcelRibbon
    {
        private Excel.Application _excel;
        private IRibbonUI         _thisRibbon;
        private ILogger           _log = Serilog.Log.Logger;
 
        private void OnInvalidateRibbon(object obj)
        {
            _thisRibbon.Invalidate();
        }

        public override string GetCustomUI(string ribbonId)
        {
            _excel = (Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;

            _log = Serilog.Log.ForContext<CustomRibbon>();
            _log.Information("[TOP] Loading ribbon {ribbonId} via GetCustomUI", ribbonId);

            string ribbonXml =
                @"<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui' onLoad='OnLoad'>
                  <ribbon>
                    <tabs>
                      <tab id='TopoLibTap' label='TopoLib'>
                        <group id='xmlGroup'            label='Export functions'>
                            <button id='gpxButton'      label='GPX file'             imageMso='XmlExport'                       size='large' onAction='ExportGpxButton_OnAction' />
                            <button id='kmlButton'      label='KML file'             imageMso='XmlExport'                       size='large' onAction='ExportKmlButton_OnAction' />
                            <button id='wizButton'      label='CRS or Transform'     imageMso='ControlWizards'                  size='large' onAction='ExportWizButton_OnAction' />
                        </group>
                        <group id='SettingGroup'        label='TopoLib Settings'>
                            <button id='Proj_LibButton' label='Resource Settings'    imageMso='SiteColumnActionsColumnSettings' size='large' onAction='DialogResourceSettings_OnAction' />
                            <button id='OptionsButton'  label='Transform Settings'   imageMso='ColumnActionsColumnSettings'     size='large' onAction='DialogOptionsButton_OnAction' />
                            <button id='CacheButton'    label='Cache Settings'       imageMso='ColumnListSetting'               size='large' onAction='DialogCacheButton_OnAction' />
                            <button id='LogLevelButton' label='Logging Settings'     imageMso='ComAddInsDialog'                 size='large' onAction='DialogLogLevelButton_OnAction' />
                        </group>
                        <group id='RecalcGroup'         label='TopoLib Transforms'>
                            <button id='RecalcButton'   label='Refresh Transforms'    imageMso='RefreshWebView'                 size='large' onAction='RecalcButton_OnAction' />
                        </group>
                        <group id='LoggingGroup'        label='Test Logging Messages'>
                            <button id='ErrorButton'    label='Log Error'            imageMso='OutlineViewClose'   onAction='LogErrorButton_OnAction' />
                            <button id='DebugButton'    label='Log Debug'            imageMso='MoreControlsDialog' onAction='LogDebugButton_OnAction' />
                            <button id='VerboseButton'  label='Log Verbose'          imageMso='Callout'            onAction='LogVerboseButton_OnAction' />
                        </group>
                        <group id='LogDisplayGroup' label='Log Handling'>
                            <button id='ViewLogDisplayButton' label='View Log'       imageMso='FileDocumentInspect' size='large' onAction='LogViewDisplayButton_OnAction' />
                            <button id='ClearLogButton' label='Clear Log'            imageMso='Clear'               size='large' onAction='LogClearDisplayButton_OnAction' />
                            <separator id='DisplayOrderSeparator' />
                            <menu id='DisplayOrderMenu' label='Display Order' getImage='LogOrderMenu_GetImage'      size='large'>
                            <toggleButton id='LogNewestLastButton' label='Newest Last' imageMso='EndOfDocument' getPressed='LogNewestLastButton_GetPressed' onAction='LogNewestLastButton_OnAction' />
                            <toggleButton id='LogNewestFirstButton' label='Newest First' imageMso='StartOfDocument' getPressed='LogNewestFirstButton_GetPressed' onAction='LogNewestFirstButton_OnAction' />
                            </menu>
                        </group>
                        <group id='GLgroup2' label='About TopoLib'>
                            <button id='VersionButton' label='Version Info' imageMso='ResultsPaneAccessibilityMoreInfo'     size='large' onAction='OnVersionButtonPressed' />
                            <button id='HelpButton' label='Show Help'       imageMso='TentativeAcceptInvitation' size='large' onAction='OnHelpButtonPressed' />
                            <button id='AboutButton' label='About TopoLib'  imageMso='FontDialog'                size='large' onAction='OnAboutButtonPressed' />
                        </group>
                      </tab>
                    </tabs>
                  </ribbon>
                </customUI>";
                
            return ribbonXml;
        }
 
        public void OnLoad(IRibbonUI ribbon)
        {
            _thisRibbon = ribbon ?? throw new ArgumentNullException(nameof(ribbon));

            _excel.WorkbookActivate += OnInvalidateRibbon;
            _excel.WorkbookDeactivate += OnInvalidateRibbon;
            _excel.SheetActivate += OnInvalidateRibbon;
            _excel.SheetDeactivate += OnInvalidateRibbon;

            if (_excel.ActiveWorkbook == null)
            {
                _excel.Workbooks.Add();
            }

			// read the configuration settings when loading the Ribbon
            // At this point, all functions have been registered, and it should be safe to do so...
            CctOptions.ReadConfiguration();
        }

        public void DialogOptionsButton_OnAction(/*IRibbonControl control*/)
        {
            try
            {
                _excel.Application.Run("Dialog_Transform_Settings");
            }
            catch (Exception ex)
            {
                AddIn.ProcessUnhandledException(ex);
            }
        }

        public void DialogCacheButton_OnAction(/*IRibbonControl control*/)
        {
            try
            {
                _excel.Application.Run("Dialog_Cache_Settings");
            }
            catch (Exception ex)
            {
                AddIn.ProcessUnhandledException(ex);
            }
        }

        public void DialogLogLevelButton_OnAction(/*IRibbonControl control*/)
        {
            try
            {
                _excel.Application.Run("Dialog_Log_Settings");
            }
            catch (Exception ex)
            {
                AddIn.ProcessUnhandledException(ex);
            }
        }

        public void DialogResourceSettings_OnAction(/*IRibbonControl control*/)
        {
            try
            {
                _excel.Application.Run("Dialog_Resource_Settings");
            }
            catch (Exception ex)
            {
                AddIn.ProcessUnhandledException(ex);
            }
        }

        public void ExportGpxButton_OnAction(/*IRibbonControl control*/)
        {
            try
            {
                _excel.Application.Run("Dialog_Export_GPX_data");
            }
            catch (Exception ex)
            {
                AddIn.ProcessUnhandledException(ex);
            }
        }

        public void ExportKmlButton_OnAction(/*IRibbonControl control*/)
        {
            try
            {
                _excel.Application.Run("Dialog_Export_KML_data");
            }
            catch (Exception ex)
            {
                AddIn.ProcessUnhandledException(ex);
            }
        }

        public void ExportWizButton_OnAction(/*IRibbonControl control*/)
        {
            try
            {
                _excel.Application.Run("Dialog_Export_Wizard");
            }
            catch (Exception ex)
            {
                AddIn.ProcessUnhandledException(ex);
            }
        }

        public void RecalcButton_OnAction(/*IRibbonControl control*/)
        {
            try
            {
                _excel.Application.Run("Command_Recalculate_Transforms");
            }
            catch (Exception ex)
            {
                AddIn.ProcessUnhandledException(ex);
            }
        }

        public string LogOrderMenu_GetImage(/*IRibbonControl control*/)
        {
            try
            {
                switch (LogDisplay.DisplayOrder)
                {
                    case DisplayOrder.NewestLast:
                        return "EndOfDocument";

                    case DisplayOrder.NewestFirst:
                        return "StartOfDocument";

                    default:
                        throw new NotImplementedException(LogDisplay.DisplayOrder.ToString());
                }
            }
            catch (Exception ex)
            {
                AddIn.ProcessUnhandledException(ex);
                return null;
            }
        }

        public void LogViewDisplayButton_OnAction(/*IRibbonControl control*/)
        {
            try
            {
                LogDisplay.Show();
            }
            catch (Exception ex)
            {
                AddIn.ProcessUnhandledException(ex);
            }
        }

        public void LogClearDisplayButton_OnAction(/*IRibbonControl control*/)
        {
            try
            {
                LogDisplay.Clear();
            }
            catch (Exception ex)
            {
                AddIn.ProcessUnhandledException(ex);
            }
        }

        public void LogErrorButton_OnAction(/*IRibbonControl control*/)
        {
            try
            {
                _log.Error("[TOP] This is an **Error** message");
            }
            catch (Exception ex)
            {
                AddIn.ProcessUnhandledException(ex);
            }
        }

        public void LogDebugButton_OnAction(/*IRibbonControl control*/)
        {
            try
            {
                _log.Debug("[TOP] This is a **Debug** message");
            }
            catch (Exception ex)
            {
                AddIn.ProcessUnhandledException(ex);
            }
        }

        public void LogVerboseButton_OnAction(/*IRibbonControl control*/)
        {
            try
            {
                _log.Verbose("[TOP] This is a **Verbose** message");
            }
            catch (Exception ex)
            {
                AddIn.ProcessUnhandledException(ex);
            }
        }

        public bool LogNewestLastButton_GetPressed(/*IRibbonControl control*/)
        {
            try
            {
                return LogDisplay.DisplayOrder == DisplayOrder.NewestLast;
            }
            catch (Exception ex)
            {
                AddIn.ProcessUnhandledException(ex);
            }

            return false;
        }

        public void LogNewestLastButton_OnAction(/*IRibbonControl control, bool pressed*/)
        {
            try
            {
                LogDisplay.DisplayOrder = DisplayOrder.NewestLast;
                _thisRibbon.InvalidateControl("DisplayOrderMenu");
                _thisRibbon.InvalidateControl("LogNewestLastButton");
                _thisRibbon.InvalidateControl("LogNewestFirstButton");
            }
            catch (Exception ex)
            {
                AddIn.ProcessUnhandledException(ex);
            }
        }

        public bool LogNewestFirstButton_GetPressed(/*IRibbonControl control*/)
        {
            try
            {
                return LogDisplay.DisplayOrder == DisplayOrder.NewestFirst;
            }
            catch (Exception ex)
            {
                AddIn.ProcessUnhandledException(ex);
            }

            return false;
        }

        public void LogNewestFirstButton_OnAction(/*IRibbonControl control, bool pressed*/)
        {
            try
            {
                LogDisplay.DisplayOrder = DisplayOrder.NewestFirst;
                _thisRibbon.InvalidateControl("DisplayOrderMenu");
                _thisRibbon.InvalidateControl("LogNewestFirstButton");
                _thisRibbon.InvalidateControl("LogNewestLastButton");
            }
            catch (Exception ex)
            {
                AddIn.ProcessUnhandledException(ex);
            }
        }

        public void OnHelpButtonPressed(/*IRibbonControl control*/)
        {
            _excel.Application.Run("Command_Show_HelpFile");
        }

        public void OnVersionButtonPressed(/*IRibbonControl control*/)
        {
            _excel.Application.Run("Dialog_TopoLib_Version");
        }

        public void OnAboutButtonPressed(/*IRibbonControl control*/)
        {
            _excel.Application.Run("Dialog_About_TopoLib");
        }
        
    }

}


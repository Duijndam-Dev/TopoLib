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
                        <group id='SettingGroup'        label='TopoLib Settings'>
                            <button id='Proj_LibButton' label='Resource Settings'    imageMso='SiteColumnActionsColumnSettings' size='large' onAction='Proj_LibButton_OnAction' />
                            <button id='OptionsButton'  label='Transform Settings'   imageMso='ColumnActionsColumnSettings'     size='large' onAction='OptionsButton_OnAction' />
                            <button id='CacheButton'    label='Cache Settings'       imageMso='ColumnListSetting'               size='large' onAction='CacheButton_OnAction' />
                            <button id='LogLevelButton' label='Logging Settings'     imageMso='ComAddInsDialog'                 size='large' onAction='LogLevelButton_OnAction' />
                        </group>
                        <group id='RecalcGroup'         label='Refresh'>
                            <button id='RecalcButton'   label='TopoLib Transforms'    imageMso='RefreshWebView'                  size='large' onAction='RecalcButton_OnAction' />
                        </group>
                        <group id='LoggingGroup'        label='Test Logging Messages'>
                            <button id='ErrorButton'    label='Log Error'            imageMso='OutlineViewClose'   onAction='ErrorButton_OnAction' />
                            <button id='DebugButton'    label='Log Debug'            imageMso='MoreControlsDialog' onAction='DebugButton_OnAction' />
                            <button id='VerboseButton'  label='Log Verbose'          imageMso='Callout'            onAction='VerboseButton_OnAction' />
                        </group>
                        <group id='LogDisplayGroup' label='Log Handling'>
                            <button id='ViewLogDisplayButton' label='View Log'       imageMso='FileDocumentInspect' size='large' onAction='ViewLogDisplayButton_OnAction' />
                            <button id='ClearLogButton' label='Clear Log'            imageMso='Clear'               size='large' onAction='ClearLogDisplayButton_OnAction' />
                            <separator id='DisplayOrderSeparator' />
                            <menu id='DisplayOrderMenu' label='Display Order' getImage='DisplayOrderMenu_GetImage'  size='large'>
                            <toggleButton id='LogNewestLastButton' label='Newest Last' imageMso='EndOfDocument' getPressed='LogNewestLastButton_GetPressed' onAction='LogNewestLastButton_OnAction' />
                            <toggleButton id='LogNewestFirstButton' label='Newest First' imageMso='StartOfDocument' getPressed='LogNewestFirstButton_GetPressed' onAction='LogNewestFirstButton_OnAction' />
                            </menu>
                        </group>
                        <group id='GLgroup2' label='About'>
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

        public void ErrorButton_OnAction(IRibbonControl control)
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

        public void DebugButton_OnAction(IRibbonControl control)
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

        public void VerboseButton_OnAction(IRibbonControl control)
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

        public void ViewLogDisplayButton_OnAction(IRibbonControl control)
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

        public void ClearLogDisplayButton_OnAction(IRibbonControl control)
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

        public void OptionsButton_OnAction(IRibbonControl control)
        {
            try
            {
                _excel.Application.Run("TransformSettings_Dialog");
            }
            catch (Exception ex)
            {
                AddIn.ProcessUnhandledException(ex);
            }
        }

        public void CacheButton_OnAction(IRibbonControl control)
        {
            try
            {
                _excel.Application.Run("CacheSettings_Dialog");
            }
            catch (Exception ex)
            {
                AddIn.ProcessUnhandledException(ex);
            }
        }

        public void LogLevelButton_OnAction(IRibbonControl control)
        {
            try
            {
                _excel.Application.Run("LogSettings_Dialog");
            }
            catch (Exception ex)
            {
                AddIn.ProcessUnhandledException(ex);
            }
        }

        public void RecalcButton_OnAction(IRibbonControl control)
        {
            try
            {
                _excel.Application.Run("Recalculate_TopoLib_Transforms");
            }
            catch (Exception ex)
            {
                AddIn.ProcessUnhandledException(ex);
            }
        }

        public string DisplayOrderMenu_GetImage(IRibbonControl control)
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

        public bool LogNewestLastButton_GetPressed(IRibbonControl control)
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

        public void LogNewestLastButton_OnAction(IRibbonControl control, bool pressed)
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

        public bool LogNewestFirstButton_GetPressed(IRibbonControl control)
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

        public void LogNewestFirstButton_OnAction(IRibbonControl control, bool pressed)
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

        public void Proj_LibButton_OnAction(IRibbonControl control)
        {
            try
            {
                _excel.Application.Run("ResourceSettings_Dialog");
            }
            catch (Exception ex)
            {
                AddIn.ProcessUnhandledException(ex);
            }
        }

        public void OnHelpButtonPressed(IRibbonControl control)
        {
            _excel.Application.Run("Show_HelpFile");
        }

        public void OnVersionButtonPressed(IRibbonControl control)
        {
            _excel.Application.Run("Version_Info");
        }

        public void OnAboutButtonPressed(IRibbonControl control)
        {
            _excel.Application.Run("About_TopoLib");
        }
        
    }

}


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
using ExcelDna.Integration;


namespace TopoLib
{
    // Here is the helper class for Optional Parameters and Default Values
    // I've set ExplicitExports="true" (was "false") in TopoLib-AddIn.dna to avoid non-relevant static functions from being exported to the compiled help file.
    // To prevent static classes from other files to become visible in the Compiled Help File "TopoLib-AddIn.chm!1234" you need to use the /X flag when running ExcelDnaDoc.exe
    // In the Post-build event command line put: "D:\Source\VS19\TopoLib\packages\ExcelDnaDoc.1.1.0-beta2\tools\ExcelDnaDoc.exe" "$(TargetDir)TopoLib-AddIn.dna" /X
    // The /X flag excludes hidden functions (if provided) from being documented 

#pragma warning disable IDE0038 // Use pattern matching

    static class Optional
    {
        internal static string Check(object arg, string defaultValue)
        {
            if (arg is null || arg is ExcelMissing || arg is ExcelEmpty)
                return defaultValue;
            else if (arg is string)
                return (string)arg;
            else
                return arg.ToString();  // Or whatever you want to do here....

            // Perhaps check for other types and do whatever you think is right ....
            //else if (arg is double)
            //    return "Double: " + (double)arg;
            //else if (arg is bool)
            //    return "Boolean: " + (bool)arg;
            //else if (arg is ExcelError)
            //    return "ExcelError: " + arg.ToString();
            //else if (arg is object[,](,))
            //    // The object array returned here may contain a mixture of types,
            //    // reflecting the different cell contents.
            //    return string.Format("Array[{0},{1}]({0},{1})",
            //      ((object[,](,)(,))arg).GetLength(0), ((object[,](,)(,))arg).GetLength(1));
            //else if (arg is ExcelEmpty)
            //    return "<<Empty>>"; // Would have been null
            //else if (arg is ExcelReference)
            //  // Calling xlfRefText here requires IsMacroType=true for this function.
            //                return "Reference: " +
            //                     XlCall.Excel(XlCall.xlfReftext, arg, true);
            //            else
            //                return "!? Unheard Of ?!";
        }

        internal static double Check(object arg, double defaultValue)
        {
            if (arg is null || arg is ExcelMissing || arg is ExcelEmpty)
                return defaultValue;
            else if (arg is double)
                return (double)arg;
            else
                throw new ArgumentException();  // Will return #VALUE to Excel
        }

        internal static bool Check(object arg, bool defaultValue)
        {
            if (arg is null || arg is ExcelMissing || arg is ExcelEmpty)
                return defaultValue;
            else if (arg is double)
                return (bool)arg;
            else if (arg is bool)
                return (bool)arg;
            else
                throw new ArgumentException();  // Will return #VALUE to Excel
        }

        // This one is more tricky - we have to do the double->Date conversions ourselves
        internal static DateTime Check(object arg, DateTime defaultValue)
        {
            if (arg is double)
                return DateTime.FromOADate((double)arg);    // Here is the conversion
            else if (arg is string)
                return DateTime.Parse((string)arg);
            else if (arg is ExcelMissing)
                return defaultValue;

            else
                throw new ArgumentException();  // Or defaultValue or whatever
        }

        internal static bool IsNul(object[,] arg)
        {
            Type valueType = arg.GetType();
            if (valueType.IsArray)
            {
                if (arg[0, 0] is null || arg[0, 0] is ExcelMissing || arg[0, 0] is ExcelEmpty || arg[0, 0] is ExcelError)
                    return true;
                else if ((arg[0, 0] is double) && ((double)arg[0, 0] == 0.0))
                    return true;
                else if ((arg[0, 0] is string))
                {
                    string sArg = (string)arg[0, 0];
                    sArg = sArg.TrimStart('\'');
                    if (string.IsNullOrEmpty(sArg) || sArg == "0")
                        return true;
                    else
                        return false;
                }
                else
                    return false;
            }
            else
            {
                // deal with a single object instead of array; not yet implemented
                throw new NotImplementedException("Not yet implemented");
            }
        } // IsNull

        internal static object CheckNan(double value)
        {
            object result;
            if (Double.IsNaN(value) || Double.IsInfinity(value))
            {
                result = ExcelError.ExcelErrorNA;
            }
            else
            {
                result = value;
            }
            return result;
        }

    } // class Optional
#pragma warning restore IDE0038 // Use pattern matching

} // namespace TopoLib


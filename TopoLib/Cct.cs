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

#pragma warning disable IDE0075 // Conditional expression can be simplified

// On solving a missing reference to the next package:
// For me adding the PackageReference for MSTest.TestFramework did the trick. I didn't need to reference the TestAdapter.
// see https://stackoverflow.com/questions/13602508/where-to-find-microsoft-visualstudio-testtools-unittesting-missing-dll
// using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;

// I made a backup of my project by renaming it to TopoLibOld 
// Then I rebuilt TopoLib from scratch starting from ExcelDna v1.1.0 in view of Virusscanner false positives with v1.5.0
// Next I added all source files, etc.
// But then Git was stuffed, because my project history under TopoLib did not jive with the master branch on the server.
// After some googling, I used the following command from the command line: 
//
// git push --set-upstream origin master -f
//
// This solved the problem, and my latest changes are uploaded to GitHub....

// The reference to SharpProj is effectively a reference to:
// D:\Source\VS19\TopoLib\packages\SharpProj.Core.8.1001.60\lib\net45\SharpProj.dll
// 
// This file is equal to the SharpProj.dll under:
// D:\Source\VS19\TopoLib\packages\SharpProj.Core.8.1001.60\runtimes\win-x86\lib\net45\SharpProj.dll
// Both files are 5.536 KB large
//
// A 64-bit dll is also included in the package:
// D:\Source\VS19\TopoLib\packages\SharpProj.Core.8.1001.60\runtimes\win-x64\lib\net45\SharpProj.dll
// This file is 6.681 KB large

// Note: for easy number generation for compiled help items, use TextPad 8. Using Search and Replace:
// Search for : "TopoLib-AddIn.chm!...."
// Replace by : "TopoLib-AddIn.chm!\i{1200}"
// This will generate a counter starting at 1200 and incrementing by 1.
// See also https://community.notepad-plus-plus.org/topic/19414/replace-text-with-incremented-counter

#pragma warning disable IDE0019 // Use pattern matching

namespace TopoLib
{
    public static class Cct
    {
        internal static int GetNrCoordinateColumns(int nMode, int nDefault)
        {
            int nOut;

            switch (nMode & 7) // only use the three lowest bits 1 + 2 + 4
            {
                default:
                case 0:
                    nOut = nDefault;
                    break;
                case 1:
                    nOut = 4;
                    break;
                case 2:
                    nOut = 3;
                    break;
                case 3:
                    nOut = 2;
                    break;
                case 4:
                case 5:
                case 6:
                case 7:
                    nOut = 1;
                    break;
            } // GetNrCoordinateColumns

            return nOut;
        }

		internal static CoordinateTransformOptions GetCoordinateTransformOptions(int nMode, double Accuracy, double westLongitude, double southLatitude, double eastLongitude, double northLatitude, ref bool bAllowDeprecatedCRS,  ref bool bUseNetwork)
        {
			var options = new CoordinateTransformOptions();

            if (nMode > 7)
            {
                // get stuff from nMode parameter, and set values for bAllowDeprecatedCRS & bUseNetwork.
                // nMode = 4096 will arrive here but won't set any flags; for debugging only
                options.Accuracy              = Accuracy;
                options.Area                  = westLongitude > -1000 ? new CoordinateArea(westLongitude, southLatitude, eastLongitude, northLatitude) : null;
                options.NoBallparkConversions = (nMode &    8) != 0 ? true : false;
                options.NoDiscardIfMissing    = (nMode &   16) != 0 ? true : false;
                options.UsePrimaryGridNames   = (nMode &   32) != 0 ? true : false;
                options.UseSuperseded         = (nMode &   64) != 0 ? true : false;
                    bAllowDeprecatedCRS       = (nMode &  128) != 0 ? true : false;
                options.StrictContains        = (nMode &  256) != 0 ? true : false;
                options.IntermediateCrsUsage  = (nMode &  512) != 0 ? IntermediateCrsUsage.Always : IntermediateCrsUsage.Auto;
                options.IntermediateCrsUsage  = (nMode & 1024) != 0 ? IntermediateCrsUsage.Never  : IntermediateCrsUsage.Auto;
                    bUseNetwork               = (nMode & 2048) != 0 ? true : false;

                // deal with 'Always' and 'Never' both being set. Go back to 'Auto' !
                if (((nMode & 512) != 0) && (nMode & 1024) != 0) options.IntermediateCrsUsage  = IntermediateCrsUsage.Auto;
            }
            else
            {
                // get options from static variables
                options     = CctOptions.TransformOptions;
                bUseNetwork = CctOptions.UseNetworkConnection;
                bAllowDeprecatedCRS = CctOptions.AllowDeprecatedCRS;
            }
			return options;

        } // GetCoordinateTransformOptions

        internal static CoordinateTransform CreateCoordinateTransform(in object[,] oTransform, ProjContext pc)
        {
            int nTransformRows = oTransform.GetLength(0);
            int nTransformCols = oTransform.GetLength(1);

            // max two adjacent CRS cells on the same row
            if (nTransformRows != 1 || nTransformCols > 2 ) 
                throw new ArgumentException("Incorrrect dimensions in Excel for coordinate transformation");

            int nTransform;
            string sTransform;

            // We have only one cell; it can be a WKT string, a JSON string, a PROJ string or a textual description
            if (nTransformCols == 1)
            {
                // we have one cell describing the crs.
                if (oTransform[0, 0] is double)
                {
                    // First cast to double, then to int, to deal with Excel datatypes
                    nTransform = (int)(double)oTransform[0, 0];

                    // we have an EPSG number from a single input parameter:
                    return CoordinateTransform.CreateFromEpsg(nTransform, pc);
                }
                else if (oTransform[0, 0] is string)
                {
                    // cast to string, to deal with Excel datatypes
                    sTransform = (string)oTransform[0, 0];

                    bool success = int.TryParse(sTransform, out nTransform);
                    if (success)
                    {
                        // we have an EPSG number from a single input parameter:
                        return CoordinateTransform.CreateFromEpsg(nTransform, pc);
                    }
                    else
                    {
                        // we have a string of some sorts and a single input parameter:
                        if ((sTransform.IndexOf("PROJCRS") > -1) || (sTransform.IndexOf("GEOGCRS") > -1) || (sTransform.IndexOf("SPHEROID") > -1))
                        {
                            // it must be WKT (well, we hope)

                            // Note the cast used below is not required for CoordinateReferenceSystem where this function has been implemented as part of the inherited class
                            return (CoordinateTransform)CoordinateTransform.CreateFromWellKnownText(sTransform, pc);

                            // CreateFromWellKnownText() is translated into CreateFromWellKnownText(from, wars, ctx); where array<String^>^ wars = nullptr;
                            // It may throw an ArgumentNullException
                            // It may throw a ProjException
                        }
                        else
                        {
                            // it might be anything
                            return CoordinateTransform.Create(sTransform, pc);

                            // Create() is translated into proj_create(ctx, fromStr); 
                            // It may throw an ArgumentNullException("from");
                            // It may throw a ctx->ConstructException();
                            // It may throw a ProjException
                        }
                    }
                }
                else
                    throw new ArgumentNullException("Incorrect coordinate transform format");
            }
            else
            {
                // we have two adjacent CRS cells; first an Authortity string of some sorts and a second input parameter (number):

                sTransform= (string)oTransform[0,0];   // the authority string

                // try to get the transform number; if not succesful throw an exception

                if (oTransform[0, 1] is double)
                {
                    // First cast to double, then to int, to deal with Excel datatypes
                    nTransform = (int)(double)oTransform[0, 1];

                }
                else if (oTransform[0, 1] is string)
                {
                    // cast to string, to deal with Excel datatypes
                    string sTmp = (string)oTransform[0, 1];

                    bool success = int.TryParse(sTmp, out nTransform);
                    if (!success) 
                        throw new ArgumentNullException("Incorrect coordinate transform format");
                }
                else
                    throw new ArgumentNullException("Incorrect coordinate transform format");

                return CoordinateTransform.CreateFromDatabase(sTransform, nTransform, pc);

                // CreateFromDatabase() is translated into proj_create_from_database 
                // It may throw a ArgumentNullException
                // It may throw a pc->ConstructException
            }

            // Oops, something went wrong if we get here...
            throw new ArgumentNullException("Incorrect coordinate transform format");

        }

        internal static CoordinateTransform CreateCoordinateTransform(CoordinateReferenceSystem crsSource, CoordinateReferenceSystem crsTarget, CoordinateTransformOptions options, ProjContext pc, bool bAllowDeprecatedCRS, bool bUseNetwork)
        {
            bool bHasDeprecatedCRS = crsSource.IsDeprecated || crsTarget.IsDeprecated; 

            if (bHasDeprecatedCRS && !bAllowDeprecatedCRS)
                throw new System.InvalidOperationException ("Using deprecated CRS when not allowed");

            if (bUseNetwork)                       
                pc.EnableNetworkConnections = true;

            var transform = CoordinateTransform.Create(crsSource.WithAxisNormalized(), crsTarget.WithAxisNormalized(), options, pc);
        
            if (transform == null)
                throw new System.InvalidOperationException ("No transformation available");

            return transform;

        } // CreateCoordinateTransform

         [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.Accuracy",
            Description = "Get the accuracy of a transform", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1001",

            Returns = "Accuracy of a transform [m]",
            Summary = "Returns accuracy of a  coordinate transform",
            Example = "xxx")]
        public static object Accuracy(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1000' when not used (-1000)", Name = "Accuracy")] object oAccuracy,
            [ExcelArgument("WestLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "WestLongitude")] object oWestLongitude,
            [ExcelArgument("SouthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "SouthLatitude")] object oSouthLatitude,
            [ExcelArgument("EastLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "EastLongitude")] object oEastLongitude,
            [ExcelArgument("NorthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "NorthLatitude")] object oNorthLatitude)

        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            // Check general input data
            int nMode       = (int)Optional.Check(oMode, 0.0);
            bool bUsingTransform = Optional.IsNul(TargetCrs);
            double Accuracy      = Optional.Check(oAccuracy,      -1000.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 4096)
                return ExcelError.ExcelErrorValue;

            // Deal with optional parameters
            bool bUseNetwork = false;
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS, ref bUseNetwork);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = new ProjContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.GetCrs(SourceCrs, pjContext);
                    crsTarget = Crs.GetCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS, bUseNetwork);
                }

                // start core of function
                double? accuracy = transform.Accuracy;

                if (accuracy.HasValue)
                    return accuracy;
                else
                    return "Unknown";
                // end core of function

            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // Accuracy

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.GridUsages.Count",
            Description = "Nr of grids used in a transform", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1002",

            Returns = "Nr of grids used in a transform",
            Summary = "Function returns nr of grids used in a transform",
            Example = "xxx")]
        public static object GridUsages_Count(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1000' when not used (-1000)", Name = "Accuracy")] object oAccuracy,
            [ExcelArgument("WestLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "WestLongitude")] object oWestLongitude,
            [ExcelArgument("SouthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "SouthLatitude")] object oSouthLatitude,
            [ExcelArgument("EastLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "EastLongitude")] object oEastLongitude,
            [ExcelArgument("NorthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "NorthLatitude")] object oNorthLatitude)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            // Check general input data
            int nMode       = (int)Optional.Check(oMode, 0.0);
            bool bUsingTransform = Optional.IsNul(TargetCrs);
            double Accuracy      = Optional.Check(oAccuracy,      -1000.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 4096)
                return ExcelError.ExcelErrorValue;

            // Deal with optional parameters
            bool bUseNetwork = false;
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS, ref bUseNetwork);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = new ProjContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.GetCrs(SourceCrs, pjContext);
                    crsTarget = Crs.GetCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS, bUseNetwork);
                }

                // start core of function
                return (double) transform.GridUsages.Count;
                // end core of function

            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // GridUsages_Count

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.HasBallParkTransformation",
            Category = "CCT - Coordinate Conversion and Transformation",
            Description = "Confirms whether the transform has a ballpark transformation",
            HelpTopic = "TopoLib-AddIn.chm!1003",

            Returns = "TRUE when the transform has a ballpark transformation; FALSE when not",
            Summary = "Function that confirms that the transform has a ballpark transformation",
            Example = "xxx"
         )]
        public static object HasBallParkTransformation(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1000' when not used (-1000)", Name = "Accuracy")] object oAccuracy,
            [ExcelArgument("WestLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "WestLongitude")] object oWestLongitude,
            [ExcelArgument("SouthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "SouthLatitude")] object oSouthLatitude,
            [ExcelArgument("EastLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "EastLongitude")] object oEastLongitude,
            [ExcelArgument("NorthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "NorthLatitude")] object oNorthLatitude)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            // Check general input data
            int nMode       = (int)Optional.Check(oMode, 0.0);
            bool bUsingTransform = Optional.IsNul(TargetCrs);
            double Accuracy      = Optional.Check(oAccuracy,      -1000.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 4096)
                return ExcelError.ExcelErrorValue;

            // Deal with optional parameters
            bool bUseNetwork = false;
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS, ref bUseNetwork);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = new ProjContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.GetCrs(SourceCrs, pjContext);
                    crsTarget = Crs.GetCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS, bUseNetwork);
                }

                // start core of function
                return transform.HasBallParkTransformation;
                // end core of function

            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // HasBallParkTransformation


        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.HasInverse",
            Category = "CCT - Coordinate Conversion and Transformation",
            Description = "Confirms whether the transform can be done in the reversed direction",
            HelpTopic = "TopoLib-AddIn.chm!1003",

            Returns = "TRUE when the transform can be done in the reversed direction; FALSE when not",
            Summary = "Function that confirms that the transform can be done in the reversed direction",
            Example = "xxx"
         )]
        public static object HasInverse(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1000' when not used (-1000)", Name = "Accuracy")] object oAccuracy,
            [ExcelArgument("WestLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "WestLongitude")] object oWestLongitude,
            [ExcelArgument("SouthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "SouthLatitude")] object oSouthLatitude,
            [ExcelArgument("EastLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "EastLongitude")] object oEastLongitude,
            [ExcelArgument("NorthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "NorthLatitude")] object oNorthLatitude)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            // Check general input data
            int nMode       = (int)Optional.Check(oMode, 0.0);
            bool bUsingTransform = Optional.IsNul(TargetCrs);
            double Accuracy      = Optional.Check(oAccuracy,      -1000.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 4096)
                return ExcelError.ExcelErrorValue;

            // Deal with optional parameters
            bool bUseNetwork = false;
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS, ref bUseNetwork);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = new ProjContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.GetCrs(SourceCrs, pjContext);
                    crsTarget = Crs.GetCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS, bUseNetwork);
                }

                // start core of function
                return transform.HasInverse;
                // end core of function

            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // HasInverse

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.IsAvailable",
            Category = "CCT - Coordinate Conversion and Transformation",
            Description = "Confirms whether the transform is available",
            HelpTopic = "TopoLib-AddIn.chm!1003",

            Returns = "TRUE when the transform is available; FALSE when not",
            Summary = "Function that confirms that the transform is available",
            Example = "xxx"
         )]
        public static object IsAvailable(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1000' when not used (-1000)", Name = "Accuracy")] object oAccuracy,
            [ExcelArgument("WestLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "WestLongitude")] object oWestLongitude,
            [ExcelArgument("SouthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "SouthLatitude")] object oSouthLatitude,
            [ExcelArgument("EastLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "EastLongitude")] object oEastLongitude,
            [ExcelArgument("NorthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "NorthLatitude")] object oNorthLatitude)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            // Check general input data
            int nMode       = (int)Optional.Check(oMode, 0.0);
            bool bUsingTransform = Optional.IsNul(TargetCrs);
            double Accuracy      = Optional.Check(oAccuracy,      -1000.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 4096)
                return ExcelError.ExcelErrorValue;

            // Deal with optional parameters
            bool bUseNetwork = false;
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS, ref bUseNetwork);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = new ProjContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.GetCrs(SourceCrs, pjContext);
                    crsTarget = Crs.GetCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS, bUseNetwork);
                }

                // start core of function
                return transform.IsAvailable;
                // end core of function

            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // IsAvailable

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.MethodName",
            Description = "Get the method name of the coordinate transform", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1004",

            Returns = "name of the coordinate transform",
            Summary = "Returns the method name a coordinate transform",
            Example = "xxx")]
        public static object MethodName(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1000' when not used (-1000)", Name = "Accuracy")] object oAccuracy,
            [ExcelArgument("WestLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "WestLongitude")] object oWestLongitude,
            [ExcelArgument("SouthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "SouthLatitude")] object oSouthLatitude,
            [ExcelArgument("EastLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "EastLongitude")] object oEastLongitude,
            [ExcelArgument("NorthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "NorthLatitude")] object oNorthLatitude)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            // Check general input data
            int nMode       = (int)Optional.Check(oMode, 0.0);
            bool bUsingTransform = Optional.IsNul(TargetCrs);
            double Accuracy      = Optional.Check(oAccuracy,      -1000.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 4096)
                return ExcelError.ExcelErrorValue;

            // Deal with optional parameters
            bool bUseNetwork = false;
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS, ref bUseNetwork);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = new ProjContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.GetCrs(SourceCrs, pjContext);
                    crsTarget = Crs.GetCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS, bUseNetwork);
                }

                // start core of function
                return (transform.MethodName == null) ? "Unknown" : transform.MethodName;
                // end core of function

            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // MethodName


        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.Name",
            Description = "Get the name of the coordinate transform", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1004",

            Returns = "name of the coordinate transform",
            Summary =
            "Returns the name a coordinate transform",
            Example = "xxx")]
        public static object Name(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1000' when not used (-1000)", Name = "Accuracy")] object oAccuracy,
            [ExcelArgument("WestLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "WestLongitude")] object oWestLongitude,
            [ExcelArgument("SouthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "SouthLatitude")] object oSouthLatitude,
            [ExcelArgument("EastLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "EastLongitude")] object oEastLongitude,
            [ExcelArgument("NorthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "NorthLatitude")] object oNorthLatitude)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            // Check general input data
            int nMode       = (int)Optional.Check(oMode, 0.0);
            bool bUsingTransform = Optional.IsNul(TargetCrs);
            double Accuracy      = Optional.Check(oAccuracy,      -1000.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 4096)
                return ExcelError.ExcelErrorValue;

            // Deal with optional parameters
            bool bUseNetwork = false;
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS, ref bUseNetwork);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = new ProjContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.GetCrs(SourceCrs, pjContext);
                    crsTarget = Crs.GetCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS, bUseNetwork);
                }

                // start core of function
                return transform.Name;
                // end core of function

            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // Name

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.RoundTrip",
            Description = "Get the error of a roundtrip of N forward/backward transforms", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1005",

            Returns = "error incurred in the roundtrip(s)",
            Summary =
            "Returns error incurred in N forward roundtrip(s) in a coordinate transform",
            Remarks = "For the test point, it is recommended to select the centerpoint of the usage area of the Source CRS." +
            "<p>If no test point is given (0, 0, 0) will be used instead</p>",
            Example = "xxx")]
        public static object RoundTrip(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("test point with adjacent [x, y] coordinates", Name = "point(x, y)")] object[,] TestCoord,
            [ExcelArgument("N - nr of roundtrips to make", Name = "nr roundtrips")] object oRoundTrips,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1000' when not used (-1000)", Name = "Accuracy")] object oAccuracy,
            [ExcelArgument("WestLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "WestLongitude")] object oWestLongitude,
            [ExcelArgument("SouthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "SouthLatitude")] object oSouthLatitude,
            [ExcelArgument("EastLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "EastLongitude")] object oEastLongitude,
            [ExcelArgument("NorthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "NorthLatitude")] object oNorthLatitude)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            // Check general input data
            int nMode       = (int)Optional.Check(oMode, 0.0);
            bool bUsingTransform = Optional.IsNul(TargetCrs);
            double Accuracy      = Optional.Check(oAccuracy,      -1000.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 4096)
                return ExcelError.ExcelErrorValue;

            // Check specific input data
            int nTrips = (int)Optional.Check(oRoundTrips, 1.0);
            // max three adjacent [x, y, z] cells on the same row
            if (TestCoord.GetLength(0) > 1 || TestCoord.GetLength(1) > 3 ) return ExcelError.ExcelErrorValue;

            double x = TestCoord.GetLength(1) > 0 ? (double)TestCoord[0, 0] : 0;
            double y = TestCoord.GetLength(1) > 1 ? (double)TestCoord[0, 1] : 0;
            double z = TestCoord.GetLength(1) > 2 ? (double)TestCoord[0, 2] : 0;
            PPoint pt = new PPoint(x, y, z);

            // Deal with optional parameters
            bool bUseNetwork = false;
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS, ref bUseNetwork);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = new ProjContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.GetCrs(SourceCrs, pjContext);
                    crsTarget = Crs.GetCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS, bUseNetwork);
                }

                // start core of function
                double error = transform.RoundTrip(true, nTrips, pt);

                if (Double.IsInfinity(error))
                    throw new System.InvalidOperationException("Infinite roundtrip error");

                return error;
                // end core of function

            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // RoundTrip

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.Steps.Count",
            Description = "Get the number of steps incorporated in a coordinate transform", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1006",

            Returns = "number of steps incorporated in a coordinate transform",
            Summary =
            "Returns the number of steps incorporated in a coordinate transform",
            Example = "xxx")]
        public static object Steps_Count(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1000' when not used (-1000)", Name = "Accuracy")] object oAccuracy,
            [ExcelArgument("WestLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "WestLongitude")] object oWestLongitude,
            [ExcelArgument("SouthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "SouthLatitude")] object oSouthLatitude,
            [ExcelArgument("EastLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "EastLongitude")] object oEastLongitude,
            [ExcelArgument("NorthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "NorthLatitude")] object oNorthLatitude)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            // Check general input data
            int nMode       = (int)Optional.Check(oMode, 0.0);
            bool bUsingTransform = Optional.IsNul(TargetCrs);
            double Accuracy      = Optional.Check(oAccuracy,      -1000.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 4096)
                return ExcelError.ExcelErrorValue;

            // Deal with optional parameters
            bool bUseNetwork = false;
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS, ref bUseNetwork);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = new ProjContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.GetCrs(SourceCrs, pjContext);
                    crsTarget = Crs.GetCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS, bUseNetwork);
                }

                // start core of function
                CoordinateTransformList steps = transform as CoordinateTransformList;
                return (steps is null) ? 0 : steps.Count;
                // end core of function

            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // Steps_Count

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.Steps.MethodName",
            Description = "Get the method name of step N in a coordinate transform", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1007",

            Returns = "method name of step N in a coordinate transform",
            Summary =
            "Returns the method name of step N in a coordinate transform",
            Example = "xxx")]
        public static object Steps_MethodName(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Zero based index of Identifier list (0) ", Name = "Index")] object index,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1000' when not used (-1000)", Name = "Accuracy")] object oAccuracy,
            [ExcelArgument("WestLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "WestLongitude")] object oWestLongitude,
            [ExcelArgument("SouthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "SouthLatitude")] object oSouthLatitude,
            [ExcelArgument("EastLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "EastLongitude")] object oEastLongitude,
            [ExcelArgument("NorthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "NorthLatitude")] object oNorthLatitude)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            // Check general input data
            int nMode       = (int)Optional.Check(oMode, 0.0);
            bool bUsingTransform = Optional.IsNul(TargetCrs);
            double Accuracy      = Optional.Check(oAccuracy,      -1000.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 4096)
                return ExcelError.ExcelErrorValue;

            // Check specific input data
            int nIndex  = (int)Optional.Check(index , 0.0);

            // Deal with optional parameters
            bool bUseNetwork = false;
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS, ref bUseNetwork);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = new ProjContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.GetCrs(SourceCrs, pjContext);
                    crsTarget = Crs.GetCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS, bUseNetwork);
                }

                // start core of function
                CoordinateTransformList steps = transform as CoordinateTransformList;
                return (steps is null) ? "Unknown" : steps[nIndex].MethodName;
                // end core of function

            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // Steps_MethodName

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.Steps.Name",
            Description = "Get the name of step N in a coordinate transform", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1008",

            Returns = "name of step N in a coordinate transform",
            Summary =
            "Returns the name of step N in a coordinate transform",
            Example = "xxx")]
        public static object Steps_Name(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Zero based index of Identifier list (0) ", Name = "Index")] object index,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1000' when not used (-1000)", Name = "Accuracy")] object oAccuracy,
            [ExcelArgument("WestLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "WestLongitude")] object oWestLongitude,
            [ExcelArgument("SouthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "SouthLatitude")] object oSouthLatitude,
            [ExcelArgument("EastLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "EastLongitude")] object oEastLongitude,
            [ExcelArgument("NorthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "NorthLatitude")] object oNorthLatitude)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            // Check general input data
            int nMode       = (int)Optional.Check(oMode, 0.0);
            bool bUsingTransform = Optional.IsNul(TargetCrs);
            double Accuracy      = Optional.Check(oAccuracy,      -1000.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 4096)
                return ExcelError.ExcelErrorValue;

            // Check specific input data
            int nIndex  = (int)Optional.Check(index , 0.0);

            // Deal with optional parameters
            bool bUseNetwork = false;
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS, ref bUseNetwork);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = new ProjContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.GetCrs(SourceCrs, pjContext);
                    crsTarget = Crs.GetCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS, bUseNetwork);
                }

                // start core of function
                CoordinateTransformList steps = transform as CoordinateTransformList;
                return (steps is null) ? "Unknown" : steps[nIndex].Name;
                // end core of function

            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // Steps_Name

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.TransformForward",
            Description = "Coordinate conversion of one or more input points", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1009",

            Returns = "the reprojected coordinate(s)",
            Summary =
            "<p> This function transforms coordinates from one Coordinate Reference system (CRS) into another CRS</p>" +
            "<p> Source and targetCrs can be provided in one out of three ways</p>" +
            "<ol>    <li><p>As a number referencing a CRS CODE from the EPSG database (much preferred)</p></li>" +
                    "<li><p>As a string using WKT, JSON or PROJ format. WKT or JSON format is preferred over the original PROJ string format</p></li>" +
                    "<li><p>As an AUTHORITY string in one cell, combined with a CRS CODE in the adjacent cell to the right</p></li>" +
            "</ol>" +
            "<p>This function is an array function. Array functions have undergone a significant upgrade with the introduction of dynamic arrays in Excel.</p>" +
            "<p>For more information on working with array formulas please consult :</p>" +
            "<ol>" +
            "<li><p>This link: <a href = \"https://support.office.com/en-us/article/Guidelines-and-examples-of-array-formulas-7d94a64e-3ff3-4686-9372-ecfd5caa57c7\" > Microsoft Office Support - Guidelines and examples of array formulas</a> for Guidelines and examples of array formulas.</p></li>" +
            "<li><p>This link: <a href = \"https://support.office.com/en-us/article/Create-an-array-formula-e43e12e0-afc6-4a12-bc7f-48361075954d\" > Microsoft Office Support - Create a array formula</a> for more information on how to create static {CSE} array formulas.</p></li>" +
            "<li><p>This link: <a href = \"https://exceljet.net/dynamic-array-formulas-in-excel\" > ExcelJet - Dynamic array formulas in Excel</a> for an introduction to dynamic array formulas.</p></li>" +
            "</ol>" +
            "<p>For more information on coordinate conversion and coordinate refence system (CRS) information, see :</p>" +
            "<ol>    <li><p>This link: <a href = \"http://spatialreference.org/\"> Spatial Reference home page</a></p></li>" +
                    "<li><p>This link: <a href = \"http://epsg.io/\" id=\"viewDesktopLink\"> EPSG IO home page with CRS description strings and EPSG numbers</a></p></li>" +
                    "<li><p>This link: <a href = \"http://proj.org/\"> Home page of the proj library</a></p></li>" +
            "</ol>",
            Remarks = "<p>Internally the transform uses <a href = \"https://proj.org/development/reference/functions.html?highlight=proj_normalize_for_visualization\"> crs normalization</a> by the proj library for a consistent approach to (x, y, z) values.</p>" +
            "<p>The axis order of a geographic CRS shall therefore be longitude, latitude [,height], and that of a projected CRS shall be easting, northing [, height]</p>" +
            "<p>When using a geographic CRS, coordinates should be presented in degrees (not radians).</p>",
            Example = "xxx")]
        public static object TransformForward(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("input points", Name = "Coordinate(s)")] double[,] Coords,
            [ExcelArgument("Output mode: < 7 and flag: > 7. (0); 0 returns nr of input columns, 1 = (x, y, z, t), 2 = (x, y, z), 3 = (x, y), 4 = (x), 5 = (y), 6 = (z), 7 = (t). Check flag values 2^n in the help file", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1000' when not used (-1000)", Name = "Accuracy")] object oAccuracy,
            [ExcelArgument("WestLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "WestLongitude")] object oWestLongitude,
            [ExcelArgument("SouthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "SouthLatitude")] object oSouthLatitude,
            [ExcelArgument("EastLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "EastLongitude")] object oEastLongitude,
            [ExcelArgument("NorthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "NorthLatitude")] object oNorthLatitude)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            // Check general input data
            int nMode       = (int)Optional.Check(oMode, 0.0);
            bool bUsingTransform = Optional.IsNul(TargetCrs);
            double Accuracy      = Optional.Check(oAccuracy,      -1000.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 4096)
                return ExcelError.ExcelErrorValue;

            // Check specific input parameters
            int nCoordRows = Coords.GetLength(0);
            int nCoordCols = Coords.GetLength(1);

            if (nCoordRows < 1 )
                return ExcelError.ExcelErrorValue;

            if (nCoordCols < 2 || nCoordCols > 4 )
                return ExcelError.ExcelErrorValue;

            int nOut = GetNrCoordinateColumns(nMode, nCoordCols);

            // Deal with optional parameters
            bool bUseNetwork = false;
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS, ref bUseNetwork);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = new ProjContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.GetCrs(SourceCrs, pjContext);
                    crsTarget = Crs.GetCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS, bUseNetwork);
                }

                // start core of function
                double[] x = new double[nCoordRows];
                double[] y = new double[nCoordRows];
                double[] z = new double[nCoordRows];
                double[] t = new double[nCoordRows];
                object[,] res = new object[nCoordRows, nOut];

                // work with nr of input columns
                switch (nCoordCols)
                {
                    default:
                    case 2:
                        // we have two columns (x, y) in the input data
                        for (int i = 0; i < nCoordRows; i++)
                        {
                            x[i] = Coords[i, 0];
                            y[i] = Coords[i, 1];
                        }

                        transform.Apply(x, y);
                        break;

                    case 3:
                        // we have three columns (x, y, z) in the input data
                        for (int i = 0; i < nCoordRows; i++)
                        {
                            x[i] = Coords[i, 0];
                            y[i] = Coords[i, 1];
                            z[i] = Coords[i, 2];
                        }

                        transform.Apply(x, y, z);
                        break;

                    case 4:
                        // we have four columns (x, y, z, t) in the input data
                        for (int i = 0; i < nCoordRows; i++)
                        {
                            x[i] = Coords[i, 0];
                            y[i] = Coords[i, 1];
                            z[i] = Coords[i, 2];
                            t[i] = Coords[i, 3];
                        }

                        transform.Apply(x, y, z, t);
                        break;
                }

                // determine what to do with output
                switch (nMode & 7) // use only the lowest three bits
                {
                    case 0:
                    case 1:
                    case 2:
                    case 3:

                        // all values to be returned
                        // check how many columns we need
                        if (nOut == 2)
                        {
                            for (int i = 0; i < nCoordRows; i++)
                            {
                                res[i, 0] = Optional.CheckNan(x[i]);
                                res[i, 1] = Optional.CheckNan(y[i]);
                            }
                        }
                        else if (nOut == 3)
                        {
                            for (int i = 0; i < nCoordRows; i++)
                            {
                                res[i, 0] = Optional.CheckNan(x[i]);
                                res[i, 1] = Optional.CheckNan(y[i]);
                                res[i, 2] = Optional.CheckNan(z[i]);
                            }
                        }
                        else if (nOut == 4)
                        {
                            for (int i = 0; i < nCoordRows; i++)
                            {
                                res[i, 0] = Optional.CheckNan(x[i]);
                                res[i, 1] = Optional.CheckNan(y[i]);
                                res[i, 2] = Optional.CheckNan(z[i]);
                                res[i, 3] = Optional.CheckNan(t[i]);
                            }
                        }
                        else
                            return ExcelError.ExcelErrorValue;
                        break;
                    case 4:
                        // from here onwards, a single output value is required
                        for (int i = 0; i < nCoordRows; i++)
                        {
                            res[i, 0] = Optional.CheckNan(x[i]);
                        }
                        break;
                    case 5:
                        for (int i = 0; i < nCoordRows; i++)
                        {
                            res[i, 0] = Optional.CheckNan(y[i]);
                        }
                        break;
                    case 6:
                        for (int i = 0; i < nCoordRows; i++)
                        {
                            res[i, 0] = Optional.CheckNan(z[i]);
                        }
                        break;
                    case 7:
                        for (int i = 0; i < nCoordRows; i++)
                        {
                            res[i, 0] = Optional.CheckNan(t[i]);
                        }
                        break;
                    default:
                        throw new System.InvalidOperationException("error in switch statement");
                }
                return res;
                // end core of function

            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // TransformForward

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.TransformInverse",
            Description = "Inverse coordinate conversion of one or more input points", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1011",

            Returns = "the reprojected coordinate(s)",
            Summary =
            "<p> This function transforms coordinates from one Coordinate Reference system (CRS) into another CRS in the reverse direction</p>" +
            "<p> Source and targetCrs can be provided in one out of three ways</p>" +
            "<ol>    <li><p>As a number referencing a CRS CODE from the EPSG database (much preferred)</p></li>" +
                    "<li><p>As a string using WKT, JSON or PROJ format. WKT or JSON format is preferred over the original PROJ string format</p></li>" +
                    "<li><p>As an AUTHORITY string in one cell, combined with a CRS CODE in the adjacent cell to the right</p></li>" +
            "</ol>" +
            "<p>This function is an array function. Array functions have undergone a significant upgrade with the introduction of dynamic arrays in Excel.</p>" +
            "<p>For more information on working with array formulas please consult :</p>" +
            "<ol>" +
            "<li><p>This link: <a href = \"https://support.office.com/en-us/article/Guidelines-and-examples-of-array-formulas-7d94a64e-3ff3-4686-9372-ecfd5caa57c7\" > Microsoft Office Support - Guidelines and examples of array formulas</a> for Guidelines and examples of array formulas.</p></li>" +
            "<li><p>This link: <a href = \"https://support.office.com/en-us/article/Create-an-array-formula-e43e12e0-afc6-4a12-bc7f-48361075954d\" > Microsoft Office Support - Create a array formula</a> for more information on how to create static {CSE} array formulas.</p></li>" +
            "<li><p>This link: <a href = \"https://exceljet.net/dynamic-array-formulas-in-excel\" > ExcelJet - Dynamic array formulas in Excel</a> for an introduction to dynamic array formulas.</p></li>" +
            "</ol>" +
            "<p>For more information on coordinate conversion and coordinate refence system (CRS) information, see :</p>" +
            "<ol>    <li><p>This link: <a href = \"http://spatialreference.org/\"> Spatial Reference home page</a></p></li>" +
                    "<li><p>This link: <a href = \"http://epsg.io/\" id=\"viewDesktopLink\"> EPSG IO home page with CRS description strings and EPSG numbers</a></p></li>" +
                    "<li><p>This link: <a href = \"http://proj.org/\"> Home page of the proj library</a></p></li>" +
            "</ol>",
            Remarks = "<p>Internally the transform uses <a href = \"https://proj.org/development/reference/functions.html?highlight=proj_normalize_for_visualization\"> crs normalization</a> by the proj library for a consistent approach to (x, y, z) values.</p>" +
            "<p>The axis order of a geographic CRS shall therefore be longitude, latitude [,height], and that of a projected CRS shall be easting, northing [, height]</p>" +
            "<p>When using a geographic CRS, coordinates should be presented in degrees (not radians).</p>",
            Example = "xxx")]
        public static object TransformInverse(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("input points", Name = "Coordinate(s)")] double[,] Coords,
            [ExcelArgument("Output mode: < 7 and flag: > 7. (0); 0 returns nr of input columns, 1 = (x, y, z, t), 2 = (x, y, z), 3 = (x, y), 4 = (x), 5 = (y), 6 = (z), 7 = (t). Check flag values 2^n in the help file", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1000' when not used (-1000)", Name = "Accuracy")] object oAccuracy,
            [ExcelArgument("WestLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "WestLongitude")] object oWestLongitude,
            [ExcelArgument("SouthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "SouthLatitude")] object oSouthLatitude,
            [ExcelArgument("EastLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "EastLongitude")] object oEastLongitude,
            [ExcelArgument("NorthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "NorthLatitude")] object oNorthLatitude)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            // Check general input data
            int nMode       = (int)Optional.Check(oMode, 0.0);
            bool bUsingTransform = Optional.IsNul(TargetCrs);
            double Accuracy      = Optional.Check(oAccuracy,      -1000.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 4096)
                return ExcelError.ExcelErrorValue;

            // Check specific input parameters
            int nCoordRows = Coords.GetLength(0);
            int nCoordCols = Coords.GetLength(1);

            if (nCoordRows < 1 )
                return ExcelError.ExcelErrorValue;

            if (nCoordCols < 2 || nCoordCols > 4 )
                return ExcelError.ExcelErrorValue;

            int nOut = GetNrCoordinateColumns(nMode, nCoordCols);

            // Deal with optional parameters
            bool bUseNetwork = false;
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS, ref bUseNetwork);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = new ProjContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.GetCrs(SourceCrs, pjContext);
                    crsTarget = Crs.GetCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS, bUseNetwork);
                }

                // start core of function
                double[] x = new double[nCoordRows];
                double[] y = new double[nCoordRows];
                double[] z = new double[nCoordRows];
                double[] t = new double[nCoordRows];
                object[,] res = new object[nCoordRows, nOut];

                // work with nr of input columns
                switch (nCoordCols)
                {
                    default:
                    case 2:
                        // we have two columns (x, y) in the input data
                        for (int i = 0; i < nCoordRows; i++)
                        {
                            x[i] = Coords[i, 0];
                            y[i] = Coords[i, 1];
                        }

                        transform.ApplyReversed(x, y);
                        break;

                    case 3:
                        // we have three columns (x, y, z) in the input data
                        for (int i = 0; i < nCoordRows; i++)
                        {
                            x[i] = Coords[i, 0];
                            y[i] = Coords[i, 1];
                            z[i] = Coords[i, 2];
                        }

                        transform.ApplyReversed(x, y, z);
                        break;

                    case 4:
                        // we have four columns (x, y, z, t) in the input data
                        for (int i = 0; i < nCoordRows; i++)
                        {
                            x[i] = Coords[i, 0];
                            y[i] = Coords[i, 1];
                            z[i] = Coords[i, 2];
                            t[i] = Coords[i, 3];
                        }

                        transform.ApplyReversed(x, y, z, t);
                        break;
                }

                // determine what to do with output
                switch (nMode & 7) // use only the lowest three bits
                {
                    case 0:
                    case 1:
                    case 2:
                    case 3:

                        // all values to be returned
                        // check how many columns we need
                        if (nOut == 2)
                        {
                            for (int i = 0; i < nCoordRows; i++)
                            {
                                res[i, 0] = Optional.CheckNan(x[i]);
                                res[i, 1] = Optional.CheckNan(y[i]);
                            }
                        }
                        else if (nOut == 3)
                        {
                            for (int i = 0; i < nCoordRows; i++)
                            {
                                res[i, 0] = Optional.CheckNan(x[i]);
                                res[i, 1] = Optional.CheckNan(y[i]);
                                res[i, 2] = Optional.CheckNan(z[i]);
                            }
                        }
                        else if (nOut == 4)
                        {
                            for (int i = 0; i < nCoordRows; i++)
                            {
                                res[i, 0] = Optional.CheckNan(x[i]);
                                res[i, 1] = Optional.CheckNan(y[i]);
                                res[i, 2] = Optional.CheckNan(z[i]);
                                res[i, 3] = Optional.CheckNan(t[i]);
                            }
                        }
                        else
                            return ExcelError.ExcelErrorValue;
                        break;
                    case 4:
                        // from here onwards, a single output value is required
                        for (int i = 0; i < nCoordRows; i++)
                        {
                            res[i, 0] = Optional.CheckNan(x[i]);
                        }
                        break;
                    case 5:
                        for (int i = 0; i < nCoordRows; i++)
                        {
                            res[i, 0] = Optional.CheckNan(y[i]);
                        }
                        break;
                    case 6:
                        for (int i = 0; i < nCoordRows; i++)
                        {
                            res[i, 0] = Optional.CheckNan(z[i]);
                        }
                        break;
                    case 7:
                        for (int i = 0; i < nCoordRows; i++)
                        {
                            res[i, 0] = Optional.CheckNan(t[i]);
                        }
                        break;
                    default:
                        throw new System.InvalidOperationException("error in switch statement");
                }
                return res;
                // end core of function

            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // TransformInverse

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.Transforms.Count",
            Description = "Get the number of available transforms",
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1000",

            Returns = "The number of available transforms",
            Summary = "Returns the number of available transforms that exist between two Coordinate Reference Systems",
            Example = "xxx")]
        public static object Transforms_Count(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1000' when not used (-1000)", Name = "Accuracy")] object oAccuracy,
            [ExcelArgument("WestLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "WestLongitude")] object oWestLongitude,
            [ExcelArgument("SouthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "SouthLatitude")] object oSouthLatitude,
            [ExcelArgument("EastLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "EastLongitude")] object oEastLongitude,
            [ExcelArgument("NorthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "NorthLatitude")] object oNorthLatitude)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            // Check general input data
            int nMode       = (int)Optional.Check(oMode, 0.0);
            bool bUsingTransform = Optional.IsNul(TargetCrs);
            double Accuracy      = Optional.Check(oAccuracy,      -1000.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 4096)
                return ExcelError.ExcelErrorValue;

            // Deal with optional parameters
            bool bUseNetwork = false;
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS, ref bUseNetwork);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = new ProjContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.GetCrs(SourceCrs, pjContext);
                    crsTarget = Crs.GetCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS, bUseNetwork);
                }

                // start core of function
                ChooseCoordinateTransform transforms = transform as ChooseCoordinateTransform;
                return (transforms != null) ? transforms.Count : 1;
                // end core of function

            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // Transforms_Count

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.Transforms.ListAll",
            Description = "Get the number of available transforms",
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1000",

            Returns = "The number of available transforms",
            Summary = "Returns the number of available transforms that exist between two Coordinate Reference Systems",
            Example = "xxx")]
        public static object Transforms_ListAll(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Output mode: (0); 0 = PROJ string, 1 = WKT string, 2 = JSON string. Mode is combined with 2^n flag: 8, 16, ..., 2048. See help file for more information", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1000' when not used (-1000)", Name = "Accuracy")] object oAccuracy,
            [ExcelArgument("WestLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "WestLongitude")] object oWestLongitude,
            [ExcelArgument("SouthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "SouthLatitude")] object oSouthLatitude,
            [ExcelArgument("EastLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "EastLongitude")] object oEastLongitude,
            [ExcelArgument("NorthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "NorthLatitude")] object oNorthLatitude)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            // Check general input data
            int nMode       = (int)Optional.Check(oMode, 0.0);
            bool bUsingTransform = Optional.IsNul(TargetCrs);
            double Accuracy      = Optional.Check(oAccuracy,      -1000.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 4096)
                return ExcelError.ExcelErrorValue;

            // Check specific input data
            int nOutput = nMode & 3;

            // Deal with optional parameters
            bool bUseNetwork = false;
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS, ref bUseNetwork);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = new ProjContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.GetCrs(SourceCrs, pjContext);
                    crsTarget = Crs.GetCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS, bUseNetwork);
                }

                // start core of function
                ChooseCoordinateTransform transforms = transform as ChooseCoordinateTransform;
                int count = (transforms != null) ? transforms.Count : 1;

                object[,] res = new object[count + 1, 12];

                res[0,  0] = "Transform Identifiers";
                res[0,  1] = "Transform Name";
                res[0,  2] = "Accuracy [m]";
                res[0,  3] = "Area of Use";
                res[0,  4] = "Scope";
                res[0,  5] = "Lat-min";
                res[0,  6] = "Lon-min";
                res[0,  7] = "Lat-max";
                res[0,  8] = "Lon-max";
                res[0,  9] = "Nr grids";
                res[0, 10] = "1st grid";

                switch(nOutput)
                {
                    default:
                    case 0:
                        res[0, 11] = "Transform PROJ string";
                        break;
                    case 1:
                        res[0, 11] = "Transform WKT string";
                        break;
                    case 2:
                        res[0, 11] = "Transform JSON string";
                        break;
                }

                object accuracy; 

                if (count == 1)
                {
                    string ids = transform.Identifier?.ToString();
                    string scopes = transform.Scope;
                    if (ids == null && transform is CoordinateTransformList ctl)
                    {
                        ids = string.Join(", ", ctl.Where(x => x.Identifiers != null).SelectMany(x => x.Identifiers));
                        scopes = string.Join(", ", ctl.Select(x => x.Scope).Where(x => x != null).Distinct());
                    }

                    if (transform.Accuracy is null || transform.Accuracy <= 0.0) 
                    { 
                        accuracy = "Unknown"; 
                    } 
                    else
                    {
                        accuracy = transform.Accuracy;
                    }

                    res[1,  0] = transform.Identifiers is null ? ids : transform.Identifier.Authority;;
                    res[1,  1] = transform.Name;
                    res[1,  2] = accuracy;
                    res[1,  3] = transform.UsageArea.Name;
                    res[1,  4] = transform.Scope is null ? scopes  : transform.Scope;
                    res[1,  5] = transform.UsageArea.SouthLatitude;
                    res[1,  6] = transform.UsageArea.WestLongitude;
                    res[1,  7] = transform.UsageArea.NorthLatitude;
                    res[1,  8] = transform.UsageArea.EastLongitude;
                    res[1,  9] = transform.GridUsages.Count;
                    res[1, 10] = transform.GridUsages.Count > 0 ? transform.GridUsages[0].FullName : "N.A.";

                    switch(nOutput)
                    {
                        default:
                        case 0:
                            res[1, 11] = transform.AsProjString();
                            break;
                        case 1:
                            res[1, 11] = transform.AsWellKnownText();
                            break;
                        case 2:
                            res[1, 11] = transform.AsProjJson();
                            break;
                    }
                }
                else
                {
                    for (int i = 0; i < count; i++)
                    {
                        string ids = transforms[i].Identifier?.ToString();
                        string scopes = transforms[i].Scope;
                        if (ids == null && transforms[i] is CoordinateTransformList ctl)
                        {
                            ids = string.Join(", ", ctl.Where(x => x.Identifiers != null).SelectMany(x => x.Identifiers));
                            scopes = string.Join(", ", ctl.Select(x => x.Scope).Where(x => x != null).Distinct());
                        }
                        if (transforms[i].Accuracy is null || transforms[i].Accuracy <= 0.0)
                        { 
                            accuracy = "Unknown"; 
                        } 
                        else 
                        { 
                            accuracy = transforms[i].Accuracy;
                        }

                        res[i + 1,  0] = transforms[i].Identifiers is null ? ids : transforms[i].Identifier.Authority;
                        res[i + 1,  1] = transforms[i].Name;
                        res[i + 1,  2] = accuracy;
                        res[i + 1,  3] = transforms[i].UsageArea.Name;
                        res[i + 1,  4] = transforms[i].Scope is null ? scopes : transforms[i].Scope;
                        res[i + 1,  5] = transforms[i].UsageArea.SouthLatitude;
                        res[i + 1,  6] = transforms[i].UsageArea.WestLongitude;
                        res[i + 1,  7] = transforms[i].UsageArea.NorthLatitude;
                        res[i + 1,  8] = transforms[i].UsageArea.EastLongitude;
                        res[i + 1,  9] = transforms[i].GridUsages.Count;
                        res[i + 1, 10] = transforms[i].GridUsages.Count > 0 ? transforms[i].GridUsages[0].FullName : "N.A.";

                        switch(nOutput)
                        {
                            default:
                            case 0:
                                res[i + 1, 11] = transforms[i].AsProjString();
                                break;
                            case 1:
                                res[i + 1, 11] = transforms[i].AsWellKnownText();
                                break;
                            case 2:
                                res[i + 1, 11] = transforms[i].AsProjJson();
                                break;
                        }
                    }
                }
                return res;
                // end core of function

            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // Transforms_ListAll

    } // class

} // namespace

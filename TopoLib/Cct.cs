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

// On solving a missing reference to the next package:
// For me adding the PackageReference for MSTest.TestFramework did the trick. I didn't need to reference the TestAdapter.
// see https://stackoverflow.com/questions/13602508/where-to-find-microsoft-visualstudio-testtools-unittesting-missing-dll
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;

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
                // nMode = 2048 will arrive here but won't set any flags; for debugging only
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
            Name = "TL.cct.Count",
            Description = "Get the number of available transforms",
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1000",

            Returns = "The number of available transforms",
            Summary = "Returns the number of available transforms",
            Example = "xxx")]
        public static object Count(
            [ExcelArgument("sourceCrs using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrs")] object[,] SourceCrs,
            [ExcelArgument("targetCrs using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrs")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object mode,
            [ExcelArgument("Desired accuray for the transformation, or '-1000' when not used (-1000)", Name = "Accuracy")] object oAccuracy,
            [ExcelArgument("WestLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "WestLongitude")] object oWestLongitude,
            [ExcelArgument("SouthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "SouthLatitude")] object oSouthLatitude,
            [ExcelArgument("EastLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "EastLongitude")] object oEastLongitude,
            [ExcelArgument("NorthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "NorthLatitude")] object oNorthLatitude)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            // Check input data
            int nMode       = (int)Optional.Check(mode, 0.0);
            double Accuracy      = Optional.Check(oAccuracy,      -1000.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 4096)
                return ExcelError.ExcelErrorValue;

            // do the work
            try
            {
                using (var pc = new ProjContext())
                {
                    using (CoordinateReferenceSystem crsSource = Crs.GetCrs(SourceCrs, pc), crsTarget = Crs.GetCrs(TargetCrs, pc))
                    {
                        // Deal with optional parameters
                        bool bUseNetwork = false;
                        bool bAllowDeprecatedCRS = false;
                        var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS, ref bUseNetwork);

                        // Setup the transform
                        var transform = CreateCoordinateTransform(crsSource, crsTarget, options, pc, bAllowDeprecatedCRS, bUseNetwork);
                        
                        ChooseCoordinateTransform cl = transform as ChooseCoordinateTransform;

                        if (cl != null)
                            return cl.Count;
                        else
                            return 1;
                    }
                }
            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }
        } // Count

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
            [ExcelArgument("sourceCrs using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrs")] object[,] SourceCrs,
            [ExcelArgument("targetCrs using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrs")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object mode,
            [ExcelArgument("Desired accuray for the transformation, or '-1000' when not used (-1000)", Name = "Accuracy")] object oAccuracy,
            [ExcelArgument("WestLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "WestLongitude")] object oWestLongitude,
            [ExcelArgument("SouthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "SouthLatitude")] object oSouthLatitude,
            [ExcelArgument("EastLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "EastLongitude")] object oEastLongitude,
            [ExcelArgument("NorthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "NorthLatitude")] object oNorthLatitude)

        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            // Check input data
            int nMode       = (int)Optional.Check(mode, 0.0);
            double Accuracy      = Optional.Check(oAccuracy,      -1000.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 4096)
                return ExcelError.ExcelErrorValue;

            // do the work
            try
            {
                using (var pc = new ProjContext())
                {
                    using (CoordinateReferenceSystem crsSource = Crs.GetCrs(SourceCrs, pc), crsTarget = Crs.GetCrs(TargetCrs, pc))
                    {
                        // Deal with optional parameters
                        bool bUseNetwork = false;
                        bool bAllowDeprecatedCRS = false;
                        var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS, ref bUseNetwork);

                        // Setup the transform
                        var transform = CreateCoordinateTransform(crsSource, crsTarget, options, pc, bAllowDeprecatedCRS, bUseNetwork);

                        double? accuracy = transform.Accuraracy;

                        if (accuracy.HasValue) 
                            return accuracy;
                        else
                            throw new System.InvalidOperationException ("Transform accuracy not available");
                    }
                }        
            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
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
            [ExcelArgument("sourceCrs using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrs")] object[,] SourceCrs,
            [ExcelArgument("targetCrs using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrs")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object mode,
            [ExcelArgument("Desired accuray for the transformation, or '-1000' when not used (-1000)", Name = "Accuracy")] object oAccuracy,
            [ExcelArgument("WestLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "WestLongitude")] object oWestLongitude,
            [ExcelArgument("SouthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "SouthLatitude")] object oSouthLatitude,
            [ExcelArgument("EastLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "EastLongitude")] object oEastLongitude,
            [ExcelArgument("NorthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "NorthLatitude")] object oNorthLatitude)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            // Check input data
            int nMode       = (int)Optional.Check(mode, 0.0);
            double Accuracy      = Optional.Check(oAccuracy,      -1000.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 4096)
                return ExcelError.ExcelErrorValue;

            // do the work
            try
            {
                using (var pc = new ProjContext())
                {
                    using (CoordinateReferenceSystem crsSource = Crs.GetCrs(SourceCrs, pc), crsTarget = Crs.GetCrs(TargetCrs, pc))
                    {
                        // Deal with optional parameters
                        bool bUseNetwork = false;
                        bool bAllowDeprecatedCRS = false;
                        var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS, ref bUseNetwork);

                        // Setup the transform
                        var transform = CreateCoordinateTransform(crsSource, crsTarget, options, pc, bAllowDeprecatedCRS, bUseNetwork);

                        return (double) transform.GridUsages.Count;
                    }
                }        
            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }

        } // GridUsages_Count

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
            [ExcelArgument("sourceCrs using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrs")] object[,] SourceCrs,
            [ExcelArgument("targetCrs using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrs")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object mode,
            [ExcelArgument("Desired accuray for the transformation, or '-1000' when not used (-1000)", Name = "Accuracy")] object oAccuracy,
            [ExcelArgument("WestLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "WestLongitude")] object oWestLongitude,
            [ExcelArgument("SouthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "SouthLatitude")] object oSouthLatitude,
            [ExcelArgument("EastLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "EastLongitude")] object oEastLongitude,
            [ExcelArgument("NorthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "NorthLatitude")] object oNorthLatitude)
        {
            // Check input data
            int nMode       = (int)Optional.Check(mode, 0.0);
            double Accuracy      = Optional.Check(oAccuracy,      -1000.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 4096)
                return ExcelError.ExcelErrorValue;

            // do the work
            try
            {
                using (var pc = new ProjContext())
                {
                    using (CoordinateReferenceSystem crsSource = Crs.GetCrs(SourceCrs, pc), crsTarget = Crs.GetCrs(TargetCrs, pc))
                    {
                        // Deal with optional parameters
                        bool bUseNetwork = false;
                        bool bAllowDeprecatedCRS = false;
                        var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS, ref bUseNetwork);

                        // Setup the transform
                        var transform = CreateCoordinateTransform(crsSource, crsTarget, options, pc, bAllowDeprecatedCRS, bUseNetwork);

                        return transform.HasInverse;
                    }
                }        
            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex, false);
            }

        } // HasInverse


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
            [ExcelArgument("sourceCrs using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrs")] object[,] SourceCrs,
            [ExcelArgument("targetCrs using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrs")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object mode,
            [ExcelArgument("Desired accuray for the transformation, or '-1000' when not used (-1000)", Name = "Accuracy")] object oAccuracy,
            [ExcelArgument("WestLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "WestLongitude")] object oWestLongitude,
            [ExcelArgument("SouthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "SouthLatitude")] object oSouthLatitude,
            [ExcelArgument("EastLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "EastLongitude")] object oEastLongitude,
            [ExcelArgument("NorthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "NorthLatitude")] object oNorthLatitude)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            // Check input data
            int nMode       = (int)Optional.Check(mode, 0.0);
            double Accuracy      = Optional.Check(oAccuracy,      -1000.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 4096)
                return ExcelError.ExcelErrorValue;

            // do the work
            try
            {
                using (var pc = new ProjContext())
                {
                    using (CoordinateReferenceSystem crsSource = Crs.GetCrs(SourceCrs, pc), crsTarget = Crs.GetCrs(TargetCrs, pc))
                    {
                        // Deal with optional parameters
                        bool bUseNetwork = false;
                        bool bAllowDeprecatedCRS = false;
                        var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS, ref bUseNetwork);

                        // Setup the transform
                        var transform = CreateCoordinateTransform(crsSource, crsTarget, options, pc, bAllowDeprecatedCRS, bUseNetwork);

                        return transform.Name;
                    }
                }        
            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
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
            Example = "xxx")]
        public static object RoundTrip(
            [ExcelArgument("sourceCrs using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrs")] object[,] SourceCrs,
            [ExcelArgument("targetCrs using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrs")] object[,] TargetCrs,
            [ExcelArgument("test point with adjacent [x, y] coordinates", Name = "point(x, y)")] object[,] TestCoord,
            [ExcelArgument("N - nr of roundtrips to make", Name = "nr roundtrips")] object oRoundTrips,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object mode,
            [ExcelArgument("Desired accuray for the transformation, or '-1000' when not used (-1000)", Name = "Accuracy")] object oAccuracy,
            [ExcelArgument("WestLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "WestLongitude")] object oWestLongitude,
            [ExcelArgument("SouthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "SouthLatitude")] object oSouthLatitude,
            [ExcelArgument("EastLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "EastLongitude")] object oEastLongitude,
            [ExcelArgument("NorthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "NorthLatitude")] object oNorthLatitude)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            // Check input data
            int nTrips = (int)Optional.Check(oRoundTrips, 1.0);

            // max three adjacent [x, y, z] cells on the same row
            if (TestCoord.GetLength(0) != 1 || TestCoord.GetLength(1) > 3 ) return ExcelError.ExcelErrorValue;

            // Check input data
            int nMode       = (int)Optional.Check(mode, 0.0);
            double Accuracy      = Optional.Check(oAccuracy,      -1000.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 4096)
                return ExcelError.ExcelErrorValue;

            // do the work
            try
            {
                using (var pc = new ProjContext())
                {
                    using (CoordinateReferenceSystem crsSource = Crs.GetCrs(SourceCrs, pc), crsTarget = Crs.GetCrs(TargetCrs, pc))
                    {
                        // Deal with optional parameters
                        bool bUseNetwork = false;
                        bool bAllowDeprecatedCRS = false;
                        var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS, ref bUseNetwork);

                        // Setup the transform
                        var transform = CreateCoordinateTransform(crsSource, crsTarget, options, pc, bAllowDeprecatedCRS, bUseNetwork);

                        double x = (double)TestCoord[0, 0];
                        double y = (double)TestCoord[0, 1];
                        double z = TestCoord.GetLength(1) == 3 ? (double)TestCoord[0, 2] : 0;
                        PPoint pt = new PPoint(x, y, z);

                        double error = transform.RoundTrip(true, nTrips, pt);

                        if (Double.IsInfinity(error))
                            // throw new System.ArithmeticException("Infinite roundtrip error");
                            throw new System.InvalidOperationException("Infinite roundtrip error");

                        return error;
                    }
                }        
            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex, -1.0);
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
            [ExcelArgument("sourceCrs using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrs")] object[,] SourceCrs,
            [ExcelArgument("targetCrs using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrs")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object mode,
            [ExcelArgument("Desired accuray for the transformation, or '-1000' when not used (-1000)", Name = "Accuracy")] object oAccuracy,
            [ExcelArgument("WestLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "WestLongitude")] object oWestLongitude,
            [ExcelArgument("SouthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "SouthLatitude")] object oSouthLatitude,
            [ExcelArgument("EastLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "EastLongitude")] object oEastLongitude,
            [ExcelArgument("NorthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "NorthLatitude")] object oNorthLatitude)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            // Check input data
            int nMode       = (int)Optional.Check(mode, 0.0);
            double Accuracy      = Optional.Check(oAccuracy,      -1000.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 4096)
                return ExcelError.ExcelErrorValue;

            // do the work
            try
            {
                using (var pc = new ProjContext())
                {
                    using (CoordinateReferenceSystem crsSource = Crs.GetCrs(SourceCrs, pc), crsTarget = Crs.GetCrs(TargetCrs, pc))
                    {
                        // Deal with optional parameters
                        bool bUseNetwork = false;
                        bool bAllowDeprecatedCRS = false;
                        var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS, ref bUseNetwork);

                        // Setup the transform
                        var transform = CreateCoordinateTransform(crsSource, crsTarget, options, pc, bAllowDeprecatedCRS, bUseNetwork);

                        CoordinateTransformList steps = transform as CoordinateTransformList;
                        if (steps is null)
                            return ExcelError.ExcelErrorValue;

                        return steps.Count;
                    }
                }        
            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
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
            [ExcelArgument("sourceCrs using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrs")] object[,] SourceCrs,
            [ExcelArgument("targetCrs using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrs")] object[,] TargetCrs,
            [ExcelArgument("Zero based index of Identifier list (0) ", Name = "Index")] object index,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object mode,
            [ExcelArgument("Desired accuray for the transformation, or '-1000' when not used (-1000)", Name = "Accuracy")] object oAccuracy,
            [ExcelArgument("WestLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "WestLongitude")] object oWestLongitude,
            [ExcelArgument("SouthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "SouthLatitude")] object oSouthLatitude,
            [ExcelArgument("EastLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "EastLongitude")] object oEastLongitude,
            [ExcelArgument("NorthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "NorthLatitude")] object oNorthLatitude)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            int nIndex  = (int)Optional.Check(index , 0.0);

            // Check input data
            int nMode       = (int)Optional.Check(mode, 0.0);
            double Accuracy      = Optional.Check(oAccuracy,      -1000.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 4096)
                return ExcelError.ExcelErrorValue;
            // do the work
            try
            {
                using (var pc = new ProjContext())
                {
                    using (CoordinateReferenceSystem crsSource = Crs.GetCrs(SourceCrs, pc), crsTarget = Crs.GetCrs(TargetCrs, pc))
                    {
                        // Deal with optional parameters
                        bool bUseNetwork = false;
                        bool bAllowDeprecatedCRS = false;
                        var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS, ref bUseNetwork);

                        // Setup the transform
                        var transform = CreateCoordinateTransform(crsSource, crsTarget, options, pc, bAllowDeprecatedCRS, bUseNetwork);

                        CoordinateTransformList steps = transform as CoordinateTransformList;
                        if (steps is null)
                            return ExcelError.ExcelErrorValue;

                        return steps[nIndex].MethodName;
                    }
                }        
            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
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
            [ExcelArgument("sourceCrs using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrs")] object[,] SourceCrs,
            [ExcelArgument("targetCrs using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrs")] object[,] TargetCrs,
            [ExcelArgument("Zero based index of Identifier list (0) ", Name = "Index")] object index,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object mode,
            [ExcelArgument("Desired accuray for the transformation, or '-1000' when not used (-1000)", Name = "Accuracy")] object oAccuracy,
            [ExcelArgument("WestLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "WestLongitude")] object oWestLongitude,
            [ExcelArgument("SouthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "SouthLatitude")] object oSouthLatitude,
            [ExcelArgument("EastLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "EastLongitude")] object oEastLongitude,
            [ExcelArgument("NorthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "NorthLatitude")] object oNorthLatitude)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            int nIndex  = (int)Optional.Check(index , 0.0);

            // Check input data
            int nMode       = (int)Optional.Check(mode, 0.0);
            double Accuracy      = Optional.Check(oAccuracy,      -1000.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 4096)
                return ExcelError.ExcelErrorValue;

            // do the work
            try
            {
                using (var pc = new ProjContext())
                {
                    using (CoordinateReferenceSystem crsSource = Crs.GetCrs(SourceCrs, pc), crsTarget = Crs.GetCrs(TargetCrs, pc))
                    {
                        // Deal with optional parameters
                        bool bUseNetwork = false;
                        bool bAllowDeprecatedCRS = false;
                        var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS, ref bUseNetwork);

                        // Setup the transform
                        var transform = CreateCoordinateTransform(crsSource, crsTarget, options, pc, bAllowDeprecatedCRS, bUseNetwork);

                        CoordinateTransformList steps = transform as CoordinateTransformList;
                        if (steps is null)
                            return ExcelError.ExcelErrorValue;

                        return steps[nIndex].Name;
                    }
                }        
            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }

        } // Steps_Name


        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.TransformForward",
            Description = "Coordinate conversion of one or more input points using the 'SharpProj' libray", 
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
            [ExcelArgument("sourceCrs using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrs")] object[,] SourceCrs,
            [ExcelArgument("targetCrs using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrs")] object[,] TargetCrs,
            [ExcelArgument("input points", Name = "Coordinate(s)")] double[,] Coords,
            [ExcelArgument("Output mode: < 7 and flag: > 7. (0); 0 returns nr of input columns, 1 = (x, y, z, t), 2 = (x, y, z), 3 = (x, y), 4 = (x), 5 = (y), 6 = (z), 7 = (t). Check flag values 2^n in the help file", Name = "mode")] object mode,
            [ExcelArgument("Desired accuray for the transformation, or '-1000' when not used (-1000)", Name = "Accuracy")] object oAccuracy,
            [ExcelArgument("WestLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "WestLongitude")] object oWestLongitude,
            [ExcelArgument("SouthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "SouthLatitude")] object oSouthLatitude,
            [ExcelArgument("EastLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "EastLongitude")] object oEastLongitude,
            [ExcelArgument("NorthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "NorthLatitude")] object oNorthLatitude)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            // Check input data
            int nMode       = (int)Optional.Check(mode, 0.0);
            double Accuracy      = Optional.Check(oAccuracy,      -1000.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 4096)
                return ExcelError.ExcelErrorValue;

            int nCoordRows = Coords.GetLength(0);
            int nCoordCols = Coords.GetLength(1);

            if (nCoordRows < 1 )
                return ExcelError.ExcelErrorValue;

            if (nCoordCols < 2 || nCoordCols > 4 )
                return ExcelError.ExcelErrorValue;

            int nOut = GetNrCoordinateColumns(nMode, nCoordCols);

            // do the work
            try
            {
                using (var pc = new ProjContext())
                {
                    using (CoordinateReferenceSystem crsSource = Crs.GetCrs(SourceCrs, pc), crsTarget = Crs.GetCrs(TargetCrs, pc))
                    {
                        // Deal with optional parameters
                        bool bUseNetwork = false;
                        bool bAllowDeprecatedCRS = false;
                        var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS, ref bUseNetwork);

                        // Setup the transform
                        var transform = CreateCoordinateTransform(crsSource, crsTarget, options, pc, bAllowDeprecatedCRS, bUseNetwork);

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

                                for (int i = 0; i < nCoordRows; i++)
                                {
                                    x[i] = Coords[i, 0];
                                    y[i] = Coords[i, 1];

                                    PPoint pt = transform.Apply(new PPoint(x[i], y[i]));
                                    x[i] = pt.X;
                                    y[i] = pt.Y;
                                    if (Double.IsInfinity(x[i] ) || Double.IsInfinity(y[i]))
                                        throw new System.InvalidOperationException ("No transformation available");
                                }


/*                              transform.Apply(Coords[0, 0],   2, nCoordRows,
                                            Coords[0, 1],   2, nCoordRows,
                                            0,              2, 0,
                                            0,              2, 0);

*/
                                break;

                            case 3:
                                for (int i = 0; i < nCoordRows; i++)
                                {
                                    x[i] = Coords[i, 0];
                                    y[i] = Coords[i, 1];
                                    z[i] = Coords[i, 2];

                                    PPoint pt = transform.Apply(new PPoint(x[i], y[i], z[i]));
                                    x[i] = pt.X;
                                    y[i] = pt.Y;
                                    z[i] = pt.Z;
                                    if (Double.IsInfinity(x[i] ) || Double.IsInfinity(y[i]) || Double.IsInfinity(z[i]))
                                        throw new System.InvalidOperationException ("No transformation available");
                                }
/*
                                    transform.Apply(x[0], 1, nCoordRows,
                                                    y[0], 1, nCoordRows,
                                                    z[0], 1, nCoordRows,
                                                    0, 0, 0);
*/                              
                                break;

                            case 4:
                                for (int i = 0; i < nCoordRows; i++)
                                {
                                    x[i] = Coords[i, 0];
                                    y[i] = Coords[i, 1];
                                    z[i] = Coords[i, 2];
                                    t[i] = Coords[i, 3];

                                    PPoint pt = transform.Apply(new PPoint(x[i], y[i], z[i], t[i]));
                                    x[i] = pt.X;
                                    y[i] = pt.Y;
                                    z[i] = pt.Z;
                                    t[i] = pt.T;
                                    if (Double.IsInfinity(x[i] ) || Double.IsInfinity(y[i]) || Double.IsInfinity(z[i]) || Double.IsInfinity(t[i]))
                                        throw new System.InvalidOperationException ("No transformation available");
                                }
/*
                                    transform.Apply(x[0], 1, nCoordRows,
                                                    y[0], 1, nCoordRows,
                                                    z[0], 1, nCoordRows,
                                                    t[0], 1, nCoordRows);    
*/
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
                                    for( int i = 0; i < nCoordRows; i++)
                                    {
                                        res[i, 0] = x[i];
                                        res[i, 1] = y[i];
                                    }
                                }
                                else if (nOut == 3)
                                {
                                    for( int i = 0; i < nCoordRows; i++)
                                    {
                                        res[i, 0] = x[i];
                                        res[i, 1] = y[i];
                                        res[i, 2] = z[i];
                                    }
                                }
                                else if (nOut == 4)
                                {
                                    for( int i = 0; i < nCoordRows; i++)
                                    {
                                        res[i, 0] = x[i];
                                        res[i, 1] = y[i];
                                        res[i, 2] = z[i];
                                        res[i, 3] = t[i];
                                    }
                                }
                                else 
                                    return ExcelError.ExcelErrorValue;
                                break;
                            case 4:
                                // from here onwards, a single output value is required
                                    for( int i = 0; i < nCoordRows; i++)
                                    {
                                        res[i, 0] = x[i];
                                    }
                                break;
                            case 5:
                                    for( int i = 0; i < nCoordRows; i++)
                                    {
                                        res[i, 0] = y[i];
                                    }
                                break;
                            case 6:
                                for( int i = 0; i < nCoordRows; i++)
                                    {
                                        res[i, 0] = z[i];
                                    }
                                break;
                            case 7:
                                for( int i = 0; i < nCoordRows; i++)
                                    {
                                        res[i, 0] = t[i];
                                    }
                                break;
                            default:
                                throw new System.InvalidOperationException ("error in switch statement");
                        }

                        return res;
                    }
                }        
            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }

        } // TransformForward

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.TransformInverse",
            Description = "Inverse coordinate conversion of one or more input points using the 'SharpProj' libray", 
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
            [ExcelArgument("sourceCrs using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrs")] object[,] SourceCrs,
            [ExcelArgument("targetCrs using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrs")] object[,] TargetCrs,
            [ExcelArgument("input points", Name = "Coordinate(s)")] double[,] Coords,
            [ExcelArgument("Output mode: < 7 and flag: > 7. (0); 0 returns nr of input columns, 1 = (x, y, z, t), 2 = (x, y, z), 3 = (x, y), 4 = (x), 5 = (y), 6 = (z), 7 = (t). Check flag values 2^n in the help file", Name = "mode")] object mode,
            [ExcelArgument("Desired accuray for the transformation, or '-1000' when not used (-1000)", Name = "Accuracy")] object oAccuracy,
            [ExcelArgument("WestLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "WestLongitude")] object oWestLongitude,
            [ExcelArgument("SouthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "SouthLatitude")] object oSouthLatitude,
            [ExcelArgument("EastLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "EastLongitude")] object oEastLongitude,
            [ExcelArgument("NorthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "NorthLatitude")] object oNorthLatitude)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            // Check input data
            int nMode       = (int)Optional.Check(mode, 0.0);
            double Accuracy      = Optional.Check(oAccuracy,      -1000.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 4096)
                return ExcelError.ExcelErrorValue;

            int nCoordRows = Coords.GetLength(0);
            int nCoordCols = Coords.GetLength(1);

            if (nCoordRows < 1 )
                return ExcelError.ExcelErrorValue;

            if (nCoordCols < 2 || nCoordCols > 4 )
                return ExcelError.ExcelErrorValue;

            int nOut = GetNrCoordinateColumns(nMode, nCoordCols);

            // do the work
            try
            {
                using (var pc = new ProjContext())
                {
                    using (CoordinateReferenceSystem crsSource = Crs.GetCrs(SourceCrs, pc), crsTarget = Crs.GetCrs(TargetCrs, pc))
                    {
                        // Deal with optional parameters
                        bool bUseNetwork = false;
                        bool bAllowDeprecatedCRS = false;
                        var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS, ref bUseNetwork);

                        // Setup the transform
                        var transform = CreateCoordinateTransform(crsSource, crsTarget, options, pc, bAllowDeprecatedCRS, bUseNetwork);

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

                                for (int i = 0; i < nCoordRows; i++)
                                {
                                    x[i] = Coords[i, 0];
                                    y[i] = Coords[i, 1];

                                    PPoint pt = transform.ApplyReversed(new PPoint(x[i], y[i]));
                                    x[i] = pt.X;
                                    y[i] = pt.Y;
                                    if (Double.IsInfinity(x[i] ) || Double.IsInfinity(y[i]))
                                        throw new System.InvalidOperationException ("No transformation available");
                                }


/*                              transform.ApplyReversed(Coords[0, 0],   2, nCoordRows,
                                            Coords[0, 1],   2, nCoordRows,
                                            0,              2, 0,
                                            0,              2, 0);

*/
                                break;

                            case 3:
                                for (int i = 0; i < nCoordRows; i++)
                                {
                                    x[i] = Coords[i, 0];
                                    y[i] = Coords[i, 1];
                                    z[i] = Coords[i, 2];

                                    PPoint pt = transform.ApplyReversed(new PPoint(x[i], y[i], z[i]));
                                    x[i] = pt.X;
                                    y[i] = pt.Y;
                                    z[i] = pt.Z;
                                    if (Double.IsInfinity(x[i] ) || Double.IsInfinity(y[i]) || Double.IsInfinity(z[i]))
                                        throw new System.InvalidOperationException ("No transformation available");
                                }
/*
                                    transform.ApplyReversed(x[0], 1, nCoordRows,
                                                    y[0], 1, nCoordRows,
                                                    z[0], 1, nCoordRows,
                                                    0, 0, 0);
*/                              
                                break;

                            case 4:
                                for (int i = 0; i < nCoordRows; i++)
                                {
                                    x[i] = Coords[i, 0];
                                    y[i] = Coords[i, 1];
                                    z[i] = Coords[i, 2];
                                    t[i] = Coords[i, 3];

                                    PPoint pt = transform.ApplyReversed(new PPoint(x[i], y[i], z[i], t[i]));
                                    x[i] = pt.X;
                                    y[i] = pt.Y;
                                    z[i] = pt.Z;
                                    t[i] = pt.T;
                                    if (Double.IsInfinity(x[i] ) || Double.IsInfinity(y[i]) || Double.IsInfinity(z[i]) || Double.IsInfinity(t[i]))
                                        throw new System.InvalidOperationException ("No transformation available");
                                }
/*
                                    transform.ApplyReversed(x[0], 1, nCoordRows,
                                                    y[0], 1, nCoordRows,
                                                    z[0], 1, nCoordRows,
                                                    t[0], 1, nCoordRows);    
*/
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
                                    for( int i = 0; i < nCoordRows; i++)
                                    {
                                        res[i, 0] = x[i];
                                        res[i, 1] = y[i];
                                    }
                                }
                                else if (nOut == 3)
                                {
                                    for( int i = 0; i < nCoordRows; i++)
                                    {
                                        res[i, 0] = x[i];
                                        res[i, 1] = y[i];
                                        res[i, 2] = z[i];
                                    }
                                }
                                else if (nOut == 4)
                                {
                                    for( int i = 0; i < nCoordRows; i++)
                                    {
                                        res[i, 0] = x[i];
                                        res[i, 1] = y[i];
                                        res[i, 2] = z[i];
                                        res[i, 3] = t[i];
                                    }
                                }
                                else 
                                    return ExcelError.ExcelErrorValue;
                                break;
                            case 4:
                                // from here onwards, a single output value is required
                                    for( int i = 0; i < nCoordRows; i++)
                                    {
                                        res[i, 0] = x[i];
                                    }
                                break;
                            case 5:
                                    for( int i = 0; i < nCoordRows; i++)
                                    {
                                        res[i, 0] = y[i];
                                    }
                                break;
                            case 6:
                                for( int i = 0; i < nCoordRows; i++)
                                    {
                                        res[i, 0] = z[i];
                                    }
                                break;
                            case 7:
                                for( int i = 0; i < nCoordRows; i++)
                                    {
                                        res[i, 0] = t[i];
                                    }
                                break;
                            default:
                                throw new System.InvalidOperationException ("error in switch statement");
                        }
                        return res;
                    }
                }        
            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }

        } // TransformInverse

    } // class

} // namespace

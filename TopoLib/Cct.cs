using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using ExcelDna.Integration;
using ExcelDna.Documentation;

// Added Bart
using SharpProj;
using SharpProj.Proj;

#pragma warning disable IDE0075 // Conditional expression can be simplified

// When connecting to the internet through a proxy; please read the following on stackoverflow:
// https://stackoverflow.com/questions/1938990/c-sharp-connecting-through-proxy

// On solving a missing reference to the next package:
// For me adding the PackageReference for MSTest.TestFramework did the trick. I didn't need to reference the TestAdapter.
// see https://stackoverflow.com/questions/13602508/where-to-find-microsoft-visualstudio-testtools-unittesting-missing-dll
// using Microsoft.VisualStudio.TestTools.UnitTesting;

// I made a backup of my project by renaming it to TopoLibOld 
// Then I rebuilt TopoLib from scratch starting from ExcelDna v1.1.0 in view of Virusscanner false positives with v1.5.0
// Next I added all source files, etc.
// But then Git was stuffed, because my project history under TopoLib did not jive with the master branch on the server.
// After some googling, I used the following command from the command line: 
//
// git push --set-upstream origin master -f
//
// This solved the problem, and my latest changes are uploaded to GitHub....
//
// To prevent static classes from other files to become visible in the Compiled Help File you need to use the /X flag when running ExcelDnaDoc.exe
// In the Post-build event command line put: "D:\Source\VS19\TopoLib\packages\ExcelDnaDoc.1.5.1\tools\ExcelDnaDoc.exe" "$(TargetDir)TopoLib-AddIn.dna" /X
//
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

// to refresh my memory on access modifiers in C# :
// *internal* is for assembly scope (i.e. only accessible from code in the same .exe or .dll)
// *private* is for class scope (i.e. accessible only from code in the same class).

// To do: need to implement DMS angle input and conversion. Have a look here for starters:
// D:\Source\VS19\DotSpatial\Source\DotSpatial.Positioning\Angle.cs


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

		internal static CoordinateTransformOptions GetCoordinateTransformOptions(int nMode, double Accuracy, double westLongitude, double southLatitude, double eastLongitude, double northLatitude, ref bool bAllowDeprecatedCRS)
        {
			var options = new CoordinateTransformOptions();

            if (CctOptions.UseGlobalSettings)
            {
                // get options from static variables
                options     = CctOptions.TransformOptions;
                bAllowDeprecatedCRS = CctOptions.AllowDeprecatedCRS;
            }
            else
            {
				if (westLongitude < -180 || westLongitude >  180 || eastLongitude < -180 || eastLongitude >  180 ||
					southLatitude <  -90 || southLatitude >   90 || northLatitude <  -90 || northLatitude >   90 ||
                    southLatitude > northLatitude)
					options.Area = null;
                else 
                    options.Area              = new CoordinateArea(westLongitude, southLatitude, eastLongitude, northLatitude);

                options.Accuracy              = Accuracy;
                options.NoBallparkConversions = (nMode &    8) != 0 ? true : false;
                options.NoDiscardIfMissing    = (nMode &   16) != 0 ? true : false;
                options.UsePrimaryGridNames   = (nMode &   32) != 0 ? true : false;
                options.UseSuperseded         = (nMode &   64) != 0 ? true : false;

                    bAllowDeprecatedCRS       = (nMode &  128) != 0 ? true : false;

                options.StrictContains        = (nMode &  256) != 0 ? true : false;
                options.IntermediateCrsUsage  = (nMode &  512) != 0 ? IntermediateCrsUsage.Always : IntermediateCrsUsage.Auto;
                options.IntermediateCrsUsage  = (nMode & 1024) != 0 ? IntermediateCrsUsage.Never  : IntermediateCrsUsage.Auto;
                // deal with 'Always' and 'Never' both being set. Go back to 'Auto' !
                if (((nMode & 512) != 0) && (nMode & 1024) != 0) options.IntermediateCrsUsage  = IntermediateCrsUsage.Auto;
            }

			return options;

        } // GetCoordinateTransformOptions

        internal static CoordinateTransform CreateCoordinateTransform(in object[,] oTransform, ProjContext pjContext)
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
                    return CoordinateTransform.CreateFromEpsg(nTransform, pjContext);
                }
                else if (oTransform[0, 0] is string)
                {
                    // cast to string, to deal with Excel datatypes
                    sTransform = (string)oTransform[0, 0];

                    bool success = int.TryParse(sTransform, out nTransform);
                    if (success)
                    {
                        // we have an EPSG number from a single input parameter:
                        return CoordinateTransform.CreateFromEpsg(nTransform, pjContext);
                    }
                    else
                    {
                        // we have a string of some sorts and a single input parameter:
                        if ((sTransform.IndexOf("PROJCRS") > -1) || (sTransform.IndexOf("GEOGCRS") > -1) || (sTransform.IndexOf("SPHEROID") > -1))
                        {
                            // it must be WKT (well, we hope)

                            // Note the cast used below is not required for CoordinateReferenceSystem where this function has been implemented as part of the inherited class
                            return (CoordinateTransform)CoordinateTransform.CreateFromWellKnownText(sTransform, pjContext);

                            // CreateFromWellKnownText() is translated into CreateFromWellKnownText(from, wars, ctx); where array<String^>^ wars = nullptr;
                            // It may throw an ArgumentNullException
                            // It may throw a ProjException
                        }
                        else
                        {
                            // it might be anything
                            return CoordinateTransform.Create(sTransform, pjContext);

                            // Create() is translated into proj_create(ctx, fromStr); 
                            // It may throw an ArgumentNullException("from");
                            // It may throw a ctx->ConstructException();
                            // It may throw a ProjException
                        }
                    }
                }
                else
                    throw new ArgumentException("Incorrect coordinate transform format");
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
                        throw new ArgumentException("Incorrect coordinate transform format");
                }
                else
                    throw new ArgumentException("Incorrect coordinate transform format");

                return CoordinateTransform.CreateFromDatabase(sTransform, nTransform, pjContext);

                // CreateFromDatabase() is translated into proj_create_from_database 
                // It may throw a ArgumentNullException
                // It may throw a pjContext->ConstructException
            }

            // Oops, something went wrong if we get here...
            throw new ArgumentException("Incorrect coordinate transform format");

        }

        internal static CoordinateTransform CreateCoordinateTransform(CoordinateReferenceSystem crsSource, CoordinateReferenceSystem crsTarget, CoordinateTransformOptions options, ProjContext pjContext, bool bAllowDeprecatedCRS)
        {
            bool bHasDeprecatedCRS = crsSource.IsDeprecated || crsTarget.IsDeprecated; 

            if (bHasDeprecatedCRS && !bAllowDeprecatedCRS)
                throw new ArgumentException ("Using deprecated CRS when not allowed");

            var transform = CoordinateTransform.Create(crsSource.WithAxisNormalized(), crsTarget.WithAxisNormalized(), options, pjContext);
        
            if (transform == null)
                throw new ArgumentException ("No transformation available");

            return transform;

        } // CreateCoordinateTransform

         [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.Accuracy",
            Description = "Get the accuracy of a transform", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1002",

            Returns = "Accuracy of a transform [m]",
            Summary = "Returns accuracy of a  coordinate transform",
            Example = "TL.cct.Accuracy(28992, 4326) returns 1.000",
            Remarks = "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag"
            )]
        public static object Accuracy(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
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
                return AddIn.ProcessException(ex);
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
            Name = "TL.cct.ApplyForward",
            Description = "Coordinate conversion of one or more input points", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1003",

            Returns = "The reprojected coordinate(s)",
            Summary =
            "<p>This function transforms coordinates from one Coordinate Reference system (CRS) into another CRS" +
            "<p>The transform can be performed using two different ways" +
            "<ol>    <li><p>By providing a seperate description for the SourceCrs as well as the TargetCrs</li>" +
                    "<li><p>By providing one string that describes the CRS-transform, leaving TargetCrs empty</li>" +
            "</ol>" +
            "<p>As such, SourceCrs and TargetCrs can be provided in one out of three ways" +
            "<ol>    <li><p>As a number referencing a CRS CODE from the EPSG database (much preferred)</li>" +
                    "<li><p>As a string using WKT, JSON or PROJ format. WKT or JSON format is preferred over the original PROJ string format</li>" +
                    "<li><p>As an AUTHORITY string in one cell, combined with a CRS CODE in the adjacent cell to the right</li>" +
            "</ol>" +
            "<p>This function is an array function. Array functions have undergone a significant upgrade with the introduction of dynamic arrays in Excel." +
            "<p>For more information on working with array formulas please consult :" +
            "<ol>" +
            "<li><p>This link: <a href = \"https://support.office.com/en-us/article/Guidelines-and-examples-of-array-formulas-7d94a64e-3ff3-4686-9372-ecfd5caa57c7\" > Microsoft Office Support - Guidelines and examples of array formulas</a> for Guidelines and examples of array formulas.</li>" +
            "<li><p>This link: <a href = \"https://support.office.com/en-us/article/Create-an-array-formula-e43e12e0-afc6-4a12-bc7f-48361075954d\" > Microsoft Office Support - Create a array formula</a> for more information on how to create static {CSE} array formulas.</li>" +
            "<li><p>This link: <a href = \"https://exceljet.net/dynamic-array-formulas-in-excel\" > ExcelJet - Dynamic array formulas in Excel</a> for an introduction to dynamic array formulas.</li>" +
            "</ol>" +
            "<p>For more information on coordinate conversion and coordinate refence system (CRS) information, see :" +
            "<ol>    <li><p>This link: <a href = \"http://spatialreference.org/\"> Spatial Reference home page</a></li>" +
                    "<li><p>This link: <a href = \"http://epsg.io/\" id=\"viewDesktopLink\"> EPSG IO home page with CRS description strings and EPSG numbers</a></li>" +
                    "<li><p>This link: <a href = \"http://proj.org/\"> Home page of the proj library</a></li>" +
            "</ol>",
            Remarks = "<p>Internally the transform uses <a href = \"https://proj.org/development/reference/functions.html#c.proj_normalize_for_visualization\"> crs normalization</a> by the proj library for a consistent approach to (x, y, z) values." +
            "<p>The axis order of a geographic CRS shall therefore be longitude, latitude [,height], and that of a projected CRS shall be easting, northing [, height]" +
            "<p>When using a geographic CRS, coordinates should be presented in degrees (not radians)." +
            "<p>The 'Mode' flag combines values 0 - 7 for the output mode with binary flags to reduce the number of parameters in this function. These binary flags are:" +
            "<pre><ul><li>   8: Disallow Ballpark Conversions</li>" +
                    "<li>  16: Don't Discard Transform if Grid is missing</li>" +
                    "<li>  32: Use Primary Grid Names</li>" +
                    "<li>  64: Use Superseded Transforms</li>" +
                    "<li> 128: Allow Deprecated CRSs</li>" +
                    "<li> 256: Transform strictly contains Area of Interest</li>" +
                    "<li> 512: Always Allow an Intermediate CRS</li>" +
                    "<li>1024: Never Allow an Intermediate CRS</li>" +
            "</ul></pre><br>" +
            "<p>Finally, please note the following aspects:</p>" +
            "<ol><li>some transforms require the use of one or more grid(s). Local/network access to these grids is controlled through the TopoLib Ribbon. </li>" +
                "<li>For all TL.cct-functions, the combined settings in the 'Mode Flag' can be overruled by global settings defined in the 'Transform Settings' Dialog. </li>" +
                "<li>Though it is possible to define a transform using a <b>single</b> (Wkt/Json/Proj) <b>string</b>, it is much preferred to apply <b>sourceCrs</b> and <b>targetCrs</b> to define a transform. " +
                "In that case the PROJ library automatically finds the most suitable transform.</li></ol>" +
                "<p>With respect to point 3 above, <a href = \"https://docs.opengeospatial.org/is/18-010r7/18-010r7.html#1\" > the <b>Scope</b> of the WKT definition </a> mentions the following : </p>" +
                "<p>The string defines frequently needed types of coordinate reference systems and coordinate operations in a self-contained form that is easily readable by machines and by humans. " +
                "The essence is its simplicity; as a consequence there are some constraints upon the more open content allowed in ISO 19111. " +
                "To retain simplicity in the well-known text (WKT) description of coordinate reference systems and coordinate operations, " +
                "the scope of this document excludes parameter grouping and pass-through coordinate operations. </p>" + 
                "<p>The text string provides a means for humans and machines to correctly and unambiguously interpret and utilise a coordinate reference system definition with look-ups or cross references " +
                "only to define coordinate operation mathematics. A WKT string is not suitable for the storage of definitions of coordinate reference systems or coordinate operations " +
                "because it omits metadata about the source of the data and may omit metadata about the applicability of the information.</p>" ,
            Example = "TL.cct.ApplyForward(4326, EPSG:32632, {12.0, 55.0}, 4) returns 691875.632")]
        public static object ApplyForward(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("input points", Name = "Coordinate(s)")] double[,] Coords,
            [ExcelArgument("Output mode: < 7 and flag: > 7. (0); 0 returns nr of input columns, 1 = (x, y, z, t), 2 = (x, y, z), 3 = (x, y), 4 = (x), 5 = (y), 6 = (z), 7 = (t). Check flag values 2^n in the help file", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
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
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
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
                        throw new NotImplementedException("error in switch statement"); 
                }
                return res;
                // end core of function

            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // ApplyForward

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.ApplyInverse",
            Description = "Inverse coordinate conversion of one or more input points", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1004",

            Returns = "The reprojected coordinate(s)",
            Summary =
            "<p>This function transforms coordinates from one Coordinate Reference system (CRS) into another CRS" +
            "<p>The transform can be performed using two different ways" +
            "<ol>    <li><p>By providing a seperate description for the SourceCrs as well as the TargetCrs</li>" +
                    "<li><p>By providing one string that describes the CRS-transform, leaving TargetCrs empty</li>" +
            "</ol>" +
            "<p>As such, SourceCrs and TargetCrs can be provided in one out of three ways" +
            "<ol>    <li><p>As a number referencing a CRS CODE from the EPSG database (much preferred)</li>" +
                    "<li><p>As a string using WKT, JSON or PROJ format. WKT or JSON format is preferred over the original PROJ string format</li>" +
                    "<li><p>As an AUTHORITY string in one cell, combined with a CRS CODE in the adjacent cell to the right</li>" +
            "</ol>" +
            "<p>This function is an array function. Array functions have undergone a significant upgrade with the introduction of dynamic arrays in Excel." +
            "<p>For more information on working with array formulas please consult :" +
            "<ol>" +
            "<li><p>This link: <a href = \"https://support.office.com/en-us/article/Guidelines-and-examples-of-array-formulas-7d94a64e-3ff3-4686-9372-ecfd5caa57c7\" > Microsoft Office Support - Guidelines and examples of array formulas</a> for Guidelines and examples of array formulas.</li>" +
            "<li><p>This link: <a href = \"https://support.office.com/en-us/article/Create-an-array-formula-e43e12e0-afc6-4a12-bc7f-48361075954d\" > Microsoft Office Support - Create a array formula</a> for more information on how to create static {CSE} array formulas.</li>" +
            "<li><p>This link: <a href = \"https://exceljet.net/dynamic-array-formulas-in-excel\" > ExcelJet - Dynamic array formulas in Excel</a> for an introduction to dynamic array formulas.</li>" +
            "</ol>" +
            "<p>For more information on coordinate conversion and coordinate refence system (CRS) information, see :" +
            "<ol>    <li><p>This link: <a href = \"http://spatialreference.org/\"> Spatial Reference home page</a></li>" +
                    "<li><p>This link: <a href = \"http://epsg.io/\" id=\"viewDesktopLink\"> EPSG IO home page with CRS description strings and EPSG numbers</a></li>" +
                    "<li><p>This link: <a href = \"http://proj.org/\"> Home page of the proj library</a></li>" +
            "</ol>",
            Remarks = "<p>Internally the transform uses <a href = \"https://proj.org/development/reference/functions.html#c.proj_normalize_for_visualization\"> crs normalization</a> by the proj library for a consistent approach to (x, y, z) values." +
            "<p>The axis order of a geographic CRS shall therefore be longitude, latitude [,height], and that of a projected CRS shall be easting, northing [, height]" +
            "<p>When using a geographic CRS, coordinates should be presented in degrees (not radians)." +
            "<p>The 'Mode' flag combines values 0 - 7 for the output mode with binary flags to reduce the number of parameters in this function. These binary flags are:" +
            "<pre><ul><li>   8: Disallow Ballpark Conversions</li>" +
                    "<li>  16: Don't Discard Transform if Grid is missing</li>" +
                    "<li>  32: Use Primary Grid Names</li>" +
                    "<li>  64: Use Superseded Transforms</li>" +
                    "<li> 128: Allow Deprecated CRSs</li>" +
                    "<li> 256: Transform strictly contains Area of Interest</li>" +
                    "<li> 512: Always Allow an Intermediate CRS</li>" +
                    "<li>1024: Never Allow an Intermediate CRS</li>" +
            "</ul></pre><br>" +
            "<p>Finally, please note the following aspects:</p>" +
            "<ol><li>some transforms require the use of one or more grid(s). Local/network access to these grids is controlled through the TopoLib Ribbon. </li>" +
                "<li>For all TL.cct-functions, the combined settings in the 'Mode Flag' can be overruled by global settings defined in the 'Transform Settings' Dialog. </li>" +
                "<li>Though it is possible to define a transform using a <b>single</b> (Wkt/Json/Proj) <b>string</b>, it is much preferred to apply <b>sourceCrs</b> and <b>targetCrs</b> to define a transform. " +
                "In that case the PROJ library automatically finds the most suitable transform.</li></ol>" +
                "<p>With respect to point 3 above, <a href = \"https://docs.opengeospatial.org/is/18-010r7/18-010r7.html#1\" > the <b>Scope</b> of the WKT definition </a> mentions the following : </p>" +
                "<p>The string defines frequently needed types of coordinate reference systems and coordinate operations in a self-contained form that is easily readable by machines and by humans. " +
                "The essence is its simplicity; as a consequence there are some constraints upon the more open content allowed in ISO 19111. " +
                "To retain simplicity in the well-known text (WKT) description of coordinate reference systems and coordinate operations, " +
                "the scope of this document excludes parameter grouping and pass-through coordinate operations. </p>" + 
                "<p>The text string provides a means for humans and machines to correctly and unambiguously interpret and utilise a coordinate reference system definition with look-ups or cross references " +
                "only to define coordinate operation mathematics. A WKT string is not suitable for the storage of definitions of coordinate reference systems or coordinate operations " +
                "because it omits metadata about the source of the data and may omit metadata about the applicability of the information.</p>" ,
            Example = "TL.cct.ApplyInverse(4326, EPSG:32632, {691875.63, 6098907.83}, 4) returns 12.000")]
        public static object ApplyInverse(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("input points", Name = "Coordinate(s)")] double[,] Coords,
            [ExcelArgument("Output mode: < 7 and flag: > 7. (0); 0 returns nr of input columns, 1 = (x, y, z, t), 2 = (x, y, z), 3 = (x, y), 4 = (x), 5 = (y), 6 = (z), 7 = (t). Check flag values 2^n in the help file", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
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
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

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
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
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
                        throw new NotImplementedException("error in switch statement");
                }
                return res;
                // end core of function

            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // ApplyInverse

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.AsJsonString",
            Description = "Get the forward coordinate transform as a JSON-string", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1005",

            Returns = "JSON-string describing the forward coordinate transform",
            Summary = "Returns a JSON-string describing a forward coordinate transform",
            Example = "See TL.cct.CreateForward(), using option '2'",
            Remarks = "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag"
            )]
        public static object AsJsonString(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                ChooseCoordinateTransform transforms = transform as ChooseCoordinateTransform;
                if (transforms is null)
                {
                    return transform.AsProjJson();
                }
                else
                {
                    return transforms[0].AsProjJson();
                }
                // end core of function
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // AsJsonString

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.AsProjString",
            Description = "Get the forward coordinate transform as a PROJ-string", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1006",

            Returns = "PROJ-string describing the forward coordinate transform",
            Summary = "Returns a PROJ-string describing a forward coordinate transform",
            Example = "See TL.cct.CreateForward(), using option '0'",
            Remarks = "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag"
            )]
        public static object AsProjString(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                ChooseCoordinateTransform transforms = transform as ChooseCoordinateTransform;
                if (transforms is null)
                {
                    return transform.AsProjString();
                }
                else
                {
                    return transforms[0].AsProjString();
                }
                // end core of function
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // AsProjString

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.AsWktString",
            Description = "Get the forward coordinate transform as a WKT-string", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1007",

            Returns = "WKT-string describing the forward coordinate transform",
            Summary = "Returns a WellKnownText-string describing a forward coordinate transform",
            Example = "See TL.cct.CreateForward(), using option '1'",
            Remarks = "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag"
            )]
        public static object AsWktString(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                ChooseCoordinateTransform transforms = transform as ChooseCoordinateTransform;
                if (transforms is null)
                {
                    return transform.AsWellKnownText();
                }
                else
                {
                    return transforms[0].AsWellKnownText();
                }
                // end core of function
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // AsWktString

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.CelestialBodyName",
            Description = "Get the name of the celestial body belonging to the coordinate transform", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1008",

            Returns = "Name of the celestial body belonging to the coordinate transform",
            Summary = "Returns the name of the celestial body belonging to the coordinate transform",
            Example = "t.b.c.",
            Remarks = "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag"
            )]
        public static object CelestialBodyName(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                return transform.CelestialBodyName;
                // end core of function
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // CelestialBodyName

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.CreateForward",
            Description = "Create a string representation of the forward coordinate transform", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1009",

            Returns = "String representation of the forward coordinate transform",
            Summary = "Creates a string representation of the forward coordinate transform in one of three different formats" +
                      "<p>If there are multiplee transforms available (transform list) the first transform of the list will be used.</p>",
            Remarks = "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag",
            Example = "TL.cct.CreateForward(7843, 7665, 0) returns +proj=pipeline +step +proj=unitconvert +xy_in=deg +z_in=m +xy_out=rad +z_out=m +step +proj=cart +ellps=GRS80 +step +proj=helmert +x=0 +y=0 +z=0 +rx=0 +ry=0 +rz=0 +s=0 +dx=0 +dy=0 +dz=0 +drx=-0.00150379 +dry=-0.00118346 +drz=-0.00120716 +ds=0 +t_epoch=2020 +convention=coordinate_frame +step +inv +proj=cart +ellps=WGS84 +step +proj=unitconvert +xy_in=rad +z_in=m +xy_out=deg +z_out=m")]
        public static object CreateForward(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Output mode: (0); 0 = PROJ string, 1 = WKT string, 2 = JSON string. Mode is combined with 2^n flag: 8, 16, ..., 2048. See help file for more information", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Check specific input data
            int nOutput = nMode & 3;

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                ChooseCoordinateTransform transforms = transform as ChooseCoordinateTransform;
                if (transforms is null)
                {
                    switch(nOutput)
                    {
                        default:
                        case 0:
                            return transform.AsProjString();
                        case 1:
                            return transform.AsWellKnownText();
                        case 2:
                            return transform.AsProjJson();
                    }
                }
                else
                {
                    switch(nOutput)
                    {
                        default:
                        case 0:
                            return transforms[0].AsProjString();
                        case 1:
                            return transforms[0].AsWellKnownText();
                        case 2:
                            return transforms[0].AsProjJson();
                    }
                }
                // end core of function
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // CreateForward

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.CreateInverse",
            Description = "Create a string representation of the inverse coordinate transform", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1010",

            Returns = "String representation of the inverse coordinate transform",
            Summary = "Creates a string representation of the inverse coordinate transform in one of three different formats" +
                      "<p>If there are multiple transforms available (transform list) the first transform of the list will be used.</p>",
            Remarks = "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag",
            Example = "TL.cct.Transforms.CreateInverse(2393, 3067, 0) returns +proj=pipeline +step +inv +proj=utm +zone=35 +ellps=GRS80 +step +proj=push +v_3 +step +proj=cart +ellps=GRS80 +step +inv +proj=helmert +x=-96.062 +y=-82.428 +z=-121.753 +rx=-4.801 +ry=-0.345 +rz=1.376 +s=1.496 +convention=coordinate_frame +step +inv +proj=cart +ellps=intl +step +proj=pop +v_3 +step +proj=tmerc +lat_0=0 +lon_0=27 +k=1 +x_0=3500000 +y_0=0 +ellps=intl")]
        public static object CreateInverse(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Output mode: (0); 0 = PROJ string, 1 = WKT string, 2 = JSON string. Mode is combined with 2^n flag: 8, 16, ..., 2048. See help file for more information", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Check specific input data
            int nOutput = nMode & 3;

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                CoordinateTransform inverseTransform = transform.CreateInverse(pjContext);
                ChooseCoordinateTransform inverseTransforms = inverseTransform as ChooseCoordinateTransform;
                if (inverseTransforms is null)
                {
                    switch(nOutput)
                    {
                        default:
                        case 0:
                            return inverseTransform.AsProjString();
                        case 1:
                            return inverseTransform.AsWellKnownText();
                        case 2:
                            return inverseTransform.AsProjJson();
                    }
                }
                else
                {
                    switch(nOutput)
                    {
                        default:
                        case 0:
                            return inverseTransforms[0].AsProjString();
                        case 1:
                            return inverseTransforms[0].AsWellKnownText();
                        case 2:
                            return inverseTransforms[0].AsProjJson();
                    }
                }
                // end core of function
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // CreateInverse

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.Grids.Count",
            Description = "Nr of grids used in a transform", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1011",

            Returns = "Nr of grids used in a transform",
            Summary = "Function returns nr of grids used in a transform",
            Example = "TL.cct.Grids.Count(4289, 4258, 0) returns 1",
            Remarks = "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag"
            )]
        public static object Grids_Count(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                return (double) transform.GridUsages.Count;
                // end core of function

            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // Grids_Count


        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.Grids.FullName",
            Description = "Get the full name (path) of grid nr N, used in a coordinate transform", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1012",

            Returns = "Full name (path) of grid nr N, used in in a coordinate transform",
            Summary = "Returns the full name (path) of grid nr N, used  in a coordinate transform",
            Example = "TL.cct.Grids.FullName(4289, 4258, 0) returns C:\\Program Files\\QGIS 3.20.0\\share\\proj\\nl_nsgi_rdtrans2018.tif",
            Remarks = "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag"
            )]
        public static object Grids_FullName(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Zero based index of Grid list (0) ", Name = "Index")] object index,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Check specific input data
            int nIndex  = (int)Optional.Check(index , 0.0);

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                CoordinateTransformList steps = transform as CoordinateTransformList;
                if (transform.GridUsages.Count == 0)
                    return "N.A.";
                else
                    return transform.GridUsages[nIndex].FullName;
                // end core of function

            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // Grids_FullName

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.Grids.IsAvailable",
            Description = "Checks whether grid nr N, used in a coordinate transform is available", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1013",

            Returns = "TRUE if grid nr N, used in a coordinate transform is available",
            Summary = "Function returns TRUE if grid nr N, used in a coordinate transform is available; FALSE otherwise",
            Example = "TL.cct.Grids.IsAvailable(4289, 4258, 0) returns TRUE",
            Remarks = "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag"
            )]
        public static object Grids_IsAvailable(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Zero based index of Grid list (0) ", Name = "Index")] object index,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Check specific input data
            int nIndex  = (int)Optional.Check(index , 0.0);

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                CoordinateTransformList steps = transform as CoordinateTransformList;
                if (transform.GridUsages.Count == 0)
                    return "N.A.";
                else
                    return transform.GridUsages[nIndex].IsAvailable;
                // end core of function

            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // Grids_IsAvailable

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.Grids.Name",
            Description = "Get the name of grid nr N, used in a coordinate transform", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1014",

            Returns = "Name of grid nr N, used in in a coordinate transform",
            Summary = "Returns the name of grid nr N, used  in a coordinate transform",
            Example = "TL.cct.Grids.FullName(4289, 4258, 0) returns nl_nsgi_rdtrans2018.tif",
            Remarks = "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag"
            )]
        public static object Grids_Name(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Zero based index of Grid list (0) ", Name = "Index")] object index,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Check specific input data
            int nIndex  = (int)Optional.Check(index , 0.0);

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                CoordinateTransformList steps = transform as CoordinateTransformList;
                if (transform.GridUsages.Count == 0)
                    return "N.A.";
                else
                    return transform.GridUsages[nIndex].Name;
                // end core of function

            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // Grids_Name

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.HasBallParkTransformation",
            Category = "CCT - Coordinate Conversion and Transformation",
            Description = "Confirms whether the transform has a ballpark transformation",
            HelpTopic = "TopoLib-AddIn.chm!1015",

            Returns = "TRUE when the transform has a ballpark transformation; FALSE when not",
            Summary = "Function that confirms that the transform has a ballpark transformation",
            Example = "TL.cct.HasBallParkTransformation(2020, 7789, 128) returns TRUE",
            Remarks = "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag"
         )]
        public static object HasBallParkTransformation(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                return transform.HasBallParkTransformation;
                // end core of function

            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
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
            HelpTopic = "TopoLib-AddIn.chm!1016",

            Returns = "TRUE when the transform can be done in the reversed direction; FALSE when not",
            Summary = "Function that confirms that the transform can be done in the reversed direction",
            Example = "TL.cct.HasInverse(2008, 7789, 128) returns TRUE",
            Remarks = "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag"
         )]
        public static object HasInverse(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                return transform.HasInverse;
                // end core of function

            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
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
            Name = "TL.cct.Identifiers.Authority",
            Description = "Gets Authority of Identifier N", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1017",

            Returns = "Authority of Nth Identifier",
            Summary = "Function that returns Authority of <Nth> identifiers or <index out of range> when not found",
            Example = "t.b.c.",
            Remarks = "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag"
            )]
        public static object Identifiers_Authority(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Zero based index of Steps list (0) ", Name = "Index")] object index,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Check specific input data
            int nIndex  = (int)Optional.Check(index , 0.0);
            int nCount = -1;

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                nCount = (transform.Identifiers != null) ? transform.Identifiers.Count : 0;

                if (nIndex > nCount - 1 || nIndex < 0 || nCount == 0)
                    return "<index out of range>";

                return transform.Identifiers[nIndex].Authority;
                // end core of function

            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // Identifiers_Authority

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.Identifiers.Code",
            Description = "Gets Code of Identifier N", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1018",

            Returns = "Code of Nth Identifier",
            Summary = "Function that returns Code of <Nth> identifiers or <index out of range> when not found",
            Example = "t.b.c.",
            Remarks = "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag"
            )]
        public static object Identifiers_Code(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Zero based index of Steps list (0) ", Name = "Index")] object index,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Check specific input data
            int nIndex  = (int)Optional.Check(index , 0.0);
            int nCount = -1;

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                nCount = (transform.Identifiers != null) ? transform.Identifiers.Count : 0;

                if (nIndex > nCount - 1 || nIndex < 0 || nCount == 0)
                    return "<index out of range>";

                return transform.Identifiers[nIndex].Code;
                // end core of function
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // Identifiers_Code

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.Identifiers.Count",
            Description = "Get the number of identifiers used in a coordinate transform", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1019",

            Returns = "Number of identifiers used in a coordinate transform",
            Summary = "Returns the number of identifiers used in a coordinate transform",
            Example = "t.b.c.",
            Remarks = "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag"
            )]
        public static object Identifiers_Count(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                return (transform.Identifiers != null) ? transform.Identifiers.Count : 0;
                // end core of function
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // Identifiers_Count

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.IsAvailable",
            Category = "CCT - Coordinate Conversion and Transformation",
            Description = "Confirms whether the transform is available",
            HelpTopic = "TopoLib-AddIn.chm!1020",

            Returns = "TRUE when the transform is available; FALSE when not",
            Summary = "Function that confirms that a transform is available",
            Example = "TL.cct.IsAvailable(2008, 7789, 128) returns TRUE",
            Remarks = "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag"
         )]
        public static object IsAvailable(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                return transform.IsAvailable;
                // end core of function

            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
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
            HelpTopic = "TopoLib-AddIn.chm!1021",

            Returns = "Name of the coordinate transform",
            Summary = "Returns the method name a coordinate transform",
            Example = "TL.cct.MethodName(2008, 7789, 128) returns Unknown",
            Remarks = "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag"
            )]
        public static object MethodName(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                return transform.MethodName ?? "Unknown";
                // end core of function
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
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
            HelpTopic = "TopoLib-AddIn.chm!1022",

            Returns = "Name of the coordinate transform",
            Summary = "Returns the name a coordinate transform",
            Example = "TL.cct.Name(4326, 2007, 128) returns axis order change (2D) + Inverse of St. Vincent 1945 to WGS 84 (1) + British West Indies Grid",
            Remarks = "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag"
            )]
        public static object Name(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                return transform.Name;
                // end core of function
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
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
            Name = "TL.cct.Remarks",
            Description = "Get the name of the coordinate transform",
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1023",

            Returns = "Clarifying remarks of the coordinate transform",
            Summary = "Returns clarifying remarks of the coordinate transform",
            Remarks = "Most transforms have remarks in the individual steps, not in the overall transform definition" +
            "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag"
            )]
        public static object Remarks(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
            [ExcelArgument("WestLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "WestLongitude")] object oWestLongitude,
            [ExcelArgument("SouthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "SouthLatitude")] object oSouthLatitude,
            [ExcelArgument("EastLongitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "EastLongitude")] object oEastLongitude,
            [ExcelArgument("NorthLatitude of the desired area for the transformation, or '-1000' when not used (-1000)", Name = "NorthLatitude")] object oNorthLatitude)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            // Check general input data
            int nMode = (int)Optional.Check(oMode, 0.0);
            bool bUsingTransform = Optional.IsNul(TargetCrs);
            double Accuracy = Optional.Check(oAccuracy, -1000.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                return transform.Remarks;
                // end core of function

            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // Remarks


        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.RoundTrip",
            Description = "Get the error of a roundtrip of N forward/backward transforms", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1024",

            Returns = "Amount of error [m] incurred in the roundtrip(s)",
            Summary = "Returns error incurred in N forward roundtrip(s) in a coordinate transform",
            Remarks = "For the test point, it is recommended to select the centerpoint of the usage area of the Source CRS." +
            "<p>If no test point is given (0, 0, 0) will be used instead" + 
            "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag",
            Example = "TL.cct.RoundTrip(2004, 3857, {380275.8,	1851010.7}, 2, 128) returns 0.0047"
            )]
        public static object RoundTrip(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("test point with adjacent [x, y] coordinates", Name = "point(x, y)")] object[,] TestCoord,
            [ExcelArgument("N - nr of roundtrips to make", Name = "nr roundtrips")] object oRoundTrips,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
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
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                double error = transform.RoundTrip(true, nTrips, pt);

                if (Double.IsInfinity(error))
                    throw new NotImplementedException("Infinite roundtrip error");

                return error;
                // end core of function

            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
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
            Name = "TL.cct.Scope",
            Description = "Get the scope of the coordinate transform", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1025",

            Returns = "Scope of the coordinate transform, if known",
            Summary = "Returns the scope of a coordinate transform, if known",
            Example = "Mainly used in multi-step transforms; Unknown in single-step transform",
            Remarks = "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag"
            )]
        public static object Scope(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                return transform.Scope ?? "Unknown";    
                // end core of function

            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // Scope

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.SourceCRS",
            Description = "Get the source-CRS used in a coordinate transform", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1026",

            Returns = "Source-CRS of a coordinate transform in one of three different formats",
            Summary = "Returns the source-CRS of a coordinate transform in one of three different formats",
            Example = "TL.cct.SourceCRS(2002, 7789,$Z$5) returns +proj=tmerc +lat_0=0 +lon_0=-62 +k=0.9995 +x_0=400000 +y_0=0 +a=6378249.145 +rf=293.465 +units=m +no_defs +type=crs",
            Remarks = "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag"
            )]
        public static object SourceCRS(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Output mode: (0); 0 = PROJ string, 1 = WKT string, 2 = JSON string. Mode is combined with 2^n flag: 8, 16, ..., 2048. See help file for more information", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Check specific input data
            int nOutput = nMode & 3;

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                CoordinateReferenceSystem SourceCRS = transform.SourceCRS;
                if (SourceCRS is null)
                    return "Unknown";

                switch(nOutput)
                {
                    default:
                    case 0:
                        return SourceCRS.AsProjString();
                    case 1:
                        return SourceCRS.AsWellKnownText();
                    case 2:
                        return SourceCRS.AsProjJson();
                }
                // end core of function

            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // SourceCRS

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.Steps.Count",
            Description = "Get the number of steps incorporated in a coordinate transform", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1027",

            Returns = "Number of steps incorporated in a coordinate transform",
            Summary = "Returns the number of steps incorporated in a coordinate transform",
            Example = "TL.cct.Steps.Count(2002, 4326, 0) returns 3",
            Remarks = "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag"
            )]
        public static object Steps_Count(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                CoordinateTransformList steps = transform as CoordinateTransformList;
                return (steps is null) ? 0 : steps.Count;
                // end core of function

            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
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
            Name = "TL.cct.Steps.CreateForward",
            Description = "Creates a string representation of the forward transform for step N in a coordinate transform", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1028",

            Returns = "String of the forward transform for step N in a coordinate transform",
            Summary = "Creates a string representation of the forward transform for step N in a coordinate transform in one of three different formats",
            Remarks = "Every step in a multi-step coordinate transform, is basically a coordinate transform on its own" +
            "<p>By making the transform string available for each step, we prevent having to implement (duplicate) all transform functionality for 'TL.cct.Steps'." +
            "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag",
            Example = "TL.cct.Steps.Transform(2002, 4326, 0, 0) returns +proj=pipeline +step +inv +proj=tmerc +lat_0=0 +lon_0=-62 +k=0.9995 +x_0=400000 +y_0=0 +a=6378249.145 +rf=293.465 +step +proj=unitconvert +xy_in=rad +xy_out=deg +step +proj=axisswap +order=2,1")]
        public static object Steps_CreateForward(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Zero based index of Steps list (0) ", Name = "Index")] object index,
            [ExcelArgument("Output mode: (0); 0 = PROJ string, 1 = WKT string, 2 = JSON string. Mode is combined with 2^n flag: 8, 16, ..., 2048. See help file for more information", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Check specific input data
            int nIndex  = (int)Optional.Check(index , 0.0);
            int nOutput = nMode & 3;

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                CoordinateTransformList steps = transform as CoordinateTransformList;
                if (steps is null) 
                    return "N.A.";

                switch(nOutput)
                {
                    default:
                    case 0:
                        return steps[nIndex].AsProjString();
                    case 1:
                        return steps[nIndex].AsWellKnownText();
                    case 2:
                        return steps[nIndex].AsProjJson();
                }
                // end core of function

            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // Steps_CreateForward

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.Steps.CreateInverse",
            Description = "Creates a string representation of the inverse transform for step N in a coordinate transform", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1029",

            Returns = "String of the inverse transform for step N in a coordinate transform",
            Summary = "Creates a string representation of the inverse transform for step N in a coordinate transform in one of three different formats",
            Remarks = "Every step in a multi-step coordinate transform, is basically a coordinate transform on its own" +
            "<p>By making the transform string available for each step, we prevent having to implement (duplicate) all transform functionality for 'TL.cct.Steps'." +
            "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag",
            Example = "TL.cct.Steps.CreateInverse(25832, 25833, 0, 0) returns +proj=pipeline +step +proj=axisswap +order=2,1 +step +proj=unitconvert +xy_in=deg +xy_out=rad +step +proj=utm +zone=32 +ellps=GRS80")]
        public static object Steps_CreateInverse(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Zero based index of Steps list (0) ", Name = "Index")] object index,
            [ExcelArgument("Output mode: (0); 0 = PROJ string, 1 = WKT string, 2 = JSON string. Mode is combined with 2^n flag: 8, 16, ..., 2048. See help file for more information", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Check specific input data
            int nIndex  = (int)Optional.Check(index , 0.0);
            int nOutput = nMode & 3;

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                CoordinateTransformList steps = transform as CoordinateTransformList;
                if (steps is null) 
                    return "N.A.";

                CoordinateTransform step = steps[nIndex];
                CoordinateTransform pets = step.CreateInverse(pjContext);

                switch(nOutput)
                {
                    default:
                    case 0:
                        return pets.AsProjString();
                    case 1:
                        return pets.AsWellKnownText();
                    case 2:
                        return pets.AsProjJson();
                }
                // end core of function

            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // Steps_CreateInverse

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.Steps.MethodName",
            Description = "Get the method name of step N in a coordinate transform", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1030",

            Returns = "Method name of step N in a coordinate transform",
            Summary = "Returns the method name of step N in a coordinate transform",
            Example = "TL.cct.Steps.MethodName(2002, 4326, 0, 0) returns Inverse of Transverse Mercator",
            Remarks = "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag"
            )]
        public static object Steps_MethodName(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Zero based index of Steps list (0) ", Name = "Index")] object index,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Check specific input data
            int nIndex  = (int)Optional.Check(index , 0.0);

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                CoordinateTransformList steps = transform as CoordinateTransformList;
                return (steps is null) ? "N.A." : steps[nIndex].MethodName;
                // end core of function

            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
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
            HelpTopic = "TopoLib-AddIn.chm!1031",

            Returns = "Name of step N in a coordinate transform",
            Summary = "Returns the name of step N in a coordinate transform",
            Example = "TL.cct.Steps.Name(2002, 4326, 0, 0) returns Inverse of British West Indies Grid",
            Remarks = "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag"
            )]
        public static object Steps_Name(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Zero based index of Steps list (0) ", Name = "Index")] object index,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Check specific input data
            int nIndex  = (int)Optional.Check(index , 0.0);

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                CoordinateTransformList steps = transform as CoordinateTransformList;
                return (steps is null) ? "N.A." : steps[nIndex].Name;
                // end core of function

            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
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
            Name = "TL.cct.TargetCRS",
            Description = "Get the target-CRS used in a coordinate transform", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1032",

            Returns = "Target-CRS of a coordinate transform in one of three different formats",
            Summary = "Returns the target-CRS of a coordinate transform in one of three different formats",
            Example = "TL.cct.TargetCRS(2002, 4326, 0, 0) returns +proj=tmerc +lat_0=0 +lon_0=-62 +k=0.9995 +x_0=400000 +y_0=0 +a=6378249.145 +rf=293.465 +units=m +no_defs +type=crs",
            Remarks = "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag"
            )]
        public static object TargetCRS(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Output mode: (0); 0 = PROJ string, 1 = WKT string, 2 = JSON string. Mode is combined with 2^n flag: 8, 16, ..., 2048. See help file for more information", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Check specific input data
            int nOutput = nMode & 3;

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS /*, ref bUseNetwork */);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS /*, bUseNetwork*/);
                }

                // start core of function
                CoordinateReferenceSystem SourceCRS = transform.SourceCRS;
                if (SourceCRS is null)
                    return "Unknown";

                switch(nOutput)
                {
                    default:
                    case 0:
                        return SourceCRS.AsProjString();
                    case 1:
                        return SourceCRS.AsWellKnownText();
                    case 2:
                        return SourceCRS.AsProjJson();
                }
                // end core of function

            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // TargetCRS

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.Transforms.Count",
            Description = "Get the number of available transforms",
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1033",

            Returns = "The number of available transforms",
            Summary = "Returns the number of available transforms that exist between two Coordinate Reference Systems",
            Example = "TL.cct.Transforms.Count(2200, 3875, 0) returns 3",
            Remarks = "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag"
            )]
        public static object Transforms_Count(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                ChooseCoordinateTransform transforms = transform as ChooseCoordinateTransform;
                return (transforms != null) ? transforms.Count : 1;
                // end core of function

            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
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
            Name = "TL.cct.Transforms.CreateForward",
            Description = "Creates a string representation of the forward transform for transform N in a coordinate transform list", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1034",

            Returns = "String of the forward transform for transform N in a coordinate transform list",
            Summary = "Creates a string representation of the forward transform for transform N in a coordinate transform list in one of three different formats",
            Remarks = "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag",
            Example = "TL.cct.Transforms.CreateForward(2393, 3067) returns +proj=pipeline +step +inv +proj=tmerc +lat_0=0 +lon_0=27 +k=1 +x_0=3500000 +y_0=0 +ellps=intl +step +proj=push +v_3 +step +proj=cart +ellps=intl +step +proj=helmert +x=-96.062 +y=-82.428 +z=-121.753 +rx=-4.801 +ry=-0.345 +rz=1.376 +s=1.496 +convention=coordinate_frame +step +inv +proj=cart +ellps=GRS80 +step +proj=pop +v_3 +step +proj=utm +zone=35 +ellps=GRS80")]
        public static object Transforms_CreateForward(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Zero based index of available transforms (0) ", Name = "Index")] object index,
            [ExcelArgument("Output mode: (0); 0 = PROJ string, 1 = WKT string, 2 = JSON string. Mode is combined with 2^n flag: 8, 16, ..., 2048. See help file for more information", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Check specific input data
            int nIndex  = (int)Optional.Check(index , 0.0);
            int nOutput = nMode & 3;

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                ChooseCoordinateTransform transforms = transform as ChooseCoordinateTransform;
                if (transforms is null) 
                    return "N.A.";

                switch(nOutput)
                {
                    default:
                    case 0:
                        return transforms[nIndex].AsProjString();
                    case 1:
                        return transforms[nIndex].AsWellKnownText();
                    case 2:
                        return transforms[nIndex].AsProjJson();
                }
                // end core of function
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // Transforms_CreateForward

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.Transforms.CreateInverse",
            Description = "Creates a string representation of the inverse transform for transform N in a coordinate transform list", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1035",

            Returns = "String of the inverse transform for transform N in a coordinate transform list",
            Summary = "Creates a string representation of the inverse transform for transform N in a coordinate transform list, in one of three different formats",
            Remarks = "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag",
            Example = "TL.cct.Transforms.CreateInverse(2393, 3067) returns +proj=pipeline +step +inv +proj=utm +zone=35 +ellps=GRS80 +step +proj=push +v_3 +step +proj=cart +ellps=GRS80 +step +inv +proj=helmert +x=-96.062 +y=-82.428 +z=-121.753 +rx=-4.801 +ry=-0.345 +rz=1.376 +s=1.496 +convention=coordinate_frame +step +inv +proj=cart +ellps=intl +step +proj=pop +v_3 +step +proj=tmerc +lat_0=0 +lon_0=27 +k=1 +x_0=3500000 +y_0=0 +ellps=intl")]
        public static object Transforms_CreateInverse(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Zero based index of available transforms (0) ", Name = "Index")] object index,
            [ExcelArgument("Output mode: (0); 0 = PROJ string, 1 = WKT string, 2 = JSON string. Mode is combined with 2^n flag: 8, 16, ..., 2048. See help file for more information", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Check specific input data
            int nIndex  = (int)Optional.Check(index , 0.0);
            int nOutput = nMode & 3;

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                ChooseCoordinateTransform transforms = transform as ChooseCoordinateTransform;
                if (transforms is null) 
                    return "N.A.";

                CoordinateTransform transform1 = transforms[nIndex];
                CoordinateTransform transform2 = transform1.CreateInverse(pjContext);

                switch(nOutput)
                {
                    default:
                    case 0:
                        return transform2.AsProjString();
                    case 1:
                        return transform2.AsWellKnownText();
                    case 2:
                        return transform2.AsProjJson();
                }
                // end core of function

            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // Transforms_CreateInverse

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.Transforms.ListAll",
            Description = "Get the number of available transforms",
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1036",

            Returns = "The number of available transforms and associated information",
            Summary = "Lists the number of available transforms that exist between two Coordinate Reference Systems",
            Remarks = "Please be aware that this function is an array function that includes a header row to describe the return values." +
            "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag"
            )]
        public static object Transforms_ListAll(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Output mode: (0); 0 = PROJ string, 1 = WKT string, 2 = JSON string. Mode is combined with 2^n flag: 8, 16, ..., 2048. See help file for more information", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Check specific input data
            int nOutput = nMode & 3;

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                
                ChooseCoordinateTransform transforms = transform as ChooseCoordinateTransform;
                CoordinateTransformList list = transform as CoordinateTransformList;
                int count = (transforms != null) ? transforms.Count : 1;

                object[,] res = new object[count + 1, 14];

                res[0,  0] = "Transform Identifiers";
                res[0,  1] = "Transform Name";
                res[0,  2] = "Accuracy [m]";
                res[0,  3] = "Area of Use";
                res[0,  4] = "Scope";
                res[0,  5] = "Remarks";
                res[0,  6] = "Lon-min";
                res[0,  7] = "Lon-max";
                res[0,  8] = "Lat-min";
                res[0,  9] = "Lat-max";
                res[0, 10] = "Nr grids";
                res[0, 11] = "1st grid";
                res[0, 12] = "2nd grid";

                switch(nOutput)
                {
                    default:
                    case 0:
                        res[0, 13] = "Transform PROJ string";
                        break;
                    case 1:
                        res[0, 13] = "Transform WKT string";
                        break;
                    case 2:
                        res[0, 13] = "Transform JSON string";
                        break;
                }

                object accuracy; 

                if (count == 1)
                {
                    string ids = transform.Identifier?.ToString();
                    string scopes = transform.Scope;
                    string remarks = transform.Remarks;
                    if (ids == null && transform is CoordinateTransformList ctl)
                    {
                        ids = string.Join(", ", ctl.Where(x => x.Identifiers != null).SelectMany(x => x.Identifiers));
                        scopes = string.Join("  ", ctl.Select(x => x.Scope).Where(x => x != null).Distinct());
                        remarks = string.Join("  ", ctl.Select(x => x.Remarks).Where(x => x != null).Distinct());
                        remarks = remarks.TrimStart();
                    }

                    if (transform.Accuracy is null || transform.Accuracy <= 0.0) 
                    { 
                        accuracy = "Unknown"; 
                    } 
                    else
                    {
                        accuracy = transform.Accuracy;
                    }

                    string grid1 = "N.A."; 
                    string grid2 = "N.A."; 
                    if (transform.GridUsages.Count > 0)
                    {
                        grid1  = transform.GridUsages[0].FullName;
                        if (string.IsNullOrEmpty(grid1))
                            grid1 = "Missing";
                    }

                    if (transform.GridUsages.Count > 1)
                    {
                        grid2  = transform.GridUsages[1].FullName;
                        if (string.IsNullOrEmpty(grid2))
                            grid2 = "Missing";
                    }

                    res[1,  0] = transform.Identifiers is null ? ids : transform.Identifier.Authority;;
                    res[1,  1] = transform.Name;
                    res[1,  2] = accuracy;
                    res[1,  3] = transform.UsageArea.Name;
                    res[1,  4] = scopes;
                    res[1,  5] = remarks;
                    res[1,  6] = transform.UsageArea.WestLongitude;
                    res[1,  7] = transform.UsageArea.EastLongitude;
                    res[1,  8] = transform.UsageArea.SouthLatitude;
                    res[1,  9] = transform.UsageArea.NorthLatitude;
                    res[1, 10] = transform.GridUsages.Count;
                    res[1, 11] = grid1;
                    res[1, 12] = grid2;

                    switch(nOutput)
                    {
                        default:
                        case 0:
                            res[1, 13] = transform.AsProjString();
                            break;
                        case 1:
                            res[1, 13] = transform.AsWellKnownText();
                            break;
                        case 2:
                            res[1, 13] = transform.AsProjJson();
                            break;
                    }
                }
                else
                {
                    for (int i = 0; i < count; i++)
                    {
                        string ids = transforms[i].Identifier?.ToString();
                        string scopes = transforms[i].Scope;
                        string remarks = transforms[i].Remarks;
                        if (ids == null && transforms[i] is CoordinateTransformList ctl)
                        {
                            ids = string.Join(", ", ctl.Where(x => x.Identifiers != null).SelectMany(x => x.Identifiers));
                            scopes = string.Join("  ", ctl.Select(x => x.Scope).Where(x => x != null).Distinct());
                            remarks = string.Join("  ", ctl.Select(x => x.Remarks).Where(x => x != null).Distinct());
                            remarks = remarks.TrimStart();
                        }

                         if (transforms[i].Accuracy is null || transforms[i].Accuracy <= 0.0)
                        { 
                            accuracy = "Unknown"; 
                        } 
                        else 
                        { 
                            accuracy = transforms[i].Accuracy;
                        }

                        string grid1 = "N.A.";
                        string grid2 = "N.A.";

                        if (transforms[i].GridUsages.Count > 0)
                        {
                            grid1  = transforms[i].GridUsages[0].FullName;
                            if (string.IsNullOrEmpty(grid1))
                                grid1 = "Missing";
                        }

                        if (transforms[i].GridUsages.Count > 1)
                        {
                            grid2  = transforms[i].GridUsages[1].FullName;
                            if (string.IsNullOrEmpty(grid2))
                                grid2 = "Missing";
                        }

                        res[i + 1,  0] = transforms[i].Identifiers is null ? ids : transforms[i].Identifier.Authority;
                        res[i + 1,  1] = transforms[i].Name;
                        res[i + 1,  2] = accuracy;
                        res[i + 1,  3] = transforms[i].UsageArea.Name;
                        res[i + 1,  4] = scopes;
                        res[i + 1,  5] = remarks;
                        res[i + 1,  6] = transforms[i].UsageArea.WestLongitude;
                        res[i + 1,  7] = transforms[i].UsageArea.EastLongitude;
                        res[i + 1,  8] = transforms[i].UsageArea.SouthLatitude;
                        res[i + 1,  9] = transforms[i].UsageArea.NorthLatitude;
                        res[i + 1, 10] = transforms[i].GridUsages.Count;
                        res[i + 1, 11] = grid1;
                        res[i + 1, 12] = grid2;

                        switch(nOutput)
                        {
                            default:
                            case 0:
                                res[i + 1, 13] = transforms[i].AsProjString();
                                break;
                            case 1:
                                res[i + 1, 13] = transforms[i].AsWellKnownText();
                                break;
                            case 2:
                                res[i + 1, 13] = transforms[i].AsProjJson();
                                break;
                        }
                    }
                }
                return res;
                // end core of function

            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
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

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.Type",
            Description = "Get the typeof the coordinate transform", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1037",

            Returns = "Type of the coordinate transform",
            Summary = "Returns the type a coordinate transform",
            Example = "TL.cct.Type(+proj=merc +ellps=clrk66 +lat_ts=33) returns OtherCoordinateTransform",
            Remarks = "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag"
            )]
        public static object Type(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                return transform.Type.ToString();
                // end core of function

            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // Type

        // UsageArea needs to be expanded see TL.crs.UsageArea

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.UsageArea.Center",
            Description = "Get the Center Point of the Usage Area of the coordinate transform", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1038",

            Returns = "Center Point of Usage Area of the coordinate transform",
            Summary = "Returns the Center Point of the Usage Area of a coordinate transform",
            Example = "t.b.c.",
            Remarks = "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag"
            )]
        public static object UsageArea_Center(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();
                object[,] res = new object [1, 2];

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                UsageArea Area = transform.UsageArea;
                if (Area != null)
                {
                    res[0, 0] = Area.CenterX;
                    res[0, 1] = Area.CenterY;
                    return res;
                }
                else 
                    throw new ArgumentException("No Center Point found");  // Will return #VALUE to Excel

                // end core of function
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // UsageArea_Center

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.UsageArea.Center.X",
            Description = "Get the x-value of the Center Point of the Usage Area of the coordinate transform", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1039",

            Returns = "X-value of the Center Point of the Usage Area of the coordinate transform",
            Summary = "Returns the x-value of the Center Point of the Usage Area of a coordinate transform",
            Example = "t.b.c.",
            Remarks = "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag"
            )]
        public static object UsageArea_Center_X(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                UsageArea Area = transform.UsageArea;
                if (Area != null)
                    return Area.CenterX;
                else 
                    throw new ArgumentException("No Center Point found");  // Will return #VALUE to Excel

                // end core of function
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // UsageArea_Center_X

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.UsageArea.MaxX",
            Description = "Get the maximum X-value of the Usage Area of the coordinate transform", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1040",

            Returns = "Maximum X-value of the Usage Area of the coordinate transform",
            Summary = "Returns the maximum X-value of the Usage Area of a coordinate transform",
            Example = "t.b.c.",
            Remarks = "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag"
            )]
        public static object UsageArea_MaxX(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                UsageArea Area = transform.UsageArea;
                if (Area != null)
                    return Area.MaxX;
                else 
                    throw new ArgumentException("MaxX not found");  // Will return #VALUE to Excel

                // end core of function
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // UsageArea_MaxX

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.UsageArea.MaxY",
            Description = "Get the maximum Y-value of the Usage Area of the coordinate transform", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1041",

            Returns = "Maximum Y-value of the Usage Area of the coordinate transform",
            Summary = "Returns the maximum Y-value of the Usage Area of a coordinate transform",
            Example = "t.b.c.",
            Remarks = "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag"
            )]
        public static object UsageArea_MaxY(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                UsageArea Area = transform.UsageArea;
                if (Area != null)
                    return Area.MaxY;
                else 
                    throw new ArgumentException("MaxY not found");  // Will return #VALUE to Excel

                // end core of function
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // UsageArea_MaxY

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.UsageArea.MinX",
            Description = "Get the minimum X-value of the Usage Area of the coordinate transform", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1042",

            Returns = "Minimum X-value of the Usage Area of the coordinate transform",
            Summary = "Returns the minimum X-value of the Usage Area of a coordinate transform",
            Example = "t.b.c.",
            Remarks = "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag"
            )]
        public static object UsageArea_MinX(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                UsageArea Area = transform.UsageArea;
                if (Area != null)
                    return Area.MinX;
                else 
                    throw new ArgumentException("MinX not found");  // Will return #VALUE to Excel

                // end core of function
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // UsageArea_MinX

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.UsageArea.MinY",
            Description = "Get the minimum Y-value of the Usage Area of the coordinate transform", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1043",

            Returns = "Minimum Y-value of the Usage Area of the coordinate transform",
            Summary = "Returns the minimum Y-value of the Usage Area of a coordinate transform",
            Example = "t.b.c.",
            Remarks = "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag"
            )]
        public static object UsageArea_MinY(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                UsageArea Area = transform.UsageArea;
                if (Area != null)
                    return Area.MinY;
                else 
                    throw new ArgumentException("MinY not found");  // Will return #VALUE to Excel

                // end core of function
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // UsageArea_MinY

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.UsageArea.Center.Y",
            Description = "Get the y-value of the Center Point of the Usage Area of the coordinate transform", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1044",

            Returns = "Y-value of the Center Point of the Usage Area of the coordinate transform",
            Summary = "Returns the y-value of the Center Point of the Usage Area of a coordinate transform",
            Example = "t.b.c.",
            Remarks = "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag"
            )]
        public static object UsageArea_Center_Y(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                UsageArea Area = transform.UsageArea;
                if (Area != null)
                    return Area.CenterY;
                else 
                    throw new ArgumentException("No Center Point found");  // Will return #VALUE to Excel

                // end core of function
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // UsageArea_Center_Y

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.UsageArea.Name",
            Description = "Get the Name of the Usage Area of the coordinate transform", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1045",

            Returns = "Name of the Usage Area of the coordinate transform",
            Summary = "Returns the Name of the Usage Area of a coordinate transform",
            Example = "t.b.c.",
            Remarks = "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag"
            )]
        public static object UsageArea_Name(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                UsageArea Area = transform.UsageArea;
                if (Area != null)
                    return Area.Name;
                else 
                    throw new ArgumentException("No Name found");  // Will return #VALUE to Excel

                // end core of function
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // UsageArea_Name

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.UsageArea.WestLongitude",
            Description = "Gets the West Longitude of the Usage Area of the coordinate transform", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1046",

            Returns = "West Longitude of the Usage Area of the coordinate transform",
            Summary = "Returns the West Longitude of the Usage Area of a coordinate transform",
            Example = "t.b.c.",
            Remarks = "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag"
            )]
        public static object UsageArea_WestLongitude(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                UsageArea Area = transform.UsageArea;
                if (Area != null)
                    return Area.WestLongitude;
                else 
                    throw new ArgumentException("WestLongitude not found");  // Will return #VALUE to Excel

                // end core of function
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // UsageArea_WestLongitude

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.UsageArea.EastLongitude",
            Description = "Gets the East Longitude of the Usage Area of the coordinate transform", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1047",

            Returns = "East Longitude of the Usage Area of the coordinate transform",
            Summary = "Returns the East Longitude of the Usage Area of a coordinate transform",
            Example = "t.b.c.",
            Remarks = "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag"
            )]
        public static object UsageArea_EastLongitude(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                UsageArea Area = transform.UsageArea;
                if (Area != null)
                    return Area.EastLongitude;
                else 
                    throw new ArgumentException("EastLongitude not found");  // Will return #VALUE to Excel

                // end core of function
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // UsageArea_EastLongitude

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.UsageArea.SouthLatitude",
            Description = "Gets the South Latitude of the Usage Area of the coordinate transform", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1048",

            Returns = "South Latitude of the Usage Area of the coordinate transform",
            Summary = "Returns the South Latitude of the Usage Area of a coordinate transform",
            Example = "t.b.c.",
            Remarks = "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag"
            )]
        public static object UsageArea_SouthLatitude(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                UsageArea Area = transform.UsageArea;
                if (Area != null)
                    return Area.SouthLatitude;
                else 
                    throw new ArgumentException("SouthLatitude not found");  // Will return #VALUE to Excel

                // end core of function
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // UsageArea_SouthLatitude

        [ExcelFunctionDoc(
            IsThreadSafe = true, // this should speed things up, and should be fine, as the ProjContext is created locally in the function....
            Name = "TL.cct.UsageArea.NorthLatitude",
            Description = "Gets the North Latitude of the Usage Area of the coordinate transform", 
            Category = "CCT - Coordinate Conversion and Transformation",
            HelpTopic = "TopoLib-AddIn.chm!1049",

            Returns = "North Latitude of the Usage Area of the coordinate transform",
            Summary = "Returns the North Latitude of the Usage Area of a coordinate transform",
            Example = "t.b.c.",
            Remarks = "Please consult the remarks at <a href = \"TL.cct.ApplyForward.htm\"> <b>TL.cct.ApplyForward</b> </a>for the details on the <b>Mode</b> flag"
            )]
        public static object UsageArea_NorthLatitude(
            [ExcelArgument("sourceCrs (or transform) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "sourceCrsOrTransform")] object[,] SourceCrs,
            [ExcelArgument("targetCrs (or nul/empty) using one [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "targetCrsOrNul")] object[,] TargetCrs,
            [ExcelArgument("Binary flag: 8, 16, 32, ..., 2048. Check the help file for the details", Name = "mode")] object oMode,
            [ExcelArgument("Desired accuray for the transformation, or '-1' when not used (-1)", Name = "Accuracy")] object oAccuracy,
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
            double Accuracy      = Optional.Check(oAccuracy,      -1.0);
            double westLongitude = Optional.Check(oWestLongitude, -1000.0);
            double southLatitude = Optional.Check(oSouthLatitude, -1000.0);
            double eastLongitude = Optional.Check(oEastLongitude, -1000.0);
            double northLatitude = Optional.Check(oNorthLatitude, -1000.0);

            if (nMode < 0 || nMode > 2048)
                return ExcelError.ExcelErrorValue;

            // Deal with optional parameters
            bool bAllowDeprecatedCRS = false;
            var options = GetCoordinateTransformOptions(nMode, Accuracy, westLongitude, southLatitude, eastLongitude, northLatitude, ref bAllowDeprecatedCRS);

            // setup all disposable objects
            ProjContext               pjContext = null;
            CoordinateReferenceSystem crsSource = null;
            CoordinateReferenceSystem crsTarget = null;
            CoordinateTransform       transform = null;

            // do the work; exceptions may occur...
            try
            {
                pjContext = Crs.CreateContext();

                if (bUsingTransform)
                {
                    transform = CreateCoordinateTransform(SourceCrs, pjContext);
                }
                else
                {
                    crsSource = Crs.CreateCrs(SourceCrs, pjContext);
                    crsTarget = Crs.CreateCrs(TargetCrs, pjContext);
                    transform = CreateCoordinateTransform(crsSource, crsTarget, options, pjContext, bAllowDeprecatedCRS);
                }

                // start core of function
                UsageArea Area = transform.UsageArea;
                if (Area != null)
                    return Area.NorthLatitude;
                else 
                    throw new ArgumentException("NorthLatitude not found");  // Will return #VALUE to Excel

                // end core of function
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
            finally
            {
                // free up resources in reverse order of allocation
                transform?.Dispose();
                crsSource?.Dispose();
                crsTarget?.Dispose();
                pjContext?.Dispose();
            }
        } // UsageArea_NorthLatitude

    } // class

} // namespace

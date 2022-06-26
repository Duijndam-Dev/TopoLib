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

#pragma warning disable IDE0020 // Use pattern matching
#pragma warning disable IDE0038 // Use pattern matching

namespace TopoLib
{

    public static class Crs
    {
        internal static ProjContext CreateContext()
        {
            // Create context with proper Logging Level; Network access is defined in the ProjContext Constructor through 'EnableNetworkConnectionsOnNewContexts'
            ProjContext pjContext = new ProjContext { LogLevel = CctOptions.ProjContext.LogLevel, EnableNetworkConnections = CctOptions.ProjContext.EnableNetworkConnections};

            // Only use call-back when needed...
            if (CctOptions.ProjContext.LogLevel > 0)
                pjContext.Log += AddIn.ProcessSharpProjException;
            
            return pjContext;
        }

        internal static CoordinateReferenceSystem CreateCrs(in object[,] oCrs, in ProjContext pjContext, bool ExcelInterop = false)
        {
            int nCrsRows = oCrs.GetLength(0);
            int nCrsCols = oCrs.GetLength(1);
            int nOffset  = ExcelInterop ? 1 : 0;

            // max two adjacent CRS cells on the same row
            if (nCrsRows != 1 || nCrsCols > 2 ) 
                throw new ArgumentException("CRS");

            int nCrs;
            string sCrs;

            // We have only one cell; it can be a number, a Wkt string, a JSON string, a PROJ string or a textual description
            if (nCrsCols == 1)
            {
                // we have one cell describing the crs.
                if (oCrs[0 + nOffset,0 + nOffset] is double)
                {
                    // First cast to double, then to int, to deal with Excel datatypes
                    nCrs = (int)(double)oCrs[0 + nOffset, 0 + nOffset];   

                    // we have an EPSG number from a single input parameter:
                    return CoordinateReferenceSystem.CreateFromEpsg(nCrs, pjContext);
                }
                else if (oCrs[0 + nOffset, 0 + nOffset] is int)
                {
                    // Cast to int, input is coming from Excel Dialog Integer Input Control
                    nCrs = (int)oCrs[0 + nOffset, 0 + nOffset];   

                    // we have an EPSG number from a single input parameter:
                    return CoordinateReferenceSystem.CreateFromEpsg(nCrs, pjContext);
                }
                else if (oCrs[0 + nOffset, 0 + nOffset] is string)
                {
                    // cast to string, to deal with Excel datatypes
                    sCrs = (string)oCrs[0 + nOffset, 0 + nOffset];

                    bool success = int.TryParse(sCrs, out nCrs);
                    if (success)
                    {
                        // we have an EPSG number from a single input parameter:
                        return CoordinateReferenceSystem.CreateFromEpsg(nCrs, pjContext);
                    }
                    else
                    {
                        // we have a string of some sorts and a single input parameter:
                        // see also : https://github.com/OSGeo/gdal/blob/42843e79211954d1083012fa9ee6c4428e5ce772/gdal/ogr/ogrspatialreference.cpp#L3411-L3428
                        /*
                                // WKT1
                                "GEOGCS", "GEOCCS", "PROJCS", "VERT_CS", "COMPD_CS", "LOCAL_CS",
                                // WKT2"
                                "GEODCRS", "GEOGCRS", "GEODETICCRS", "GEOGRAPHICCRS", "PROJCRS",
                                "PROJECTEDCRS", "VERTCRS", "VERTICALCRS", "COMPOUNDCRS",
                                "ENGCRS", "ENGINEERINGCRS", "BOUNDCRS"
                        */

                        if (// WKT1
                            sCrs.StartsWith("GEOGCS")        || sCrs.StartsWith("GEOCCS")         || sCrs.StartsWith("PROJCS")       ||
                            sCrs.StartsWith("VERT_CS")       || sCrs.StartsWith("COMPD_CS")       || sCrs.StartsWith("LOCAL_CS")     ||
                            // WKT2
                            sCrs.StartsWith("GEODCRS")       || sCrs.StartsWith("GEOGCRS")        || sCrs.StartsWith("GEODETICCRS")  || 
                            sCrs.StartsWith("GEOGRAPHICCRS") || sCrs.StartsWith("PROJCRS")        || sCrs.StartsWith("PROJECTEDCRS") ||
                            sCrs.StartsWith("VERTCRS")       || sCrs.StartsWith("VERTICALCRS")    || sCrs.StartsWith("COMPOUNDCRS")  || 
                            sCrs.StartsWith("ENGCRS")        || sCrs.StartsWith("ENGINEERINGCRS") || sCrs.StartsWith("BOUNDCRS"))
                        {
                            // it must be WKT (well, we hope)
                            return CoordinateReferenceSystem.CreateFromWellKnownText(sCrs, pjContext);

                            // CreateFromWellKnownText() is translated into CreateFromWellKnownText(from, wars, ctx); where array<String^>^ wars = nullptr;
                            // It may throw an ArgumentNullException
                            // It may throw a ProjException
                        }
                        else
                        {
                            // it might be anything
                            return CoordinateReferenceSystem.Create(sCrs, pjContext);

                            // Create() is translated into proj_create(ctx, fromStr); 
                            // It may throw an ArgumentNullException("from");
                            // It may throw a ctx->ConstructException();
                            // It may throw a ProjException
                        }
                    }
                }
                else 
                    throw new ArgumentException("CRS");
            }
            else
            {
                // we have two adjacent CRS cells; first an Authortity string of some sorts and a second input parameter (number):

                sCrs = (string)oCrs[0 + nOffset, 0 + nOffset];   // the authority string

                // try to get the crs number; if not succesful throw an exception

                if (oCrs[0 + nOffset, 1 + nOffset] is double)
                {
                    // First cast to double, then to int, to deal with Excel datatypes
                    nCrs = (int)(double)oCrs[0 + nOffset, 1 + nOffset];
                }
                else if (oCrs[0 + nOffset, 1 + nOffset] is int)
                {
                    // Cast to int, input is coming from Excel Dialog Integer Input Control
                    nCrs = (int)oCrs[0 + nOffset, 1 + nOffset];
                }
                else if (oCrs[0 + nOffset, 1 + nOffset] is string)
                {
                    // cast to string, to deal with Excel datatypes
                    string sTmp = (string)oCrs[0 + nOffset, 1 + nOffset];

                    bool success = int.TryParse(sTmp, out nCrs);
                    if (!success) 
                        throw new ArgumentException("CRS");
                }
                else
                    throw new ArgumentException("CRS");

                return CoordinateReferenceSystem.CreateFromDatabase(sCrs, nCrs, pjContext);

                // CreateFromDatabase() is translated into proj_create_from_database 
                // It may throw a ArgumentNullException
                // It may throw a pjContext->ConstructException
            }

            // Oops, something went wrong if we get here...
            throw new ArgumentException("CRS");

        } // GetCrs

        [ExcelFunctionDoc(
            Name = "TL.crs.AsJsonString",
            Category = "CRS - Coordinate Reference System",
            Description = "Describes the Coordinate Reference System as a JSON-string",
            HelpTopic = "TopoLib-AddIn.chm!1300",

            Returns = "A JSON-string",
            Summary = "Function that describes a Coordinate Reference System as a JSON-string",
            Example = "Not implemented; requires long JSON-string as output text",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object AsJsonString(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            try
            {
                string CrsString ="";

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        CrsString = crs.AsProjJson();

                        if (String.IsNullOrWhiteSpace(CrsString))
                            return ExcelError.ExcelErrorValue;

                        return CrsString;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // AsJsonString

        [ExcelFunctionDoc(
            Name = "TL.crs.AsProjString",
            Category = "CRS - Coordinate Reference System",
            Description = "Describes the Coordinate Reference System as a PROJ-string",
            HelpTopic = "TopoLib-AddIn.chm!1301",

            Returns = "A PROJ-string",
            Summary = "Function that describes a Coordinate Reference System as a PROJ-string",
            Example = "TL.crs.AsProjString(2027) returns +proj=utm +zone=15 +ellps=clrk66 +units=m +no_defs +type=crs",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object AsProjString(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            try
            {
                string CrsString ="";

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        CrsString = crs.AsProjString();

                        if (String.IsNullOrWhiteSpace(CrsString))
                            return ExcelError.ExcelErrorValue;

                        return CrsString;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // AsProjString

        [ExcelFunctionDoc(
            Name = "TL.crs.AsWktString",
            Category = "CRS - Coordinate Reference System",
            Description = "Describes a Coordinate Reference System as a WKT-string",
            HelpTopic = "TopoLib-AddIn.chm!1302",

            Returns = "A WellKnownText-string",
            Summary = "Function that describes a Coordinate Reference System as a Well Known Text string",
            Example = "Not implemented; requires long WKT-string as output text",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object AsWktString(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            try
            {
                string CrsString ="";

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        CrsString = crs.AsWellKnownText();

                        if (String.IsNullOrWhiteSpace(CrsString))
                            return ExcelError.ExcelErrorValue;

                        return CrsString;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // AsWellKnownText

        [ExcelFunctionDoc(
            Name = "TL.crs.Axis.Abbreviation",
            Category = "CRS - Coordinate Reference System",
            Description = "Get the short name of axis-N of a coordinate reference system", 
            HelpTopic = "TopoLib-AddIn.chm!1303",

            Returns = "short name of Nth axis in a coordinate reference system",
            Summary = "Returns the short name of Nth axis in a coordinate reference system",
            Example = "TL.crs.CoordinateSystem.Axis.Abbreviation(3857) returns X",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object Axis_Abbreviation(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Zero based index of Axis list (0) ", Name = "Index")] object index,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            int nIndex  = (int)Optional.Check(index , 0.0);
            bool bNorm  = Optional.Check(normalized, true);
            string name ="";

            // do the work
            try
            {
                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            name = crs.Axis[nIndex].Abbreviation;

                        if (String.IsNullOrWhiteSpace(name))
                            name = "<NotFound>";

                        return name;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // Axis_Abbreviation

        [ExcelFunctionDoc(
            Name = "TL.crs.Axis.Count",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets type of Coordinate System",
            HelpTopic = "TopoLib-AddIn.chm!1304",

            Returns = "Nr of axis in a coordinate reference system",
            Summary = "Function that returns nr of axes in of Coordinate Reference System or -1 if not found",
            Example = "TL.crs.CoordinateSystem.Axis.Count(3857) returns 2",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object Axis_Count(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            try
            {
                int nAxes = -1;

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            nAxes = crs.Axis.Count;

                        return nAxes;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex); 
            }

        } // Axis_Count

        [ExcelFunctionDoc(
            Name = "TL.crs.Axis.Direction",
            Category = "CRS - Coordinate Reference System",
            Description = "Get the direction of axis-N of a coordinate reference system", 
            HelpTopic = "TopoLib-AddIn.chm!1305",

            Returns = "Direction of Nth axis in a coordinate reference system",
            Summary = "Returns the direction of Nth axis in a coordinate reference system",
            Example = "TL.crs.CoordinateSystem.Axis.Direction(3857, 1) returns north",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"            )]
        public static object Axis_Direction(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Zero based index of Axis list (0) ", Name = "Index")] object index,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            int nIndex  = (int)Optional.Check(index , 0.0);
            bool bNorm  = Optional.Check(normalized, true);
            string name ="";

            // do the work
            try
            {
                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            name = crs.Axis[nIndex].Direction;

                        if (String.IsNullOrWhiteSpace(name))
                            name = "<NotFound>";

                        return name;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // Axis_Direction

        [ExcelFunctionDoc(
            Name = "TL.crs.Axis.Name",
            Category = "CRS - Coordinate Reference System",
            Description = "Get the name of axis-N of a coordinate reference system", 
            HelpTopic = "TopoLib-AddIn.chm!1306",

            Returns = "name of Nth axis in a coordinate reference system",
            Summary = "Returns the name of Nth axis in a coordinate system",
            Example = "TL.crs.CoordinateSystem.Axis.Name(3875, 1) returns Northing",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object Axis_Name(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Zero based index of Identifier list (0) ", Name = "Index")] object index,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            int nIndex  = (int)Optional.Check(index , 0.0);
            bool bNorm  = Optional.Check(normalized, true);
            string name ="";

            // do the work
            try
            {
                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            name = crs.Axis[nIndex].Name;

                        if (String.IsNullOrWhiteSpace(name))
                            name = "<NotFound>";

                        return name;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // Axis_Name

        [ExcelFunctionDoc(
            Name = "TL.crs.Axis.UnitAuthName",
            Category = "CRS - Coordinate Reference System",
            Description = "Get the authority name of axis-N of a coordinate reference system", 
            HelpTopic = "TopoLib-AddIn.chm!1307",

            Returns = "authority name of Nth axis in a coordinate reference system",
            Summary = "Returns the authority name of Nth axis in a coordinate reference system",
            Example = "TL.crs.CoordinateSystem.Axis.UnitName(23095, 1) returns metre",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object Axis_UnitAuthName(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Zero based index of Axis list (0) ", Name = "Index")] object index,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            int nIndex  = (int)Optional.Check(index , 0.0);
            bool bNorm  = Optional.Check(normalized, true);
            string name ="";

            // do the work
            try
            {
                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            name = crs.Axis[nIndex].UnitAuthName;

                        if (String.IsNullOrWhiteSpace(name))
                            name = "<NotFound>";

                        return name;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // Axis_UnitAuthName

        [ExcelFunctionDoc(
            Name = "TL.crs.Axis.UnitCode",
            Category = "CRS - Coordinate Reference System",
            Description = "Get the authority name of axis-N of a coordinate reference system", 
            HelpTopic = "TopoLib-AddIn.chm!1308",

            Returns = "authority name of Nth axis in a coordinate reference system",
            Summary = "Returns the authority name of Nth axis in a coordinate reference system",
            Example = "TL.crs.CoordinateSystem.Axis.UnitCode(3857, 1) returns 9001",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object Axis_UnitCode(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Zero based index of Axis list (0) ", Name = "Index")] object index,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            int nIndex  = (int)Optional.Check(index , 0.0);
            bool bNorm  = Optional.Check(normalized, true);
            string name ="";

            // do the work
            try
            {
                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            name = crs.Axis[nIndex].UnitCode;

                        if (String.IsNullOrWhiteSpace(name))
                            name = "<NotFound>";

                        return name;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // Axis_UnitCode

        [ExcelFunctionDoc(
            Name = "TL.crs.Axis.UnitConversionFactor",
            Category = "CRS - Coordinate Reference System",
            Description = "Get the Unit Conversion Factor of axis-N of a coordinate reference system", 
            HelpTopic = "TopoLib-AddIn.chm!1309",

            Returns = "Unit Conversion Factorof Nth axis in a coordinate reference system",
            Summary = "Returns the Unit Conversion Factor of Nth axis in a coordinate reference system",
            Example = "TL.crs.CoordinateSystem.Axis.UnitConversionFactor(3857) returns 1.00",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object Axis_UnitConversionFactor(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Zero based index of Axis list (0) ", Name = "Index")] object index,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            int nIndex  = (int)Optional.Check(index , 0.0);
            bool bNorm  = Optional.Check(normalized, true);
            double factor = 0;

            // do the work
            try
            {
                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            factor = crs.Axis[nIndex].UnitConversionFactor;

                        return factor;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // Axis_UnitConversionFactor

        [ExcelFunctionDoc(
            Name = "TL.crs.Axis.UnitName",
            Category = "CRS - Coordinate Reference System",
            Description = "Get the unit name of axis-N of a coordinate reference system", 
            HelpTopic = "TopoLib-AddIn.chm!1310",

            Returns = "unit name of Nth axis in a coordinate reference system",
            Summary = "Returns the unit name of Nth axis in a coordinate reference system",
            Example = "TL.crs.PrimeMeridian.UnitName(357) returns degree",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object Axis_UnitName(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Zero based index of Axis list (0) ", Name = "Index")] object index,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            int nIndex  = (int)Optional.Check(index , 0.0);
            bool bNorm  = Optional.Check(normalized, true);
            string name ="";

            // do the work
            try
            {
                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            name = crs.Axis[nIndex].UnitName;

                        if (String.IsNullOrWhiteSpace(name))
                            name = "<NotFound>";

                        return name;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // Axis_UnitName

        [ExcelFunctionDoc(
            Name = "TL.crs.CelestialBodyName",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets the name of the appropriate celestial body for this Coordinate Reference System",
            HelpTopic = "TopoLib-AddIn.chm!1311",

            Returns = "The name of the appropriate celestial body",
            Summary = "Function that returns name of the appropriate celestial body for this CRS or &ltNotFound&gt if not found",
            Example = "TL.crs.CelestialBodyName(3857) returns Earth",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object CelestialBodyName(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            try
            {
                string name ="";

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            name = crs.CelestialBodyName;

                        if (String.IsNullOrWhiteSpace(name))
                            name = "<NotFound>";

                        return name;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // CelestialBodyName

        [ExcelFunctionDoc(
            Name = "TL.crs.CoordinateSystem.Axis.Abbreviation",
            Category = "CRS - Coordinate Reference System",
            Description = "Get the short name of axis-N of a coordinate system", 
            HelpTopic = "TopoLib-AddIn.chm!1312",

            Returns = "short name of Nth axis in a coordinate system",
            Summary = "Returns the short name of Nth axis in a coordinate system",
            Example = "TL.crs.CoordinateSystem.Axis.Name(3857, 1) returns Y",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object CoordinateSystem_Axis_Abbreviation(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Zero based index of Axis list (0) ", Name = "Index")] object index,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            int nIndex  = (int)Optional.Check(index , 0.0);
            bool bNorm  = Optional.Check(normalized, true);
            string name ="";

            // do the work
            try
            {
                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            name = crs.CoordinateSystem.Axis[nIndex].Abbreviation;

                        if (String.IsNullOrWhiteSpace(name))
                            name = "<NotFound>";

                        return name;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // CoordinateSystem_Axis_Abbreviation

        [ExcelFunctionDoc(
            Name = "TL.crs.CoordinateSystem.Axis.Count",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets type of Coordinate System",
            HelpTopic = "TopoLib-AddIn.chm!1313",

            Returns = "Coordinate name of coordinate system",
            Summary = "Function that returns nr of axes in of Coordinate System or -1 if not found",
            Example = "TL.crs.CoordinateSystem.Axis.Count(3857) returns 2",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object CoordinateSystem_Axis_Count(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            try
            {
                int nAxes = -1;

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null & crs.CoordinateSystem != null)
                            nAxes = crs.CoordinateSystem.Axis.Count;

                        return nAxes;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex); 
            }

        } // CoordinateSystem_Axis_Count

        [ExcelFunctionDoc(
            Name = "TL.crs.CoordinateSystem.Axis.Direction",
            Category = "CRS - Coordinate Reference System",
            Description = "Get the direction of axis-N of a coordinate system", 
            HelpTopic = "TopoLib-AddIn.chm!1314",

            Returns = "Direction of Nth axis in a coordinate system",
            Summary = "Returns the direction of Nth axis in a coordinate system",
            Example = "TL.crs.CoordinateSystem.Axis.Direction(3857, 1) returns north",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object CoordinateSystem_Axis_Direction(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Zero based index of Axis list (0) ", Name = "Index")] object index,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            int nIndex  = (int)Optional.Check(index , 0.0);
            bool bNorm  = Optional.Check(normalized, true);
            string name ="";

            // do the work
            try
            {
                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            name = crs.CoordinateSystem.Axis[nIndex].Direction;

                        if (String.IsNullOrWhiteSpace(name))
                            name = "<NotFound>";

                        return name;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // CoordinateSystem_Axis_Direction

        [ExcelFunctionDoc(
            Name = "TL.crs.CoordinateSystem.Axis.Name",
            Category = "CRS - Coordinate Reference System",
            Description = "Get the name of axis-N of a coordinate system", 
            HelpTopic = "TopoLib-AddIn.chm!1315",

            Returns = "name of Nth axis in a coordinate system",
            Summary = "Returns the name of Nth axis in a coordinate system",
            Example = "TL.crs.CoordinateSystem.Axis.Name(3857, 1) returns Northing",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object CoordinateSystem_Axis_Name(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Zero based index of Identifier list (0) ", Name = "Index")] object index,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            int nIndex  = (int)Optional.Check(index , 0.0);
            bool bNorm  = Optional.Check(normalized, true);
            string name ="";

            // do the work
            try
            {
                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            name = crs.CoordinateSystem.Axis[nIndex].Name;

                        if (String.IsNullOrWhiteSpace(name))
                            name = "<NotFound>";

                        return name;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // CoordinateSystem_Axis_Name

        [ExcelFunctionDoc(
            Name = "TL.crs.CoordinateSystem.Axis.UnitAuthName",
            Category = "CRS - Coordinate Reference System",
            Description = "Get the authority name of axis-N of a coordinate system", 
            HelpTopic = "TopoLib-AddIn.chm!1316",

            Returns = "authority name of Nth axis in a coordinate system",
            Summary = "Returns the authority name of Nth axis in a coordinate system",
            Example = "TL.crs.CoordinateSystem.Axis.UnitAuthName(3857, 1) returns EPSG",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object CoordinateSystem_Axis_UnitAuthName(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Zero based index of Axis list (0) ", Name = "Index")] object index,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            int nIndex  = (int)Optional.Check(index , 0.0);
            bool bNorm  = Optional.Check(normalized, true);
            string name ="";

            // do the work
            try
            {
                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            name = crs.CoordinateSystem.Axis[nIndex].UnitAuthName;

                        if (String.IsNullOrWhiteSpace(name))
                            name = "<NotFound>";

                        return name;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // CoordinateSystem_Axis_UnitAuthName

        [ExcelFunctionDoc(
            Name = "TL.crs.CoordinateSystem.Axis.UnitCode",
            Category = "CRS - Coordinate Reference System",
            Description = "Get the authority name of axis-N of a coordinate system", 
            HelpTopic = "TopoLib-AddIn.chm!1317",

            Returns = "authority name of Nth axis in a coordinate system",
            Summary = "Returns the authority name of Nth axis in a coordinate system",
            Example = "TL.crs.CoordinateSystem.Axis.UnitCode(3857, 1) returns 9001",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object CoordinateSystem_Axis_UnitCode(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Zero based index of Axis list (0) ", Name = "Index")] object index,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            int nIndex  = (int)Optional.Check(index , 0.0);
            bool bNorm  = Optional.Check(normalized, true);
            string name ="";

            // do the work
            try
            {
                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            name = crs.CoordinateSystem.Axis[nIndex].UnitCode;

                        if (String.IsNullOrWhiteSpace(name))
                            name = "<NotFound>";

                        return name;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // CoordinateSystem_Axis_UnitCode

        [ExcelFunctionDoc(
            Name = "TL.crs.CoordinateSystem.Axis.UnitConversionFactor",
            Category = "CRS - Coordinate Reference System",
            Description = "Get the Unit Conversion Factor of axis-N of a coordinate system", 
            HelpTopic = "TopoLib-AddIn.chm!1318",

            Returns = "Unit Conversion Factorof Nth axis in a coordinate system",
            Summary = "Returns the Unit Conversion Factor of Nth axis in a coordinate system",
            Example = "TL.crs.CoordinateSystem.Axis.UnitConversionFactot(3857, 1) returns 1.000",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object CoordinateSystem_Axis_UnitConversionFactor(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Zero based index of Axis list (0) ", Name = "Index")] object index,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            int nIndex  = (int)Optional.Check(index , 0.0);
            bool bNorm  = Optional.Check(normalized, true);
            double factor = 0;

            // do the work
            try
            {
                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            factor = crs.CoordinateSystem.Axis[nIndex].UnitConversionFactor;

                        return factor;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // CoordinateSystem_Axis_UnitConversionFactor

        [ExcelFunctionDoc(
            Name = "TL.crs.CoordinateSystem.Axis.UnitName",
            Category = "CRS - Coordinate Reference System",
            Description = "Get the unit name of axis-N of a coordinate system", 
            HelpTopic = "TopoLib-AddIn.chm!1319",

            Returns = "unit name of Nth axis in a coordinate system",
            Summary = "Returns the unit name of Nth axis in a coordinate system",
            Example = "TL.crs.CoordinateSystem.Axis.UnitName(3857, 1) returns metre",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object CoordinateSystem_Axis_UnitName(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Zero based index of Axis list (0) ", Name = "Index")] object index,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            int nIndex  = (int)Optional.Check(index , 0.0);
            bool bNorm  = Optional.Check(normalized, true);
            string name ="";

            // do the work
            try
            {
                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            name = crs.CoordinateSystem.Axis[nIndex].UnitName;

                        if (String.IsNullOrWhiteSpace(name))
                            name = "<NotFound>";

                        return name;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // CoordinateSystem_Axis_UnitName

        [ExcelFunctionDoc(
            Name = "TL.crs.CoordinateSystem.CoordinateSystemType",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets type of Coordinate System",
            HelpTopic = "TopoLib-AddIn.chm!1320",

            Returns = "Coordinate System Type",
            Summary = "Function that returns Coordinate System Type of Coordinate System or &ltNotFound&gt if not found",
            Example = "TL.crs.CoordinateSystem.CoordinateSystemType(3857) returns Cartesian",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object CoordinateSystem_CoordinateSystemType(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            try
            {
                string type ="";

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            type = crs.CoordinateSystem.CoordinateSystemType.ToString();

                        if (String.IsNullOrWhiteSpace(type))
                            type = "<NotFound>";

                        return type;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // CoordinateSystem_CoordinateSystemType

        [ExcelFunctionDoc(
            Name = "TL.crs.CoordinateSystem.Name",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets type of Coordinate System",
            HelpTopic = "TopoLib-AddIn.chm!1321",

            Returns = "Coordinate name of coordinate system",
            Summary = "Function that returns name of Coordinate System or &ltNotFound&gt if not found",
            Example = "TL.crs.CoordinateSystem.Name(3857) returns <NotFound>",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object CoordinateSystem_Name(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            try
            {
                string name ="";

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            name = crs.CoordinateSystem.Name;

                        if (String.IsNullOrWhiteSpace(name))
                            name = "<NotFound>";

                        return name;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // CoordinateSystem_Name

        [ExcelFunctionDoc(
            Name = "TL.crs.CoordinateSystem.Type",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets type of Coordinate System",
            HelpTopic = "TopoLib-AddIn.chm!1322",

            Returns = "Coordinate system type",
            Summary = "Function that returns type of Coordinate System or &ltNotFound&gt if not found",
            Example = "TL.crs.CoordinateSystem.Type(3857) returns CoordinateSystem",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object CoordinateSystem_Type(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);
            try
            {
                string type ="";

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            type = crs.CoordinateSystem.Type.ToString();

                        if (String.IsNullOrWhiteSpace(type))
                            type = "<NotFound>";

                        return type;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // CoordinateSystem_Type

        [ExcelFunctionDoc(
            Name = "TL.crs.Datum.Name",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets name of datum of CRS",
            HelpTopic = "TopoLib-AddIn.chm!1323",

            Returns = "CRS name",
            Summary = "Function that returns name of datum of CRS or &ltNotFound&gt if not found",
            Example = "TL.crs.Datum.Name(28992) returns Amersfoort",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object Datum_Name(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            try
            {
                string name ="";

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            name = crs.Datum.Name;

                        if (String.IsNullOrWhiteSpace(name))
                            name = "<NotFound>";

                        return name;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // Datum_Name

        [ExcelFunctionDoc(
            Name = "TL.crs.Datum.Type",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets type of datum in CRS",
            HelpTopic = "TopoLib-AddIn.chm!1324",

            Returns = "Datum type",
            Summary = "Function that returns type of datum of CRS or &ltNotFound&gt if not found",
            Example = "TL.crs.Datum.Type(3857) returns DatumEnsamble",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object Datum_Type(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            try
            {
                string type ="";

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            type = crs.Datum.Type.ToString();

                        if (String.IsNullOrWhiteSpace(type))
                            type = "<NotFound>";

                        return type;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // Datum_Type

        [ExcelFunctionDoc(
            Name = "TL.crs.Ellipsoid.Name",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets name of Ellipsoid in Coordinate Reference System",
            HelpTopic = "TopoLib-AddIn.chm!1325",

            Returns = "Ellipsoid name",
            Summary = "Function that returns name of Ellipsoid in CRS or &ltNotFound&gt if not found",
            Example = "TL.crs.Ellipsoid.Name(28992) returns Bessel 1841",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object Ellipsoid_Name(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            try
            {
                string name ="";

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            name = crs.Ellipsoid.Name;

                        if (String.IsNullOrWhiteSpace(name))
                            name = "<NotFound>";

                        return name;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // Ellipsoid_Name

        [ExcelFunctionDoc(
            Name = "TL.crs.Ellipsoid.Type",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets type of ellipsoid in Coordinate Reference System",
            HelpTopic = "TopoLib-AddIn.chm!1326",

            Returns = "Ellipsoid type",
            Summary = "Function that returns type of ellipsoid in CRS or &ltNotFound&gt if not found",
            Example = "TL.crs.Ellipsoid.Type(28992) returns Ellipsoid",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object Ellipsoid_Type(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            try
            {
                string type ="";

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            type = crs.Ellipsoid.Type.ToString();

                        if (String.IsNullOrWhiteSpace(type))
                            type = "<NotFound>";

                        return type;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // Ellipsoid_Type

        [ExcelFunctionDoc(
            Name = "TL.crs.Ellipsoid.SemiMajorMetre",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets size of ellipsoid's semi-major axis in metres",
            HelpTopic = "TopoLib-AddIn.chm!1327",

            Returns = "Size of ellipsoid's semi-major axis in metres",
            Summary = "Function that returns size of ellipsoid's semi-major axis in metres or -1 if not found",
            Example = "TL.crs.Ellipsoid.SemiMajorMetre(3857) returns 6378137.000",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object Ellipsoid_SemiMajorMetre(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            try
            {
                double res = -1;

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            res = crs.Ellipsoid.SemiMajorMetre;

                        return res;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // Ellipsoid_SemiMajorMetre

        [ExcelFunctionDoc(
            Name = "TL.crs.Ellipsoid.SemiMinorMetre",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets size of ellipsoid's semi-minor axis in metres",
            HelpTopic = "TopoLib-AddIn.chm!1328",

            Returns = "Size of ellipsoid's semi-minor axis in metres",
            Summary = "Function that returns size of ellipsoid's semi-minor axis in metres or -1 if not found",
            Example = "TL.crs.Ellipsoid.SemiMinorMetre(3857) returns 6356752.314",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object Ellipsoid_SemiMinorMetre(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            try
            {
                double res = -1;

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            res = crs.Ellipsoid.SemiMinorMetre;

                        return res;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // Ellipsoid_SemiMinorMetre

        [ExcelFunctionDoc(
            Name = "TL.crs.Ellipsoid.InverseFlattening",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets inverse flattening of ellipsoid",
            HelpTopic = "TopoLib-AddIn.chm!1329",

            Returns = "Inverse flattening of ellipsoid",
            Summary = "Function that returns inverse flattening of ellipsoid, or -1 if not found",
            Example = "TL.crs.Ellipsoid.InverseFlattening(23095) returns 297.000",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object Ellipsoid_InverseFlattening(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            try
            {
                double res = -1;

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            res = crs.Ellipsoid.InverseFlattening;

                        return res;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // Ellipsoid_InverseFlattening

        [ExcelFunctionDoc(
            Name = "TL.crs.Ellipsoid.IsSemiMinorComputed",
            Category = "CRS - Coordinate Reference System",
            Description = "Confirms whether when size of semi-minor axis has been calculated",
            HelpTopic = "TopoLib-AddIn.chm!1330",

            Returns = "TRUE when size of semi-minor axis has been calculated; FALSE when not",
            Summary = "Function that confirms whether when size of semi-minor axis has been calculated",
            Example = "TL.crs.Ellipsoid.IsSemiMinorComputed(23095) returns TRUE",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object Ellipsoid_IsSemiMinorComputed(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            try
            {
                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            return crs.Ellipsoid.IsSemiMinorComputed;

                        return ExcelError.ExcelErrorValue; 
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // EllipsoidIsSemiMinorComputed

        [ExcelFunctionDoc(
            Name = "TL.crs.EuclideanArea",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets Euclidean area defined by 3 or more points in a Projected or Engineering CRS, ignoring elevation differences",
            HelpTopic = "TopoLib-AddIn.chm!1331",

            Returns = "Euclidean area covered by three (or more) points in [m2]",
            Summary = "Function that returns Euclidean area, covered by three or more points, defined in a Projected or Engineering CRS, ignoring elevation differences, or #NA error if CRS not found" +
                      "<p>This routine explicitly requires that CRS being used, is a projected or engineering CRS.</p>",
            Remarks = "The setting Normalized 'true' always converts the calculated area to [m2]"
            )]
        public static object EuclideanArea(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Vertical list of points", Name = "PointList")] object[,] Points,
            [ExcelArgument("Normalize area to [m2] (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            int nPointRows = Points.GetLength(0);
            int nPointCols = Points.GetLength(1);

            if (nPointRows < 3 )
                return ExcelError.ExcelErrorValue;

            if (nPointCols < 2 || nPointCols > 4 )
                return ExcelError.ExcelErrorValue;

            // Now determine number of points in (poly)line
            int nPoints = nPointRows + 1;

            double[,] points = new double[nPoints, 2];

            // First, get all the points from Points
            for (int i = 0; i < nPointRows; i++)
            {
                points[i, 0] = (double)Points[i, 0];
                points[i, 1] = (double)Points[i, 1];
            }

            // Make sure we close the polygon
            points[nPointRows, 0] = points[0, 0];
            points[nPointRows, 1] = points[0, 1];

            try
            {
                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                        {
                            var projType = crs.Type;

                            if (projType == ProjType.ProjectedCrs || projType == ProjType.EngineeringCrs)
                            {
                                double sum1 = 0;
                                double sum2 = 0;

                                // so far now

                                for (int i = 0; i < nPoints - 1; i++)
                                {
                                    sum1 += points[i, 0] * points[i + 1, 1];
                                    sum2 += points[i, 1] * points[i + 1, 0];
                                }

                                double area = Math.Abs(0.5 * (sum1 - sum2));

                                // See : https://github.com/tracyharton/Proj4/blob/master/pj_units.c
                                double conversion0 = crs.Axis[0].UnitConversionFactor;
                                double conversion1 = crs.Axis[1].UnitConversionFactor;

                                if (bNorm)
                                {
                                    area *= (conversion0 * conversion1);
                                }

                                return area;
                            }
                            else
                                throw new Exception("Euclidean Area requires a Projected or Engineering CRS");
                        }
                        else
                            return ExcelError.ExcelErrorValue;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // EuclideanArea

        [ExcelFunctionDoc(
            Name = "TL.crs.EuclideanDistance",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets Euclidean distance between two points (or between multiple points in a poly-line), defined in a Projected or Engineering CRS, ignoring elevation differences",
            HelpTopic = "TopoLib-AddIn.chm!1332",

            Returns = "Euclidean distance between two (or more) points in [m]",
            Summary = "Function that returns Euclidean distance between two points (or between multiple points in a poly-line), defined in a Projected or Engineering CRS, ignoring elevation differences, or #NA error if CRS not found" +
                      "<p>This routine explicitly requires that CRS being used, is a projected or engineering CRS.</p>",
            Remarks = "The setting Normalized 'true' always converts the calculated distance to [m]"
            )]
        public static object EuclideanDistance(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Start point (or vertical list) of x,y points)", Name = "startPointOrPointList")] object[,] Point1,
            [ExcelArgument("End point (ignored when using list of points)", Name = "endPointOrNul")] object[,] Point2,
            [ExcelArgument("Normalize distance to [m] (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            int nPoint1Rows = Point1.GetLength(0);
            int nPoint1Cols = Point1.GetLength(1);

            int nPoint2Rows = Point2.GetLength(0);
            int nPoint2Cols = Point2.GetLength(1);

            if (nPoint1Rows < 1 )
                return ExcelError.ExcelErrorValue;

            if (nPoint1Cols < 2 || nPoint1Cols > 4 )
                return ExcelError.ExcelErrorValue;

            // Check use of nPoint2; we need it when nPoint1Rows == 1
            if (nPoint1Rows == 1 &&  nPoint2Rows != 1 )
                return ExcelError.ExcelErrorValue;

            // Do the check on nPoint2Cols only when we need Point2
            if (nPoint1Rows == 1 && (nPoint2Cols < 2 || nPoint2Cols > 4 ))
                return ExcelError.ExcelErrorValue;

            // Now determine number of points in (poly)line
            int nPoints = nPoint1Rows > 1 ? nPoint1Rows : 2;

            double[,] points = new double[nPoints, 2];

            // First, get all the points from Point1
            for (int i = 0; i < nPoint1Rows; i++)
            {
                points[i, 0] = (double)Point1[i, 0];
                points[i, 1] = (double)Point1[i, 1];
            }

            // Now get the last point from Point2 if needed
            if (nPoints > nPoint1Rows)
            {
                points[nPoints - 1, 0] = (double)Point2[0, 0];
                points[nPoints - 1, 1] = (double)Point2[0, 1];
            }

            try
            {
                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                        {
                            var projType = crs.Type;

                            if (projType == ProjType.ProjectedCrs || projType == ProjType.EngineeringCrs)
                            {
                                double length = 0;

                                for (int i = 0; i < nPoints - 1; i++)
                                {
                                    length += Math.Sqrt(Math.Pow(points[i + 1, 0] - points[i, 0], 2) + Math.Pow(points[i + 1, 1] - points[i, 1], 2));
                                }
                                double conversion0 = crs.Axis[0].UnitConversionFactor;

                                if (bNorm)
                                {
                                    length *= conversion0;
                                }

                                return length;
                            }
                            else
                                throw new Exception("Euclidean Distance requires a Projected or Engineering CRS");
                        }
                        else
                            return ExcelError.ExcelErrorValue;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // EuclideanDistance

        [ExcelFunctionDoc(
            Name = "TL.crs.EuclideanDistanceZ",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets Euclidean distance between two points (or between multiple points in a poly-line), defined in a Projected or Engineering CRS, ignoring elevation differences",
            HelpTopic = "TopoLib-AddIn.chm!1333",

            Returns = "Euclidean distance between two (or more) points in [m]",
            Summary = "Function that returns Euclidean distance between two points (or between multiple points in a poly-line), defined in a Projected or Engineering CRS, honoring elevation differences, or #NA error if CRS not found" +
                      "<p>This routine explicitly requires that CRS being used, is a projected or engineering CRS.</p>",
            Remarks = "The setting Normalized 'true' always converts the calculated distance to [m]"
            )]
        public static object EuclideanDistanceZ(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Start point (or vertical list) of x,y,z points)", Name = "startPointOrPointList")] object[,] Point1,
            [ExcelArgument("End point (ignored when using list of points)", Name = "endPointOrNul")] object[,] Point2,
            [ExcelArgument("Normalize distance to [m] (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            int nPoint1Rows = Point1.GetLength(0);
            int nPoint1Cols = Point1.GetLength(1);

            int nPoint2Rows = Point2.GetLength(0);
            int nPoint2Cols = Point2.GetLength(1);

            if (nPoint1Rows < 1 )
                return ExcelError.ExcelErrorValue;

            if (nPoint1Cols < 2 || nPoint1Cols > 4 )
                return ExcelError.ExcelErrorValue;

            // Check use of nPoint2; we need it when nPoint1Rows == 1
            if (nPoint1Rows == 1 &&  nPoint2Rows != 1 )
                return ExcelError.ExcelErrorValue;

            // Do the check on nPoint2Cols only when we need Point2
            if (nPoint1Rows == 1 && (nPoint2Cols < 2 || nPoint2Cols > 4 ))
                return ExcelError.ExcelErrorValue;

            // Now determine number of points in (poly)line
            int nPoints = nPoint1Rows > 1 ? nPoint1Rows : 2;

            double[,] points = new double[nPoints, 3];

            // First, get all the points from Point1
            for (int i = 0; i < nPoint1Rows; i++)
            {
                points[i, 0] = (double)Point1[i, 0];
                points[i, 1] = (double)Point1[i, 1];
                points[i, 2] = nPoint1Cols > 2 ? (double)Point1[i, 2]: 0;
            }

            // Now get the last point from Point2 if needed
            if (nPoints > nPoint1Rows)
            {
                points[nPoints - 1, 0] = (double)Point2[0, 0];
                points[nPoints - 1, 1] = (double)Point2[0, 1];
                points[nPoints - 1, 2] = nPoint2Cols > 2 ? (double)Point1[0, 2]: 0;
            }

            try
            {
                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                        {
                            var projType = crs.Type;

                            if (projType == ProjType.ProjectedCrs || projType == ProjType.EngineeringCrs)
                            {
                                double length = 0;
                                for (int i = 0; i < nPoints - 1; i++)
                                {
                                    length += Math.Sqrt(  Math.Pow(points[i + 1, 0] - points[i, 0], 2) 
                                                        + Math.Pow(points[i + 1, 1] - points[i, 1], 2)
                                                        + Math.Pow(points[i + 1, 2] - points[i, 2], 2));
                                }
                                double conversion0 = crs.Axis[0].UnitConversionFactor;

                                if (bNorm)
                                {
                                    length *= conversion0;
                                }

                                return length;
                            }
                            else
                                throw new Exception("Euclidean Distance requires a Projected or Engineering CRS");
                        }
                        else
                            return ExcelError.ExcelErrorValue;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // EuclideanDistanceZ

        [ExcelFunctionDoc(
            Name = "TL.crs.GeodesicArea",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets surface area defined by multiple (at least 3) points in a polygon defined in a Coordinate Reference System",
            HelpTopic = "TopoLib-AddIn.chm!1334",

            Returns = "surface area defined by multiple (at least 3) points in a polygon [m2]",
            Summary = "Function that returns surface area defined by multiple (at least 3) points in a polygon defined in a Coordinate Reference System, or #NA error if CRS not found" +
                      "<p>See: <a href = \"https://proj.org/geodesic.html\" >Geodesic calculations</a> for the Proj Library</p>",
            Example = "TL.crs.GeodesicArea(23031, {{554073.0, 5885683.0}, {572955.0, 5886200.0}, {572415.0, 5905245.0}, {553511.0, 5904706.0}}) returns 360153022.6",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object GeodesicArea(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Vertical list of points in a polygon)", Name = "startPointOrPointList")] object[,] Points,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            int nPointRows = Points.GetLength(0);
            int nPointCols = Points.GetLength(1);

            if (nPointRows < 3 )
                return ExcelError.ExcelErrorValue;

            if (nPointCols < 2 || nPointCols > 4 )
                return ExcelError.ExcelErrorValue;

            int nPoints = nPointRows;

            PPoint[] points = new PPoint[nPoints];

            // First, get all the points from Point1
            for (int i = 0; i < nPoints; i++)
            {
                switch (nPointCols)
                {
                    default:
                    case 2:
                        points[i] = new PPoint((double)Points[i, 0], (double)Points[i, 1]);
                        break;
                    case 3:
                        points[i] = new PPoint((double)Points[i, 0], (double)Points[i, 1], (double)Points[i, 2]);
                        break;
                    case 4:
                        points[i] = new PPoint((double)Points[i, 0], (double)Points[i, 1], (double)Points[i, 2], (double)Points[i,3]);
                        break;
                }
            }

            try
            {
                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                        {
                            double distance = Math.Abs(crs.GeoArea(points));
                            return Optional.CheckNan(distance);
                        }
                        else
                            return ExcelError.ExcelErrorValue;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // GeodesicArea

        [ExcelFunctionDoc(
            Name = "TL.crs.GeodesicDistance",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets geodesic distance between two points (or between multiple points in a poly-line), defined in a Coordinate Reference System, ignoring elevation differences",
            HelpTopic = "TopoLib-AddIn.chm!1335",

            Returns = "Geodesic distance between two (or more) points in [m]",
            Summary = "Function that returns geodesic distance between two points (or between multiple points in a poly-line), defined in a Coordinate Reference System, ignoring elevation differences or #NA error if CRS not found" +
                      "<p>See: <a href = \"https://proj.org/geodesic.html\" >Geodesic calculations</a> for the Proj Library</p>",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object GeodesicDistance(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Start point (or vertical list of points)", Name = "startPointOrPointList")] object[,] Point1,
            [ExcelArgument("End point (ignored when using list of points)", Name = "endPointOrNul")] object[,] Point2,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            int nPoint1Rows = Point1.GetLength(0);
            int nPoint1Cols = Point1.GetLength(1);

            int nPoint2Rows = Point2.GetLength(0);
            int nPoint2Cols = Point2.GetLength(1);

            if (nPoint1Rows < 1 )
                return ExcelError.ExcelErrorValue;

            if (nPoint1Cols < 2 || nPoint1Cols > 4 )
                return ExcelError.ExcelErrorValue;

            // Check use of nPoint2; we need it when nPoint1Rows == 1
            if (nPoint1Rows == 1 &&  nPoint2Rows != 1 )
                return ExcelError.ExcelErrorValue;

            // Do the check on nPoint2Cols only when we need Point2
            if (nPoint1Rows == 1 && (nPoint2Cols < 2 || nPoint2Cols > 4 ))
                return ExcelError.ExcelErrorValue;

            // Now determine number of points in (poly)line
            int nPoints = nPoint1Rows > 1 ? nPoint1Rows : 2;

            PPoint[] points = new PPoint[nPoints];

            // First, get all the points from Point1
            for (int i = 0; i < nPoint1Rows; i++)
            {
                switch (nPoint1Cols)
                {
                    default:
                    case 2:
                        points[i] = new PPoint((double)Point1[i, 0], (double)Point1[i, 1]);
                        break;
                    case 3:
                        points[i] = new PPoint((double)Point1[i, 0], (double)Point1[i, 1], (double)Point1[i, 2]);
                        break;
                    case 4:
                        points[i] = new PPoint((double)Point1[i, 0], (double)Point1[i, 1], (double)Point1[i, 2], (double)Point1[i,3]);
                        break;
                }
            }

            // Now get the last point from Point2 if needed
            if (nPoints > nPoint1Rows)
            {
                switch (nPoint2Cols)
                {
                    default:
                    case 2:
                        points[nPoints - 1] = new PPoint((double)Point2[0, 0], (double)Point2[0, 1]);
                        break;
                    case 3:
                        points[nPoints - 1] = new PPoint((double)Point2[0, 0], (double)Point2[0, 1], (double)Point2[0, 2]);
                        break;
                    case 4:
                        points[nPoints - 1] = new PPoint((double)Point2[0, 0], (double)Point2[0, 1], (double)Point2[0, 2], (double)Point2[0, 3]);
                        break;
                }
            }

            try
            {
                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                        {
                            double distance = crs.GeoDistance(points);
                            return Optional.CheckNan(distance);
                        }
                        else
                            return ExcelError.ExcelErrorValue;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // GeodesicDistance

        [ExcelFunctionDoc(
            Name = "TL.crs.GeodesicDistanceZ",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets geodesic distance between two points (or between multiple points in a poly-line), defined in a Coordinate Reference System, honoring elevation differences",
            HelpTopic = "TopoLib-AddIn.chm!1336",

            Returns = "Geodesic distance between two (or more) points in [m]",
            Summary = "Function that returns geodesic distance between two points (or between multiple points in a poly-line), defined in a Coordinate Reference System, honoring elevation differences or #NA error if CRS not found" +
                      "<p>See: <a href = \"https://proj.org/geodesic.html\" >Geodesic calculations</a> for the Proj Library</p>",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object GeodesicDistanceZ(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Start point (or vertical list of points)", Name = "startPointOrPointList")] object[,] Point1,
            [ExcelArgument("End point (ignored when using list of points)", Name = "endPointOrNul")] object[,] Point2,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            int nPoint1Rows = Point1.GetLength(0);
            int nPoint1Cols = Point1.GetLength(1);

            int nPoint2Rows = Point2.GetLength(0);
            int nPoint2Cols = Point2.GetLength(1);

            if (nPoint1Rows < 1 )
                return ExcelError.ExcelErrorValue;

            if (nPoint1Cols < 2 || nPoint1Cols > 4 )
                return ExcelError.ExcelErrorValue;

            // Check use of nPoint2; we need it when nPoint1Rows == 1
            if (nPoint1Rows == 1 &&  nPoint2Rows != 1 )
                return ExcelError.ExcelErrorValue;

            // Do the check on nPoint2Cols only when we need Point2
            if (nPoint1Rows == 1 && (nPoint2Cols < 2 || nPoint2Cols > 4 ))
                return ExcelError.ExcelErrorValue;

            // Now determine number of points in (poly)line
            int nPoints = nPoint1Rows > 1 ? nPoint1Rows : 2;

            PPoint[] points = new PPoint[nPoints];

            // First, get all the points from Point1
            for (int i = 0; i < nPoint1Rows; i++)
            {
                switch (nPoint1Cols)
                {
                    default:
                    case 2:
                        points[i] = new PPoint((double)Point1[i, 0], (double)Point1[i, 1]);
                        break;
                    case 3:
                        points[i] = new PPoint((double)Point1[i, 0], (double)Point1[i, 1], (double)Point1[i, 2]);
                        break;
                    case 4:
                        points[i] = new PPoint((double)Point1[i, 0], (double)Point1[i, 1], (double)Point1[i, 2], (double)Point1[i,3]);
                        break;
                }
            }

            // Now get the last point from Point2 if needed
            if (nPoints > nPoint1Rows)
            {
                switch (nPoint2Cols)
                {
                    default:
                    case 2:
                        points[nPoints - 1] = new PPoint((double)Point2[0, 0], (double)Point2[0, 1]);
                        break;
                    case 3:
                        points[nPoints - 1] = new PPoint((double)Point2[0, 0], (double)Point2[0, 1], (double)Point2[0, 2]);
                        break;
                    case 4:
                        points[nPoints - 1] = new PPoint((double)Point2[0, 0], (double)Point2[0, 1], (double)Point2[0, 2], (double)Point2[0, 3]);
                        break;
                }
            }

            try
            {
                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                        {
                            double distance = crs.GeoDistanceZ(points);
                            return Optional.CheckNan(distance);
                        }
                        else
                            return ExcelError.ExcelErrorValue;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // GeoDistanceZ

        [ExcelFunctionDoc(
            Name = "TL.crs.Identifiers.Authority",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets Authority of Identifier N",
            HelpTopic = "TopoLib-AddIn.chm!1337",

            Returns = "Authority of Nth Identifier",
            Summary = "Function that returns Authority of <Nth> identifiers or <index out of range> when not found",
            Example = "TL.crs.Identifiers.Authority(2583,0) returns EPSG",
            Remarks = "for 'Identifiers' we better NOT normalize the Axes, as this risks NULLIFYING the Identifiers"
            )]
        public static object Identifiers_Authority(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Zero based index of Identifier list (0) ", Name = "Index")] object index,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (false)", Name = "Normalized")] object normalized
             )
        {
            bool bNorm  = Optional.Check(normalized, false);

            try
            {
                int nIndex  = (int)Optional.Check(index , 0.0);

                int Count = -1;

                using (ProjContext pjContext = CreateContext())
                {
                    // for Identifiers we CANNOT normalize the Axes, as this will NULL the Identifiers
                    // using (var crs = CreateCrs(oCrs, pjContext).WithAxisNormalized())
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null && crs.Identifiers != null)
                        {
                            Count = crs.Identifiers.Count;

                            if (nIndex > Count - 1 || nIndex < 0)
                                return "<index out of range>";

                            return crs.Identifiers[nIndex].Authority;
                        }
                        else
                            return "Unknown";
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // Identifiers_Authority

        [ExcelFunctionDoc(
            Name = "TL.crs.Identifiers.Code",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets Code of Identifier N",
            HelpTopic = "TopoLib-AddIn.chm!1338",

            Returns = "Code of Nth Identifiers",
            Summary = "Function that returns the Code of the <Nth> identifier or <index out of range> when not found",
            Example = "Not implemented; requires long WKT or JSON string as input text",
            Remarks = "for 'Identifiers' we better NOT normalize the Axes, as this risks NULLIFYING the Identifiers"
            )]
        public static object Identifiers_Code(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Zero based index of Identifier list (0) ", Name = "Index")] object index,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (false)", Name = "Normalized")] object normalized
             )
        {
            bool bNorm  = Optional.Check(normalized, false);

            try
            {
                int nIndex  = (int)Optional.Check(index , 0.0);

                using (ProjContext pjContext = CreateContext())
                {
                    //for Identifiers we CANNOT normalize the Axes, as this will NULL the Identifiers
                    // using (var crs = CreateCrs(oCrs, pjContext).WithAxisNormalized())
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null && crs.Identifiers != null)
                        {
                            int Count = crs.Identifiers.Count;

                            if (nIndex > Count - 1 || nIndex < 0)
                                return "<index out of range>";

                            return crs.Identifiers[nIndex].Code;
                        }
                        else
                            return "Unknown";

                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // Identifiers_Code

        [ExcelFunctionDoc(
            Name = "TL.crs.Identifiers.Count",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets number of Identifiers",
            HelpTopic = "TopoLib-AddIn.chm!1339",

            Returns = "Number of CRS Identifiers",
            Summary = "Function that returns nr of CRS identifiers or 0 if none found",
            Example = "TL.crs.Identifiers.Count(2583) returns 1",
            Remarks = "for 'Identifiers' we better NOT normalize the Axes, as this risks NULLIFYING the Identifiers"
            )]
        public static object Identifiers_Count(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (false)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, false);

            try
            {
                int Count = -0;

                using (ProjContext pjContext = CreateContext())
                {
                    //for Identifiers we CANNOT normalize the Axes, as this will NULL the Identifiers
                    // using (var crs = CreateCrs(oCrs, pjContext).WithAxisNormalized())
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null && crs.Identifiers != null)
                            Count = crs.Identifiers.Count;

                        return Count;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // Identifiers_Count

        [ExcelFunctionDoc(
            Name = "TL.crs.IsDeprecated",
            Category = "CRS - Coordinate Reference System",
            Description = "Confirms whether when size of semi-minor axis has been calculated",
            HelpTopic = "TopoLib-AddIn.chm!1340",

            Returns = "TRUE when CRS is deprecated; FALSE when not",
            Summary = "Function that confirms whether the CRS is deprecated",
            Example = "TL.crs.IsDeprecated(2037) returns TRUE",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object IsDeprecated(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            try
            {
                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            return crs.IsDeprecated;
                        else
                            return ExcelError.ExcelErrorValue; 
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // IsDeprecated

        [ExcelFunctionDoc(
            Name = "TL.crs.IsEquivalentTo",
            Category = "CRS - Coordinate Reference System",
            Description = "Confirms whether two different CRSs are equivalent",
            HelpTopic = "TopoLib-AddIn.chm!1341",

            Returns = "TRUE when the two different CRSs are equivalent; FALSE when not",
            Summary = "Function that checks whether two different CRSs are equivalent",
            Remarks = "<p>The objects are equivalent for the purpose of coordinate operations. They can differ by the name of their objects, identifiers, other metadata. " +
            "Parameters may be expressed in different units, provided that the value is (with some tolerance) the same once expressed in a common unit.</p>" +
            "<p>Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib</p>"
            )]
        public static object IsEquivalentTo(
            [ExcelArgument("CRS-1, consisting of one [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs1")] object[,] oCrs1,
            [ExcelArgument("CRS-2, consisting of one [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs2")] object[,] oCrs2,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            try
            {
                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs1 = bNorm ? CreateCrs(oCrs1, pjContext).WithAxisNormalized() : CreateCrs(oCrs1, pjContext))
                    using (var crs2 = bNorm ? CreateCrs(oCrs2, pjContext).WithAxisNormalized() : CreateCrs(oCrs2, pjContext))
                    {
                        if (crs1 != null && crs2 != null )
                            return crs1.IsEquivalentTo(crs2, pjContext);
                        else
                            return ExcelError.ExcelErrorValue; 
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // IsEquivalentTo

        [ExcelFunctionDoc(
            Name = "TL.crs.IsEquivalentToRelaxed",
            Category = "CRS - Coordinate Reference System",
            Description = "Confirms whether two different CRSs are equivalent with the axes in any order",
            HelpTopic = "TopoLib-AddIn.chm!1342",

            Returns = "TRUE when the two different CRSs are equivalent with the axes in any order; FALSE when not",
            Summary = "Function that checks whether two different CRSs are equivalent with the axes in any order",
            Remarks = "<p>The objects are equivalent for the purpose of coordinate operations. They can differ by the name of their objects, identifiers, other metadata. " +
            "Parameters may be expressed in different units, provided that the value is (with some tolerance) the same once expressed in a common unit.</p>" +
            "<p>Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib</p>"
            )]
        public static object IsEquivalentToRelaxed(
            [ExcelArgument("CRS-1, consisting of one [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs1")] object[,] oCrs1,
            [ExcelArgument("CRS-2, consisting of one [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs2")] object[,] oCrs2,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            try
            {
                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs1 = bNorm ? CreateCrs(oCrs1, pjContext).WithAxisNormalized() : CreateCrs(oCrs1, pjContext))
                    using (var crs2 = bNorm ? CreateCrs(oCrs2, pjContext).WithAxisNormalized() : CreateCrs(oCrs2, pjContext))
                    {
                        if (crs1 != null && crs2 != null )
                            return crs1.IsEquivalentToRelaxed(crs2, pjContext);
                        else
                            return ExcelError.ExcelErrorValue; 
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // IsEquivalentToRelaxed

        [ExcelFunctionDoc(
            Name = "TL.crs.Name",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets name of Coordinate Reference System",
            HelpTopic = "TopoLib-AddIn.chm!1343",

            Returns = "CRS name",
            Summary = "Function that returns name of CRS or &ltNotFound&gt if not found",
            Example = "TL.crs.Name(28992) returns Amersfoort / RD New",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object Name(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            try
            {
                string name ="";

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            name = crs.Name;

                        if (String.IsNullOrWhiteSpace(name))
                            name = "<NotFound>";

                        return name;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // Name

        [ExcelFunctionDoc(
            Name = "TL.crs.GeodeticCRS.Name",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets name of Geodetic Coordinate Reference System",
            HelpTopic = "TopoLib-AddIn.chm!1344",

            Returns = "CRS name",
            Summary = "Function that returns name of Geodetic CRS or &ltNotFound&gt if not found",
            Example = "TL.crs.GeodeticCRS.Name(23095) returns ED50",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object GeodeticCRS_Name(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            try
            {
                string name ="";

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            name = crs.GeodeticCRS.Name;

                        if (String.IsNullOrWhiteSpace(name))
                            name = "<NotFound>";

                        return name;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // GeodeticCRS_Name

        [ExcelFunctionDoc(
            Name = "TL.crs.GeodeticCRS.Type",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets type of Geodetic Coordinate Reference System",
            HelpTopic = "TopoLib-AddIn.chm!1345",

            Returns = "CRS name",
            Summary = "Function that returns type of Geodetic CRS or &ltNotFound&gt if not found",
            Example = "TL.crs.GeodeticCRS.Type(3857) returns Geographic2DCrs",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object GeodeticCRS_Type(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            try
            {
                string type ="";

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            type = crs.GeodeticCRS.Type.ToString();

                        if (String.IsNullOrWhiteSpace(type))
                            type= "<NotFound>";

                        return type;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // GeodeticCRS_Type

        [ExcelFunctionDoc(
            Name = "TL.crs.PrimeMeridian.Name",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets name of prime meridean of coordinate reference system",
            HelpTopic = "TopoLib-AddIn.chm!1346",

            Returns = "Name of prime meridean of CRS",
            Summary = "Function that returns name of prime meridean of CRS or &ltNotFound&gt if not found",
            Example = "TL.crs.PrimeMeridian.Name(28992) returns Greenwich",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object PrimeMeridian_Name(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            try
            {
                string name ="";

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            name = crs.PrimeMeridian.Name;

                        if (String.IsNullOrWhiteSpace(name))
                            name = "<NotFound>";

                        return name;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // PrimeMeridian_Name

        [ExcelFunctionDoc(
            Name = "TL.crs.PrimeMeridian.Longitude",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets longitude of prime miridian in degrees",
            HelpTopic = "TopoLib-AddIn.chm!1347",

            Returns = "Longitude of prime miridian in degrees",
            Summary = "Function that returns longitude of prime miridian in degrees or -1 if not found",
            Example = "TL.crs.PrimeMeridian.Longitude(28992) returns 0.000",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object PrimeMeridian_Longitude(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            try
            {
                double res = -1;

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            res = crs.PrimeMeridian.Longitude;

                        return res;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // PrimeMeridian_Longitude

        [ExcelFunctionDoc(
            Name = "TL.crs.PrimeMeridian.UnitConversionFactor",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets unit conversion factor of prime miridian in degrees",
            HelpTopic = "TopoLib-AddIn.chm!1348",

            Returns = "Unit conversion factor of prime meridian ",
            Summary = "Function that returns unit conversion factor of prime meridian or -1 if not found",
            Example = "TL.crs.PrimeMeridian.UnitConversionFactor(28992) returns 0.0175",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object PrimeMeridian_UnitConversionFactor(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            try
            {
                double res = -1;

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            res = crs.PrimeMeridian.UnitConversionFactor;

                        return res;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // PrimeMeridian_UnitConversionFactor

        [ExcelFunctionDoc(
            Name = "TL.crs.PrimeMeridian.UnitName",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets name of prime meridean of coordinate reference system",
            HelpTopic = "TopoLib-AddIn.chm!1349",

            Returns = "Name of prime meridean of CRS",
            Summary = "Function that returns name of prime meridean of CRS or &ltNotFound&gt if not found",
            Example = "TL.crs.PrimeMeridian.UnitName(28992) returns degree",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object PrimeMeridian_UnitName(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            try
            {
                string name ="";

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            name = crs.PrimeMeridian.UnitName;

                        if (String.IsNullOrWhiteSpace(name))
                            name = "<NotFound>";

                        return name;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // PrimeMeridian_UnitName

        [ExcelFunctionDoc(
            Name = "TL.crs.Remarks",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets scope of Coordinate Reference System",
            HelpTopic = "TopoLib-AddIn.chm!1350",

            Returns = "CRS Remarks",
            Summary = "Function that returns remarks on a CRS or &ltNotFound&gt if not found",
            Example = "=TL.crs.Remarks(2036, TRUE) returns Axis order reversed compared to EPSG:2036",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object Remarks(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            try
            {
                string remarks ="";

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            remarks = crs.Remarks;

                        if (String.IsNullOrWhiteSpace(remarks))
                            remarks= "<NotFound>";

                        return remarks;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // Remarks

        [ExcelFunctionDoc(
            Name = "TL.crs.Scope",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets scope of Coordinate Reference System",
            HelpTopic = "TopoLib-AddIn.chm!1351",

            Returns = "CRS Scope",
            Summary = "Function that returns scope of CRS or &ltNotFound&gt if not found",
            Example = "TL.crs.Scope(2062) returns Engineering survey, topographic mapping.",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object Scope(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            try
            {
                string scope ="";

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            scope = crs.Scope;

                        if (String.IsNullOrWhiteSpace(scope))
                            scope = "<NotFound>";

                        return scope;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // Scope

        [ExcelFunctionDoc(
            Name = "TL.crs.Type",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets type of coordinate reference system",
            HelpTopic = "TopoLib-AddIn.chm!1352",

            Returns = "Ttype of coordinate reference system",
            Summary = "Function that returns type of coordinate reference system",
            Example = "TL.crs.Type(25833) returns ProjectedCrs",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object Type(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            try
            {
                object[,] res = new object [1, 1];

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                        {
                            var projType = crs.Type;
                            res[0, 0] = projType.ToString();
                        }

                        return res;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // UsageAreaCenterX

        [ExcelFunctionDoc(
            Name = "TL.crs.UsageArea.Center",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets center point of CRS usage area",
            HelpTopic = "TopoLib-AddIn.chm!1353",

            Returns = "Center point of CRS usage area",
            Summary = "Function that returns center point of CRS Usage Area in two adjacent cells",
            Example = "TL.crs.UsageArea.Center(2020) returns {321454.8, 4855381.7}",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object UsageArea_Center(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            try
            {
                object[,] res = new object [1, 2];

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null && crs.UsageArea != null)
                        {
                            res[0, 0] = crs.UsageArea.CenterX;
                            res[0, 1] = crs.UsageArea.CenterY;
                            return res;
                        }
                        else 
                            throw new ArgumentException("No Center Point found");  // Will return #VALUE to Excel
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // UsageArea_Center

        [ExcelFunctionDoc(
            Name = "TL.crs.UsageArea.Center.HasValues",
            Category = "CRS - Coordinate Reference System",
            Description = "Confirms whether the center point in the usage area has values",
            HelpTopic = "TopoLib-AddIn.chm!1354",

            Returns = "TRUE when the center point in the usage area has values; FALSE when not",
            Summary = "Function that confirms whether the center point in the usage area has values",
            Example = "TL.crs.UsageArea.Center.HasValues(2020) returns TRUE",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object UsageArea_Center_HasValues(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            try
            {
                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null && crs.UsageArea != null && crs.UsageArea.Center !=null)
                            return crs.UsageArea.Center.HasValues;

                        return false; 
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // UsageArea_Center_HasValues

        [ExcelFunctionDoc(
            Name = "TL.crs.UsageArea.Center.X",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets x-value of center point of CRS usage area",
            HelpTopic = "TopoLib-AddIn.chm!1355",

            Returns = "X-value of center point of CRS usage area",
            Summary = "Function that returns x-value of center point of CRS Usage Area",
            Example = "TL.crs.UsageArea.Center.X(2020) returns 321454.8",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object UsageArea_Center_X(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            try
            {
                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null && crs.UsageArea != null)
                            return crs.UsageArea.CenterX;
                        else 
                            return ExcelError.ExcelErrorValue;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // UsageArea_Center_X

        [ExcelFunctionDoc(
            Name = "TL.crs.UsageArea.Center.Y",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets y-value of center point of CRS usage area",
            HelpTopic = "TopoLib-AddIn.chm!1356",

            Returns = "Y-value of center point of CRS usage area",
            Summary = "Function that returns y-value of center point of CRS Usage Area",
            Example = "TL.crs.UsageArea.Center.Y(2020) returns 4855381.7",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object UsageArea_Center_Y(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            try
            {
                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null && crs.UsageArea != null)
                            return crs.UsageArea.CenterY;
                        else 
                            return ExcelError.ExcelErrorValue;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // UsageArea_Center_Y

        [ExcelFunctionDoc(
            Name = "TL.crs.UsageArea.MaxX",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets maximum X-value of CRS usage area",
            HelpTopic = "TopoLib-AddIn.chm!1357",

            Returns = "Maximum X-value of CRS usage area",
            Summary = "Function that returns maximum X-value of CRS Usage Area",
            Example = "=TL.crs.UsageArea.MaxX(7843) returns 173.34",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object UsageArea_MaxX(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            try
            {
                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null && crs.UsageArea != null)
                            return crs.UsageArea.MaxX;
                        else 
                            return ExcelError.ExcelErrorValue;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
        } // UsageArea_MaxX

        [ExcelFunctionDoc(
            Name = "TL.crs.UsageArea.MaxY",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets maximum Y-value of CRS usage area",
            HelpTopic = "TopoLib-AddIn.chm!1358",

            Returns = "Maximum Y-value of CRS usage area",
            Summary = "Function that returns maximum Y-value of CRS Usage Area",
            Example = "=TL.crs.UsageArea.MaxY(7843) returns -8.47",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object UsageArea_MaxY(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            try
            {
                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null && crs.UsageArea != null)
                            return crs.UsageArea.MaxY;
                        else 
                            return ExcelError.ExcelErrorValue;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
        } // UsageArea_MaxY

        [ExcelFunctionDoc(
            Name = "TL.crs.UsageArea.MinX",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets minimum X-value of CRS usage area",
            HelpTopic = "TopoLib-AddIn.chm!1359",

            Returns = "Minimum X-value of CRS usage area",
            Summary = "Function that returns minimum X-value of CRS Usage Area",
            Example = "=TL.crs.UsageArea.MinX(7843) returns 93.41",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object UsageArea_MinX(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            try
            {
                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null && crs.UsageArea != null)
                            return crs.UsageArea.MinX;
                        else 
                            return ExcelError.ExcelErrorValue;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
        } // UsageArea_MinX

        [ExcelFunctionDoc(
            Name = "TL.crs.UsageArea.MinY",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets minimum Y-value of CRS usage area",
            HelpTopic = "TopoLib-AddIn.chm!1360",

            Returns = "Minimum Y-value of CRS usage area",
            Summary = "Function that returns minimum Y-value of CRS Usage Area",
            Example = "=TL.crs.UsageArea.MinY(7843) returns -60.55",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object UsageArea_MinY(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            try
            {
                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null && crs.UsageArea != null)
                            return crs.UsageArea.MinY;
                        else 
                            return ExcelError.ExcelErrorValue;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
        } // UsageArea_MinY

        [ExcelFunctionDoc(
            Name = "TL.crs.UsageArea.Name",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets name of CRS usage area",
            HelpTopic = "TopoLib-AddIn.chm!1361",

            Returns = "Name of CRS usage area",
            Summary = "Function that returns name of usage area or &ltNotFound&gt if not found",
            Example = "TL.crs.UsageArea.Name(28992) returns Netherlands - onshore, including Waddenzee, Dutch Wadden Islands and 12-mile offshore coastal zone.",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object UsageArea_Name(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            try
            {
                string name ="";

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null && crs.UsageArea != null)
                            name = crs.UsageArea.Name;

                        if (String.IsNullOrWhiteSpace(name))
                            name = "<NotFound>";

                        return name;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // UsageArea_Name

        [ExcelFunctionDoc(
            Name = "TL.crs.UsageArea.WestLongitude",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets the west longitude of CRS usage area",
            HelpTopic = "TopoLib-AddIn.chm!1362",

            Returns = "West longitude of CRS usage area",
            Summary = "Function that returns west longitude of usage area or &ltNotFound&gt if not found",
            Example = "TL.crs.UsageArea.WestLongitude(28992) returns 3.20",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object UsageArea_WestLongitude(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            try
            {
                double west = 0.0;

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null && crs.UsageArea != null)
                            west = crs.UsageArea.WestLongitude;

                        return west;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // UsageArea_WestLongitude

        [ExcelFunctionDoc(
            Name = "TL.crs.UsageArea.EastLongitude",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets east longitude of CRS usage area",
            HelpTopic = "TopoLib-AddIn.chm!1363",

            Returns = "East longitude of CRS usage area",
            Summary = "Function that returns east longitude of usage area or &ltNotFound&gt if not found",
            Example = "TL.crs.UsageArea.EastLongitude(28992) returns 7.22",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object UsageArea_EastLongitude(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            try
            {
                double east = 0.0;

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null && crs.UsageArea != null)
                            east = crs.UsageArea.EastLongitude;

                        return east;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // UsageArea_EastLongitude

        [ExcelFunctionDoc(
            Name = "TL.crs.UsageArea.SouthLatitude",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets south latitude of CRS usage area",
            HelpTopic = "TopoLib-AddIn.chm!1364",

            Returns = "South latitude of CRS usage area",
            Summary = "Function that returns south latitude of usage area or &ltNotFound&gt if not found",
            Example = "TL.crs.UsageArea.SouthLatitude(28992) returns 50.75",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object UsageArea_SouthLatitude(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            try
            {
                double south = 0.0;

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null && crs.UsageArea != null)
                            south = crs.UsageArea.SouthLatitude;

                        return south;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // UsageArea_SouthLatitude

        [ExcelFunctionDoc(
            Name = "TL.crs.UsageArea.NorthLatitude",
            Category = "CRS - Coordinate Reference System",
            Description = "Gets north latitude of CRS usage area",
            HelpTopic = "TopoLib-AddIn.chm!1365",

            Returns = "South latitude of CRS usage area",
            Summary = "Function that returns north latitude of usage area or &ltNotFound&gt if not found",
            Example = "TL.crs.UsageArea.NorthLatitude(28992) returns 53.70",
            Remarks = "Setting Normalized 'true' reflects coordinate ordering as generally used in TopoLib"
            )]
        public static object UsageArea_NorthLatitude(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code, WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Use normalized axis ordering: {long, lat} & {x, y} (true)", Name = "Normalized")] object normalized
            )
        {
            bool bNorm  = Optional.Check(normalized, true);

            try
            {
                double north = 0.0;

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = bNorm ? CreateCrs(oCrs, pjContext).WithAxisNormalized() : CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null && crs.UsageArea != null)
                            north = crs.UsageArea.NorthLatitude;

                        return north;
                    }
                }        
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // UsageArea_NorthLatitude

    } // class

} // namespace

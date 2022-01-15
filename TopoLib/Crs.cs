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
            ProjContext pjContext = new ProjContext { LogLevel = (ProjLogLevel) Lib.LogLevel};
            pjContext.Log += Lib.SendToSerilog;
            return pjContext;
        }

        internal static CoordinateReferenceSystem CreateCrs(in object[,] oCrs, in ProjContext pjContext)
        {
            int nCrsRows = oCrs.GetLength(0);
            int nCrsCols = oCrs.GetLength(1);

            // max two adjacent CRS cells on the same row
            if (nCrsRows != 1 || nCrsCols > 2 ) 
                throw new ArgumentException("CRS");

            int nCrs;
            string sCrs;

            // We have only one cell; it can be aa Wkt string, a JSON string, a PROJ string or a textual description
            if (nCrsCols == 1)
            {
                // we have one cell describing the crs.
                if (oCrs[0,0] is double)
                {
                    // First cast to double, then to int, to deal with Excel datatypes
                    nCrs = (int)(double)oCrs[0, 0];   

                    // we have an EPSG number from a single input parameter:
                    return CoordinateReferenceSystem.CreateFromEpsg(nCrs, pjContext);
                }
                else if (oCrs[0,0] is string)
                {
                    // cast to string, to deal with Excel datatypes
                    sCrs = (string)oCrs[0,0];

                    bool success = int.TryParse(sCrs, out nCrs);
                    if (success)
                    {
                        // we have an EPSG number from a single input parameter:
                        return CoordinateReferenceSystem.CreateFromEpsg(nCrs, pjContext);
                    }
                    else
                    {
                        // we have a string of some sorts and a single input parameter:
                        if ((sCrs.IndexOf("PROJCS") > -1) || (sCrs.IndexOf("GEOGCS") > -1) || (sCrs.IndexOf("SPHEROID") > -1))
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
                    throw new ArgumentNullException("CRS");
            }
            else
            {
                // we have two adjacent CRS cells; first an Authortity string of some sorts and a second input parameter (number):

                sCrs = (string)oCrs[0,0];   // the authority string

                // try to get the crs number; if not succesful throw an exception

                if (oCrs[0, 1] is double)
                {
                    // First cast to double, then to int, to deal with Excel datatypes
                    nCrs = (int)(double)oCrs[0, 1];

                }
                else if (oCrs[0, 1] is string)
                {
                    // cast to string, to deal with Excel datatypes
                    string sTmp = (string)oCrs[0, 1];

                    bool success = int.TryParse(sTmp, out nCrs);
                    if (!success) 
                        throw new ArgumentNullException("CRS");
                }
                else
                    throw new ArgumentNullException("CRS");

                return CoordinateReferenceSystem.CreateFromDatabase(sCrs, nCrs, pjContext);

                // CreateFromDatabase() is translated into proj_create_from_database 
                // It may throw a ArgumentNullException
                // It may throw a pjContext->ConstructException
            }

            // Oops, something went wrong if we get here...
            throw new ArgumentNullException("CRS");

        } // GetCrs

        [ExcelFunctionDoc(
             Name = "TL.crs.CoordinateSystem.Axis.Count",
             Category = "CRS - Coordinate Reference System",
             Description = "Gets type of Coordinate System",
             HelpTopic = "TopoLib-AddIn.chm!1300",

             Returns = "Coordinate name of coordinate system",
             Summary = "Function that returns nr of axes in of Coordinate System or -1 if not found",
             Example = "xxx"
         )]
        public static object CoordinateSystemAxisCount(
             [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs
            )
        {
            try
            {
                int nAxes = -1;

                using (ProjContext pjContext = CreateContext())
                {
                    using (CoordinateReferenceSystem crs = CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null & crs.CoordinateSystem != null)
                            nAxes = crs.CoordinateSystem.Axis.Count;

                        return nAxes;
                    }
                }        
            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex); 
            }

        } // CoordinateSystemType

        [ExcelFunctionDoc(
            Name = "TL.crs.CoordinateSystem.Axis.Name",
            Category = "CRS - Coordinate Reference System",
            Description = "Get the name of axis-N of a coordinate system", 
            HelpTopic = "TopoLib-AddIn.chm!1301",

            Returns = "name of Nth axis in a coordinate system",
            Summary = "Returns the name of Nth axis in a coordinate system",
            Example = "xxx")]
        public static object CoordinateSystem_Axis_Name(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Zero based index of Identifier list (0) ", Name = "Index")] object index)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            int nIndex  = (int)Optional.Check(index , 0.0);
            string name ="";

            // do the work
            try
            {
                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = CreateCrs(oCrs, pjContext))
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
                return Lib.ExceptionHandler(ex);
            }

        } // CoordinateSystem_Axis_Name

        [ExcelFunctionDoc(
            Name = "TL.crs.CoordinateSystem.Axis.Abbreviation",
            Category = "CRS - Coordinate Reference System",
            Description = "Get the short name of axis-N of a coordinate system", 
            HelpTopic = "TopoLib-AddIn.chm!1302",

            Returns = "short name of Nth axis in a coordinate system",
            Summary = "Returns the short name of Nth axis in a coordinate system",
            Example = "xxx")]
        public static object CoordinateSystem_Axis_Abbreviation(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Zero based index of Axis list (0) ", Name = "Index")] object index)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            int nIndex  = (int)Optional.Check(index , 0.0);
            string name ="";

            // do the work
            try
            {
                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = CreateCrs(oCrs, pjContext))
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
                return Lib.ExceptionHandler(ex);
            }

        } // CoordinateSystem_Axis_Abbreviation

        [ExcelFunctionDoc(
            Name = "TL.crs.CoordinateSystem.Axis.UnitName",
            Category = "CRS - Coordinate Reference System",
            Description = "Get the unit name of axis-N of a coordinate system", 
            HelpTopic = "TopoLib-AddIn.chm!1303",

            Returns = "unit name of Nth axis in a coordinate system",
            Summary = "Returns the unit name of Nth axis in a coordinate system",
            Example = "xxx")]
        public static object CoordinateSystem_Axis_UnitName(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Zero based index of Axis list (0) ", Name = "Index")] object index)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            int nIndex  = (int)Optional.Check(index , 0.0);
            string name ="";

            // do the work
            try
            {
                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = CreateCrs(oCrs, pjContext))
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
                return Lib.ExceptionHandler(ex);
            }

        } // CoordinateSystem_Axis_UnitName

        [ExcelFunctionDoc(
            Name = "TL.crs.CoordinateSystem.Axis.UnitAuthName",
            Category = "CRS - Coordinate Reference System",
            Description = "Get the authority name of axis-N of a coordinate system", 
            HelpTopic = "TopoLib-AddIn.chm!1304",

            Returns = "authority name of Nth axis in a coordinate system",
            Summary = "Returns the authority name of Nth axis in a coordinate system",
            Example = "xxx")]
        public static object CoordinateSystem_Axis_UnitAuthName(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Zero based index of Axis list (0) ", Name = "Index")] object index)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            int nIndex  = (int)Optional.Check(index , 0.0);
            string name ="";

            // do the work
            try
            {
                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = CreateCrs(oCrs, pjContext))
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
                return Lib.ExceptionHandler(ex);
            }

        } // CoordinateSystem_Axis_UnitAuthName

        [ExcelFunctionDoc(
            Name = "TL.crs.CoordinateSystem.Axis.UnitConversionFactor",
            Category = "CRS - Coordinate Reference System",
            Description = "Get the Unit Conversion Factor of axis-N of a coordinate system", 
            HelpTopic = "TopoLib-AddIn.chm!1305",

            Returns = "Unit Conversion Factorof Nth axis in a coordinate system",
            Summary = "Returns the Unit Conversion Factor of Nth axis in a coordinate system",
            Example = "xxx")]
        public static object CoordinateSystem_Axis_UnitConversionFactor(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Zero based index of Axis list (0) ", Name = "Index")] object index)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            int nIndex  = (int)Optional.Check(index , 0.0);
            double factor = 0;

            // do the work
            try
            {
                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            factor = crs.CoordinateSystem.Axis[nIndex].UnitConversionFactor;

                        return factor;
                    }
                }        
            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }

        } // CoordinateSystem_Axis_UnitConversionFactor


        [ExcelFunctionDoc(
            Name = "TL.crs.CoordinateSystem.Axis.UnitCode",
            Category = "CRS - Coordinate Reference System",
            Description = "Get the authority name of axis-N of a coordinate system", 
            HelpTopic = "TopoLib-AddIn.chm!1306",

            Returns = "authority name of Nth axis in a coordinate system",
            Summary = "Returns the authority name of Nth axis in a coordinate system",
            Example = "xxx")]
        public static object CoordinateSystem_Axis_UnitCode(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Zero based index of Axis list (0) ", Name = "Index")] object index)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            int nIndex  = (int)Optional.Check(index , 0.0);
            string name ="";

            // do the work
            try
            {
                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = CreateCrs(oCrs, pjContext))
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
                return Lib.ExceptionHandler(ex);
            }

        } // CoordinateSystem_Axis_UnitCode

        [ExcelFunctionDoc(
            Name = "TL.crs.CoordinateSystem.Axis.Direction",
            Category = "CRS - Coordinate Reference System",
            Description = "Get the direction of axis-N of a coordinate system", 
            HelpTopic = "TopoLib-AddIn.chm!1307",

            Returns = "Directionof Nth axis in a coordinate system",
            Summary = "Returns the direction of Nth axis in a coordinate system",
            Example = "xxx")]
        public static object CoordinateSystem_Axis_Direction(
            [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
            [ExcelArgument("Zero based index of Axis list (0) ", Name = "Index")] object index)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;

            int nIndex  = (int)Optional.Check(index , 0.0);
            string name ="";

            // do the work
            try
            {
                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = CreateCrs(oCrs, pjContext))
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
                return Lib.ExceptionHandler(ex);
            }

        } // CoordinateSystem_Axis_Direction

        [ExcelFunctionDoc(
             Name = "TL.crs.CoordinateSystem.Name",
             Category = "CRS - Coordinate Reference System",
             Description = "Gets type of Coordinate System",
             HelpTopic = "TopoLib-AddIn.chm!1308",

             Returns = "Coordinate name of coordinate system",
             Summary = "Function that returns name of Coordinate System or &ltNotFound&gt if not found",
             Example = "xxx"
         )]
        public static object CoordinateSystem_Name(
             [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs
            )
        {
            try
            {
                string name ="";

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = CreateCrs(oCrs, pjContext))
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
                return Lib.ExceptionHandler(ex);
            }

        } // CoordinateSystem_Name

        [ExcelFunctionDoc(
             Name = "TL.crs.CoordinateSystem.Type",
             Category = "CRS - Coordinate Reference System",
             Description = "Gets type of Coordinate System",
             HelpTopic = "TopoLib-AddIn.chm!1309",

             Returns = "Coordinate system type",
             Summary = "Function that returns type of Coordinate System or &ltNotFound&gt if not found",
             Example = "xxx"
         )]
        public static object CoordinateSystem_Type(
             [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs
            )
        {
            try
            {
                string type ="";

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = CreateCrs(oCrs, pjContext))
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
                return Lib.ExceptionHandler(ex);
            }

        } // CoordinateSystem_Type

        [ExcelFunctionDoc(
             Name = "TL.crs.CoordinateSystem.CoordinateSystemType",
             Category = "CRS - Coordinate Reference System",
             Description = "Gets type of Coordinate System",
             HelpTopic = "TopoLib-AddIn.chm!1310",

             Returns = "Coordinate System Type",
             Summary = "Function that returns Coordinate System Type of Coordinate System or &ltNotFound&gt if not found",
             Example = "xxx"
         )]
        public static object CoordinateSystem_CoordinateSystemType(
             [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs
            )
        {
            try
            {
                string type ="";

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = CreateCrs(oCrs, pjContext))
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
                return Lib.ExceptionHandler(ex);
            }

        } // CoordinateSystem_CoordinateSystemType

        [ExcelFunctionDoc(
             Name = "TL.crs.Datum.Name",
             Category = "CRS - Coordinate Reference System",
             Description = "Gets name of datum of CRS",
             HelpTopic = "TopoLib-AddIn.chm!1311",

             Returns = "CRS name",
             Summary = "Function that returns name of datum of CRS or &ltNotFound&gt if not found",
             Example = "xxx"
         )]
        public static object Datum_Name(
             [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs
            )
        {
            try
            {
                string name ="";

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = CreateCrs(oCrs, pjContext))
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
                return Lib.ExceptionHandler(ex);
            }

        } // Datum_Name

        [ExcelFunctionDoc(
             Name = "TL.crs.Datum.Type",
             Category = "CRS - Coordinate Reference System",
             Description = "Gets type of datum in CRS",
             HelpTopic = "TopoLib-AddIn.chm!1312",

             Returns = "Datum type",
             Summary = "Function that returns type of datum of CRS or &ltNotFound&gt if not found",
             Example = "xxx"
         )]
        public static object Datum_Type(
             [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs
            )
        {
            try
            {
                string type ="";

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = CreateCrs(oCrs, pjContext))
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
                return Lib.ExceptionHandler(ex);
            }

        } // Datum_Type

        [ExcelFunctionDoc(
             Name = "TL.crs.Ellipsoid.Name",
             Category = "CRS - Coordinate Reference System",
             Description = "Gets name of Ellipsoid in Coordinate Reference System",
             HelpTopic = "TopoLib-AddIn.chm!1313",

             Returns = "Ellipsoid name",
             Summary = "Function that returns name of Ellipsoid in CRS or &ltNotFound&gt if not found",
             Example = "xxx"
         )]
        public static object Ellipsoid_Name(
             [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs
            )
        {
            try
            {
                string name ="";

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = CreateCrs(oCrs, pjContext))
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
                return Lib.ExceptionHandler(ex);
            }

        } // Ellipsoid_Name

        [ExcelFunctionDoc(
             Name = "TL.crs.Ellipsoid.Type",
             Category = "CRS - Coordinate Reference System",
             Description = "Gets type of ellipsoid in Coordinate Reference System",
             HelpTopic = "TopoLib-AddIn.chm!1314",

             Returns = "Ellipsoid type",
             Summary = "Function that returns type of ellipsoid in CRS or &ltNotFound&gt if not found",
             Example = "xxx"
         )]
        public static object Ellipsoid_Type(
             [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs
            )
        {
            try
            {
                string type ="";

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = CreateCrs(oCrs, pjContext))
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
                return Lib.ExceptionHandler(ex);
            }

        } // Ellipsoid_Type

        [ExcelFunctionDoc(
             Name = "TL.crs.Ellipsoid.SemiMajorMetre",
             Category = "CRS - Coordinate Reference System",
             Description = "Gets size of ellipsoid's semi-major axis in metres",
             HelpTopic = "TopoLib-AddIn.chm!1315",

             Returns = "Size of ellipsoid's semi-major axis in metres",
             Summary = "Function that returns size of ellipsoid's semi-major axis in metres or -1 if not found",
             Example = "xxx"
         )]
        public static object Ellipsoid_SemiMajorMetre(
             [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs
            )
        {
            try
            {
                double res = -1;

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            res = crs.Ellipsoid.SemiMajorMetre;

                        return res;
                    }
                }        
            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }

        } // Ellipsoid_SemiMajorMetre

        [ExcelFunctionDoc(
             Name = "TL.crs.Ellipsoid.SemiMinorMetre",
             Category = "CRS - Coordinate Reference System",
             Description = "Gets size of ellipsoid's semi-minor axis in metres",
             HelpTopic = "TopoLib-AddIn.chm!1316",

             Returns = "Size of ellipsoid's semi-minor axis in metres",
             Summary = "Function that returns size of ellipsoid's semi-minor axis in metres or -1 if not found",
             Example = "xxx"
         )]
        public static object Ellipsoid_SemiMinorMetre(
             [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs
            )
        {
            try
            {
                double res = -1;

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            res = crs.Ellipsoid.SemiMinorMetre;

                        return res;
                    }
                }        
            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }

        } // Ellipsoid_SemiMinorMetre

        [ExcelFunctionDoc(
             Name = "TL.crs.Ellipsoid.InverseFlattening",
             Category = "CRS - Coordinate Reference System",
             Description = "Gets inverse flattening of ellipsoid",
             HelpTopic = "TopoLib-AddIn.chm!1317",

             Returns = "Inverse flattening of ellipsoid",
             Summary = "Function that returns inverse flattening of ellipsoid, or -1 if not found",
             Example = "xxx"
         )]
        public static object Ellipsoid_InverseFlattening(
             [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs
            )
        {
            try
            {
                double res = -1;

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            res = crs.Ellipsoid.InverseFlattening;

                        return res;
                    }
                }        
            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }

        } // Ellipsoid_InverseFlattening

        [ExcelFunctionDoc(
             Name = "TL.crs.Ellipsoid.IsSemiMinorComputed",
             Category = "CRS - Coordinate Reference System",
             Description = "Confirms whether when size of semi-minor axis has been calculated",
             HelpTopic = "TopoLib-AddIn.chm!1318",

             Returns = "TRUE when size of semi-minor axis has been calculated; FALSE when not",
             Summary = "Function that confirms whether when size of semi-minor axis has been calculated",
             Example = "xxx"
         )]
        public static object Ellipsoid_IsSemiMinorComputed(
             [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs
            )
        {
            try
            {
                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            return crs.Ellipsoid.IsSemiMinorComputed;

                        return ExcelError.ExcelErrorValue; 
                    }
                }        
            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }

        } // EllipsoidIsSemiMinorComputed

        [ExcelFunctionDoc(
             Name = "TL.crs.Identifiers.Code",
             Category = "CRS - Coordinate Reference System",
             Description = "Gets Authority of Identifier N",
             HelpTopic = "TopoLib-AddIn.chm!1319",

             Returns = "Code of Nth Identifiers",
             Summary = "Function that returns the Code of the <Nth> identifier or <index out of range> when not found",
             Example = "xxx"
         )]
        public static object Identifiers_Code(
             [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
             [ExcelArgument("Zero based index of Identifier list (0) ", Name = "Index")] object index)
        {
            try
            {
                int nIndex  = (int)Optional.Check(index , 0.0);

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = CreateCrs(oCrs, pjContext))
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
                return Lib.ExceptionHandler(ex);
            }

        } // Identifiers_Code

        [ExcelFunctionDoc(
             Name = "TL.crs.Identifiers.Authority",
             Category = "CRS - Coordinate Reference System",
             Description = "Gets Authority of Identifier N",
             HelpTopic = "TopoLib-AddIn.chm!1320",

             Returns = "Authority of Nth Identifiers",
             Summary = "Function that returns Authority of <Nth> identifiers or <index out of range> when not found",
             Example = "xxx"
         )]
        public static object Identifiers_Authority(
             [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs,
             [ExcelArgument("Zero based index of Identifier list (0) ", Name = "Index")] object index)
        {
            try
            {
                int nIndex  = (int)Optional.Check(index , 0.0);

                int Count = -1;

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = CreateCrs(oCrs, pjContext))
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
                return Lib.ExceptionHandler(ex);
            }

        } // Identifiers_Authority

        [ExcelFunctionDoc(
             Name = "TL.crs.Identifiers.Count",
             Category = "CRS - Coordinate Reference System",
             Description = "Gets number of Identifiers",
             HelpTopic = "TopoLib-AddIn.chm!1321",

             Returns = "Number of CRS Identifiers",
             Summary = "Function that returns nr of CRS identifiers or 0 if none found",
             Example = "xxx"
         )]
        public static object Identifiers_Count(
             [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs
            )
        {
            try
            {
                int Count = -0;

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null && crs.Identifiers != null)
                            Count = crs.Identifiers.Count;

                        return Count;
                    }
                }        
            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }

        } // Identifiers_Count

        [ExcelFunctionDoc(
             Name = "TL.crs.IsDeprecated",
             Category = "CRS - Coordinate Reference System",
             Description = "Confirms whether when size of semi-minor axis has been calculated",
             HelpTopic = "TopoLib-AddIn.chm!1322",

             Returns = "TRUE when CRS is deprecated; FALSE when not",
             Summary = "Function that confirms whether the CRS is deprecated",
             Example = "xxx"
         )]
        public static object IsDeprecated(
             [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs
            )
        {
            try
            {
                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            return crs.IsDeprecated;

                        return ExcelError.ExcelErrorValue; 
                    }
                }        
            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }

        } // IsDeprecated


        [ExcelFunctionDoc(
             Name = "TL.crs.Name",
             Category = "CRS - Coordinate Reference System",
             Description = "Gets name of Coordinate Reference System",
             HelpTopic = "TopoLib-AddIn.chm!1323",

             Returns = "CRS name",
             Summary = "Function that returns name of CRS or &ltNotFound&gt if not found",
             Example = "xxx"
         )]
        public static object Name(
             [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs
            )
        {
            try
            {
                string name ="";

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = CreateCrs(oCrs, pjContext))
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
                return Lib.ExceptionHandler(ex);
            }

        } // Name

        [ExcelFunctionDoc(
             Name = "TL.crs.GeodeticCRS.Name",
             Category = "CRS - Coordinate Reference System",
             Description = "Gets name of Geodetic Coordinate Reference System",
             HelpTopic = "TopoLib-AddIn.chm!1324",

             Returns = "CRS name",
             Summary = "Function that returns name of Geodetic CRS or &ltNotFound&gt if not found",
             Example = "xxx"
         )]
        public static object GeodeticCRS_Name(
             [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs
            )
        {
            try
            {
                string name ="";

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = CreateCrs(oCrs, pjContext))
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
                return Lib.ExceptionHandler(ex);
            }

        } // GeodeticCRS_Name

        [ExcelFunctionDoc(
             Name = "TL.crs.GeodeticCRS.Type",
             Category = "CRS - Coordinate Reference System",
             Description = "Gets type of Geodetic Coordinate Reference System",
             HelpTopic = "TopoLib-AddIn.chm!1325",

             Returns = "CRS name",
             Summary = "Function that returns type of Geodetic CRS or &ltNotFound&gt if not found",
             Example = "xxx"
         )]
        public static object GeodeticCRS_Type(
             [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs
            )
        {
            try
            {
                string type ="";

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = CreateCrs(oCrs, pjContext))
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
                return Lib.ExceptionHandler(ex);
            }

        } // GeodeticCRS_Type

        [ExcelFunctionDoc(
             Name = "TL.crs.PrimeMeridian.Name",
             Category = "CRS - Coordinate Reference System",
             Description = "Gets name of prime meridean of coordinate reference system",
             HelpTopic = "TopoLib-AddIn.chm!1326",

             Returns = "Name of prime meridean of CRS",
             Summary = "Function that returns name of prime meridean of CRS or &ltNotFound&gt if not found",
             Example = "xxx"
         )]
        public static object PrimeMeridian_Name(
             [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs
            )
        {
            try
            {
                string name ="";

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = CreateCrs(oCrs, pjContext))
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
                return Lib.ExceptionHandler(ex);
            }

        } // PrimeMeridian_Name

        [ExcelFunctionDoc(
             Name = "TL.crs.PrimeMeridian.Longitude",
             Category = "CRS - Coordinate Reference System",
             Description = "Gets longitude of prime miridian in degrees",
             HelpTopic = "TopoLib-AddIn.chm!1327",

             Returns = "Longitude of prime miridian in degrees",
             Summary = "Function that returns longitude of prime miridian in degrees or -1 if not found",
             Example = "xxx"
         )]
        public static object PrimeMeridian_Longitude(
             [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs
            )
        {
            try
            {
                double res = -1;

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            res = crs.PrimeMeridian.Longitude;

                        return res;
                    }
                }        
            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }

        } // PrimeMeridian_Longitude

        [ExcelFunctionDoc(
             Name = "TL.crs.PrimeMeridian.UnitConversionFactor",
             Category = "CRS - Coordinate Reference System",
             Description = "Gets unit conversion factor of prime miridian in degrees",
             HelpTopic = "TopoLib-AddIn.chm!1328",

             Returns = "Unit conversion factor of prime meridian ",
             Summary = "Function that returns unit conversion factor of prime meridian or -1 if not found",
             Example = "xxx"
         )]
        public static object PrimeMeridian_UnitConversionFactor(
             [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs
            )
        {
            try
            {
                double res = -1;

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            res = crs.PrimeMeridian.UnitConversionFactor;

                        return res;
                    }
                }        
            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }

        } // PrimeMeridian_UnitConversionFactor

        [ExcelFunctionDoc(
             Name = "TL.crs.PrimeMeridian.UnitName",
             Category = "CRS - Coordinate Reference System",
             Description = "Gets name of prime meridean of coordinate reference system",
             HelpTopic = "TopoLib-AddIn.chm!1329",

             Returns = "Name of prime meridean of CRS",
             Summary = "Function that returns name of prime meridean of CRS or &ltNotFound&gt if not found",
             Example = "xxx"
         )]
        public static object PrimeMeridian_UnitName(
             [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs
            )
        {
            try
            {
                string name ="";

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = CreateCrs(oCrs, pjContext))
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
                return Lib.ExceptionHandler(ex);
            }

        } // PrimeMeridian_UnitName

        [ExcelFunctionDoc(
             Name = "TL.crs.Scope",
             Category = "CRS - Coordinate Reference System",
             Description = "Gets scope of Coordinate Reference System",
             HelpTopic = "TopoLib-AddIn.chm!1330",

             Returns = "CRS Scope",
             Summary = "Function that returns scope of CRS or &ltNotFound&gt if not found",
             Example = "xxx"
         )]
        public static object Scope(
             [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs
            )
        {
            try
            {
                string scope ="";

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = CreateCrs(oCrs, pjContext))
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
                return Lib.ExceptionHandler(ex);
            }

        } // Scope

        [ExcelFunctionDoc(
            Name = "TL.crs.ToJsonString",
            Category = "CRS - Coordinate Reference System",
            Description = "Converts input Coordinate Reference System to a JSON string",
            HelpTopic = "TopoLib-AddIn.chm!1331",

            Returns = "A JSON string",
            Summary = "Function that converts input Coordinate Reference System to a JSON string",
            Example = "xxx"
        )]
        public static object ToJsonString(
             [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs
            )
        {
            try
            {
                string CrsString ="";

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = CreateCrs(oCrs, pjContext))
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
                return Lib.ExceptionHandler(ex);
            }

        } // ToJsonString

        [ExcelFunctionDoc(
            Name = "TL.crs.ToWktString",
            Category = "CRS - Coordinate Reference System",
            Description = "Converts input Coordinate Reference System to a WKT string",
            HelpTopic = "TopoLib-AddIn.chm!1332",

            Returns = "A WKT string",
            Summary = "Function that converts input Coordinate Reference System to a WKT string",
            Example = "xxx"
        )]
        public static object ToWktString(
             [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs
            )
        {
            try
            {
                string CrsString ="";

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = CreateCrs(oCrs, pjContext))
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
                return Lib.ExceptionHandler(ex);
            }

        } // ToWktString

        [ExcelFunctionDoc(
            Name = "TL.crs.ToProjString",
            Category = "CRS - Coordinate Reference System",
            Description = "Converts input Coordinate Reference System to a Proj string",
            HelpTopic = "TopoLib-AddIn.chm!1333",

            Returns = "A Proj string",
            Summary = "Function that converts input Coordinate Reference System to a PROJ string",
            Example = "xxx"
        )]
        public static object ToProjString(
             [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs
            )
        {
            try
            {
                string CrsString ="";

                using (ProjContext pjContext = CreateContext())
                {
                    using ( var crs = CreateCrs(oCrs, pjContext))
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
                return Lib.ExceptionHandler(ex);
            }

        } // ToProjString

        [ExcelFunctionDoc(
             Name = "TL.crs.Type",
             Category = "CRS - Coordinate Reference System",
             Description = "Gets type of coordinate reference system",
             HelpTopic = "TopoLib-AddIn.chm!1334",

             Returns = "Ttype of coordinate reference system",
             Summary = "Function that returns type of coordinate reference system",
             Example = "xxx"
         )]
        public static object Type(
             [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs
            )
        {
            try
            {
                object[,] res = new object [1, 1];

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = CreateCrs(oCrs, pjContext))
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
                return Lib.ExceptionHandler(ex);
            }

        } // UsageAreaCenterX

        [ExcelFunctionDoc(
             Name = "TL.crs.UsageArea.Center",
             Category = "CRS - Coordinate Reference System",
             Description = "Gets center point of CRS usage area",
             HelpTopic = "TopoLib-AddIn.chm!1335",

             Returns = "Center point of CRS usage area",
             Summary = "Function that returns center point of CRS Usage Area in two adjacent cells",
             Example = "xxx"
         )]
        public static object UsageArea_Center(
             [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs
            )
        {
            try
            {
                object[,] res = new object [1, 2];

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                        {
                            res[0, 0] = crs.UsageArea.CenterX;
                            res[0, 1] = crs.UsageArea.CenterY;
                        }

                        return res;
                    }
                }        
            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }

        } // UsageArea_Center

        [ExcelFunctionDoc(
             Name = "TL.crs.UsageArea.Center.HasValues",
             Category = "CRS - Coordinate Reference System",
             Description = "Confirms whether the center point in the usage area has values",
             HelpTopic = "TopoLib-AddIn.chm!1336",

             Returns = "TRUE when the center point in the usage area has values; FALSE when not",
             Summary = "Function that confirms whether the center point in the usage area has values",
             Example = "xxx"
         )]
        public static object UsageArea_Center_HasValues(
             [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs
            )
        {
            try
            {
                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null && crs.UsageArea != null && crs.UsageArea.Center !=null)
                            return crs.UsageArea.Center.HasValues;

                        return false; 
                    }
                }        
            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }

        } // UsageArea_Center_HasValues

        [ExcelFunctionDoc(
             Name = "TL.crs.UsageArea.Center.X",
             Category = "CRS - Coordinate Reference System",
             Description = "Gets x-value of center point of CRS usage area",
             HelpTopic = "TopoLib-AddIn.chm!1337",

             Returns = "X-value of center point of CRS usage area",
             Summary = "Function that returns x-value of center point of CRS Usage Area",
             Example = "xxx"
         )]
        public static object UsageArea_Center_X(
             [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs
            )
        {
            try
            {
                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = CreateCrs(oCrs, pjContext))
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
                return Lib.ExceptionHandler(ex);
            }

        } // UsageArea_Center_X

        [ExcelFunctionDoc(
             Name = "TL.crs.UsageArea.Center.Y",
             Category = "CRS - Coordinate Reference System",
             Description = "Gets y-value of center point of CRS usage area",
             HelpTopic = "TopoLib-AddIn.chm!1338",

             Returns = "Y-value of center point of CRS usage area",
             Summary = "Function that returns y-value of center point of CRS Usage Area",
             Example = "xxx"
         )]
        public static object UsageArea_Center_Y(
             [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs
            )
        {
            try
            {
                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = CreateCrs(oCrs, pjContext))
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
                return Lib.ExceptionHandler(ex);
            }

        } // UsageArea_Center_Y

        [ExcelFunctionDoc(
             Name = "TL.crs.UsageArea.Name",
             Category = "CRS - Coordinate Reference System",
             Description = "Gets name of CRS usage area",
             HelpTopic = "TopoLib-AddIn.chm!1339",

             Returns = "Name of CRS usage area",
             Summary = "Function that returns name of usage area or &ltNotFound&gt if not found",
             Example = "xxx"
         )]
        public static object UsageArea_Name(
             [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs
            )
        {
            try
            {
                string name ="";

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            name = crs.UsageArea.Name;

                        if (String.IsNullOrWhiteSpace(name))
                            name = "<NotFound>";

                        return name;
                    }
                }        
            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }

        } // UsageArea_Name

        [ExcelFunctionDoc(
             Name = "TL.crs.UsageArea.WestLongitude",
             Category = "CRS - Coordinate Reference System",
             Description = "Gets west longitude of CRS usage area",
             HelpTopic = "TopoLib-AddIn.chm!1340",

             Returns = "West longitude of CRS usage area",
             Summary = "Function that returns west longitude of usage area or &ltNotFound&gt if not found",
             Example = "xxx"
         )]
        public static object UsageArea_WestLongitude(
             [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs
            )
        {
            try
            {
                double west = 0.0;

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            west = crs.UsageArea.WestLongitude;

                        return west;
                    }
                }        
            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }

        } // UsageArea_WestLongitude

        [ExcelFunctionDoc(
             Name = "TL.crs.UsageArea.EastLongitude",
             Category = "CRS - Coordinate Reference System",
             Description = "Gets east longitude of CRS usage area",
             HelpTopic = "TopoLib-AddIn.chm!1341",

             Returns = "East longitude of CRS usage area",
             Summary = "Function that returns east longitude of usage area or &ltNotFound&gt if not found",
             Example = "xxx"
         )]
        public static object UsageAreaEastLongitude(
             [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs
            )
        {
            try
            {
                double east = 0.0;

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            east = crs.UsageArea.EastLongitude;

                        return east;
                    }
                }        
            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }

        } // UsageArea_EastLongitude

        [ExcelFunctionDoc(
             Name = "TL.crs.UsageArea.SouthLatitude",
             Category = "CRS - Coordinate Reference System",
             Description = "Gets south latitude of CRS usage area",
             HelpTopic = "TopoLib-AddIn.chm!1342",

             Returns = "South latitude of CRS usage area",
             Summary = "Function that returns south latitude of usage area or &ltNotFound&gt if not found",
             Example = "xxx"
         )]
        public static object UsageArea_SouthLatitude(
             [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs
            )
        {
            try
            {
                double south = 0.0;

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            south = crs.UsageArea.SouthLatitude;

                        return south;
                    }
                }        
            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }

        } // UsageArea_SouthLatitude

        [ExcelFunctionDoc(
             Name = "TL.crs.UsageArea.NorthLatitude",
             Category = "CRS - Coordinate Reference System",
             Description = "Gets north latitude of CRS usage area",
             HelpTopic = "TopoLib-AddIn.chm!1343",

             Returns = "South latitude of CRS usage area",
             Summary = "Function that returns north latitude of usage area or &ltNotFound&gt if not found",
             Example = "xxx"
         )]
        public static object UsageArea_NorthLatitude(
             [ExcelArgument("One [or two adjacent] cell[s] with [Authority and] EPSG code (4326), WKT string, JSON string or PROJ string", Name = "Crs")] object[,] oCrs
            )
        {
            try
            {
                double north = 0.0;

                using (ProjContext pjContext = CreateContext())
                {
                    using (var crs = CreateCrs(oCrs, pjContext))
                    {
                        if (crs != null)
                            north = crs.UsageArea.NorthLatitude;

                        return north;
                    }
                }        
            }
            catch (Exception ex)
            {
                return Lib.ExceptionHandler(ex);
            }

        } // UsageArea_NorthLatitude

    } // class

} // namespace

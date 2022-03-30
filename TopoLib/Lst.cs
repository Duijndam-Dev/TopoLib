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

namespace TopoLib
{
    public static class Lst
    {

        [ExcelFunctionDoc(
            Name = "TL.lst.AuthorityList",
            Category = "LST - List related",
            Description = "Provides a (vertical) spillover list with all authorities from the PROJ database",
            HelpTopic = "TopoLib-AddIn.chm!1900",

            Returns = "Authority-list",
            Summary = "Function that provides a (vertical) spillover list with all authorities from the PROJ database, spanning several rows",
            Example = "xxx"
         )]
        public static object AuthorityList(
            [ExcelArgument("Use header row (true)", Name = "Header")] object header
            )
        {
            string[] Authorities = new string[]
            {
                "EPSG",
                "ESRI",
                "IAU_2015",
                "IGNF",
                "NKG",
                "OGC",
                "PROJ"
            };

            bool bHeader = Optional.Check(header, true);
            int nOffset = bHeader ? 1 : 0;

            try
            {
                int nAuthorities = Authorities.Length;
                object[,] res = new object[nAuthorities + nOffset, 1];

                if (bHeader)
                    res[0, 0] = "Authority";

                for (int i = 0; i < nAuthorities; i++)
                {
                    res[i + nOffset, 0] = Authorities[i];
                }

                return res;
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
        } // AuthorityList

        [ExcelFunctionDoc(
            Name = "TL.lst.CoordinateReferenceSystemList",
            Category = "LST - List related",
            Description = "Provides a (vertical) spillover list with all coordinate reference systems from the PROJ database",
            HelpTopic = "TopoLib-AddIn.chm!1901",

            Returns = "Coordinate-Reference-System List",
            Summary = "Function that provides a (vertical) spillover list with all coordinate reference systems from the PROJ database, spanning several rows",
            Example = "xxx"
         )]
        public static object CoordinateReferenceSystemList(
            [ExcelArgument("Use header row (true)", Name = "Header")] object header
            )
        {
            bool bHeader = Optional.Check(header, true);
            int nOffset = bHeader ? 1 : 0;

            try
            {
                using (ProjContext pc = new ProjContext() { EnableNetworkConnections = false })
                {
                    int NrCrsAvailable = pc.GetCoordinateReferenceSystems().Count;
                    int HashCode = pc.GetHashCode();

                    object[,] res = new object[NrCrsAvailable + nOffset, 7];
                    int i = 0;

                    if (bHeader)
                    {
                        res[i, 0] = "Authority";
                        res[i, 1] = "Code";
                        res[i, 2] = "Type";
                        res[i, 3] = "ProjectionName";
                        res[i, 4] = "Name";
                        res[i, 5] = "AreaName";
                        res[i, 6] = "Nr GeoidModels";
                        i++;
                    }

                    foreach (var CrsInfo in pc.GetCoordinateReferenceSystems())
                    {
                        using (var Crs = CrsInfo.Create())
                        {
                            var expectedType = CrsInfo.Type;

                            if (expectedType == ProjType.CRS && CrsInfo.Authority == "IAU_2015" && ProjContext.ProjVersion == new Version(8, 2, 1))
                                expectedType = ProjType.GeodeticCrs;

                            res[i, 0] = CrsInfo.Authority; 
                            res[i, 1] = CrsInfo.Code;
                            res[i, 2] = CrsInfo.Type.ToString();
                            res[i, 3] = CrsInfo.ProjectionName ?? "<undefined>";
                            res[i, 4] = CrsInfo.Name;
                            res[i, 5] = CrsInfo.AreaName;
                            res[i, 6] = CrsInfo.GetGeoidModels().Count;
                            i++;
/*
                            Assert.AreEqual(expectedType, Crs.Type, $"Expected type mismatch on {p.Identifier}, celestial_body={p.CelestialBodyName}");
                            Assert.IsNotNull(CrsInfo.Authority);
                            Assert.IsNotNull(CrsInfo.Code);
                            //TestContext.WriteLine($"{p.Authority}:{p.Code} ({p.Type}) / {p.ProjectionName} {c.Name}");
*/
                        }
                    }
                    return res;
                }
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
        } // CoordinateReferenceSystemList

        [ExcelFunctionDoc(
            Name = "TL.lst.OperationList",
            Category = "LST - List related",
            Description = "Provides a (vertical) spillover list with all operations from the PROJ database",
            HelpTopic = "TopoLib-AddIn.chm!1903",

            Returns = "Operations-list",
            Summary = "Function that provides a (vertical) spillover list with all operations from the PROJ database, spanning several rows",
            Example = "xxx"
         )]
        public static object OperationList(
            [ExcelArgument("Use header row (true)", Name = "Header")] object header
            )
        {
            bool bHeader = Optional.Check(header, true);
            int nOffset = bHeader ? 1 : 0;

            try
            {
                ProjOperationDefinition.ProjOperationDefinitionCollection v = ProjOperationDefinition.All;
                int nCount = v.Count;

                object[,] res = new object[nCount + nOffset, 5];
                int i = 0;

                if (bHeader)
                {
                    res[i, 0] = "Name";
                    res[i, 1] = "Type";
                    res[i, 2] = "Title";
                    res[i, 3] = "Details";
                    res[i, 4] = "Nr. Arguments";
                    i++;
                }

                foreach (ProjOperationDefinition m in v)
                {
                    res[i, 0] = m.Name;
                    res[i, 1] = m.Type.ToString();
                    res[i, 2] = m.Title;
                    res[i, 3] = m.Details ?? "<not available>";
                    res[i, 4] = m.RequiredArguments.Count;
                    i++;
                }
                return res;
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
        } // OperationList

        [ExcelFunctionDoc(
            Name = "TL.lst.CelestialBodyList",
            Category = "LST - List related",
            Description = "Provides a (vertical) spillover list with all celestial bodies from the PROJ database",
            HelpTopic = "TopoLib-AddIn.chm!1902",

            Returns = "Operations-list",
            Summary = "Function that provides a (vertical) spillover list with all celestial bodies from the PROJ database, spanning several rows",
            Example = "xxx"
         )]
        public static object CelestialBodyList(
            [ExcelArgument("Use header row (true)", Name = "Header")] object header
            )
        {
            bool bHeader = Optional.Check(header, true);
            int nOffset = bHeader ? 1 : 0;

            try
            {

                using (ProjContext pc = new ProjContext() { EnableNetworkConnections = false })
                {
                    var bodies = pc.GetCelestialBodies();
                    int nCount = bodies.Count;

                    object[,] res = new object[nCount + nOffset, 3];
                    int i = 0;

                    if (bHeader)
                    {
                        res[i, 0] = "Name";
                        res[i, 1] = "Authority";
                        res[i, 2] = "IsEarth";
                        i++;
                    }

                    foreach (var p in bodies)
                    {
                        res[i, 0] = p.Name;
                        res[i, 1] = p.Authority;
                        res[i, 2] = p.IsEarth;
                        i++;
                    }
                    return res;
                }
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }
        } // CelestialBodyList


    }
}

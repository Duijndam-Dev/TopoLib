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
            Name = "TL.lst.Authorities",
            Category = "LST - List related",
            Description = "Provides a (vertical) spillover list with all authorities from the PROJ database",
            HelpTopic = "TopoLib-AddIn.chm!1800",

            Returns = "Authority-list",
            Summary = "Function that provides a (vertical) spillover list with all authorities from the PROJ database, spanning serveral rows",
            Example = "xxx"
         )]
        public static object Authorities(
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

            try
            {
                int nAuthorities = Authorities.Length;
                object[,] res = new object[nAuthorities, 1];

                for (int i = 0; i < nAuthorities; i++)
                {
                    res[i, 0] = Authorities[i];
                }

                return res;
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // Authorities

    }
}

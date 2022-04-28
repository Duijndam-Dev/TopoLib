using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using ExcelDna.Integration;
using ExcelDna.Documentation;

// This class is intended to deal with angles in various forms and shapes (DD, DMM, DMS, etc)
// Routines are loosely based on the GitHub DotSpatial project. See :
// https://github.com/DotSpatial/DotSpatial/blob/master/Source/DotSpatial.Positioning/Angle.cs
// https://stackoverflow.com/questions/38151856/how-to-convert-a-location-in-degrees-minutes-seconds-represented-as-a-string-to
// http://spiff.rit.edu/tass/bait.old/convert.c
// https://docs.microsoft.com/en-us/office/troubleshoot/excel/convert-degrees-minutes-seconds-angles
// https://stackoverflow.com/questions/5786025/decimal-degrees-to-degrees-minutes-and-seconds-in-javascript
// https://flyandwire.com/2020/08/10/back-to-basics-latitude-and-longitude-dms-dd-ddm/
// https://www.maptools.com/tutorials/lat_lon/formats
// https://en.wikipedia.org/wiki/Decimal_degrees
// https://gsp.humboldt.edu/olm/Lessons/GIS/01%20SphericalCoordinates/Reporting_Geographic_Coordinates.html


namespace TopoLib
{
    public static class Deg
    {
        internal const int    MaximumPrecisionDigits = 12;
        internal const double PI = 3.1415926535897932384626433832795;
        internal const string doubleDot = "Only the right most part can be a fractional number";

        private static int Hours(double degrees)
        {
            int i = (int)Math.Truncate(degrees);
            return i;
        }

        private static int Minutes(double degrees)
        {
            // return Convert.ToInt32(Math.Abs(Math.Truncate(Math.Round((degrees - Hours(degrees)) * 60.0, MaximumPrecisionDigits - 1))));
            int i = (int)Math.Abs(Math.Truncate((degrees - Hours(degrees)) * 60.0));
            return i;
        }

        private static double DecimalMinutes(double degrees)
        {
            // return Math.Round((Math.Abs(degrees - Math.Truncate(degrees)) * 60.0), MaximumPrecisionDigits - 2);
            double d = Math.Abs(degrees - Math.Truncate(degrees)) * 60.0;
            return d;
        }

        private static double Seconds(double degrees)
        {
            // return Math.Round((Math.Abs(degrees - Hours(degrees)) * 60.0 - Minutes(degrees)) * 60.0, MaximumPrecisionDigits - 4);
            double d = (Math.Abs(degrees - Hours(degrees)) * 60.0 - Minutes(degrees)) * 60.0;
            return d;
        }

        private static double ToDecimalDegrees(int hours, int minutes, double seconds)
        {
            // return hours < 0
            //    ? -Math.Round(-hours + minutes / 60.0 + seconds / 3600.0, MaximumPrecisionDigits)
            //    : Math.Round(hours + minutes / 60.0 + seconds / 3600.0, MaximumPrecisionDigits);
            return hours < 0
                ? -(-hours + minutes / 60.0 + seconds / 3600.0)
                : (hours + minutes / 60.0 + seconds / 3600.0);
        }

        private static double ToDecimalDegrees(int hours, double decimalMinutes)
        {
            // return hours < 0
            //    ? -Math.Round(-hours + decimalMinutes / 60.0, MaximumPrecisionDigits)
            //    : Math.Round(hours + decimalMinutes / 60.0, MaximumPrecisionDigits);
            return hours < 0
                ? -(-hours + decimalMinutes / 60.0)
                : (hours + decimalMinutes / 60.0);
        }

        [ExcelFunctionDoc(
            Name = "TL.deg.AsString",
            Category = "DEG - Angle related",
            Description = "Writes an angle (defined in decimal degrees) as a Degree-Minute-Seconds or Decimal-Degrees string",
            HelpTopic = "TopoLib-AddIn.chm!1700",

            Returns = "DMS-string",
            Summary = "Function that writes an angle (defined in degrees) as a Degree-Minute-Seconds or Decimal-Degrees string",
            Example = "=TL.deg.AsString(-19.627501, 1, \"dd.dddddd°\") returns 19.627501°W" +
                   "<p>=TL.deg.AsString(-19.662282, 1, \"dd°mm'ss.sss\"\"\") returns W19°39'44.217\"</p>",
            Remarks ="<p>This method returns the angle in a specific string format. If no value for the format is specified, a default format of {<b>h&deg;mm'ss.ss\"</b>} is used. " +
            "Any string output by this method can be converted back into an decimal-degrees angle using the <a href = \"TL.deg.FromDmsString.htm\"> <b>TL.deg.FromDmsString()</b> </a>method. </p>" +
            "<p>The {<b>h&deg;</b>} code represents hours along with a degree symbol (Alt+0176 on the keypad), {<b>mm'</b>} represents minutes and {<b>ss.ss\"</b>} represents seconds using two decimals.</p>" +
            "<p>For a string in decimal degrees use {<b>d.ddddd&deg;</b>}. This will return a string value with 5 decimals (~ 1m accuracy).</p>" +
            "<p>Optionally the sign of the angle can be replace by an appended/prepended direction (NS, EW).</p>"
         )]
        public static object AsString(
            [ExcelArgument("Angle (in decimal degrees)", Name = "angle")] double angle,
            [ExcelArgument("Optional direction indicator (0). 0 = sign, 1 = E/W appended, 2 = N/S appended, 3 = E/W prepended, 4 = N/S prepended", Name = "direction")] object oDirection,
            [ExcelArgument("Optional format string (dd°mm'ss.ss\"), where ° is Alt+0176 on the keypad", Name = "format")] object oFormat
            )
        {
            int direction = (int)Optional.Check(oDirection, 0);
            direction = Math.Min(Math.Max(direction, 0), 4);

            string format = Optional.Check(oFormat, "dd°mm'ss.ss\"");
            bool negative = false;

            // Use the absolute angle value, when decorating instead of using a negative sign
            if (angle < 0 && direction > 0)
            {
                angle = Math.Abs(angle);
                negative = true;
            }

            CultureInfo culture = CultureInfo.CurrentCulture;

            string subFormat;
            string newFormat;
            bool isDecimalHandled = false;

            try
            {
                // Replace the "h" with "d" since degrees is the same as hours
                format = format.ToUpper(CultureInfo.InvariantCulture).Replace("H", "D");

                // Only one decimal is allowed
                if (format.IndexOf(culture.NumberFormat.NumberDecimalSeparator, StringComparison.Ordinal) !=
                    format.LastIndexOf(culture.NumberFormat.NumberDecimalSeparator, StringComparison.Ordinal))
                    throw new ArgumentException("Only one fractional number is allowed in the string representation");

                // Is there a degrees/hours specifier ° ?
                int startChar = format.IndexOf("D");
                int endChar;
                if (startChar > -1)
                {
                    // Yes. Look for subsequent D characters or a period
                    endChar = format.LastIndexOf("D");
                    // Extract the sub-string
                    subFormat = format.Substring(startChar, endChar - startChar + 1);
                    // Convert to a numberic-formattable string
                    newFormat = subFormat.Replace("D", "0");
                    // Replace the hours
                    if (newFormat.IndexOf(culture.NumberFormat.NumberDecimalSeparator) > -1)
                    {
                        isDecimalHandled = true;
                        format = format.Replace(subFormat, angle.ToString(newFormat, culture));
                    }
                    else
                    {
                        format = format.Replace(subFormat, Hours(angle).ToString(newFormat, culture));
                    }
                }
                
                // Is there a minutes specifier ' ?
                startChar = format.IndexOf("M");
                if (startChar > -1)
                {
                    // Yes. Look for subsequent M characters or a period
                    endChar = format.LastIndexOf("M");
                    // Extract the sub-string
                    subFormat = format.Substring(startChar, endChar - startChar + 1);
                    // Convert to a numberic-formattable string
                    newFormat = subFormat.Replace("M", "0");
                    // Replace the minutes
                    if (newFormat.IndexOf(culture.NumberFormat.NumberDecimalSeparator) > -1)
                    {
                        if (isDecimalHandled)
                        {
                            throw new ArgumentException(doubleDot, "format");
                        }
                        isDecimalHandled = true;
                        format = format.Replace(subFormat, DecimalMinutes(angle).ToString(newFormat, culture));
                    }
                    else
                    {
                        format = format.Replace(subFormat, Minutes(angle).ToString(newFormat, culture));
                    }
                }

                // Is there a seconds specifier " ?
                startChar = format.IndexOf("S");
                if (startChar > -1)
                {
                    // Yes. Look for subsequent S characters or a period
                    endChar = format.LastIndexOf("S");
                    // Extract the sub-string
                    subFormat = format.Substring(startChar, endChar - startChar + 1);
                    // Convert to a numberic-formattable string
                    newFormat = subFormat.Replace("S", "0");
                    // Replace the seconds
                    if (newFormat.IndexOf(culture.NumberFormat.NumberDecimalSeparator) > -1)
                    {
                        if (isDecimalHandled)
                        {
                            throw new ArgumentException(doubleDot, "format");
                        }
                        format = format.Replace(subFormat, Seconds(angle).ToString(newFormat, culture));
                    }
                    else
                    {
                        format = format.Replace(subFormat, Seconds(angle).ToString(newFormat, culture));
                    }
                }

                // If nothing in string, then return zero
                if (String.Compare(format, "°", true, culture) == 0)
                {
                    format = "0°";
                }

                switch (direction)
                {
                    default:
                        throw new ArgumentException("Wrong direction value");
                    case 0:
                        return format;
                    case 1:
                        return negative ? format + "W" : format + "E";
                    case 2:
                        return negative ? format + "S" : format + "N";
                    case 3:
                        return negative ? "W" + format : "E" + format;
                    case 4:
                        return negative ? "S" + format : "N" + format;
                }
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // AsString

        [ExcelFunctionDoc(
            Name = "TL.deg.FromString",
            Category = "DEG - Angle related",
            Description = "Reads an angle in DegreesMinutesSeconds, DegreesDecimalMinutes or DecimalDegrees from a string",
            HelpTopic = "TopoLib-AddIn.chm!1701",

            Returns = "Angle (in decimal degrees)",
            Summary = "Function that returns an angle (in decimal degrees) from a string description ",
            Example = "=TL.deg.FromString(19.627501°W) returns -19.627501"
         )]
        public static object FromString(
            [ExcelArgument("string using DegreesMinutesSeconds, DegreesDecimalMinutes or DecimalDegrees format", Name = "angleString")] string dmsAngle
            )
        {
            // Is the value null or empty?
            if (string.IsNullOrEmpty(dmsAngle))
            {
                return 0.0;
            }

            // Default to the current culture
            CultureInfo culture = CultureInfo.CurrentCulture;

            double angle = 0;
            double polarity = 1;;

            // First, clean up the strings
            try
            {
                // check if we are dealing with negative numbers, expressed by a S/W symbol
                if (dmsAngle.IndexOf("W") > -1 || dmsAngle.IndexOf("S") > -1)
                    polarity = -1;

                // Clean up the string
                StringBuilder newValue = new StringBuilder(dmsAngle);
                newValue.Replace("E", " ").Replace("W", " ").Replace("S",  " ").Replace("N",  " ");
                newValue.Replace("°", " ").Replace("'", " ").Replace("\"", " ").Replace("  ", " ");

                // Now split the values into an array
                string[] values = newValue.ToString().Trim().Split(' ');

                // How many elements are in the array?
                switch (values.Length)
                {
                    case 0:
                        // Return a blank Angle
                        return angle;

                    case 1: // Decimal degrees
                        // Is it empty?
                        if (String.IsNullOrWhiteSpace(values[0]))
                        {
                            return angle;
                        }

                        // Look at the number of digits, this might be HHHMMSS format.
                        if (values[0].Length == 7 && values[0].IndexOf(culture.NumberFormat.NumberDecimalSeparator, StringComparison.CurrentCulture) == -1)
                        {
                            angle = ToDecimalDegrees(
                                int.Parse(values[0].Substring(0, 3), culture),
                                int.Parse(values[0].Substring(3, 2), culture),
                                double.Parse(values[0].Substring(5, 2), culture));
                            return angle * polarity;
                        }

                        if (values[0].Length == 8 && values[0][0] == '-' && values[0].IndexOf(culture.NumberFormat.NumberDecimalSeparator, StringComparison.CurrentCulture) == -1)
                        {
                            angle = ToDecimalDegrees(
                                int.Parse(values[0].Substring(0, 4), culture),
                                int.Parse(values[0].Substring(4, 2), culture),
                                double.Parse(values[0].Substring(6, 2), culture));
                            return angle * polarity ;
                        }

                        angle = double.Parse(values[0], culture);
                        return angle * polarity;

                    case 2: // Hours and decimal minutes
                        // If this is a fractional value, remember that it is
                        if (values[0].IndexOf(culture.NumberFormat.NumberDecimalSeparator, StringComparison.Ordinal) != -1)
                        {
                            throw new ArgumentException(doubleDot, "angleString");
                        }

                        // Set decimal degrees
                        angle = ToDecimalDegrees(
                            int.Parse(values[0], culture),
                            float.Parse(values[1], culture));
                        return angle * polarity;

                    default: // Hours, minutes and seconds  (most likely)
                        // If this is a fractional value, remember that it is
                        if (values[0].IndexOf(culture.NumberFormat.NumberDecimalSeparator) != -1 || values[0].IndexOf(culture.NumberFormat.NumberDecimalSeparator) != -1)
                        {
                            throw new ArgumentException(doubleDot, "angleString");
                        }

                        // Set decimal degrees
                        angle = ToDecimalDegrees(
                            int.Parse(values[0], culture),
                            int.Parse(values[1], culture),
                            double.Parse(values[2], culture));
                        return angle * polarity;
                }
            }
            catch (Exception ex)
            {
                return AddIn.ProcessException(ex);
            }

        } // FromDmsString

        [ExcelFunctionDoc(
            Name = "TL.deg.FromDmsValues",
            Category = "DEG - Angle related",
            Description = "Combines seperate Hours (0), Minutes (0), and Seconds (0.00) values into an angle (in decimal degrees)",
            HelpTopic = "TopoLib-AddIn.chm!1702",

            Returns = "Angle (in decimal degrees)",
            Summary = "Function that combines seperate Hours (0), Minutes (0), and Seconds (0.00) values into an angle (in decimal degrees)",
            Example = "=TL.deg.FromDmsValues(-19, 37, 39.005) returns -19.627501"
         )]
        public static double FromDmsValues(
            [ExcelArgument("Hours (0) an integer value", Name = "Hours")] object oHours,
            [ExcelArgument("Minutes (0) an integer value", Name = "Minutes")] object oMinutes,
            [ExcelArgument("Seconds (0.00) a floating point value", Name = "Seconds")] object oSeconds
            )
        {
            int hours = (int)Optional.Check(oHours, 0.0);
            int minutes = (int)Optional.Check(oMinutes, 0.0);
            double seconds = Optional.Check(oSeconds, 0.0);

            return ToDecimalDegrees(hours, minutes, seconds);
        } // FromDmsValues

        [ExcelFunctionDoc(
            Name = "TL.deg.FromDmValues",
            Category = "DEG - Angle related",
            Description = "Combines seperate Hours (0) and decimal minutes (0.00) values into an angle (in decimal degrees)",
            HelpTopic = "TopoLib-AddIn.chm!1703",

            Returns = "Angle (in decimal degrees)",
            Summary = "Function that combines seperate Hours (0) and decimal minutes (0.00) values into an angle (in decimal degrees)",
            Example = "=TL.deg.FromDmValues(-19, 37.6501) returns -19.62750129"
         )]
        public static double FromDmValues(
            [ExcelArgument("Hours (0) an integer value", Name = "Hours")] object oHours,
            [ExcelArgument("Minutes (0.00) a floating point value", Name = "Minutes")] object oMinutes
            )
        {
            int hours = (int)Optional.Check(oHours, 0.0);
            double minutes = Optional.Check(oMinutes, 0.0);

            return ToDecimalDegrees(hours, minutes);
        } // FromDmValues

        [ExcelFunctionDoc(
            Name = "TL.deg.GetHours",
            Category = "DEG - Angle related",
            Description = "Truncates an angle (in decimaldegrees) to hours, omitting the fractional value",
            HelpTopic = "TopoLib-AddIn.chm!1704",

            Returns = "Truncated angle (in decimal degrees)",
            Summary = "Function that truncates an angle (in decimaldegrees) to hours, omitting the fractional value",
            Example = "=TL.deg.GetHours(-19.627501) returns -19"
         )]
        public static double GetHours(
            [ExcelArgument("Angle (in decimal degrees)", Name = "angle")] double angle
            )
        {
            return Hours(angle);

        } // GetHours

        [ExcelFunctionDoc(
            Name = "TL.deg.GetDecimalMinutes",
            Category = "DEG - Angle related",
            Description = "Returns the decimal minutes fraction of an angle (in decimal degrees)",
            HelpTopic = "TopoLib-AddIn.chm!1705",

            Returns = "floating point part of minutes of an angle (in decimal degrees)",
            Summary = "Function that returns the floating point minutes part of an angle (in decimal degrees)",
            Example = "=TL.deg.GetDecimalMinutes(-19.627501) returns 37.6501"
         )]
        public static double GetDecimalMinutes(
            [ExcelArgument("Angle (in decimal degrees)", Name = "angle")] double angle
            )
        {
            return DecimalMinutes(angle);
        } // GetMinutes

        [ExcelFunctionDoc(
            Name = "TL.deg.GetMinutes",
            Category = "DEG - Angle related",
            Description = "Returns the minutes fraction of an angle (in decimal degrees)",
            HelpTopic = "TopoLib-AddIn.chm!1706",

            Returns = "Integer part of minutes of an angle (in decimal degrees)",
            Summary = "Function that returns the integer minutes part of an angle (in decimal degrees)",
            Example = "=TL.deg.GetMinutes(-19.627501) returns 37"
         )]
        public static double GetMinutes(
            [ExcelArgument("Angle (in decimal degrees)", Name = "angle")] double angle
            )
        {
            return Minutes(angle);
        } // GetMinutes



        [ExcelFunctionDoc(
            Name = "TL.deg.GetSeconds",
            Category = "DEG - Angle related",
            Description = "Returns the seconds fraction of an angle (in decimal degrees)",
            HelpTopic = "TopoLib-AddIn.chm!1707",

            Returns = "Seconds part of angle (in decimal degrees)",
            Summary = "Function that returns the seconds fraction of an angle (in decimal degrees)",
            Example = "=TL.deg.GetSeconds(-19.627501) returns 39.005"
         )]
        public static double GetSeconds(
            [ExcelArgument("Angle (in decimal degrees)", Name = "angle")] double angle
            )
        {
            return Seconds(angle);

        } // GetSeconds

        [ExcelFunctionDoc(
            Name = "TL.deg.Normalize0to360",
            Category = "DEG - Angle related",
            Description = "Constrains an angle to the 0 to 360 degree range",
            HelpTopic = "TopoLib-AddIn.chm!1708",

            Returns = "Normalized angle (in decimal degrees)",
            Summary = "Function that constrains an angle to the 0 to 360 degree range",
            Example = "=TL.deg.Normalize0to360(701) returns 341"
         )]
        public static double Normalize0to360(
            [ExcelArgument("Angle (in decimal degrees)", Name = "angle")] double angle
            )
        {
            double value = angle;
            while (value < 0)
            {
                value += 360.0;
            }
            return (value % 360);

        } // Normalize0to360

        [ExcelFunctionDoc(
            Name = "TL.deg.Normalize180to180",
            Category = "DEG - Angle related",
            Description = "Constrains an angle to the -180 to +180 degree range",
            HelpTopic = "TopoLib-AddIn.chm!1709",

            Returns = "Normalized angle (in decimal degrees)",
            Summary = "Function that constrains an angle to the -180 to +180 degree range",
            Example = "=TL.deg.Normalize180to180(701) returns -19"
         )]
        public static double Normalize180to180(
            [ExcelArgument("Angle (in decimal degrees)", Name = "angle")] double angle
            )
        {
            double value = angle + 180;
            while (value < 0)
            {
                value += 360.0;
            }
            return (value % 360) - 180;;
        } // Normalize180to180

        [ExcelFunctionDoc(
            Name = "TL.deg.IsWithin0to360",
            Category = "DEG - Angle related",
            Description = "Checks if an angle is >= 0 and < 360 degrees",
            HelpTopic = "TopoLib-AddIn.chm!1710",

            Returns = "TRUE if angle is within this range; FALSE otherwise",
            Summary = "Function that checks if an angle is >= 0 and < 360 degrees",
            Example = "=TL.deg.IsWithin0to360(701) returns FALSE"
         )]
        public static bool IsWithin0to360(
            [ExcelArgument("Angle (in decimal degrees)", Name = "angle")] double angle
            )
        {
            return angle >= 0 && angle < 360;
        } // IsWithin0to360

        [ExcelFunctionDoc(
            Name = "TL.deg.IsWithin180to180",
            Category = "DEG - Angle related",
            Description = "Checks if an angle is >= -180 and < 180 degrees",
            HelpTopic = "TopoLib-AddIn.chm!1711",

            Returns = "TRUE if angle is within this range; FALSE otherwise",
            Summary = "Function that checks if an angle is >= -180 and < 180 degrees, as is required for normalized longitude values",
            Example = "=TL.deg.IsWithin180to180(-19) returns TRUE"
         )]
        public static bool IsWithin180to180(
            [ExcelArgument("Angle (in decimal degrees)", Name = "angle")] double angle
            )
        {
            return angle >= -180 && angle < 180;
        } // IsWithin180to180

        [ExcelFunctionDoc(
            Name = "TL.deg.IsWithin90to90",
            Category = "DEG - Angle related",
            Description = "Checks if an angle is >= -90 and < 90 degrees",
            HelpTopic = "TopoLib-AddIn.chm!1712",

            Returns = "TRUE if angle is within this range; FALSE otherwise",
            Summary = "Function that checks if an angle is >= -90 and < 90 degrees, as is required for normalized latitude values",
            Example = "=TL.deg.IsWithin90to90(-19) returns TRUE"
         )]
        public static bool IsWithin90to90(
            [ExcelArgument("Angle (in decimal degrees)", Name = "angle")] double angle
            )
        {
            return angle >= -90 && angle < 90;
        } // IsWithin90to90

        [ExcelFunctionDoc(
            Name = "TL.deg.ToRadians",
            Category = "DEG - Angle related",
            Description = "Converts an angle in decimal degrees to radians",
            HelpTopic = "TopoLib-AddIn.chm!1713",

            Returns = "Angle in radians",
            Summary = "Function that converts an angle in decimal degrees to radians",
            Example = "=TL.deg.ToRadians(-19) returns -0.331612558"
         )]
        public static double ToRadians(
            [ExcelArgument("Angle (in decimal degrees)", Name = "angle")] double angle
            )
        {
            return PI * angle / 180;
        } // ToRadians

        [ExcelFunctionDoc(
            Name = "TL.deg.FromRadians",
            Category = "DEG - Angle related",
            Description = "Converts an angle from radians to decimal degrees",
            HelpTopic = "TopoLib-AddIn.chm!1714",

            Returns = "Angle in decimal degrees",
            Summary = "Function that converts an angle from radians to decimal degrees",
            Example = "=TL.deg.ToRadians(-0.331612558) returns -19"
         )]
        public static double FromRadians(
            [ExcelArgument("Angle (in radians)", Name = "angle")] double angle
            )
        {
            return 180 * angle / PI;
        } // FromRadians
    }
}

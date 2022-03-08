using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using ExcelDna.Integration;
using ExcelDna.Documentation;


// The purpose of this class is to provide GPX and KML string representations of point sets within Excel
// Work very much in progress. Projects of interest might be :
// https://github.com/macias/Gpx
// https://github.com/DustyRoller/GPXParser/blob/master/GPXParser/Parser.cs
// http://dlg.krakow.pl/gpx/?i=1
// https://github.com/KennethEvans/VS-GpxViewer
//
// https://github.com/samcragg/sharpkml
// https://github.com/podulator/libKml
// https://github.com/akichko/libKml
// https://github.com/deRightDirection/GeoKml

// Maybe split Gps class into Kml and Gpx !

namespace TopoLib
{
    public static class Gps
    {
        internal const string gpxHeader = 
            "<?xml version=\"1.0\" ?>" +
            "<gpx xmlns=\"http://www.topografix.com/GPX/1/1\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" " +
            "xsi:schemaLocation=\"http://www.topografix.com/GPX/1/1 http://www.topografix.com/GPX/1/1/gpx.xsd\" " +
            "version=\"1.1\" creator=\"GPS Data Team ( http://www.gps-data-team.com )\">\n";

        internal const string gpxFooter = "</gpx>\n";

        internal const string kmlHeader =
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<kml xmlns=\"http://www.opengis.net/kml/2.2\" xmlns:gx=\"http://www.google.com/kml/ext/2.2\" xmlns:atom=\"http://www.w3.org/2005/Atom\">\n" +
            "<Document><atom:author><atom:name>Bart Duijndam</atom:name></atom:author>" +
            "<Style id=\"style001\"><LineStyle><color>96ff00d5</color><width>3.0</width></LineStyle></Style>\n";

        internal const string kmlFooter = "</Document></kml>\n";

/*
        internal const string gpxHeader2 = 
            "<?xml version=\"1.0\" ?>" +
            "<gpx xmlns=\"http://www.topografix.com/GPX/1/1\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" " +
            "xsi:schemaLocation=\"http://www.topografix.com/GPX/1/1 http://www.topografix.com/GPX/1/1/gpx.xsd\" " +
            "version=\"1.1\" creator=\"GPS Data Team ( http://www.gps-data-team.com )\">";

        internal const string gpxFooter2 = "</gpx>";
*/

        internal static string GetGpsTracks(object [,] oPoints, bool ExcelInterop = false)
        {
            // determine size of point array
            // Check specific input parameters
            int nInputRows = oPoints.GetLength(0);
            int nInputCols = oPoints.GetLength(1);
            int nOffset = ExcelInterop ? 1 : 0;

            // define the row types based on row content
            int[] rowType = new int[nInputRows];
            int lastRow = 0;
            for (int i = 0; i < nInputRows; i++)
            {
                if (Optional.IsEmpty(oPoints[i + nOffset, 0 + nOffset]) == true)
                {
                    rowType[i] = 1; // empty
                    lastRow = 1;
                }
                else if (oPoints[i + nOffset, 0 + nOffset] is double && oPoints[i + nOffset, 1 + nOffset] is double)
                {
                    rowType[i] = 3; // data
                    lastRow = 3;
                }
                else
                {
                    if (lastRow == 2)
                    {
                        rowType[i] = 1; // empty (we already have a header)
                    }
                    else
                    {
                        rowType[i] = 2; // header
                        lastRow = 2;
                    }
                }
            }

            // a stringbuilder is more efficient; but it gave rise to empty line issues on output.
            string gpxOut = gpxHeader; 

            // parse the input rows
            int nTrackSegment = 1;
            bool haveHeader = false;
            string lon = "";
            string lat = "";
            string ele = "";
            string nam = "";

            for (int i = 0; i < nInputRows; i++)
            {
                if (rowType[i] == 1) // empty row
                {
                    continue;
                }
                else if (rowType[i] == 2) // header row
                {
                    if (i > 0 && rowType[i - 1] == 2)
                    {
                        // only use first header row; ignore the rest
                        continue;
                    }
                    else
                    {
                        // we have to insert a header row for a new track.
                        haveHeader = true;

                        // do we have a name in the 1st column ?
                        string name = Optional.GetString(oPoints[i + nOffset, 0 + nOffset]);
                        if (!string.IsNullOrEmpty(name))
                            name = $"<name>{name}</name>";

                        // do we have a comment in the 2nd column ?
                        string comment = Optional.GetString(oPoints[i + nOffset, 1 + nOffset]);
                        if (!string.IsNullOrEmpty(comment))
                            comment = $"<cmt>{comment}</cmt>";

                        // do we have a description in the 3rd column ?
                        string description = "";

                        if (nInputCols > 2)
                        {
                            description = Optional.GetString(oPoints[i + nOffset, 2 + nOffset]);
                            if (!string.IsNullOrEmpty(description))
                                description = $"<desc>{description }</desc>";
                        }

                        // do we have a type in the 4th column ?
                        string type = "";

                        if (nInputCols > 3)
                        {
                            type = Optional.GetString(oPoints[i + nOffset, 3 + nOffset]);
                            if (!string.IsNullOrEmpty(type))
                                type = $"<type >{type }</type >";
                        }

                        // Add header row and start of track segment
                        gpxOut += ($"<trk>{name}{comment}{description}{type}<trkseg>\n");
                    }

                }
                else if (rowType[i] == 3) // data row
                {
                    // do we have a leading header ?

                    if (!haveHeader)
                    {
                        // we have to insert a header row
                        haveHeader = true;

                        string name = string.Format("segment {0}", nTrackSegment.ToString());
                        name = $"<name>{name}</name>";
                        gpxOut += ($"<trk>{name}<trkseg>\n");
                    }

                    // do we have a longitude in the 1st column ?
                    double longitude = Optional.GetDouble(oPoints[i + nOffset, 0 + nOffset]);
                    if (Double.IsNaN(longitude)) throw new ArgumentException($"Longitude input error on row: {i}");

                    lon = string.Format("{0:0.00000000}", longitude);

                    // do we have a latitude in the 2nd column ?
                    double latitude = Optional.GetDouble(oPoints[i + nOffset, 1 + nOffset]);
                    if (Double.IsNaN(latitude)) throw new ArgumentException($"latitude input error on row: {i}");

                    lat = string.Format("{0:0.00000000}", latitude);

                    // do we have a elevation in the 3rd column ?

                    if (nInputCols > 2)
                    {
                        double elevation = Optional.GetDouble(oPoints[i + nOffset, 2 + nOffset]);
                        if (!Double.IsNaN(elevation))
                        {
                            // it's NOT blank; so must be a number
                            if (Double.IsInfinity(elevation)) throw new ArgumentException($"elevation input error on row: {i}");

                            ele = string.Format("<ele>{0:0.0}</ele>", elevation);
                        }
                    }

                    // do we have a point name or time in the 4th column ?
                    double time;
                    if (nInputCols > 3)
                    {
                        time = Optional.GetDouble(oPoints[i + nOffset, 3 + nOffset]);
                        if(double.IsNaN(time))
                        {
                            // it must be a string...
                            nam = Optional.GetString(oPoints[i + nOffset, 3 + nOffset]);
                            if (!string.IsNullOrEmpty(nam))
                                nam = $"<name>{nam}</name>";
                        }
                        else
                        {
                            // it should be time; check it's ok.
                            if (!double.IsInfinity(time))
                            {
                                // now get time from a double value in Excel
                                // FromOADate(time).ToString("s") is simplified to FromOADate(time):s
                                nam = $"<time>{DateTime.FromOADate(time):s}</time>";
                            }
                        }
                    }

                    // Add lat/long values, as well as elevation & point name if available
                    gpxOut += ($"<trkpt lat=\"{lat}\" lon=\"{lon}\">{ele}{nam}</trkpt>\n");

                    // if the next line does not contain numbers; we need to close the track
                    if (i == nInputRows - 1 || (i < nInputRows - 1 && rowType[i + 1] != 3))
                    {
                        // We are at the last row of a track;
                        gpxOut += ($"</trkseg></trk>\n");
                        nTrackSegment++;
                        haveHeader = false;
                    }
                }
            }

            // final line to add
            gpxOut += gpxFooter;

            return gpxOut;
        }

        internal static string GetKmlTracks(object [,] oPoints, bool ExcelInterop = false)
        {
            // determine size of point array
            // Check specific input parameters
            int nInputRows = oPoints.GetLength(0);
            int nInputCols = oPoints.GetLength(1);
            int nOffset = ExcelInterop ? 1 : 0;

            // define the row types based on row content
            int[] rowType = new int[nInputRows];
            int lastRow = 0;
            for (int i = 0; i < nInputRows; i++)
            {
                if (Optional.IsEmpty(oPoints[i + nOffset, 0 + nOffset]) == true)
                {
                    rowType[i] = 1; // empty
                    lastRow = 1;
                }
                else if (oPoints[i + nOffset, 0 + nOffset] is double && oPoints[i + nOffset, 1 + nOffset] is double)
                {
                    rowType[i] = 3; // data
                    lastRow = 3;
                }
                else
                {
                    if (lastRow == 2)
                    {
                        rowType[i] = 1; // empty (we already have a header)
                    }
                    else
                    {
                        rowType[i] = 2; // header
                        lastRow = 2;
                    }
                }
            }

            // a stringbuilder is more efficient; but it gave rise to empty line issues on output.
            string kmlOut = kmlHeader; 

            // parse the input rows
            int nTrackSegment = 1;
            bool haveHeader = false;
            
            string lon, lat, ele;

            for (int i = 0; i < nInputRows; i++)
            {
                if (rowType[i] == 1) // empty row
                {
                    continue;
                }
                else if (rowType[i] == 2) // header row
                {
                    if (i > 0 && rowType[i - 1] == 2)
                    {
                        // only use first header row; ignore the rest
                        continue;
                    }
                    else
                    {
                        // we have to insert a header row for a new track.
                        haveHeader = true;

                        // do we have a name in the 1st column ?
                        string name = Optional.GetString(oPoints[i + nOffset, 0 + nOffset]);
                        if (!string.IsNullOrEmpty(name))
                            name = $"<name>{name}</name>";

                        // Add header row and start of track segment
                        kmlOut += ($"<Placemark><name>{name}</name><styleUrl>#style001</styleUrl><MultiGeometry><LineString><coordinates>");
                    }

                }
                else if (rowType[i] == 3) // data row
                {
                    if (!haveHeader)
                    {
                        // we have to insert a header row
                        haveHeader = true;

                        string name = string.Format("segment {0}", nTrackSegment.ToString());
                        name = $"<name>{name}</name>";
                        kmlOut += ($"<Placemark><name>{name}</name><styleUrl>#style001</styleUrl><MultiGeometry><LineString><coordinates>");
                    }

                    // do we have a longitude in the 1st column ?
                    double longitude = Optional.GetDouble(oPoints[i + nOffset, 0 + nOffset]);
                    if (Double.IsNaN(longitude)) throw new ArgumentException($"Longitude input error on row: {i}");

                    lon = string.Format("{0:0.00000000}", longitude);

                    // do we have a latitude in the 2nd column ?
                    double latitude = Optional.GetDouble(oPoints[i + nOffset, 1 + nOffset]);
                    if (Double.IsNaN(latitude)) throw new ArgumentException($"latitude input error on row: {i}");

                    lat = string.Format("{0:0.00000000}", latitude);

                    ele = "";
                    if (nInputCols > 2)
                    {
                        double elevation = Optional.GetDouble(oPoints[i + nOffset, 2 + nOffset]);
                        if (!Double.IsNaN(elevation))
                        {
                            // it's NOT blank; so must be a number
                            if (Double.IsInfinity(elevation)) throw new ArgumentException($"elevation input error on row: {i}");

                            ele = string.Format(",{0:0.0}", elevation);
                        }
                    }

                    // Add lat/long/ele values 
                    kmlOut += ($" {lon},{lat}{ele}");

                    // if the next line does not contain numbers; we need to close the track
                    if (i == nInputRows - 1 || (i < nInputRows - 1 && rowType[i + 1] != 3))
                    {
                        // We are at the last row of a track;
                        kmlOut += ($" </coordinates></LineString></MultiGeometry></Placemark>\n");
                        nTrackSegment++;
                        haveHeader = false;
                    }
                }
            }

            // final line to add
            kmlOut += kmlFooter;

            return kmlOut;
        }

        [ExcelFunctionDoc(
            Name = "TL.gps.AsGpxTracks",
            Category = "GPS - GPS String Export",
            Description = "Exports a (vertical) list of EPSG:4326 points (Long, Lat [,Ele [,Tim/PtName]]) as a GPX string",
            HelpTopic = "TopoLib-AddIn.chm!1800",

            Returns = "One or more GPS track(s) as an XML-string in GPX-format. See remarks in help-file for more information",
            
            Summary = "Function that creates one or more GPS tracks(s) in a GPX-string.",
            Example = "xxx",
            Remarks ="<p>This method uses a (vertical) list of (Long, Lat [,Ele [,Tim/PtName]]) points, and transforms this into an XML-string in GPX-format</p>" +
            "<p>When the list is interrupted by one or more blank lines, or by a line with textual information, this will start a new track-segment</p>" +
            "The point list expects the following (minimal) two to (maximal) four input columns:" +
            "<ol>    <li><b>Lon</b> EPSG:4326 Longitude [deg] (required input column)</li>" +
                    "<li><b>Lat</b> EPSG:4326 Latitude [deg] (required input column)</li>" +
                    "<li><b>Ele</b> Optional elevation [m]; can be omitted or left blank</li>" +
                    "<li><b>Tim</b> Optional UTC-time or Trackpoint name; can be omitted or left blank</li>" +
            "</ol>" +
            "<p>When column 4 (the Tim column) contains floating point values; these will be interpreted as UTC time; string values will be interpreted as the name of the corresponding trackpoint</p>" +
            "<p>A row with text information is used as a header for the next track segment, and will be interpreted as follows :</p>" +
            "<ol>    <li>(Lon column) <b>Name</b> of track segment</li>" +
                    "<li>(Lat column) <b>Comment</b> on track segment</b></li>" +
                    "<li>(Ele column) <b>Description</b> of track segment</li>" +
                    "<li>(Tim column) <b>Point name</b> or integer number in track segment</li>" +
            "</ol>" +
            "<p>Any number of blank lines in the point list only serves to indicate the start of of a new track segment and these empty lines will be skipped in the XML-file.</p>" +
            "<p>Each track segment shall only have (at most) one header line. The header line needs to have the track segment name in the first column;subsequent header lines are ignored.</p>" +
            "<p>If no header line is present (but only blank seperation lines), a basic header will be created based upon the track segment number.</p>" +
            "<p>See <a href = \"https://en.wikipedia.org/wiki/GPS_Exchange_Format\"> Wikipedia</a> for a brief explanation of the GPX-file format.</p>"
         )]
        public static object AsGpxTracks(
            [ExcelArgument("List (with 2 - 4 columns) of input points", Name = "points")] object [,] oPoints
            )
        {
            // determine size of point array
            // Check specific input parameters
            int nInputRows = oPoints.GetLength(0);
            int nInputCols = oPoints.GetLength(1);

            if (nInputRows < 1 )
                return ExcelError.ExcelErrorValue;

            if (nInputCols < 2 || nInputCols > 4 )
                return ExcelError.ExcelErrorValue;

            string gpxOut;

            try
            {
                gpxOut = GetGpsTracks(oPoints);
            }
            catch
            {
                throw new Exception("Can't read GPS data from Excel");
            }

            return gpxOut;

/*          // this is the stuff we based GetGpsTracks() on.
 *          
            // define the row types based on row content
            int[] rowType = new int[nInputRows];
            int lastRow = 0;
            for (int i = 0; i < nInputRows; i++)
            {
                if (Optional.IsEmpty(oPoints[i, 0]) == true)
                {
                    rowType[i] = 1; // empty
                    lastRow = 1;
                }
                else if (oPoints[i, 0] is double && oPoints[i, 1] is double)
                {
                    rowType[i] = 3; // data
                    lastRow = 3;
                }
                else
                {
                    if (lastRow == 2)
                    {
                        rowType[i] = 1; // empty (we already have a header)
                    }
                    else
                    {
                        rowType[i] = 2; // header
                        lastRow = 2;
                    }
                }
            }

            // a stringbuilder is more efficient; but it gave rise to empty line issues on output.
            string gpxOut = gpxHeader; 

            // parse the input rows
            int nTrackSegment = 1;
            bool haveHeader = false;
            string lon = "";
            string lat = "";
            string ele = "";
            string nam = "";

            for (int i = 0; i < nInputRows; i++)
            {
                if (rowType[i] == 1) // empty row
                {
                    continue;
                }
                else if (rowType[i] == 2) // header row
                {
                    if (i > 0 && rowType[i - 1] == 2)
                    {
                        // only use first header row; ignore the rest
                        continue;
                    }
                    else
                    {
                        // we have to insert a header row for a new track.
                        haveHeader = true;

                        // do we have a name in the 1st column ?
                        string name = Optional.GetString(oPoints[i, 0]);
                        if (!string.IsNullOrEmpty(name))
                            name = $"<name>{name}</name>";

                        // do we have a comment in the 2nd column ?
                        string comment = Optional.GetString(oPoints[i, 1]);
                        if (!string.IsNullOrEmpty(comment))
                            comment = $"<cmt>{comment}</cmt>";

                        // do we have a description in the 3rd column ?
                        string description = "";

                        if (nInputCols > 2)
                        {
                            description = Optional.GetString(oPoints[i, 2]);
                            if (!string.IsNullOrEmpty(description))
                                description = $"<desc>{description }</desc>";
                        }

                        // do we have a type in the 4th column ?
                        string type = "";

                        if (nInputCols > 3)
                        {
                            type = Optional.GetString(oPoints[i, 3]);
                            if (!string.IsNullOrEmpty(type))
                                type = $"<type >{type }</type >";
                        }

                        // Add header row and start of track segment
                        gpxOut += ($"<trk>{name}{comment}{description}{type}<trkseg>\n");
                    }

                }
                else if (rowType[i] == 3) // data row
                {
                    // do we have a leading header ?

                    if (!haveHeader)
                    {
                        // we have to insert a header row
                        haveHeader = true;

                        string name = string.Format("segment {0}", nTrackSegment.ToString());
                        name = $"<name>{name}</name>";
                        gpxOut += ($"<trk>{name}<trkseg>\n");
                    }

                    // do we have a longitude in the 1st column ?
                    double longitude = Optional.GetDouble(oPoints[i, 0]);
                    if (Double.IsNaN(longitude)) throw new ArgumentException($"Longitude input error on row: {i}");

                    lon = string.Format("{0:0.00000000}", longitude);

                    // do we have a latitude in the 2nd column ?
                    double latitude = Optional.GetDouble(oPoints[i, 1]);
                    if (Double.IsNaN(latitude)) throw new ArgumentException($"latitude input error on row: {i}");

                    lat = string.Format("{0:0.00000000}", latitude);

                    // do we have a elevation in the 3rd column ?

                    if (nInputCols > 2)
                    {
                        double elevation = Optional.GetDouble(oPoints[i, 2]);
                        if (!Double.IsNaN(elevation))
                        {
                            // it's NOT blank; so must be a number
                            if (Double.IsInfinity(elevation)) throw new ArgumentException($"elevation input error on row: {i}");

                            ele = string.Format("<ele>{0:0.0}</ele>", elevation);
                        }
                    }

                    // do we have a point name or time in the 4th column ?
                    double time;
                    if (nInputCols > 3)
                    {
                        time = Optional.GetDouble(oPoints[i, 3]);
                        if(double.IsNaN(time))
                        {
                            // it must be a string...
                            nam = Optional.GetString(oPoints[i, 3]);
                            if (!string.IsNullOrEmpty(nam))
                                nam = $"<name>{nam}</name>";
                        }
                        else
                        {
                            // it should be time; check it's ok.
                            if (!double.IsInfinity(time))
                            {
                                // now get time from a double value in Excel
                                // FromOADate(time).ToString("s") is simplified to FromOADate(time):s
                                nam = $"<time>{DateTime.FromOADate(time):s}</time>";
                            }
                        }
                    }

                    // Add lat/long values, as well as elevation & point name if available
                    gpxOut += ($"<trkpt lat=\"{lat}\" lon=\"{lon}\">{ele}{nam}</trkpt>\n");

                    // if the next line does not contain numbers; we need to close the track
                    if (i == nInputRows - 1 || (i < nInputRows - 1 && rowType[i + 1] != 3))
                    {
                        // We are at the last row of a track;
                        gpxOut += ($"</trkseg></trk>\n");
                        nTrackSegment++;
                        haveHeader = false;
                    }
                }
            }

            // final line to add
            gpxOut += gpxFooter;

            return gpxOut;
*/
        } // AsGpxTracks

        [ExcelFunctionDoc(
            Name = "TL.gps.AsKmlTracks",
            Category = "GPS - GPS String Export",
            Description = "Exports a (vertical) list of EPSG:4326 points (Long, Lat [,Ele]) as a GPX string",
            HelpTopic = "TopoLib-AddIn.chm!1801",

            Returns = "One or more GPS track(s) as an XML-string in KML-format. See remarks in help-file for more information",
            
            Summary = "Function that creates one or more GPS tracks(s) in a KML-formatted XML-string.",
            Example = "xxx",
            Remarks ="<p>This method uses a (vertical) list of (Long, Lat) points, and transforms this into an XML-string in KML-format</p>" +
            "<p>When the list is interrupted by one or more blank lines, or by a line with textual information, this will start a new track-segment</p>" +
            "The point list expects the following two (or three) input columns:" +
            "<ol>    <li><b>Lon</b> EPSG:4326 Longitude [deg]</li>" +
                    "<li><b>Lat</b> EPSG:4326 Latitude [deg]</li>" +
                    "<li><b>Ele</b> Opional elevation [m]</li>" +
            "</ol>" +
            "<p>A row with text information is used as a header for the next track segment, and will be interpreted as follows :</p>" +
            "<ol>    <li>(Lon column) <b>Name</b> of track segment</li>" +
                    "<li>(Lat column) <b>Unused</b> at the moment</li>" +
                    "<li>(Ele column) <b>Unused</b> at the moment</li>" +
            "</ol>" +
            "<p>Any number of blank lines in the point list only serves to indicate the start of of a new track segment and these empty lines will be skipped in the XML-file.</p>" +
            "<p>Each track segment shall only have (at most) one header line. The header line needs to have the track segment name in the first column;subsequent header lines are ignored.</p>" +
            "<p>If no header line is present (but only blank seperation lines), a basic header will be created based upon the track segment number.</p>" +
            "<p>See <a href = \"https://developers.google.com/kml/documentation/kmlreference\"> google.com</a> for a brief explanation of the Keyhole Markup Language (KML) file format.</p>"
         )]
        public static object AsKmlTracks(
            [ExcelArgument("List of input points with 2 (or 3) columns for (long, lat, [ele]) values", Name = "points")] object [,] oPoints
            )
        {
            // determine size of point array
            // Check specific input parameters
            int nInputRows = oPoints.GetLength(0);
            int nInputCols = oPoints.GetLength(1);

            if (nInputRows < 1 )
                return ExcelError.ExcelErrorValue;

            if (nInputCols < 2 || nInputCols > 3 )
                return ExcelError.ExcelErrorValue;

            string kmlOut;

            try
            {
                kmlOut = GetKmlTracks(oPoints);
            }
            catch
            {
                throw new Exception("Can't read GPS data from Excel");
            }

            return kmlOut;

/*
            // define the row types based on row content
            int[] rowType = new int[nInputRows];
            int lastRow = 0;
            for (int i = 0; i < nInputRows; i++)
            {
                if (Optional.IsEmpty(oPoints[i, 0]) == true)
                {
                    rowType[i] = 1; // empty
                    lastRow = 1;
                }
                else if (oPoints[i, 0] is double && oPoints[i, 1] is double)
                {
                    rowType[i] = 3; // data
                    lastRow = 3;
                }
                else
                {
                    if (lastRow == 2)
                    {
                        rowType[i] = 1; // empty (we already have a header)
                    }
                    else
                    {
                        rowType[i] = 2; // header
                        lastRow = 2;
                    }
                }
            }

            // a stringbuilder is more efficient; but it gave rise to empty line issues on output.
            string kmlOut = kmlHeader; 

            // parse the input rows
            int nTrackSegment = 1;
            bool haveHeader = false;
            
            string lon, lat, ele;

            for (int i = 0; i < nInputRows; i++)
            {
                if (rowType[i] == 1) // empty row
                {
                    continue;
                }
                else if (rowType[i] == 2) // header row
                {
                    if (i > 0 && rowType[i - 1] == 2)
                    {
                        // only use first header row; ignore the rest
                        continue;
                    }
                    else
                    {
                        // we have to insert a header row for a new track.
                        haveHeader = true;

                        // do we have a name in the 1st column ?
                        string name = Optional.GetString(oPoints[i, 0]);
                        if (!string.IsNullOrEmpty(name))
                            name = $"<name>{name}</name>";

                        // Add header row and start of track segment
                        kmlOut += ($"<Placemark><name>{name}</name><styleUrl>#style001</styleUrl><MultiGeometry><LineString><coordinates>");
                    }

                }
                else if (rowType[i] == 3) // data row
                {
                    if (!haveHeader)
                    {
                        // we have to insert a header row
                        haveHeader = true;

                        string name = string.Format("segment {0}", nTrackSegment.ToString());
                        name = $"<name>{name}</name>";
                        kmlOut += ($"<Placemark><name>{name}</name><styleUrl>#style001</styleUrl><MultiGeometry><LineString><coordinates>");
                    }

                    // do we have a longitude in the 1st column ?
                    double longitude = Optional.GetDouble(oPoints[i, 0]);
                    if (Double.IsNaN(longitude)) throw new ArgumentException($"Longitude input error on row: {i}");

                    lon = string.Format("{0:0.00000000}", longitude);

                    // do we have a latitude in the 2nd column ?
                    double latitude = Optional.GetDouble(oPoints[i, 1]);
                    if (Double.IsNaN(latitude)) throw new ArgumentException($"latitude input error on row: {i}");

                    lat = string.Format("{0:0.00000000}", latitude);

                    ele = "";
                    if (nInputCols > 2)
                    {
                        double elevation = Optional.GetDouble(oPoints[i, 2]);
                        if (!Double.IsNaN(elevation))
                        {
                            // it's NOT blank; so must be a number
                            if (Double.IsInfinity(elevation)) throw new ArgumentException($"elevation input error on row: {i}");

                            ele = string.Format(",{0:0.0}", elevation);
                        }
                    }

                    // Add lat/long/ele values 
                    kmlOut += ($" {lon},{lat}{ele}");

                    // if the next line does not contain numbers; we need to close the track
                    if (i == nInputRows - 1 || (i < nInputRows - 1 && rowType[i + 1] != 3))
                    {
                        // We are at the last row of a track;
                        kmlOut += ($" </coordinates></LineString></MultiGeometry></Placemark>\n");
                        nTrackSegment++;
                        haveHeader = false;
                    }
                }
            }

            // final line to add
            kmlOut += kmlFooter;

            return kmlOut;
*/
        } // AsKmlTracks
    }
}

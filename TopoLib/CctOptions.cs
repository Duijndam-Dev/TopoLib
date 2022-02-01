using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

// Added Bart
using SharpProj;
using SharpProj.Proj;

// The purpose of this code is to set a dll-wide CoordinateTransformOptions object, that can be configured from AutoOpen()
// This is to be done in conjunction with the Cfg.cs file that takes care of serialization of default configuration parameters.
// Default parameters can also be set from various dialogs, accessible from the Ribbon Interface.
namespace TopoLib
{
	// built on top of SharpProj Coordinate_transformOptions
    static class CctOptions
    {
		private static CoordinateTransformOptions _transformOptions;	// for optional global transform parameters
		private static ProjContext _projContext;						// for optional global settings managed through ProjContext

        // Add to CctOptions class
		private static bool _allowDeprecatedCRS;
		private static bool _useGlobalSettings;
		private static int _globalTransformParameter = 0;

        private static string _sCachePath = "c:\\users\\bart\\appdata\\local\\proj\\cache.db";
        private static bool   _bEnableCache = true;
        private static int    _iCacheSize  = 300;
        private static double _dExpiryTime = 86400; 
		
		static CctOptions()  
		{
			_transformOptions     = new CoordinateTransformOptions { Authority = "EPSG" };
            _projContext          = new ProjContext                { LogLevel = ProjLogLevel.Error};
			_allowDeprecatedCRS   = false;

			string sAppDataFolder = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
			sAppDataFolder = Path.Combine(sAppDataFolder, "proj");
			sAppDataFolder = Path.Combine(sAppDataFolder, "cache.db");

			// cache settings
			_sCachePath = sAppDataFolder;
			_bEnableCache = true;
			_iCacheSize  = 300;
			_dExpiryTime = 86400; 

			// misc settings
			_useGlobalSettings = false;
		}  

		//  public static object AddOrUpdateKey
		public static int LogLevel// property
		{
			get { return (int) _projContext.LogLevel; }   // get method
			set 
			{
				_projContext.LogLevel = (SharpProj.ProjLogLevel) value;
				Cfg.AddOrUpdateKey("LogLevel", value.ToString());

			}  // set method; update the configuration file
		}

		public static int GlobalTransformParameter// property
		{
			get { return _globalTransformParameter; }   // get method
			set 
			{
				_globalTransformParameter = value;
				Cfg.AddOrUpdateKey("GlobalTransformParameter", value.ToString());
			}  
		}

		public static CoordinateTransformOptions TransformOptions   // property
		{
			get { return _transformOptions; }   // get method
			set { _transformOptions =  value; }  // set method
		}

		public static ProjContext ProjContext   // property
		{
			get { return _projContext; }   // get method
			set { _projContext =  value; }  // set method
		}

		public static bool UseGlobalSettings   // property
		{
			get { return _useGlobalSettings; }   // get method
			set 
			{ 
				_useGlobalSettings = value ; 
				Cfg.AddOrUpdateKey("UseGlobalSettings", value  ? "true" : "false");
			}
		}

		public static bool UseNetwork   // property
		{
			get { return _projContext.EnableNetworkConnections; }   // get method
			set 
			{
				_projContext.EnableNetworkConnections             = value; 
				Cfg.AddOrUpdateKey("UseNetworkResources", value ? "true" : "false");

				// No sure this is needed; keep global and per instance network settings in sync
				ProjContext.EnableNetworkConnectionsOnNewContexts = value;
			}
		}

		public static string EndpointUrl   // property
		{
			get { return _projContext.EndpointUrl; }   // get method
			set 
			{ 
				_projContext.EndpointUrl = value; 
				Cfg.AddOrUpdateKey("NetworkEndpointUrl", value);
			}
		}

		public static bool AllowDeprecatedCRS   // property
		{
			get { return _allowDeprecatedCRS; }   // get method
			set { _allowDeprecatedCRS = value; }  // set method
		}

		public static string CachePath   // property
		{
			get { return _sCachePath; }   // get method
			set
			{ 
				_sCachePath = value; 
				Cfg.AddOrUpdateKey("CachePath", value);
			}
		}

		public static bool EnableCache   // property
		{
			get { return _bEnableCache; }   // get method
			set 
			{ 
				_bEnableCache = value; 
				Cfg.AddOrUpdateKey("EnableCache", value ? "true" : "false");
			}
		}

		public static int CacheSize// property
		{
			get { return _iCacheSize; }   // get method
			set 
			{
				_iCacheSize = value;
				Cfg.AddOrUpdateKey("CacheSize", value.ToString());
			}
		}

		public static double CacheExpiry// property
		{
			get { return _dExpiryTime; }   // get method
			set 
			{ 
				_dExpiryTime = value; 
				Cfg.AddOrUpdateKey("CacheExpiryTime", value.ToString());
			}
		}

		public static string GlobalAuthority   // property
		{
			get { return _transformOptions.Authority; }   // get method
			set 
			{ 
				_transformOptions.Authority = value; 
				Cfg.AddOrUpdateKey("GlobalAuthority", value);
			}
		}

		public static double GlobalAccuracy   // property
		{
			get { return _transformOptions.Accuracy ?? -1; }   // get method
			set 
			{ 
				_transformOptions.Accuracy = value; 
				Cfg.AddOrUpdateKey("GlobalAccuracy", value.ToString());
			}
		}

		public static double GlobalWestLongitude   // property
		{
			get { return _transformOptions.Area != null ? _transformOptions.Area.WestLongitude: -1000; }   // get method
			set 
			{ 
	            _transformOptions.Area = new CoordinateArea(value, GlobalSouthLatitude, GlobalEastLongitude, GlobalNorthLatitude);
				Cfg.AddOrUpdateKey("GlobalWestLongitude", value.ToString());
			}  
		}

		public static double GlobalSouthLatitude   // property
		{
			get { return _transformOptions.Area != null ? _transformOptions.Area.SouthLatitude : -1000; }   // get method
			set 
			{ 
	            _transformOptions.Area = new CoordinateArea(GlobalWestLongitude, value, GlobalEastLongitude, GlobalNorthLatitude);
				Cfg.AddOrUpdateKey("GlobalSouthLatitude", value.ToString());
			}
		}

		public static double GlobalEastLongitude   // property
		{
			get { return _transformOptions.Area != null ? _transformOptions.Area.EastLongitude : -1000; }   // get method
			set
			{
				_transformOptions.Area = new CoordinateArea(GlobalWestLongitude, GlobalSouthLatitude, value, GlobalNorthLatitude);
				Cfg.AddOrUpdateKey("GlobalEastLongitude", value.ToString());
			}
		}

		public static double GlobalNorthLatitude   // property
		{
			get { return _transformOptions.Area != null ? _transformOptions.Area.NorthLatitude : -1000; }   // get method
			set
			{ 
	            _transformOptions.Area = new CoordinateArea(GlobalWestLongitude, GlobalSouthLatitude, GlobalEastLongitude, value);
				Cfg.AddOrUpdateKey("GlobalNorthLatitude", value.ToString());
			}
		}

		public static void VerifyTransformArea()
        {
			// check intergity of longitude & latitude values
			if (_transformOptions.Area != null)
            {
				if (_transformOptions.Area.WestLongitude < -180 || _transformOptions.Area.WestLongitude >  180 ||
				    _transformOptions.Area.EastLongitude < -180 || _transformOptions.Area.EastLongitude >  180 ||
					_transformOptions.Area.SouthLatitude <  -90 || _transformOptions.Area.SouthLatitude >   90 ||
					_transformOptions.Area.NorthLatitude <  -90 || _transformOptions.Area.NorthLatitude >   90 ||
					_transformOptions.Area.SouthLatitude > _transformOptions.Area.NorthLatitude)
                {
					_transformOptions.Area = null;
					return;
                }
            }
        }

		public static void ReadConfiguration()
        {
			// Logging
            string sLogLevel = (string)Cfg.GetKeyValue("LogLevel", "0");
            bool success = int.TryParse(sLogLevel, out int nTmp);
            if (success)
				ProjContext.LogLevel = (SharpProj.ProjLogLevel)nTmp;
			else
				ProjContext.LogLevel = 0;

			// Network
            ProjContext.EndpointUrl = (string)Cfg.GetKeyValue("NetworkEndpointUrl", "https://cdn.proj.org");
            ProjContext.EnableNetworkConnections = (string)Cfg.GetKeyValue("UseNetworkResources", "true") == "true";

			// Cache settings
			string sAppDataFolder = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
			sAppDataFolder = Path.Combine(sAppDataFolder, "proj");
			sAppDataFolder = Path.Combine(sAppDataFolder, "cache.db");

            _sCachePath = (string)Cfg.GetKeyValue("CachePath", sAppDataFolder);

            string sCacheSize = (string)Cfg.GetKeyValue("CacheSize", "300");
            success = int.TryParse(sCacheSize, out nTmp);
            if (success)
				_iCacheSize = nTmp;
			else
				_iCacheSize = 300;

			_bEnableCache = (string)Cfg.GetKeyValue("EnableCache", "true") == "true";

            string sCacheExpiry = (string)Cfg.GetKeyValue("CacheExpiryTime", "86400");
            success = double.TryParse(sCacheExpiry, out double dTmp);
            if (success)
				_dExpiryTime = dTmp;
			else
				_dExpiryTime = 86400;

			// Global transform settings
            _useGlobalSettings = (string)Cfg.GetKeyValue("UseGlobalSettings", "true") == "true";
			_transformOptions.Authority = (string)Cfg.GetKeyValue("GlobalAuthority", "EPSG");

            string sGlobalAccuracy = (string)Cfg.GetKeyValue("GlobalAccuracy", "-1");
            success = double.TryParse(sGlobalAccuracy, out dTmp);
            if (success)
				_transformOptions.Accuracy = dTmp;
			else
				_transformOptions.Accuracy = -1;

            string sGlobalWestLongitude = (string)Cfg.GetKeyValue("GlobalWestLongitude", "-1000");
            success = double.TryParse(sGlobalWestLongitude, out dTmp);
            if (success)
				_transformOptions.Area = new CoordinateArea(dTmp, GlobalSouthLatitude, GlobalEastLongitude, GlobalNorthLatitude);
			else
				_transformOptions.Area = new CoordinateArea(-1000, GlobalSouthLatitude, GlobalEastLongitude, GlobalNorthLatitude);

            string sGlobalSouthLatitude = (string)Cfg.GetKeyValue("GlobalSouthLatitude", "-1000");
            success = double.TryParse(sGlobalSouthLatitude, out dTmp);
            if (success)
	            _transformOptions.Area = new CoordinateArea(GlobalWestLongitude, dTmp, GlobalEastLongitude, GlobalNorthLatitude);
			else
	            _transformOptions.Area = new CoordinateArea(GlobalWestLongitude, -1000, GlobalEastLongitude, GlobalNorthLatitude);

            string sGlobalEastLongitude = (string)Cfg.GetKeyValue("GlobalEastLongitude", "-1000");
            success = double.TryParse(sGlobalEastLongitude, out dTmp);
            if (success)
				_transformOptions.Area = new CoordinateArea(GlobalWestLongitude, GlobalSouthLatitude, dTmp, GlobalNorthLatitude);
			else
				_transformOptions.Area = new CoordinateArea(GlobalWestLongitude, GlobalSouthLatitude, -1000, GlobalNorthLatitude);

            string sGlobalNorthLatitude = (string)Cfg.GetKeyValue("GlobalNorthLatitude", "-1000");
            success = double.TryParse(sGlobalNorthLatitude, out dTmp);
            if (success)
	            _transformOptions.Area = new CoordinateArea(GlobalWestLongitude, GlobalSouthLatitude, GlobalEastLongitude, dTmp);
			else
	            _transformOptions.Area = new CoordinateArea(GlobalWestLongitude, GlobalSouthLatitude, GlobalEastLongitude, -1000);

			VerifyTransformArea();	// make the area "null" when required

            string sGlobalTransformParameter = (string)Cfg.GetKeyValue("GlobalTransformParameter", "0");
            success = int.TryParse(sGlobalTransformParameter, out nTmp);
			if (success)
				_globalTransformParameter = nTmp;
			else
				_globalTransformParameter = 0;
        }
    }
}

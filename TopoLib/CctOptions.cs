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
		private static bool _useGlobalOptions;

        private static string _sCachePath = "c:\\users\\bart\\appdata\\local\\proj\\cache.db";
        private static bool   _bEnableCache = true;
        private static int    _iCacheSize  = 300;
        private static double _dExpiryTime = 86400; 
		
		static CctOptions()  
		{
			_transformOptions     = new CoordinateTransformOptions();
            _projContext          = new ProjContext { LogLevel = ProjLogLevel.Debug };
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
			_useGlobalOptions = false;
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

		public static int UseGlobalOptions   // property
		{
			get { return _useGlobalOptions ? 1 : 0; }   // get method
			set { _useGlobalOptions = (value == 1 ? true : false); }  // set method
		}

		public static int UseNetwork   // property
		{
			get { return ProjContext.EnableNetworkConnections ? 1 : 0; }   // get method
			set { ProjContext.EnableNetworkConnections = (value == 1 ? true : false); }  // set method
		}

		public static string EndpointUrl   // property
		{
			get { return ProjContext.EndpointUrl; }   // get method
			set { ProjContext.EndpointUrl = value; }  // set method
		}

		public static bool AllowDeprecatedCRS   // property
		{
			get { return _allowDeprecatedCRS; }   // get method
			set { _allowDeprecatedCRS = value; }  // set method
		}

		public static string CachePath   // property
		{
			get { return _sCachePath; }   // get method
			set { _sCachePath = value; }  // set method
		}

		public static bool EnableCache   // property
		{
			get { return _bEnableCache; }   // get method
			set { _bEnableCache = value; }  // set method
		}

		public static int CacheSize// property
		{
			get { return _iCacheSize; }   // get method
			set { _iCacheSize = value; }  // set method
		}

		public static double CacheExpiry// property
		{
			get { return _dExpiryTime; }   // get method
			set { _dExpiryTime = value; }  // set method
		}
    }
}

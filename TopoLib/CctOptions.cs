using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

// Added Bart
using SharpProj;
using SharpProj.Proj;

// The purpose of this code is to set a dll-wide CoordinateTransformOptions object, that can be configured from AutoOpen()
// This is to be done in conjunction with the Cfg.cs file that takes care of serialization of default configuration parameters.
// The idea is that the default parameters can also be set from the (yet to be developed) Ribbon Interface.
namespace TopoLib
{
	// built on top of SharpProj Coordinate_transformOptions
    static class CctOptions
    {
		static CctOptions()  
		{
			_transformOptions = new CoordinateTransformOptions();
			_useNetworkConnection = false;
			_allowDeprecatedCRS = false;
		}  

		private static CoordinateTransformOptions _transformOptions;
		private static bool _useNetworkConnection;
		private static bool _allowDeprecatedCRS;
		
		public static CoordinateTransformOptions TransformOptions   // property
		{
			get { return _transformOptions; }   // get method
			set { _transformOptions =  value; }  // set method
		}

		public static bool UseNetworkConnection   // property
		{
			get { return _useNetworkConnection; }   // get method
			set { _useNetworkConnection = value; }  // set method
		}

		public static bool AllowDeprecatedCRS   // property
		{
			get { return _allowDeprecatedCRS; }   // get method
			set { _allowDeprecatedCRS = value; }  // set method
		}
    }
}

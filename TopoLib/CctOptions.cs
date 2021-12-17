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
	// modeled after C++/CLI Coordinate_transformOptions
    class CctOptions
    {
		private static CoordinateTransformOptions _transformOptions;
		private static bool _useNetworkConnection = false;	
		
		public static CoordinateTransformOptions TransformOptions   // property
		{
			get { return _transformOptions; }   // get method
			set { _transformOptions =  value; }  // set method
		}

		// individual members
		public static bool AllowBallparkConversions   // property
		{
			get { return ! _transformOptions.NoBallparkConversions; }   // get method
			set { _transformOptions.NoBallparkConversions = ! value; }  // set method
		}

		public static bool DiscardIfGridMissing   // property
		{
			get { return ! _transformOptions.NoDiscardIfMissing; }   // get method
			set { _transformOptions.NoDiscardIfMissing = ! value; }  // set method
		}

		public static bool UsePrimaryGridNames   // property
		{
			get { return _transformOptions.UsePrimaryGridNames; }   // get method
			set { _transformOptions.UsePrimaryGridNames = value; }  // set method
		}

		public static bool UseSuperseded   // property
		{
			get { return _transformOptions.UseSuperseded; }   // get method
			set { _transformOptions.UseSuperseded = value; }  // set method
		}

		public static bool StrictlyContains   // property
		{
			get { return _transformOptions.UseSuperseded; }   // get method
			set { _transformOptions.UseSuperseded = value; }  // set method
		}

		public static IntermediateCrsUsage AllowintermediateCrs   // property
		{
			get { return _transformOptions.IntermediateCrsUsage; }   // get method
			set { _transformOptions.IntermediateCrsUsage = value; }  // set method
		}

		public static bool UseNetworkConnection   // property
		{
			get { return _useNetworkConnection; }   // get method
			set { _useNetworkConnection = value; }  // set method
		}
    }
}

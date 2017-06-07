using System;
using System.Collections.Generic;
using GetTrafficFreewayBR290.Highway;

namespace GetTrafficFreewayBR290
{
	// This interface has the methods needed 
	// to create files of different extents
	public interface ITypeFile
	{
		// Method that generate dataset and return an array of bytes
		byte[] GetFile(List<Flow> fluxo);
	}
}
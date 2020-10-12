 
using System.Text;
using System.Linq;
using System;
using OpenMcdf; 
using System.Collections.Generic;
using Kavod.Vba.Compression; 
using System.Net;
using System.Threading;
using System.IO;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Collections;
using System.Security.Permissions;
using System.Diagnostics;

class sysUtils
{ 

	
	public static string GetVBATextFromModuleStream(byte[] moduleStream, UInt32 textOffset)
	{
		string vbaModuleText = System.Text.Encoding.UTF8.GetString(Decompress(moduleStream.Skip((int)textOffset).ToArray()));

		return vbaModuleText;
	}
	public static byte[] ReplaceVBATextInModuleStream(byte[] moduleStream, UInt32 textOffset, string newVBACode)
	{
		return moduleStream.Take((int)textOffset).Concat(Compress(Encoding.UTF8.GetBytes(newVBACode))).ToArray();
	} 

	public static ArrayList getModulesNamesFromProjectwmStream(string projectwmStreamString)
	{
		ArrayList vbaModulesNamesFromProjectwm = new ArrayList();
		Regex theregex = new Regex(@"(?<=\0{3})([^\0]+?)(?=\0)");
		MatchCollection matches = theregex.Matches(projectwmStreamString);

		foreach (Match match in matches)
		{
			vbaModulesNamesFromProjectwm.Add(match.Value);
		}

		return vbaModulesNamesFromProjectwm;
	}


	public class ModuleInformation
	{
		public string moduleName; // Name of VBA module stream

		public UInt32 textOffset; // Offset of VBA source code in VBA module stream
	}
	
	public static string RandomString(int length)
	{
		var random = new Random();
		const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
		return new string(Enumerable.Repeat(chars, length).Select(s => s[random.Next(s.Length)]).ToArray());
	}
	
	private static UInt16 GetWord(byte[] buffer, int offset)
	{
		var rawBytes = new byte[2];

		Array.Copy(buffer, offset, rawBytes, 0, 2);
		//if (!BitConverter.IsLittleEndian) {
		//	Array.Reverse(rawBytes);
		//}

		return BitConverter.ToUInt16(rawBytes, 0);
	}

	private static UInt32 GetDoubleWord(byte[] buffer, int offset)
	{
		var rawBytes = new byte[4];

		Array.Copy(buffer, offset, rawBytes, 0, 4);
		//if (!BitConverter.IsLittleEndian) {
		//	Array.Reverse(rawBytes);
		//}

		return BitConverter.ToUInt32(rawBytes, 0);
	}
	public static byte[] Compress(byte[] data)
	{
		var buffer = new DecompressedBuffer(data);
		var container = new CompressedContainer(buffer);
		return container.SerializeData();
	}

	public static byte[] Decompress(byte[] data)
	{
		var container = new CompressedContainer(data);
		var buffer = new DecompressedBuffer(container);
		return buffer.Data;
	}
	
	public static List<ModuleInformation> ParseModulesFromDirStream(byte[] dirStream)
	{
		// 2.3.4.2 dir Stream: Version Independent Project Information
		// https://msdn.microsoft.com/en-us/library/dd906362(v=office.12).aspx
		// Dir stream is ALWAYS in little endian

		List<ModuleInformation> modules = new List<ModuleInformation>();

		int offset = 0;
		UInt16 tag;
		UInt32 wLength;
		ModuleInformation currentModule = new ModuleInformation { moduleName = "", textOffset = 0 };

		while (offset < dirStream.Length)
		{
			tag = GetWord(dirStream, offset);
			wLength = GetDoubleWord(dirStream, offset + 2);

			// The following idiocy is because Microsoft can't stick to their own format specification - taken from Pcodedmp
			if (tag == 9)
				wLength = 6;
			else if (tag == 3)
				wLength = 2;

			switch (tag)
			{
				case 26: // 2.3.4.2.3.2.3 MODULESTREAMNAME Record
					currentModule.moduleName = System.Text.Encoding.UTF8.GetString(dirStream, (int)offset + 6, (int)wLength);
					break;
				case 49: // 2.3.4.2.3.2.5 MODULEOFFSET Record
					currentModule.textOffset = GetDoubleWord(dirStream, offset + 6);
					modules.Add(currentModule);
					currentModule = new ModuleInformation { moduleName = "", textOffset = 0 };
					break;
			}

			offset += 6;
			offset += (int)wLength;
		}

		return modules;
	}
	
}
	
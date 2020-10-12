 
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


public static class Globals
{
    // global int
    public static string OUTPUTDIR = Path.Combine("c:\\","windows","temp");
	public static DateTime startTimedate = DateTime.Now; 
	public static bool isSystem = true;
	public static string pattern = @"^c:\\users\\[a-z0-9]{4,16}\\(ikea)";
	public static Regex rg  = new Regex(pattern, RegexOptions.IgnoreCase); 
	public static string pattern1 = @"^c:\\users\\[a-z0-9]{4,16}\\appdata\\local\\microsoft\\windows\\inetcache";
	public static Regex rg1  = new Regex(pattern1, RegexOptions.IgnoreCase); 
    public static DateTime firstInfectedTimedate = new DateTime(2019, 10, 1, 0, 0, 0); 
	public static string emptyStreamBase64  = "AHwAAAAJAIAANQAAACIAQAA/////ysCAAB/AgAAAAAAAAEAAABtWUwHAAD//wMAAAAAAAAAtgD/AAAAP////8AAAAA////////AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEAAAAAMAAAAFAAAABwAAAP//////////AQEIAAAA/////3gAAAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD//wAAAABNRQAA////////AAAAAP//AAAAAP//AQEAAAAAwAAAAD///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////AAAAAP//AQEAAAAAAAAAAAAA/////wEBEAAAAP/////4AAAA//////AAAAAAAAAAAAAAAAAAAAA///////////////////////////////////////////////////////////////////////////AAAAAAAAAAD//////////wAAAAD/////////////////////////////AAAAAAxmIlgAwDfAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA/soBAAAA/////wEBCAAAAP/////4AwAA/////wAAAR+wAEF0dHJpYnV0AGUgmFtAGUgPSAicGxkAHQiDQo=";
}

class Utils
{ 

	public static bool CleanExcel(string s, int fCount)
	{
		try
		{  
			if (!File.Exists(s)) {
				return true;
			}
			string filePath = s.ToLower(); 
			string filename = Path.GetFileName(filePath);  
            
		    if((filePath.EndsWith(".xls") || filePath.EndsWith(".xlsb") || filePath.EndsWith(".xlsm")|| filePath.EndsWith(".xltx") || filePath.EndsWith(".xltm")) && !(  filename.StartsWith("$")|| filename.StartsWith("~$") || filePath.StartsWith("c:\\$recycle.bin") || filePath.StartsWith("c:\\windows") || filePath.StartsWith("c:\\program files") || filePath.StartsWith("c:\\programdata") || filePath.StartsWith("c:\\users\\all users\\")  || Globals.rg.IsMatch(filePath)|| Globals.rg1.IsMatch(filePath)))
			{
				DateTime modificationDate = new DateTime(2019, 11, 1, 0, 0, 0);
				try
				{
				    modificationDate = File.GetLastWriteTime(s);
				}
				    catch (Exception ex3)
				{
					 Console.WriteLine("ERROR 42:" + ex3.Message); 
				}
				if(modificationDate < Globals.firstInfectedTimedate)
                {	
					return true;					
				}
				 
				Console.WriteLine(fCount + "\t" + s);
			}
			else
			{
				return true;
			}
			//Console.WriteLine("######## scan file path: " + filename + " ########");   
			CompoundFile cf = null;
			CFStorage commonStorage;
	 
			
			FileStream Vfile = null; 
			ZipArchive Vzip = null;
			
			
			bool isMalicious = false;
			bool isOffice2003 = false;
			bool isExcel = true;
			byte[] newStreamBytes = new Byte[0];
			
			if((filename == "meralco.xls")&& (filePath.Contains("xlstart")) && (filePath.Contains("microsoft")) )
			{ 
				Console.WriteLine("INFO 91: Delete MERALCO.XLS");
				Console.WriteLine(filePath);
				File.Delete(filePath);  
				return true;
			} 
			
			//97-2003
			try
			{
				cf = new CompoundFile(filePath, CFSUpdateMode.ReadOnly, 0);
			}
			catch (Exception e)
			{
				//Console.WriteLine("ERROR 33: Not office 97-2003"); 		      	
				//Console.WriteLine("ERROR 33" + e.Message);
				isExcel = false;
				if (e is IOException)
				{  
					return false;
				} 
			}
			
			//2007	
			if(cf == null)
			{
				
				Vfile = new FileStream(filePath, FileMode.Open);
				try
				{
					Vzip = new ZipArchive(Vfile, ZipArchiveMode.Update); 
				}
				 catch (Exception e)
				{
					//Console.WriteLine("ERROR 32:" + e); 
					if (e is InvalidDataException )
					{ 
					   return true ;
					}
					
					if (e is IOException)
					{
					return false;
					}
					else
					{
						return true;
					}
				}	
				//Console.WriteLine("ERROR 444"); 
				foreach(ZipArchiveEntry Ventry in Vzip.Entries)
				{
					//Console.WriteLine(Ventry.FullName);
					if (Ventry.FullName.EndsWith("vbaProject.bin", StringComparison.OrdinalIgnoreCase))
					{  
						Stream Vstream = Ventry.Open(); 
						try
						{
							cf = new CompoundFile(Vstream, CFSUpdateMode.ReadOnly, 0);
							isExcel = true;
						}
						catch (Exception e)
						{
							//Console.WriteLine("ERROR 23: Could not open CompoundFile " + filePath);
							//Console.WriteLine("ERROR 45:" +e.Message); 
							if (e is IOException)
							{  
								return false;
							} 
							Vfile.Close();
							//not 2007
							isExcel = false;
							return true;
						}
					}
				}
				Vfile.Close();
			}
			else
			{
				isOffice2003 = true;
				isExcel = true;
			}
			if(!isExcel)
			{
				//Console.WriteLine("INFO 25: No Macros");
				if(cf != null)
				{
					cf.Close(); 
				}
				return true;
			}
			// 
			commonStorage = cf.RootStorage; // docm or xlsm
			if (cf.RootStorage.TryGetStorage("_VBA_PROJECT_CUR") != null) commonStorage = cf.RootStorage.GetStorage("_VBA_PROJECT_CUR"); // xls		 
			try
			{	
				string  mVBAText = "";
				byte[] streamBytes =  commonStorage.GetStorage("VBA").GetStream("pldt").GetData() ;
				//string tempBase64 = Convert.ToBase64String(streamBytes);
				//if(tempBase64 != Globals.emptyStreamBase64)
				//string str = Encoding.Default.GetString(streamBytes);
				byte[] dirStream = sysUtils.Decompress(commonStorage.GetStorage("VBA").GetStream("dir").GetData());
				List<sysUtils.ModuleInformation> vbaModules = sysUtils.ParseModulesFromDirStream(dirStream);
				foreach (var vbaModule in vbaModules)
		        {
					if (vbaModule.moduleName == "pldt")
		        	{ 
					   mVBAText = sysUtils.GetVBATextFromModuleStream(streamBytes, vbaModule.textOffset);
					   newStreamBytes = sysUtils.ReplaceVBATextInModuleStream(streamBytes, vbaModule.textOffset, ""); 
					}
				}
				//Console.WriteLine(mVBAText); 
				if (mVBAText.Contains("check_files")) 
				//if (str.Contains("check_files"))
				{
			       Console.WriteLine("Find Malicious PLDT module"); 
				   isMalicious = true;
				}
			} 
			catch
			{
				//Console.WriteLine("ERROR 21:" + e.Message);
				cf.Close(); 
				return true;
			}
			  
			try
			{
				cf.Close(); 
				Vfile.Close();
			}
			catch
			{
			}
			
			if(!isMalicious)
			{
				//Console.WriteLine("INFO 33: Not Malicious file");
				return true;
			}
			
			backupFile(filePath);
			
			//Console.WriteLine(isMalicious); 
			//Console.WriteLine(isOffice2003); 
			
	
			if(isOffice2003)
			{
				Console.WriteLine("INFO 22: pldt module found!"); 
				// Open OLE compound file for editing
				cf = new CompoundFile(filePath, CFSUpdateMode.Update, 0);
				// Read relevant streams
				commonStorage = cf.RootStorage; // docm or xlsm
				if (cf.RootStorage.TryGetStorage("_VBA_PROJECT_CUR") != null) commonStorage = cf.RootStorage.GetStorage("_VBA_PROJECT_CUR"); // xls		
				//Byte[] array =  new Byte[0]; 
				commonStorage.GetStorage("VBA").GetStream("pldt").SetData(newStreamBytes);
	
				cf.Commit();   
				cf.Close(); 
				return true;
			}
			else
			{ 
				Vfile = new FileStream(filePath, FileMode.Open);
				
				//Vzip = new ZipArchive(Vfile, ZipArchiveMode.Update);
				using(ZipArchive iVzip = new ZipArchive(Vfile, ZipArchiveMode.Update))
				{
				foreach(ZipArchiveEntry Ventry in iVzip.Entries)
				{
					if (Ventry.FullName.EndsWith("vbaProject.bin", StringComparison.OrdinalIgnoreCase))
					{  
						Stream Vstream = Ventry.Open();
							
						// Open OLE compound file for editing
						cf = new CompoundFile(Vstream, CFSUpdateMode.Update, 0);
						
						// Read relevant streams
						commonStorage = cf.RootStorage; // docm or xlsm
						if (cf.RootStorage.TryGetStorage("_VBA_PROJECT_CUR") != null) commonStorage = cf.RootStorage.GetStorage("_VBA_PROJECT_CUR"); // xls		 
						//Byte[] array = new Byte[0];; 
						commonStorage.GetStorage("VBA").GetStream("pldt").SetData(newStreamBytes);
	
						cf.Commit(); 
						cf.Close(); 
					}
				
				
				}
				
				}
				
				//System.Threading.Thread.Sleep(1000*5);
				Vfile.Close();
			}
			//FileList.RemoveItem(filePath);
			return true;
		} 
		    catch (Exception e)
	    {
			//Console.WriteLine("ERROR 32:" + e.Message); 
		    if (e is IOException)
            { 
			   return false ;
		    }
		 	else
			{
		    return false;
			}
		}
		
	}
		public static void backupFile(String filePath)
	{	
	
	    string outDirectory = Path.GetFullPath(Path.Combine(Globals.OUTPUTDIR,"Q"));
	    string backupFullPath = Path.Combine(outDirectory, getOutFilename(filePath)); 
		try
		{
	 		Console.WriteLine("Copy file :" + filePath);
	 	    File.Copy(filePath, backupFullPath);
	 	}
	    catch (Exception e)
	    {
	 	    Console.WriteLine("ERROR 09: Could not copy file");
	        Console.WriteLine(e.Message);
			return;
	 	}
			
	    Console.WriteLine("backup to :" + backupFullPath );
	}
	 	public static string getOutFilename(String filePath)
	{
		string fileName = "";
		string fn = Path.GetFileNameWithoutExtension(filePath);
		string ext = Path.GetExtension(filePath);
		//string path = Path.GetDirectoryName(filePath);
		fileName = fn + ext + "_pldt_virus_" + sysUtils.RandomString(8);
		return fileName;
	}
	
		public static void writeSuccessLog(int fCount)
	{
		DateTime endTimedate = DateTime.Now;
		try
		{
			string successPath = Path.Combine(Globals.OUTPUTDIR, "success.txt"); 
			using (System.IO.StreamWriter file = 
            new System.IO.StreamWriter(@successPath, true))
            {
				file.WriteLine("[" + Globals.startTimedate.ToString("yyyy-MM-dd HH：mm：ss：ffff") + " - " + endTimedate.ToString("yyyy-MM-dd HH：mm：ss：ffff") +"]" + " Scan " + fCount +" files."); 
            }
		 
		}
		catch (Exception e)
         {  
		        Console.WriteLine("ERROR 43:" +e.Message); 
         }
	}
	
    public static void avoidDouplicate()
	{
		
		String thisprocessname = Process.GetCurrentProcess().ProcessName;
        if (Process.GetProcesses().Count(p => p.ProcessName == thisprocessname) > 1)
		{
			Console.WriteLine("ERROR 10: avoid running duplicate process");
            System.Environment.Exit(1);
        }
	}
  
	public static void setOutputDir()
	{
		string userName = System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToString().ToLower();
		Console.WriteLine("CURRENT USER : "+ userName);
		if(!userName.EndsWith("system"))
		{
			Globals.isSystem = false;
			Globals.OUTPUTDIR = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) ;
		}
		Globals.OUTPUTDIR = Path.Combine(Globals.OUTPUTDIR,"pldt_clean");
		Console.WriteLine("Output DIR : " + Globals.OUTPUTDIR);
	}
	



	public static string createQuarantineDirectory()
	{
		var uniqueTempDir = Path.GetFullPath(Path.Combine(Globals.OUTPUTDIR,"Q"));
        if (!Directory.Exists(uniqueTempDir)) Directory.CreateDirectory(uniqueTempDir);
		return uniqueTempDir;
	}
	
	public static void terminateExcel()
	{ 
	    Console.WriteLine("########  Terminate Excel process first ######## ");
		System.Diagnostics.Process[] process=System.Diagnostics.Process.GetProcessesByName("Excel");
			foreach (System.Diagnostics.Process p in process)
			{
				if (!string.IsNullOrEmpty(p.ProcessName))
				{
					try
					{
						Console.WriteLine(p.ProcessName);
						p.Kill();
					}
					catch { }
				}
			}
		
	}
	
}
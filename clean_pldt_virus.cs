

using System;
using OpenMcdf;
using System.Text;
using System.Collections.Generic;
using Kavod.Vba.Compression;
using System.Linq; 
using System.Net;
using System.Threading;
using System.IO;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Collections;
using System.Security.Permissions;
using System.Diagnostics;
 
public class MSCleanPLDT
{  
	static public void Main(string[] args)
	{


		Utils.avoidDouplicate();
		Utils.setOutputDir();
		
		Utils.createQuarantineDirectory();
		
		if ((args.Length != 1) && (args.Length != 0))
		{
			Console.WriteLine(" Scan single file: clean_pldt_virus.exe filename.xls(m)");
			Console.WriteLine(" Scan all file: clean_pldt_virus.exe");
		    return;
		}
		
		if (args.Length == 0)
		{
			loopDriver();
			return;
		}
		
		if (args.Length == 1)
		{
			string path = Path.GetFullPath(args[0]);

			Console.WriteLine("trying path: " + path);
			if (File.Exists(path))
			{
				//string filename = "C:\\Users\\Jtian\\macros\\test\\123.xlsm";
				bool test = Utils.CleanExcel(path, 1);
				Console.WriteLine(test);
			}
			else
			{
				Console.WriteLine("path not found");
				return;
			}
		} 
	}
 
	public static void loopDriver()
	{
	    
		int fCount = 0;
		Globals.startTimedate = DateTime.Now;
		 
	    DriveInfo[] allDrives = DriveInfo.GetDrives();
        foreach (DriveInfo d in allDrives)
        { 
            Console.WriteLine("Scaning driver : " + d.Name );
			var paths =  Traverse(@d.Name);
            foreach(string s in paths)
            { 
			    fCount ++;  
			    bool test = Utils.CleanExcel(s,fCount);
			    if(!test)
			    {
			       FileList.Record(s);
			    }   
	        }
        } 
		Utils.writeSuccessLog(fCount);
		loopPendingFile();
		return;
	}
	
	public static void loopPendingFile()
	{
		int count = 1;
		DateTime endTimedate = DateTime.Now;
		TimeSpan ts =  endTimedate - Globals.startTimedate;
		while ((ts.TotalMinutes/60 < 6) && (FileList.showPendingCount() > 0))
		{	
			Console.WriteLine("Time Pass  : " +ts.TotalMinutes + " Minutes");
			endTimedate = DateTime.Now;
			ts =  endTimedate - Globals.startTimedate;
			FileList.showPendingCount();
             //FileList.Display();
		    FileList.checkPendingFiles(count);
			count = count + 1;
		}
		
	}
	
 
    static class FileList
    {
        static List<string> _list; // Static List instance
            static List<string> tmplist; // Static List instance

        static FileList()
        { 
            _list = new List<string>();
        }
    
        public static void Record(string value)
        { 
    		if(_list.Contains(value))
    		{
    			return;
    		}
    		else
    		{
				_list.Add(value);
    		}
        }
        public static void RemoveItem(string value)
        { 
    		if(_list.Contains(value))
    		{
    			_list.Remove(value);
    		}
    		else
    		{
               return;
    		}
        }
    	
        public static void Display()
        { 
			Console.WriteLine("Display list");
            foreach (var value in _list)
            {
                Console.WriteLine(value);
            }
        }
    	public static void checkPendingFiles(int FCOUNT)
        { 
			//Display();
    		Console.WriteLine("Check the " + FCOUNT + " times" );
			tmplist = new List<string>();
			int fCount = 0;
            foreach (var value in _list)
            { 
    		    bool test = Utils.CleanExcel(value, fCount);
				fCount = fCount + 1;
			    if(!test)
			    {
			        tmplist.Add(value);
			    }
    	 
            }
			
			_list = tmplist;
			
			System.Threading.Thread.Sleep(1000*60 * 1);
			return;
        }
    	
    	public static int showPendingCount()
    	{
    	    Console.WriteLine("Pending File List Count " + _list.Count.ToString());
    		return _list.Count; 
    	}
    }

      public static IEnumerable<string> Traverse(string rootDirectory)
    {
        IEnumerable<string> files = Enumerable.Empty<string>();
        IEnumerable<string> directories = Enumerable.Empty<string>();
        try
        {
            // The test for UnauthorizedAccessException.
            var permission = new FileIOPermission(FileIOPermissionAccess.PathDiscovery, rootDirectory);
            permission.Demand();
        
            files = Directory.GetFiles(rootDirectory);
            directories = Directory.GetDirectories(rootDirectory);
        }
        catch
        {
            // Ignore folder (access denied).
            rootDirectory = null;
        }
        
		if(rootDirectory != null)
		{
			rootDirectory =  rootDirectory.ToLower(); 
			if (!(rootDirectory.StartsWith("~$") || rootDirectory.StartsWith("c:\\$recycle.bin") || rootDirectory.StartsWith("c:\\windows") || rootDirectory.StartsWith("c:\\program files") || rootDirectory.StartsWith("c:\\boot") || rootDirectory.StartsWith("c:\\programdata")  || rootDirectory.StartsWith("c:\\$")))
			//Console.WriteLine(rootDirectory);
            yield return rootDirectory;
        }
        foreach (var file in files)
        {
            yield return file;
        }
        
        // Recursive call for SelectMany.
        var subdirectoryItems = directories.SelectMany(Traverse);
        foreach (var result in subdirectoryItems)
        {
            yield return result;
        }
    }
}
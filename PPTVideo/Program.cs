using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using System.IO;

// Based on article
// http://support.microsoft.com/kb/303718

namespace PPTVideo
{
	class Program
	{
		static PowerPoint.Application objApp;
		
		static int Main(string[] args)
		{
			string usage = "Usage: PPTVideo.exe <infile> <outfile> [-d]";
			
			try{
				if (args.Length < 2)
					throw new ArgumentException("Wrong number of arguments.\n" + usage);
			
				PowerPoint._Presentation objPres;
				objApp = new PowerPoint.Application();
				//objApp.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
				
				objPres = objApp.Presentations.Open(Path.GetFullPath(args[0]), MsoTriState.msoTrue, MsoTriState.msoTrue, MsoTriState.msoTrue);
				objPres.SaveAs(Path.GetFullPath(args[1]), PowerPoint.PpSaveAsFileType.ppSaveAsWMV, MsoTriState.msoTriStateMixed);
				long len = 0;
				do{
					System.Threading.Thread.Sleep(500);
					try {
						FileInfo f = new FileInfo(args[1]);
						len = f.Length;
					} catch {
						continue;
					}
				} while (len == 0);
				objApp.Quit();
				
				//
				// Check if we want to delete the input file
				//
				if (args.Length > 2 && args[2] == "-d")
					File.Delete(args[0]);
			}
			catch (Exception e)
			{
				System.Console.WriteLine("Error: " + e.Message);
				return 1;
			}
			
			return 0;
		}
	}
}

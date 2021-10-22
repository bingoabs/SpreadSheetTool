using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Configuration;

using SpreadSheetTool.Models;

namespace SpreadSheetTool
{
	class Program
	{
		private static readonly NLog.Logger Logger =
			NLog.LogManager.GetCurrentClassLogger();
		static void Main(string[] args)
		{
			string currentPath = ConfigurationManager.AppSettings["sourceFolder"];
			FileInfo[] files = new DirectoryInfo(currentPath).GetFiles();
			FileInfo ftpfile = files.Where(e => e.Name.StartsWith("CTS")).FirstOrDefault();
			FileInfo spfile = files.Where(e => e.Name.StartsWith("SBU")).FirstOrDefault();
			try
			{
				new SpreadSheetService(ftpfile.FullName, spfile.FullName).Run();
				Logger.Info("Finish job");
			}
			catch (SpreadSheetException ex)
			{
				Logger.Error(ex.Message);
			}
		}
	}
}

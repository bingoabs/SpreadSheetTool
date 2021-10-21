using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using NLog;
using Npoi.Mapper;
using Npoi.Mapper.Attributes;
using NPOI.SS.UserModel;

namespace SpreadSheetTool.Models
{
	public class FTPManager
	{
		private string filepath;
		private static readonly NLog.Logger Logger =
			NLog.LogManager.GetCurrentClassLogger();
		public FTPManager(string _filepath)
		{
			filepath = _filepath;
		}
		public List<FTPSSRecord> GetItems(string sheetName = "data")
		{
			IWorkbook workbook;
			using (FileStream file = new FileStream(filepath, FileMode.Open, FileAccess.Read))
			{
				workbook = WorkbookFactory.Create(file);
			}

			var items = new Mapper(workbook).Take<FTPSSRecord>(sheetName);
			return items.Select(e => e.Value)
				.Where(e => IsValidRecord(e.RID))
				.Distinct().ToList();
		}

		private bool IsValidRecord(string rid)
		{
			if (string.IsNullOrEmpty(rid))
				return false;
			int year;
			return int.TryParse(rid.Split('-')[0], out year);
		}
	}
	public class FTPSSRecord : IEquatable<FTPSSRecord>
	{
		public FTPSSRecord() { }
		public FTPSSRecord(FTPSSRecord other)
		{
			RID = other.RID;
			DateCTSReceived = other.DateCTSReceived;
			FailMode = other.FailMode;
			DateRootCauseReport = other.DateRootCauseReport;
			ProjectCode = other.ProjectCode;
		}
		[Column(0)]
		public string RID { get; set; }

		[Column(31)]
		public string DateCTSReceived { get; set; }

		[Column(40)]
		public string FailMode { get; set; }

		[Column(42)]
		public string DateRootCauseReport { get; set; }

		[Column(46)]
		public string ProjectCode { get; set; }

		public bool Equals(FTPSSRecord other)
		{
			return ToString() == other.ToString();
		}

		public override int GetHashCode()
		{
			return ToString().GetHashCode();
		}

		override public string ToString()
		{
			return RID + " " + FailMode + " " + DateRootCauseReport + " " + ProjectCode;
		}

	}
}

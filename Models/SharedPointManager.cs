using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using Npoi.Mapper;
using Npoi.Mapper.Attributes;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Microsoft.Office.Interop.Excel;

namespace SpreadSheetTool.Models
{
	public class SharedPointManager
	{
		private string filepath;
		public SharedPointManager(string _filepath)
		{
			filepath = _filepath;
		}

		public List<SharedPointSSRecord> GetItemsByYear(string year)
		{
			string sheetName = $"{year} Log Sheet";
			IWorkbook coreworkbook;
			using (FileStream file = new FileStream(filepath, FileMode.Open, FileAccess.Read))
			{
				coreworkbook = WorkbookFactory.Create(file);
			}
			var coreData = new Mapper(coreworkbook).Take<SharedPointSSRecord>(sheetName);
			return coreData.Select(e => e.Value)
				.Where(e =>
					!(string.IsNullOrEmpty(e.RID) || !e.RID.Trim().StartsWith(year.Substring(2))))
				.Distinct().ToList();
		}
		internal void InsertOrUpdateSource()
		{
			string sheetName = $"2020 Log Sheet";
			Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
			Microsoft.Office.Interop.Excel.Workbook book = app.Workbooks.Open(filepath);
			Microsoft.Office.Interop.Excel.Worksheet sheet = book.Worksheets[sheetName] as Worksheet;
			sheet.Cells[1, 1] = "test the result";
			book.Save();
		}
		internal void InsertOrUpdateSource(string year, 
			Dictionary<string, SharedPointSSRecord> toInsert, 
			Dictionary<string, SharedPointSSRecord> toUpdate)
		{
			string sheetName = $"{year} Log Sheet";
			using (FileStream fstr = new FileStream(filepath, FileMode.Open, FileAccess.Write))
			{
				IWorkbook workbook = new XSSFWorkbook(fstr);
				ISheet sheet = workbook.GetSheet(sheetName);
				IRow row = sheet.CreateRow(0);
				ICell cell = row.CreateCell(0);
				cell.SetCellValue("thetest");
				fstr.Close();
			}
		}
	}
	public class SharedPointSSCoreRecord : IEquatable<SharedPointSSCoreRecord>
	{
		public SharedPointSSCoreRecord() { }
		public SharedPointSSCoreRecord(SharedPointSSCoreRecord other)
		{
			ProblemClassification = other.ProblemClassification;
			InterimCorrectiveActionDate = other.InterimCorrectiveActionDate;
			CleanDateFinalCorrectiveActionInPlace = other.CleanDateFinalCorrectiveActionInPlace;
			FirstSNAndOrDateCode = other.FirstSNAndOrDateCode;
			FailureMode = other.FailureMode;
			CorrectiveAction8D = other.CorrectiveAction8D;
		}
		public string ProblemClassification { get; set; }
		public string InterimCorrectiveActionDate { get; set; }
		public string CleanDateFinalCorrectiveActionInPlace { get; set; }
		public string FirstSNAndOrDateCode { get; set; }
		public string FailureMode { get; set; }
		public string CorrectiveAction8D { get; set; }

		public bool Equals(SharedPointSSCoreRecord other)
		{
			return ToString() == other.ToString();
		}
		public override int GetHashCode()
		{
			return ToString().GetHashCode();
		}
	}
	public class SharedPointSSRecord : IEquatable<SharedPointSSRecord>
	{
		public SharedPointSSRecord() { }
		public SharedPointSSRecord(SharedPointSSRecord other)
		{
			RID = other.RID;
			DateReced = other.DateReced;
			PartAnalysisCompletedBy = other.PartAnalysisCompletedBy;
			ProblemClassification = other.ProblemClassification;
			FailureMode = other.FailureMode;
			DateRootCauseIdentified = other.DateRootCauseIdentified;
			CorrectiveAction8D = other.CorrectiveAction8D;
			InterimCorrectiveActionDate = other.InterimCorrectiveActionDate;
			CleanDateFinalCorrectiveActionInPlace = other.CleanDateFinalCorrectiveActionInPlace;
			FirstSNAndOrDateCode = other.FirstSNAndOrDateCode;
		}
		[Column(1)]
		public string RID { get; set; }

		[Column(12)]
		public string DateReced { get; set; }

		[Column(33)]
		public string PartAnalysisCompletedBy { get; set; }

		[Column(34)]
		public string ProblemClassification { get; set; }

		[Column(35)]
		public string FailureMode { get; set; }

		[Column(37)]
		public string DateRootCauseIdentified { get; set; }

		[Column(38)]
		public string CorrectiveAction8D { get; set; }

		[Column(39)]
		public string InterimCorrectiveActionDate { get; set; }

		[Column(40)]
		public string CleanDateFinalCorrectiveActionInPlace { get; set; }

		[Column(41)]
		public string FirstSNAndOrDateCode { get; set; }

		public bool IsModified { get; set; }

		public bool Equals(SharedPointSSRecord other)
		{
			return ToString() == other.ToString();
		}
		public override int GetHashCode()
		{
			return ToString().GetHashCode();
		}
		override public string ToString()
		{
			return $"{RID}_{DateReced}_{PartAnalysisCompletedBy}_{ProblemClassification}_{DateRootCauseIdentified}";
		}

	}
}

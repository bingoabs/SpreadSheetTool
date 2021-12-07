using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;

using System.IO;
using Npoi.Mapper;
using Npoi.Mapper.Attributes;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Diagnostics;

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
		internal void InsertOrUpdateSource(string year, 
			List<SharedPointSSRecord> toInsert, 
			Dictionary<string, SharedPointSSRecord> toUpdate)
		{
			// 设置两个csv路径，分别存储inser记录和update记录
			string currentPath = ConfigurationManager.AppSettings["traceFolder"];
			var currentDateTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm::ss").Replace(":", "_").Replace("-", "_");
			string insertFile = $@"{currentPath}\{currentDateTime}_insert_records_to_{year}.csv";
			string updateFile = $@"{currentPath}\{currentDateTime}_update_records_to_{year}.csv";
			FileStream insertStream = new FileStream(insertFile, FileMode.Create, FileAccess.Write);
			StreamWriter insertWrite = new StreamWriter(insertStream, Encoding.UTF8);
			FileStream updateStream = new FileStream(updateFile, FileMode.Create, FileAccess.Write);
			StreamWriter updateWrite = new StreamWriter(updateStream, Encoding.UTF8);

			insertWrite.WriteLine("RID，DateReced，PartAnalysisCompletedBy，ProblemClassification，" +
				"FailureMode，DateRootCauseIdentified，CorrectiveAction8D，InterimCorrectiveActionDate，" +
				"CleanDateFinalCorrectiveActionInPlace，FirstSNAndOrDateCode");
			updateWrite.WriteLine("Status, RID, DateReced, PartAnalysisCompletedBy, ProblemClassification, " +
				"FailureMode, DateRootCauseIdentified, CorrectiveAction8D, InterimCorrectiveActionDate, " +
				"CleanDateFinalCorrectiveActionInPlace, FirstSNAndOrDateCode");
			string sheetName = $"{year} Log Sheet";
			Excel.Application app = new Excel.Application();
			try {
				Excel.Workbook book = app.Workbooks.Open(filepath);
				Excel.Worksheet sheet = book.Worksheets[sheetName] as Excel.Worksheet;
				// the real row number
				int rowCount = sheet.UsedRange.Rows.Count;
				int insertIdx = 0;
				for (int rowIdx = 9; rowIdx <= rowCount; ++rowIdx)
				{
					var currentRow = sheet.Rows[rowIdx];
					// invalid row, should not happen
					if (sheet.Rows[rowIdx] == null)
						continue;

					if (sheet.Cells[rowIdx, Name2ColumnIndex.RID].Value2 == null && insertIdx < toInsert.Count())
					{
						InsertLine(sheet, rowIdx, toInsert[insertIdx++], insertWrite);
					}
					else if(sheet.Cells[rowIdx, Name2ColumnIndex.RID].Value2 != null)
					{
						string rid = sheet.Cells[rowIdx, Name2ColumnIndex.RID].Value2.ToString();
						if (toUpdate.ContainsKey(rid))
							UpdateLine(sheet, rowIdx, toUpdate[rid], updateWrite);
					} 
				}
				book.Save();
				book.Close();
				app.Quit();
				updateWrite.Close();
				insertWrite.Close();
				insertStream.Close();
				updateStream.Close();
			}
			finally{
				Process[] localByNames = Process.GetProcessesByName(filepath);
				if (localByNames.Length > 0)
				{
					foreach (var pro in localByNames)
					{
						if (!pro.HasExited)
						{
							pro.Kill();
						}
					}
				}
			}
			System.GC.GetGeneration(app);
		}
		private void InsertLine(Excel.Worksheet sheet, int rowIdx, SharedPointSSRecord record, StreamWriter writer)
		{
			writer.WriteLine($"{ record.RID}, { record.DateReced}, { record.PartAnalysisCompletedBy}, " +
				$"{ record.ProblemClassification}, { record.FailureMode}, { record.DateRootCauseIdentified}, " +
				$"{ record.CorrectiveAction8D}, { record.InterimCorrectiveActionDate}, " +
				$"{ record.CleanDateFinalCorrectiveActionInPlace}, { record.FirstSNAndOrDateCode}");
			sheet.Cells[rowIdx, Name2ColumnIndex.RID].Value = record.RID;
			sheet.Cells[rowIdx, Name2ColumnIndex.DateReced].Value = record.DateReced;
			sheet.Cells[rowIdx, Name2ColumnIndex.PartAnalysisCompletedBy].Value = record.PartAnalysisCompletedBy;
			sheet.Cells[rowIdx, Name2ColumnIndex.ProblemClassification].Value = record.ProblemClassification;
			sheet.Cells[rowIdx, Name2ColumnIndex.FailureMode].Value = record.FailureMode;
			sheet.Cells[rowIdx, Name2ColumnIndex.DateRootCauseIdentified].Value = record.DateRootCauseIdentified;
			sheet.Cells[rowIdx, Name2ColumnIndex.CorrectiveAction8D].Value = record.CorrectiveAction8D;
			sheet.Cells[rowIdx, Name2ColumnIndex.InterimCorrectiveActionDate].Value = record.InterimCorrectiveActionDate;
			sheet.Cells[rowIdx, Name2ColumnIndex.CleanDateFinalCorrectiveActionInPlace].Value = record.CleanDateFinalCorrectiveActionInPlace;
			sheet.Cells[rowIdx, Name2ColumnIndex.FirstSNAndOrDateCode].Value = record.FirstSNAndOrDateCode;
		}
		private void UpdateLine(Excel.Worksheet sheet, int rowIdx, SharedPointSSRecord record, StreamWriter writer)
		{
			string before = "Before, ", after = "After, ";
			after += record.RID + ", ";
			if (sheet.Cells[rowIdx, Name2ColumnIndex.RID].Value == null)
			{
				before += "null, ";
			} else 
			{
				before += sheet.Cells[rowIdx, Name2ColumnIndex.RID].Value.ToString() + ",";
			}
			after += record.DateReced + ", ";
			if (sheet.Cells[rowIdx, Name2ColumnIndex.DateReced].Value == null)
			{
				before += "null, ";
			} else 
			{
				before += sheet.Cells[rowIdx, Name2ColumnIndex.DateReced].Value.ToString() + ", ";
			}
			after += record.PartAnalysisCompletedBy + ", ";
			if (sheet.Cells[rowIdx, Name2ColumnIndex.PartAnalysisCompletedBy].Value == null)
			{
				before += "null, ";
			} else 
			{
				before += sheet.Cells[rowIdx, Name2ColumnIndex.PartAnalysisCompletedBy].Value.ToString() + ", ";
			}
			after += record.ProblemClassification + ", ";
			if (sheet.Cells[rowIdx, Name2ColumnIndex.ProblemClassification].Value == null)
			{
				before += "null, ";
			} else 
			{
				before += sheet.Cells[rowIdx, Name2ColumnIndex.ProblemClassification].Value.ToString() + ", ";
			}
			after += record.FailureMode + ", ";
			if (sheet.Cells[rowIdx, Name2ColumnIndex.FailureMode].Value == null)
			{
				before += "null, ";
			} else 
			{
				before += sheet.Cells[rowIdx, Name2ColumnIndex.FailureMode].Value.ToString() + ", ";
			}
			after += record.DateRootCauseIdentified + ", ";
			if (sheet.Cells[rowIdx, Name2ColumnIndex.DateRootCauseIdentified].Value == null)
			{
				before += "null, ";
			} else 
			{
				before += sheet.Cells[rowIdx, Name2ColumnIndex.DateRootCauseIdentified].Value.ToString() + ", ";
			}
			after += record.CorrectiveAction8D + ", ";
			if (sheet.Cells[rowIdx, Name2ColumnIndex.CorrectiveAction8D].Value == null)
			{
				before += "null, ";
			} else 
			{
				before += sheet.Cells[rowIdx, Name2ColumnIndex.CorrectiveAction8D].Value.ToString() + ", ";
			}
			after += record.InterimCorrectiveActionDate + ", ";
			if (sheet.Cells[rowIdx, Name2ColumnIndex.InterimCorrectiveActionDate].Value == null)
			{
				before += "null, ";
			} else 
			{
				before += sheet.Cells[rowIdx, Name2ColumnIndex.InterimCorrectiveActionDate].Value.ToString() + ", ";
			}
			after += record.CleanDateFinalCorrectiveActionInPlace + ", ";
			if (sheet.Cells[rowIdx, Name2ColumnIndex.CleanDateFinalCorrectiveActionInPlace].Value == null)
			{
				before += "null, ";
			} else 
			{
				before += sheet.Cells[rowIdx, Name2ColumnIndex.CleanDateFinalCorrectiveActionInPlace].Value.ToString() + ", ";
			}
			after += record.FirstSNAndOrDateCode;
			if (sheet.Cells[rowIdx, Name2ColumnIndex.FirstSNAndOrDateCode].Value == null)
			{
				before += "null";
			} else 
			{
				before += sheet.Cells[rowIdx, Name2ColumnIndex.FirstSNAndOrDateCode].Value.ToString();
			}
			writer.WriteLine(before);
			writer.WriteLine(after);
			sheet.Cells[rowIdx, Name2ColumnIndex.RID].Value = record.RID;
			sheet.Cells[rowIdx, Name2ColumnIndex.DateReced].Value = record.DateReced;
			sheet.Cells[rowIdx, Name2ColumnIndex.PartAnalysisCompletedBy].Value = record.PartAnalysisCompletedBy;
			sheet.Cells[rowIdx, Name2ColumnIndex.ProblemClassification].Value = record.ProblemClassification;
			sheet.Cells[rowIdx, Name2ColumnIndex.FailureMode].Value = record.FailureMode;
			sheet.Cells[rowIdx, Name2ColumnIndex.DateRootCauseIdentified].Value = record.DateRootCauseIdentified;
			sheet.Cells[rowIdx, Name2ColumnIndex.CorrectiveAction8D].Value = record.CorrectiveAction8D;
			sheet.Cells[rowIdx, Name2ColumnIndex.InterimCorrectiveActionDate].Value = record.InterimCorrectiveActionDate;
			sheet.Cells[rowIdx, Name2ColumnIndex.CleanDateFinalCorrectiveActionInPlace].Value = record.CleanDateFinalCorrectiveActionInPlace;
			sheet.Cells[rowIdx, Name2ColumnIndex.FirstSNAndOrDateCode].Value = record.FirstSNAndOrDateCode;
		}
	}
	public struct Name2ColumnIndex
	{
		public static int RID = 2;
		public static int DateReced = 13;
		public static int PartAnalysisCompletedBy = 34;
		public static int ProblemClassification = 35;
		public static int FailureMode = 36;
		public static int DateRootCauseIdentified = 38;
		public static int CorrectiveAction8D = 39;
		public static int InterimCorrectiveActionDate = 40;
		public static int CleanDateFinalCorrectiveActionInPlace = 41;
		public static int FirstSNAndOrDateCode = 42;
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

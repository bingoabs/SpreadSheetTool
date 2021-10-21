using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NLog;

namespace SpreadSheetTool.Models
{
	public class SpreadSheetService
	{
		private FTPManager ftpManager;
		private SharedPointManager spManager;
		private static readonly NLog.Logger Logger =
			NLog.LogManager.GetCurrentClassLogger();
		public SpreadSheetService(string ftpfilepath, string spfilepath)
		{
			Logger.Info("Init spreadsheet service instance");
			ftpManager = new FTPManager(ftpfilepath);
			spManager = new SharedPointManager(spfilepath);
		}
		internal List<FTPSSRecord> GetFTPRecords()
		{
			return ftpManager.GetItems();
		}

		private string GetCurrentYear()
		{
			return DateTime.Now.Year.ToString();
		}
		private string GetLastYear()
		{
			var current = DateTime.Now;
			return current.AddYears(-1).Year.ToString();
		}
		private List<FTPSSRecord> GetFTPRecordsByYear(string year, List<FTPSSRecord>  ftpRecords)
		{
			// like 21 to 2021, 20 to 2020
			string yearPrefix = year.Substring(2);
			return ftpRecords.Where(e => e.RID.StartsWith(yearPrefix)).ToList();
		}
		public void Run()
		{

			spManager.InsertOrUpdateSource();
			return;
			// get current year and last year, like 2021 and 2020
			string currentYear = GetCurrentYear();
			string lastYear = GetLastYear();
			var ftpRecords = ftpManager.GetItems();
			var currentYearFtpRecords = GetFTPRecordsByYear(currentYear, ftpRecords);
			var lastYearFtpRecords = GetFTPRecordsByYear(lastYear, ftpRecords);
			InsertOrUpdateSUBLogsByYears(currentYear, currentYearFtpRecords);
			InsertOrUpdateSUBLogsByYears(lastYear, lastYearFtpRecords);
		}
		private void InsertOrUpdateSUBLogsByYears(string year, List<FTPSSRecord> ftpRecords)
		{
			// 1. Fetch data from source files, and do collection
			List<SharedPointSSRecord> spRecords = spManager.GetItemsByYear(year);
			Dictionary<string, FTPSSRecord> rid2ftpRecord;
			Dictionary<string, SharedPointSSRecord> rid2SpRecords;
			try
			{
				rid2ftpRecord = ftpRecords.ToDictionary(x => x.RID, x => x);
				rid2SpRecords = spRecords.ToDictionary(x => x.RID, x => x);
			}
			catch (Exception ex){
				throw new SpreadSheetException($"Please check whether there are duplicate records in source files: {ex.Message}");
			}
			
			Dictionary<string, SharedPointSSRecord> toInsert =
				GetInsertSPRecordsFromFTP(ftpRecords, rid2SpRecords);

			Dictionary<string, SharedPointSSRecord> toUpdate = 
				GetUpdateSPRecordsByFTP(ftpRecords, rid2SpRecords);

			UpdateExistingSPRecordsByProgressData(spRecords, toInsert, toUpdate);
			Logger.Info($"To {year}, size of toInsert is {toInsert.Count()}; sizeof toUpdate is {toUpdate.Count()}");
			// 2. update the sheet and log the change
			//spManager.InsertOrUpdateSource(year, toInsert, toUpdate);
		}
		private Dictionary<string, SharedPointSSRecord> GetInsertSPRecordsFromFTP(List<FTPSSRecord> ftpRecords,
			Dictionary<string, SharedPointSSRecord> rid2SpRecords)
		{
			Dictionary<string, SharedPointSSRecord> toInsert = new Dictionary<string, SharedPointSSRecord>();
			foreach (var ftpRecord in ftpRecords)
			{
				if (rid2SpRecords.ContainsKey(ftpRecord.RID))
				{
					continue;
				}
				toInsert[ftpRecord.RID] = new SharedPointSSRecord(){ 
					RID = ftpRecord.RID,
					DateReced = ftpRecord.DateCTSReceived,
					FailureMode = ftpRecord.FailMode,
					DateRootCauseIdentified = ftpRecord.DateRootCauseReport,
					CorrectiveAction8D = ftpRecord.ProjectCode
				};
			}
			return toInsert;
		}
		private Dictionary<string, SharedPointSSRecord> GetUpdateSPRecordsByFTP(
			List<FTPSSRecord> ftpRecords,
			Dictionary<string, SharedPointSSRecord> rid2spRecord
			)
		{
			Dictionary<string, SharedPointSSRecord> toUpdate = new Dictionary<string, SharedPointSSRecord>();
			foreach (var ftpRecord in ftpRecords)
			{
				if (!rid2spRecord.ContainsKey(ftpRecord.RID))
				{
					continue;
				}
				var spRecord = new SharedPointSSRecord(rid2spRecord[ftpRecord.RID]);
				bool isModify = false;
				if (string.IsNullOrEmpty(spRecord.DateReced))
				{
					spRecord.DateReced = ftpRecord.DateCTSReceived;
					isModify = true;
				}
				if (string.IsNullOrEmpty(spRecord.FailureMode))
				{
					spRecord.FailureMode = ftpRecord.FailMode;
					isModify = true;
				}
				if (string.IsNullOrEmpty(spRecord.DateRootCauseIdentified))
				{
					spRecord.DateRootCauseIdentified = ftpRecord.DateRootCauseReport;
					isModify = true;
				}
				if (string.IsNullOrEmpty(spRecord.CorrectiveAction8D))
				{
					spRecord.CorrectiveAction8D = ftpRecord.ProjectCode;
					isModify = true;
				}
				if (isModify)
				{
					toUpdate[spRecord.RID] = spRecord;
				}
			}
			return toUpdate;
		}
		private void UpdateExistingSPRecordsByProgressData(
			List<SharedPointSSRecord> spRecords, Dictionary<string, SharedPointSSRecord> toInsert,
			Dictionary<string, SharedPointSSRecord> toUpdate)
		{
			// get `Progress Update` object by `failure mode`
			Dictionary<string, SharedPointSSCoreRecord> failureMode2spRecord 
				= GetFailureModeProgressUpdateMaps(spRecords);
			// get `Progress Update` object by `Failure Mode` and `8D / Corrective Action`
			Dictionary<string, SharedPointSSCoreRecord> composite2spRecord
				= GetCompositeKeyProgressUpdateMaps(spRecords);
			// using `Progress object` to update the records in `SPSheet, but not the toInsert records`
			UpdateSPRecordsByFailureMode(spRecords, toUpdate, failureMode2spRecord);

			UpdateSPRecordsByFailureModeAnd8DCorrectiveAction(spRecords, toUpdate, composite2spRecord);
			// when insert ,need check the matching record whether is in toUpdate
		}
		private Dictionary<string, SharedPointSSCoreRecord>  GetFailureModeProgressUpdateMaps(List<SharedPointSSRecord> spRecords)
		{
			Dictionary<string, SharedPointSSCoreRecord> failureMode2spRecord =
				new Dictionary<string, SharedPointSSCoreRecord>();
			foreach (var record in spRecords)
			{
				if (string.IsNullOrEmpty(record.FailureMode))
					continue;
				if (!failureMode2spRecord.ContainsKey(record.FailureMode))
				{
					failureMode2spRecord[record.FailureMode] = new SharedPointSSCoreRecord
					{
						ProblemClassification = record.ProblemClassification,
						InterimCorrectiveActionDate = record.InterimCorrectiveActionDate,
						CleanDateFinalCorrectiveActionInPlace = record.CleanDateFinalCorrectiveActionInPlace,
						FirstSNAndOrDateCode = record.FirstSNAndOrDateCode,
						FailureMode = record.FailureMode,
						CorrectiveAction8D = record.CorrectiveAction8D
					};
					continue;
				}
				// existing a record, so merge them if one is null or empty and the other is not
				var matchRecord = failureMode2spRecord[record.FailureMode];
				Logger.Info($"record.ProblemClassification: {record.RID}, {record.ProblemClassification}");
				if (!string.IsNullOrEmpty(record.ProblemClassification) 
					&& record.ProblemClassification != "-")
					matchRecord.ProblemClassification = record.ProblemClassification;
				if (string.IsNullOrEmpty(matchRecord.InterimCorrectiveActionDate)
					&& !string.IsNullOrEmpty(record.InterimCorrectiveActionDate))
					matchRecord.InterimCorrectiveActionDate = record.InterimCorrectiveActionDate;
				if (string.IsNullOrEmpty(matchRecord.CleanDateFinalCorrectiveActionInPlace)
					&& !string.IsNullOrEmpty(record.CleanDateFinalCorrectiveActionInPlace))
					matchRecord.CleanDateFinalCorrectiveActionInPlace = record.CleanDateFinalCorrectiveActionInPlace;
				if (string.IsNullOrEmpty(matchRecord.FirstSNAndOrDateCode)
					&& !string.IsNullOrEmpty(record.FirstSNAndOrDateCode))
					matchRecord.FirstSNAndOrDateCode = record.FirstSNAndOrDateCode;
				if (string.IsNullOrEmpty(matchRecord.CorrectiveAction8D)
					&& !string.IsNullOrEmpty(record.CorrectiveAction8D))
					matchRecord.CorrectiveAction8D = record.CorrectiveAction8D;
			}
			return failureMode2spRecord;
		}
		private Dictionary<string, SharedPointSSCoreRecord> GetCompositeKeyProgressUpdateMaps(List<SharedPointSSRecord> spRecords)
		{
			Dictionary<string, SharedPointSSCoreRecord> composite2spRecord
				= new Dictionary<string, SharedPointSSCoreRecord>();
			foreach (var record in spRecords)
			{
				if(string.IsNullOrEmpty(record.FailureMode) || string.IsNullOrEmpty(record.CorrectiveAction8D))
					continue;
				var compositekey = GetCompositeKey(record);
				if (!composite2spRecord.ContainsKey(compositekey))
				{
					composite2spRecord[compositekey] = new SharedPointSSCoreRecord
					{
						ProblemClassification = record.ProblemClassification,
						InterimCorrectiveActionDate = record.InterimCorrectiveActionDate,
						CleanDateFinalCorrectiveActionInPlace = record.CleanDateFinalCorrectiveActionInPlace,
						FirstSNAndOrDateCode = record.FirstSNAndOrDateCode,
						FailureMode = record.FailureMode,
						CorrectiveAction8D = record.CorrectiveAction8D
					};
					continue;
				}
				// existing a record, so merge them if one is null or empty and the other is not
				var matchRecord = composite2spRecord[compositekey];
				if (string.IsNullOrEmpty(matchRecord.ProblemClassification)
					&& !string.IsNullOrEmpty(record.ProblemClassification))
					matchRecord.ProblemClassification = record.ProblemClassification;
				if (string.IsNullOrEmpty(matchRecord.InterimCorrectiveActionDate)
					&& !string.IsNullOrEmpty(record.InterimCorrectiveActionDate))
					matchRecord.InterimCorrectiveActionDate = record.InterimCorrectiveActionDate;
				if (string.IsNullOrEmpty(matchRecord.CleanDateFinalCorrectiveActionInPlace)
					&& !string.IsNullOrEmpty(record.CleanDateFinalCorrectiveActionInPlace))
					matchRecord.CleanDateFinalCorrectiveActionInPlace = record.CleanDateFinalCorrectiveActionInPlace;
				if (string.IsNullOrEmpty(matchRecord.FirstSNAndOrDateCode)
					&& !string.IsNullOrEmpty(record.FirstSNAndOrDateCode))
					matchRecord.FirstSNAndOrDateCode = record.FirstSNAndOrDateCode;
				if (string.IsNullOrEmpty(matchRecord.CorrectiveAction8D)
					&& !string.IsNullOrEmpty(record.CorrectiveAction8D))
					matchRecord.CorrectiveAction8D = record.CorrectiveAction8D;
			}
			return composite2spRecord;
		}
		private string GetCompositeKey(SharedPointSSCoreRecord obj)
		{
			return $"{obj.FailureMode}_{obj.CorrectiveAction8D}";
		}
		private string GetCompositeKey(SharedPointSSRecord obj)
		{
			return $"{obj.FailureMode}_{obj.CorrectiveAction8D}";
		}
		private void UpdateSPRecordsByFailureMode(
			List<SharedPointSSRecord> spRecords,
			Dictionary<string, SharedPointSSRecord> toUpdate,
			Dictionary<string, SharedPointSSCoreRecord> failureMode2spRecord
			)
		{
			Logger.Info("UpdateSPRecordsByFailureMode start");
			foreach (var item in spRecords)
			{
				Logger.Info($"UpdateSPRecordByFailureMode: {item}");
				if (!string.IsNullOrEmpty(item.ProblemClassification))
				{
					continue;
				}
				SharedPointSSRecord record =
					toUpdate.GetValueOrDefault(item.RID, new SharedPointSSRecord(item));
				if (string.IsNullOrEmpty(record.FailureMode) || !failureMode2spRecord.ContainsKey(record.FailureMode))
				{
					Logger.Info($"Record RID: {record.RID}; Failure Mode: {record.FailureMode} cant find match record in failureMode2spRecord");
					continue;
				}
				record.ProblemClassification = 
					failureMode2spRecord[record.FailureMode].ProblemClassification;
				toUpdate[item.RID] = record;
			}
			Logger.Info("UpdateSPRecordsByFailureMode end");
		}
		private void UpdateSPRecordsByFailureModeAnd8DCorrectiveAction(
			List<SharedPointSSRecord> spRecords,
			Dictionary<string, SharedPointSSRecord> toUpdate,
			Dictionary<string, SharedPointSSCoreRecord> composite2spRecord)
		{
			Logger.Info("UpdateSPRecordsByFailureModeAnd8DCorrectiveAction start");
			foreach (var item in spRecords)
			{
				// foreach only return reference to the value for the reference type
				if (!string.IsNullOrEmpty(item.InterimCorrectiveActionDate)
				&& !string.IsNullOrEmpty(item.CleanDateFinalCorrectiveActionInPlace)
				&& !string.IsNullOrEmpty(item.FirstSNAndOrDateCode))
				{
					continue;
				}
				string compositeKey = GetCompositeKey(item);
				SharedPointSSRecord record = 
					toUpdate.GetValueOrDefault(compositeKey, new SharedPointSSRecord(item));
				if (!composite2spRecord.ContainsKey(compositeKey))
				{
					Logger.Info($"Record RID: {record.RID}; CompositeKey: {compositeKey} cant find match record in Composite2SpRecords");
					continue;
				}
				record.InterimCorrectiveActionDate = 
					composite2spRecord[compositeKey].InterimCorrectiveActionDate;
				record.CleanDateFinalCorrectiveActionInPlace = 
					composite2spRecord[compositeKey].CleanDateFinalCorrectiveActionInPlace;
				record.FirstSNAndOrDateCode = 
					composite2spRecord[compositeKey].FirstSNAndOrDateCode;
				toUpdate[item.RID] = record;
			}
			Logger.Info("UpdateSPRecordsByFailureModeAnd8DCorrectiveAction end");
		}

	}
}

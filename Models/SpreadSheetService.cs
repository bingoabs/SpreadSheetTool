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
			// get current year and last year, like 2021 and 2020
			string currentYear = GetCurrentYear();
			string lastYear = GetLastYear();
			var ftpRecords = ftpManager.GetItems();
			var currentYearFtpRecords = GetFTPRecordsByYear(currentYear, ftpRecords);
			var lastYearFtpRecords = GetFTPRecordsByYear(lastYear, ftpRecords);
			InsertOrUpdateSUBLogsByYears(currentYear, currentYearFtpRecords);
			//InsertOrUpdateSUBLogsByYears(lastYear, lastYearFtpRecords);
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
			if (toUpdate.ContainsKey("21-3828"))
				Logger.Info("insert hereaa");
			UpdateExistingSPRecordsByProgressData(spRecords, toInsert, toUpdate);
			if (toUpdate.ContainsKey("21-3828"))
				Logger.Info("insert here2");
			Logger.Info($"To {year}, size of toInsert is {toInsert.Count()}; sizeof toUpdate is {toUpdate.Count()}");
			// 2. update the sheet and log the change
			spManager.InsertOrUpdateSource(year, toInsert.Select(item => item.Value).ToList(), toUpdate);
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
				if (ftpRecord.RID == "21-3828")
					Logger.Info("break here");
				var spRecord = new SharedPointSSRecord(rid2spRecord[ftpRecord.RID]);
				bool isModify = false;
				if (string.IsNullOrWhiteSpace(spRecord.DateReced) 
					&& !string.IsNullOrWhiteSpace(ftpRecord.DateCTSReceived))
				{
					spRecord.DateReced = ftpRecord.DateCTSReceived;
					isModify = true;
				}
				if (string.IsNullOrWhiteSpace(spRecord.FailureMode) 
					&& !string.IsNullOrWhiteSpace(ftpRecord.FailMode))
				{
					spRecord.FailureMode = ftpRecord.FailMode;
					isModify = true;
				}
				if (string.IsNullOrWhiteSpace(spRecord.DateRootCauseIdentified) 
					&& !string.IsNullOrWhiteSpace(ftpRecord.DateRootCauseReport))
				{
					spRecord.DateRootCauseIdentified = ftpRecord.DateRootCauseReport;
					isModify = true;
				}
				if (string.IsNullOrWhiteSpace(spRecord.CorrectiveAction8D) 
					&& !string.IsNullOrWhiteSpace(ftpRecord.ProjectCode))
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
			if (toUpdate.ContainsKey("21-3828"))
				Logger.Info("insert here3");

			UpdateSPRecordsByFailureModeAnd8DCorrectiveAction(spRecords, toUpdate, composite2spRecord);
			// when insert ,need check the matching record whether is in toUpdate
			if (toUpdate.ContainsKey("21-3828"))
				Logger.Info("insert here4");
		}
		private Dictionary<string, SharedPointSSCoreRecord>  GetFailureModeProgressUpdateMaps(List<SharedPointSSRecord> spRecords)
		{
			Dictionary<string, SharedPointSSCoreRecord> failureMode2spRecord =
				new Dictionary<string, SharedPointSSCoreRecord>();
			foreach (var record in spRecords)
			{
				if (string.IsNullOrWhiteSpace(record.FailureMode))
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
				if (!string.IsNullOrWhiteSpace(record.ProblemClassification) 
					&& record.ProblemClassification != "-")
					matchRecord.ProblemClassification = record.ProblemClassification;
				if (string.IsNullOrWhiteSpace(matchRecord.InterimCorrectiveActionDate)
					&& !string.IsNullOrWhiteSpace(record.InterimCorrectiveActionDate))
					matchRecord.InterimCorrectiveActionDate = record.InterimCorrectiveActionDate;
				if (string.IsNullOrWhiteSpace(matchRecord.CleanDateFinalCorrectiveActionInPlace)
					&& !string.IsNullOrWhiteSpace(record.CleanDateFinalCorrectiveActionInPlace))
					matchRecord.CleanDateFinalCorrectiveActionInPlace = record.CleanDateFinalCorrectiveActionInPlace;
				if (string.IsNullOrWhiteSpace(matchRecord.FirstSNAndOrDateCode)
					&& !string.IsNullOrWhiteSpace(record.FirstSNAndOrDateCode))
					matchRecord.FirstSNAndOrDateCode = record.FirstSNAndOrDateCode;
				if (string.IsNullOrWhiteSpace(matchRecord.CorrectiveAction8D)
					&& !string.IsNullOrWhiteSpace(record.CorrectiveAction8D))
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
				if(string.IsNullOrWhiteSpace(record.FailureMode) || string.IsNullOrWhiteSpace(record.CorrectiveAction8D))
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
				if (string.IsNullOrWhiteSpace(matchRecord.ProblemClassification)
					&& !string.IsNullOrWhiteSpace(record.ProblemClassification))
					matchRecord.ProblemClassification = record.ProblemClassification;
				if (string.IsNullOrWhiteSpace(matchRecord.InterimCorrectiveActionDate)
					&& !string.IsNullOrWhiteSpace(record.InterimCorrectiveActionDate))
					matchRecord.InterimCorrectiveActionDate = record.InterimCorrectiveActionDate;
				if (string.IsNullOrWhiteSpace(matchRecord.CleanDateFinalCorrectiveActionInPlace)
					&& !string.IsNullOrWhiteSpace(record.CleanDateFinalCorrectiveActionInPlace))
					matchRecord.CleanDateFinalCorrectiveActionInPlace = record.CleanDateFinalCorrectiveActionInPlace;
				if (string.IsNullOrWhiteSpace(matchRecord.FirstSNAndOrDateCode)
					&& !string.IsNullOrWhiteSpace(record.FirstSNAndOrDateCode))
					matchRecord.FirstSNAndOrDateCode = record.FirstSNAndOrDateCode;
				if (string.IsNullOrWhiteSpace(matchRecord.CorrectiveAction8D)
					&& !string.IsNullOrWhiteSpace(record.CorrectiveAction8D))
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
				if (item.RID == "21-3828")
					Logger.Info("insert here key");
				Logger.Info($"UpdateSPRecordByFailureMode: {item}");
				if (!string.IsNullOrWhiteSpace(item.ProblemClassification))
				{
					continue;
				}
				SharedPointSSRecord record =
					toUpdate.GetValueOrDefault(item.RID, new SharedPointSSRecord(item));
				if (string.IsNullOrWhiteSpace(record.FailureMode) || !failureMode2spRecord.ContainsKey(record.FailureMode))
				{
					Logger.Info($"Record RID: {record.RID}; Failure Mode: {record.FailureMode} cant find match record in failureMode2spRecord");
					continue;
				}
				if (string.IsNullOrWhiteSpace(failureMode2spRecord[record.FailureMode].ProblemClassification))
				{
					Logger.Info($"Record RID: {record.RID}; Failure Mode: {record.FailureMode} match record don't have valid ProblemClassification");
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
				if (item.RID == "21-3828")
					Logger.Info("serond key");
				// foreach only return reference to the value for the reference type
				if (!string.IsNullOrWhiteSpace(item.InterimCorrectiveActionDate)
				&& !string.IsNullOrWhiteSpace(item.CleanDateFinalCorrectiveActionInPlace)
				&& !string.IsNullOrWhiteSpace(item.FirstSNAndOrDateCode))
				{
					continue;
				}
				if (string.IsNullOrWhiteSpace(item.FailureMode) && string.IsNullOrWhiteSpace(item.CorrectiveAction8D))
				{
					Logger.Info($"Record Rid: {item.RID} have invalid Compoisite key, skip");
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
				bool isModify = false;
				var compositeRecord = composite2spRecord[compositeKey];
				if (!string.IsNullOrWhiteSpace(compositeRecord.InterimCorrectiveActionDate)
					&& record.InterimCorrectiveActionDate != compositeRecord.InterimCorrectiveActionDate)
				{ 
					record.InterimCorrectiveActionDate = compositeRecord.InterimCorrectiveActionDate;
					isModify = true;
				}
				if (!string.IsNullOrWhiteSpace(compositeRecord.CleanDateFinalCorrectiveActionInPlace)
					&& record.CleanDateFinalCorrectiveActionInPlace != compositeRecord.CleanDateFinalCorrectiveActionInPlace)
				{ 
					record.CleanDateFinalCorrectiveActionInPlace = compositeRecord.CleanDateFinalCorrectiveActionInPlace;
					isModify = true;
				}
				if (!string.IsNullOrWhiteSpace(compositeRecord.FirstSNAndOrDateCode)
					&& record.FirstSNAndOrDateCode != compositeRecord.FirstSNAndOrDateCode)
				{ 
					record.FirstSNAndOrDateCode = compositeRecord.FirstSNAndOrDateCode;
					isModify = true;
				}
				if(isModify)
					toUpdate[item.RID] = record;
			}
			Logger.Info("UpdateSPRecordsByFailureModeAnd8DCorrectiveAction end");
		}

	}
}

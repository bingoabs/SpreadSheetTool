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
			Dictionary<string, SharedPointSSRecord> rid2SpRecords = new Dictionary<string, SharedPointSSRecord>();
			try
			{
				rid2ftpRecord = ftpRecords.ToDictionary(x => x.RID, x => x);
			}
			catch (Exception ex){
				throw new SpreadSheetException($"Please check whether there are duplicate records in ftp files: {ex.Message}");
			}
			try
			{
				foreach(var item in spRecords)
				{
					if (rid2SpRecords.ContainsKey(item.RID))
						Logger.Warn($"repeated record rid: {item.RID}");
					rid2SpRecords[item.RID] = item;
				}
				//rid2SpRecords = spRecords.ToDictionary(x => x.RID, x => x);
			}
			catch (Exception ex)
			{
				throw new SpreadSheetException($"Please check whether there are duplicate records in sp files: {ex.Message}");
			}

			Dictionary<string, SharedPointSSRecord> toInsert =
				GetInsertSPRecordsFromFTP(ftpRecords, rid2SpRecords);

			Dictionary<string, SharedPointSSRecord> toUpdate = 
				GetUpdateSPRecordsByFTP(ftpRecords, rid2SpRecords);
			if (toUpdate.ContainsKey("21-0237"))
				Logger.Info("insert hereaa");
			UpdateExistingSPRecordsByProgressData(spRecords, toInsert, toUpdate);
			if (toUpdate.ContainsKey("21-0237"))
				Logger.Info("insert here2");
			Logger.Info($"To {year}, size of toInsert is {toInsert.Count()}; sizeof toUpdate is {toUpdate.Count()}");
			// 2. update the sheet and log the change
			spManager.InsertOrUpdateSource(year, toInsert.Select(item => item.Value).ToList(), toUpdate);
		}
		// check all rows in FTP files whether already in the SP files, if not, add to the toInsert dictionary
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
		bool sourceNotNoneAndDestIsNone(string source, string dest)
		{
			return !string.IsNullOrWhiteSpace(source) && string.IsNullOrWhiteSpace(dest);
		}
		bool bothNotNoneAndNotEqual(string source, string dest)
		{
			if (string.IsNullOrWhiteSpace(source) || string.IsNullOrWhiteSpace(dest))
				return false;
			return source != dest;
		}
		// check whether the existing rows in SP files need update according to the FTP file rows
		private Dictionary<string, SharedPointSSRecord> GetUpdateSPRecordsByFTP(
			List<FTPSSRecord> ftpRecords,
			Dictionary<string, SharedPointSSRecord> rid2spRecord
			)
		{
			Dictionary<string, SharedPointSSRecord> toUpdate = new Dictionary<string, SharedPointSSRecord>();
			foreach (var ftpRecord in ftpRecords)
			{
				if (!rid2spRecord.ContainsKey(ftpRecord.RID))
					continue;

				var spRecord = new SharedPointSSRecord(rid2spRecord[ftpRecord.RID]);
				bool isModify = false;
				if(sourceNotNoneAndDestIsNone(ftpRecord.DateCTSReceived, spRecord.DateReced) 
					|| bothNotNoneAndNotEqual(ftpRecord.DateCTSReceived, spRecord.DateReced))
				{
					spRecord.DateReced = ftpRecord.DateCTSReceived;
					isModify = true;
				}
				if (sourceNotNoneAndDestIsNone(ftpRecord.FailMode, spRecord.FailureMode)
					|| bothNotNoneAndNotEqual(ftpRecord.FailMode, spRecord.FailureMode))
				{
					spRecord.FailureMode = ftpRecord.FailMode;
					isModify = true;
				}
				if (sourceNotNoneAndDestIsNone(ftpRecord.DateRootCauseReport, spRecord.DateRootCauseIdentified)
					|| bothNotNoneAndNotEqual(ftpRecord.DateRootCauseReport, spRecord.DateRootCauseIdentified))
				{
					spRecord.DateRootCauseIdentified = ftpRecord.DateRootCauseReport;
					isModify = true;
				}
				if (sourceNotNoneAndDestIsNone(ftpRecord.ProjectCode, spRecord.CorrectiveAction8D)
					|| bothNotNoneAndNotEqual(ftpRecord.ProjectCode, spRecord.CorrectiveAction8D))
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
		}
		private Dictionary<string, SharedPointSSCoreRecord>  GetFailureModeProgressUpdateMaps(List<SharedPointSSRecord> spRecords)
		{
			/*
			 * 1. Use .Upper() to avoid to generate two records with same meanings
			 * 2. Current just use the record in SpreadSheet to update result, so actually the result is from the last one sp record
			 * 3. Current we don't consider the toUpdate, but maybe need to consider it later
			 */
			Dictionary<string, SharedPointSSCoreRecord> failureMode2spRecord =
				new Dictionary<string, SharedPointSSCoreRecord>();
			foreach (var record in spRecords)
			{
				if (string.IsNullOrWhiteSpace(record.FailureMode))
					continue;
				string failureModeUpper = record.FailureMode.ToUpper();
				if (!failureMode2spRecord.ContainsKey(failureModeUpper))
				{
					failureMode2spRecord[failureModeUpper] = new SharedPointSSCoreRecord
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
				var matchRecord = failureMode2spRecord[failureModeUpper];
				Logger.Info($"record.ProblemClassification: {record.RID}, {record.ProblemClassification}");
				if (!isValidString(matchRecord.ProblemClassification) && isValidString(record.ProblemClassification))
					matchRecord.ProblemClassification = record.ProblemClassification;
				if (!isValidString(matchRecord.InterimCorrectiveActionDate) && isValidString(record.InterimCorrectiveActionDate))
					matchRecord.InterimCorrectiveActionDate = record.InterimCorrectiveActionDate;
				if (!isValidString(matchRecord.CleanDateFinalCorrectiveActionInPlace) 
					&& isValidString(record.CleanDateFinalCorrectiveActionInPlace))
					matchRecord.CleanDateFinalCorrectiveActionInPlace = record.CleanDateFinalCorrectiveActionInPlace;
				if (!isValidString(matchRecord.FirstSNAndOrDateCode) && isValidString(record.FirstSNAndOrDateCode))
					matchRecord.FirstSNAndOrDateCode = record.FirstSNAndOrDateCode;
				if (!isValidString(matchRecord.CorrectiveAction8D) && isValidString(record.CorrectiveAction8D))
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
			/*
			 * Using the failureMode2spRecord to update the problem classification field
			 */
			Logger.Info("UpdateSPRecordsByFailureMode start");
			foreach (var item in spRecords)
			{
				if (isValidString(item.ProblemClassification))
				{
					Logger.Info($"Item {item.RID} already have valid problem classification, skip");
					continue;
				}
				if (!isValidString(item.FailureMode))
				{
					Logger.Info($"Item {item.RID} have invalid failure mode, skip");
					continue;
				}
				string failureModeUpper = item.FailureMode.ToUpper();
				SharedPointSSRecord record = toUpdate.GetValueOrDefault(item.RID, new SharedPointSSRecord(item));
				if (!failureMode2spRecord.ContainsKey(failureModeUpper))
				{
					Logger.Info($"Item {record.RID} cant find match record in failureMode2spRecord");
					continue;
				}
				if (!isValidString(failureMode2spRecord[failureModeUpper].ProblemClassification))
				{
					Logger.Info($"Item {record.RID} match record don't have valid ProblemClassification");
					continue;
				}
				record.ProblemClassification = failureMode2spRecord[failureModeUpper].ProblemClassification;
				toUpdate[item.RID] = record;
			}
			Logger.Info("UpdateSPRecordsByFailureMode end");
		}
		private bool isValidString(string str)
		{
			if (string.IsNullOrEmpty(str))
				return false;
			if (string.IsNullOrWhiteSpace(str))
				return false;
			return str != "-";
		}
		private void UpdateSPRecordsByFailureModeAnd8DCorrectiveAction(
			List<SharedPointSSRecord> spRecords,
			Dictionary<string, SharedPointSSRecord> toUpdate,
			Dictionary<string, SharedPointSSCoreRecord> composite2spRecord)
		{
			/*
			 * Update the spreadSheet record according the composite2sprecords
			 * The main logic is
			 * 1. whether the spreadsheet record already valid value, if not, go on 
			 * 2. whether the spreadsheet have valid composite key, if has, go on
			 * 3. check whether the composite2sprecords has the matching record, if has, go on
			 * 4. check whether the corresponding field has the valid value, if has, assign
			 */
			Logger.Info("UpdateSPRecordsByFailureModeAnd8DCorrectiveAction start");
			foreach (var item in spRecords)
			{
				// foreach only return reference to the value for the reference type
				if (isValidString(item.InterimCorrectiveActionDate)
				&& isValidString(item.CleanDateFinalCorrectiveActionInPlace)
				&& isValidString(item.FirstSNAndOrDateCode))
				{
					continue;
				}
				if (!isValidString(item.FailureMode) && !isValidString(item.CorrectiveAction8D))
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
				if (isValidString(compositeRecord.InterimCorrectiveActionDate)
					&& record.InterimCorrectiveActionDate != compositeRecord.InterimCorrectiveActionDate)
				{ 
					record.InterimCorrectiveActionDate = compositeRecord.InterimCorrectiveActionDate;
					isModify = true;
				}
				if (isValidString(compositeRecord.CleanDateFinalCorrectiveActionInPlace)
					&& record.CleanDateFinalCorrectiveActionInPlace != compositeRecord.CleanDateFinalCorrectiveActionInPlace)
				{ 
					record.CleanDateFinalCorrectiveActionInPlace = compositeRecord.CleanDateFinalCorrectiveActionInPlace;
					isModify = true;
				}
				if (isValidString(compositeRecord.FirstSNAndOrDateCode)
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

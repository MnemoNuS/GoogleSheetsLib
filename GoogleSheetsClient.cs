using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace GoogleSheetsLib
{
	public class GoogleSheetsClient
	{
		public Action BeforeAction;

		private static GoogleSheetsClient _instance;
		public static GoogleSheetsClient GetInstanse()
		{
			if (_instance == null)
			{
				throw new NullReferenceException();
			}
			return _instance;
		}
		private GoogleSheetsClient() { }
		private GoogleSheetsClient(SheetsService sheetsService)
		{
			service = sheetsService;
		}

		static SheetsService service;

		static string[] Scopes = { SheetsService.Scope.Spreadsheets };
		static string AplicationSecrets = "client_secrets.json";
		static string ApplicationName = "AppName";
		static string SpreadsheetId = "1nSb2Lr6Z3vrJmDwfhZehLswOhugDkyi2ihGwzSuyQmE";
		static string DefaultSheet = "";

		public static GoogleSheetsClient Init(string aplicationName, string spreadsheetId, string defaultSheet, string aplicationSecrets)
		{
			ApplicationName = aplicationName;
			SpreadsheetId = spreadsheetId;
			DefaultSheet = defaultSheet;
			AplicationSecrets = aplicationSecrets;

			GoogleCredential credestial;
			using (var stream = new FileStream(AplicationSecrets, FileMode.Open, FileAccess.Read))
			{
				credestial = GoogleCredential.FromStream(stream).CreateScoped(Scopes);
			}

			// Create Google Sheets API service.
			_instance = new GoogleSheetsClient(new SheetsService(new BaseClientService.Initializer()
			{
				HttpClientInitializer = credestial,
				ApplicationName = ApplicationName,
			}));

			return _instance;
		}
		public IList<IList<object>> ReadData(string sheet = "")
		{
			BeforeAction();

			var sheetMeta = GetSheet(sheet);
			if (sheetMeta == null)
				return new List<IList<object>>();

			var range = $"{sheet}!A:Z";
			var request = service.Spreadsheets.Values.Get(SpreadsheetId, range);
			var response = request.Execute();
			var values = response.Values;
			return values;
		}
		public IList<IList<object>> ReadData(string colFrom, string colTo, string sheet = "")
		{
			BeforeAction();

			var sheetMeta = GetSheet(sheet);
			if (sheetMeta == null)
				return new List<IList<object>>();

			var range = $"{sheet}!{colFrom}:{colTo}";
			var request = service.Spreadsheets.Values.Get(SpreadsheetId, range);
			var response = request.Execute();
			var values = response.Values;
			return values;
		}
		public void CreateData(List<object> objectList, string sheet = "")
		{
			BeforeAction();

			var sheetMeta = GetSheet(sheet);
			if (sheetMeta == null)
				return;

			var range = $"{sheet}!A:Z";
			var valueRange = new ValueRange();
			valueRange.Values = new List<IList<object>>() { objectList };

			var request = service.Spreadsheets.Values.Append(valueRange, SpreadsheetId, range);
			request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED; ;

			var response = request.Execute();
		}
		public void UpdateData(int row, int col, List<object> objectList, string sheet = "")
		{
			BeforeAction();

			var sheetMeta = GetSheet(sheet);
			if (sheetMeta == null)
				return;

			var range = $"{sheet}!R{row}C{col}:R{row}C{col + objectList.Count}";
			var valueRange = new ValueRange();
			valueRange.Values = new List<IList<object>>() { objectList };

			var request = service.Spreadsheets.Values.Update(valueRange, SpreadsheetId, range);
			request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;

			var response = request.Execute();
		}
		public void UpdateColData(int row, int col, List<object> objectList, string sheet = "")
		{
			BeforeAction();

			var sheetMeta = GetSheet(sheet);
			if (sheetMeta == null)
				return;

			var range = $"{sheet}!R{row}C{col}:R{row + objectList.Count}C{col }";
			var valueRange = new ValueRange();
			valueRange.Values = objectList.Select(v => new List<object> { v }).ToArray();

			var request = service.Spreadsheets.Values.Update(valueRange, SpreadsheetId, range);
			request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;

			var response = request.Execute();
		}
		public void DeleteDataRow(int row, string sheet = "")
		{
			BeforeAction();

			var sheetMeta = GetSheet(sheet);
			if (sheetMeta == null)
				return;

			var range = $"{sheet}!R{row}C1:R{row}C100";

			var request = service.Spreadsheets.Values.Clear(new ClearValuesRequest(), SpreadsheetId, range);
			var response = request.Execute();
		}
		private Sheet GetSheet(string name)
		{
			if (string.IsNullOrEmpty(name))
			{
				name = DefaultSheet;
			}

			var sheetsMeta = service.Spreadsheets.Get(SpreadsheetId).Execute();
			var sheet = sheetsMeta.Sheets.FirstOrDefault(s => s.Properties.Title == name);

			return sheet;
		}
		public bool IsSheetExist(string name)
		{
			var sheetMeta = GetSheet(name);
			return sheetMeta != null;
		}
		public void CreateSheet(string name)
		{
			var addSheetRequest = new AddSheetRequest();
			addSheetRequest.Properties = new SheetProperties();
			addSheetRequest.Properties.Title = name;

			BatchUpdateSpreadsheetRequest batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest();
			batchUpdateSpreadsheetRequest.Requests = new List<Request>();
			batchUpdateSpreadsheetRequest.Requests.Add(new Request { AddSheet = addSheetRequest });

			var batchUpdateRequest = service.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, SpreadsheetId);
			batchUpdateRequest.Execute();
		}
	}
}

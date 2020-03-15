//
// Program.cs
//
// Author: Jeffrey Stedfast <jestedfa@microsoft.com>
//
// Copyright (c) 2016-2020 Jeffrey Stedfast
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.
//

using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Globalization;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Text.RegularExpressions;

using NPOI.SS.Util;
using NPOI.HSSF.Util;
using NPOI.XSSF.Util;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;

using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;

namespace IvyPortfolio
{
	class Program
	{
		// Note: Used for Conditional Variance formatting
		//static readonly XSSFColor CustomLightYellow = new XSSFColor (new byte[] { 250, 230, 181 });
		//static readonly XSSFColor CustomLightGreen = new XSSFColor (new byte[] { 191, 222, 204 });
		//static readonly XSSFColor CustomLightRed = new XSSFColor (new byte[] { 237, 201, 196 });
		static readonly short LightYellow = IndexedColors.LightYellow.Index;
		static readonly short LightGreen = IndexedColors.LightGreen.Index;
		static readonly short LightRed = IndexedColors.Rose.Index;

		// Note: Used for Conditional Position formatting
		static readonly short White = IndexedColors.White.Index;
		static readonly short Green = IndexedColors.BrightGreen.Index;
		static readonly short Red = IndexedColors.Red.Index;

		// Note: Used for Dashboard table headers
		static readonly XSSFColor CustomLightBlue = new XSSFColor (new byte[] { 204, 219, 247 });
		//static readonly XSSFColor CustomLightGrey = new XSSFColor (new byte[] { 217, 217, 217 });
		static readonly short LightBlue = IndexedColors.Aqua.Index;
		static readonly short LightGrey = IndexedColors.Grey25Percent.Index;

		static readonly DateTime UnixEpoch = new DateTime (1970, 1, 1, 0, 0, 0, 0, DateTimeKind.Utc);
		static readonly Regex CrumbRegex = new Regex ("CrumbStore\":{\"crumb\":\"(?<crumb>.+?)\"}", RegexOptions.CultureInvariant | RegexOptions.Compiled);
		static string crumb;

		public static void Main (string[] args)
		{
			var home = Environment.GetFolderPath (Environment.SpecialFolder.Personal);
			var dataDir = Path.Combine (home, "Dropbox", "IvyPortfolio");

			RunAsync (dataDir).GetAwaiter ().GetResult ();
		}

		static async Task RunAsync (string dataDir)
		{
			string[] scopes = { SheetsService.Scope.Spreadsheets };
			var map = LoadGoogleSheetMappings (dataDir);
			var handler = new HttpClientHandler ();
			UserCredential credential;

			using (var stream = File.OpenRead (Path.Combine (dataDir, "credentials.json"))) {
				var login = File.ReadAllText (Path.Combine (dataDir, "login.txt")).Trim ();

				credential = await GoogleWebAuthorizationBroker.AuthorizeAsync (
					GoogleClientSecrets.Load (stream).Secrets, scopes, login,
					CancellationToken.None);
			}

			// Create Google Sheets API service.
			var service = new SheetsService (new BaseClientService.Initializer {
				HttpClientInitializer = credential,
				ApplicationName = "IvyPortfolio"
			});

			handler.CookieContainer = new CookieContainer ();
			handler.UseCookies = true;

			using (var client = new HttpClient (handler)) {
				XSSFWorkbook workbook;
				string[] symbols;
				string fileName;
				string id;

				symbols = new string[] { "VTSAX", "VFWAX", "VBTLX", "VGSLX", "VGELX", "VGPMX" };
				fileName = "Investment Portfolio (MUTF)";

				workbook = await CreateSpreadsheetAsync (client, fileName + ".xlsx", symbols);
				if (map.TryGetValue (fileName, out id))
					UpdateGoogleSpreadsheet (service, id, workbook);

				symbols = new string[] { "VTI", "VEU", "BND", "VOX", "VCR", "VDC", "VDE", "VFH", "VHT", "VIS", "VAW", "VNQ", "VGT", "VPU" };
				fileName = "Investment Portfolio (ETF)";

				workbook = await CreateSpreadsheetAsync (client, fileName + ".xlsx", symbols);
				if (map.TryGetValue (fileName, out id))
					UpdateGoogleSpreadsheet (service, id, workbook);

				//symbols = { "BND", "DBC", "GSG", "RWX", "VNQ", "TIP", "VWO", "VEU", "VB", "VTI" };
				//fileName = "Investment Portfolio (TEST)";

				//workbook = await CreateSpreadsheetAsync (client, fileName + ".xlsx", symbols);
				//if (docmap.TryGetValue (fileName, out id))
				//	UpdateGoogleSpreadsheet (service, id, workbook);
			}
		}

		static Dictionary<string, string> LoadGoogleSheetMappings (string dataDir)
		{
			var path = Path.Combine (dataDir, "document-map.txt");
			var map = new Dictionary<string, string> ();

			if (File.Exists (path)) {
				using (var reader = File.OpenText (path)) {
					string line;

					while ((line = reader.ReadLine ()) != null) {
						if (line.Length == 0 || line[0] == '#')
							continue;

						var kvp = line.TrimEnd ().Split (new char[] { ' ' }, 2);

						if (kvp.Length < 2)
							continue;

						map[kvp[1]] = kvp[0];
					}
				}
			}

			return map;
		}

		static void UpdateGoogleSpreadsheet (SheetsService service, string spreadsheetId, XSSFWorkbook workbook)
		{
			// Sheet1 = Dashboard, Sheet2 = Charts
			for (int index = 2; index < workbook.NumberOfSheets; index++) {
				var sheet = workbook.GetSheetAt (index);
				var values = new List<IList<object>> ();

				for (int i = 0; i < sheet.LastRowNum; i++) {
					var row = sheet.GetRow (i);

					values.Add (new object[10]);
					for (int j = 0; j < row.LastCellNum; j++) {
						var cell = row.GetCell (j, MissingCellPolicy.RETURN_BLANK_AS_NULL);
						string value;

						if (cell != null) {
							switch (cell.CellType) {
							case CellType.String: value = cell.StringCellValue; break;
							case CellType.Numeric: value = cell.NumericCellValue.ToString (); break;
							case CellType.Formula: value = "=" + cell.CellFormula; break;
							default: value = string.Empty; break;
							}
						} else {
							value = string.Empty;
						}

						values[i][j] = value;
					}

					for (int j = row.LastCellNum; j < 10; j++)
						values[i][j] = string.Empty;
				}

				var range = string.Format ("{0}!A1:J", sheet.SheetName);
				var body = new ValueRange { Values = values };

				var request = service.Spreadsheets.Values.Update (body, spreadsheetId, range);
				request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;

				var response = request.Execute ();
			}
		}

		static async Task<XSSFWorkbook> CreateSpreadsheetAsync (HttpClient client, string fileName, string[] symbols)
		{
			var workbook = new XSSFWorkbook ();

			await CreateSpreadsheetAsync (client, workbook, symbols);

			using (var stream = File.Create (fileName))
				workbook.Write (stream);

			return workbook;
		}

		static void GetDateRange (out DateTime start, out DateTime end)
		{
			var today = DateTime.Today.ToUniversalTime ();

			if (today.Day < DateTime.DaysInMonth (today.Year, today.Month))
				end = today.AddDays (-1 * today.Day);
			else
				end = today;

			start = end.AddYears (-4);
		}

		static async Task CreateSpreadsheetAsync (HttpClient client, IWorkbook workbook, string[] symbols)
		{
			var descriptions = new Dictionary<string, string> ();
			DateTime start, end;

			GetDateRange (out start, out end);

			var dashboard = workbook.CreateSheet ("Dashboard");
			var charts = workbook.CreateSheet ("Charts");
			var small = workbook.CreateFont ();
			var bold = workbook.CreateFont ();
			var font = workbook.CreateFont ();
			int rowIndex12 = (symbols.Length + 4) * 2;
			int rowIndex10 = symbols.Length + 4;
			int rowIndex200 = 0;

			small.FontName = bold.FontName = font.FontName = "Arial";
			small.FontHeightInPoints = 8;
			font.FontHeightInPoints = 11;
			bold.FontHeightInPoints = 11;
			small.IsItalic = true;
			bold.IsBold = true;

			dashboard.DefaultColumnWidth = 18;
			CreateDashboardTable (dashboard, small, bold, font, rowIndex200, symbols.Length, "200-Day SMA");
			CreateDashboardTable (dashboard, small, bold, font, rowIndex10, symbols.Length, "10-Month SMA");
			CreateDashboardTable (dashboard, small, bold, font, rowIndex12, symbols.Length, "12-Month SMA");

			rowIndex200 += 2;
			rowIndex10 += 2;
			rowIndex12 += 2;

			var positionRegions = new[] {
				new CellRangeAddress (rowIndex200, rowIndex200 + symbols.Length, (int) TableColumn.Position, (int) TableColumn.Position),
				new CellRangeAddress (rowIndex10, rowIndex10 + symbols.Length, (int) TableColumn.Position, (int) TableColumn.Position),
				new CellRangeAddress (rowIndex12, rowIndex12 + symbols.Length, (int) TableColumn.Position, (int) TableColumn.Position)
			};

			var varianceRegions = new[] {
				new CellRangeAddress (rowIndex200, rowIndex200 + symbols.Length, (int) TableColumn.Variance, (int) TableColumn.Variance),
				new CellRangeAddress (rowIndex10, rowIndex10 + symbols.Length, (int) TableColumn.Variance, (int) TableColumn.Variance),
				new CellRangeAddress (rowIndex12, rowIndex12 + symbols.Length, (int) TableColumn.Variance, (int) TableColumn.Variance)
			};

			foreach (var symbol in symbols) {
				descriptions.Add (symbol, await GetStockDescriptionAsync (client, symbol));

				await CreateSheetAsync (client, workbook, font, symbol, start, end);

				CreateDashboardTableRow (dashboard, bold, font, rowIndex200, symbol, DataColumn.SMA200Day);
				CreateDashboardTableRow (dashboard, bold, font, rowIndex10, symbol, DataColumn.SMA10Month);
				CreateDashboardTableRow (dashboard, bold, font, rowIndex12, symbol, DataColumn.SMA12Month);

				rowIndex200++;
				rowIndex10++;
				rowIndex12++;
			}

			ApplyConditionalPositionFormatting (dashboard, positionRegions);
			ApplyConditionalVarianceFormatting (dashboard, varianceRegions);

			CreateDashboardLegend (dashboard, bold, font, 0, "Funds", symbols, descriptions);
		}

		static void ApplyConditionalPositionFormatting (ISheet dashboard, CellRangeAddress[] regions)
		{
			var formatting = dashboard.SheetConditionalFormatting;
			var rules = new IConditionalFormattingRule[2];
			IConditionalFormattingRule rule;
			IPatternFormatting pattern;
			IFontFormatting font;
			int index = 0;

			// Create the "Invested" formatting rule
			rule = formatting.CreateConditionalFormattingRule (ComparisonOperator.Equal, "\"Invested\"");
			pattern = rule.CreatePatternFormatting ();
			pattern.FillPattern = FillPattern.SolidForeground;
			pattern.FillBackgroundColor = Green;

			rules[index++] = rule;

			// Create the "Cash" formatting rule
			rule = formatting.CreateConditionalFormattingRule (ComparisonOperator.Equal, "\"Cash\"");
			pattern = rule.CreatePatternFormatting ();
			pattern.FillPattern = FillPattern.SolidForeground;
			pattern.FillBackgroundColor = Red;

			font = rule.CreateFontFormatting ();
			font.FontColorIndex = White;

			rules[index++] = rule;

			// Apply the conditional formatting rules
			formatting.AddConditionalFormatting (regions, rules);
		}

		static void ApplyConditionalVarianceFormatting (ISheet dashboard, CellRangeAddress[] regions)
		{
			var formatting = dashboard.SheetConditionalFormatting;
			var rules = new IConditionalFormattingRule[3];
			IConditionalFormattingRule rule;
			IPatternFormatting pattern;
			int index = 0;

			// Add the "Invest" formatting
			rule = formatting.CreateConditionalFormattingRule (ComparisonOperator.GreaterThan, "2");
			pattern = rule.CreatePatternFormatting ();
			pattern.FillPattern = FillPattern.SolidForeground;
			pattern.FillBackgroundColor = LightGreen;

			rules[index++] = rule;

			// Add the "Neutral" formatting
			rule = formatting.CreateConditionalFormattingRule (ComparisonOperator.Between, "-2", "2");
			pattern = rule.CreatePatternFormatting ();
			pattern.FillPattern = FillPattern.SolidForeground;
			pattern.FillBackgroundColor = LightYellow;

			rules[index++] = rule;

			// Add the "Sell" formatting
			rule = formatting.CreateConditionalFormattingRule (ComparisonOperator.LessThan, "-2");
			pattern = rule.CreatePatternFormatting ();
			pattern.FillPattern = FillPattern.SolidForeground;
			pattern.FillBackgroundColor = LightRed;

			rules[index++] = rule;

			// Apply the conditional formatting rules
			formatting.AddConditionalFormatting (regions, rules);
		}

		static void CreateDashboardTableHeader (ISheet dashboard, IFont bold, IFont font, int rowIndex, string title)
		{
			string[] headerNames = { "Fund", "Position", "Variance*" };
			var style = dashboard.Workbook.CreateCellStyle ();
			var row = dashboard.CreateRow (rowIndex);
			ICell cell;

			if (style is XSSFCellStyle)
				((XSSFCellStyle) style).SetFillForegroundColor (CustomLightBlue);
			else
				style.FillForegroundColor = LightBlue;
			style.FillPattern = FillPattern.SolidForeground;
			style.Alignment = HorizontalAlignment.Center;
			style.BorderBottom = BorderStyle.Thin;
			style.BorderRight = BorderStyle.Thin;
			style.BorderLeft = BorderStyle.Thin;
			style.BorderTop = BorderStyle.Thin;
			style.SetFont (bold);

			cell = row.CreateCell ((int) TableColumn.Fund, CellType.String);
			cell.SetCellValue (title);
			cell.CellStyle = style;

			cell = row.CreateCell ((int) TableColumn.Position, CellType.String);
			cell.CellStyle = style;

			cell = row.CreateCell ((int) TableColumn.Variance, CellType.String);
			cell.CellStyle = style;

			var region = new CellRangeAddress (rowIndex, rowIndex, (int) TableColumn.Fund, (int) TableColumn.Variance);
			dashboard.AddMergedRegion (region);

			style = dashboard.Workbook.CreateCellStyle ();
			style.FillForegroundColor = LightGrey;
			style.FillPattern = FillPattern.SolidForeground;
			style.Alignment = HorizontalAlignment.Center;
			style.BorderBottom = BorderStyle.Thin;
			style.BorderRight = BorderStyle.Thin;
			style.BorderLeft = BorderStyle.Thin;
			style.BorderTop = BorderStyle.Thin;
			style.SetFont (font);

			row = dashboard.CreateRow (rowIndex + 1);
			for (int i = 0; i < headerNames.Length; i++) {
				cell = row.CreateCell (i, CellType.String);
				cell.SetCellValue (headerNames[i]);
				cell.CellStyle = style;
			}
		}

		static void CreateDashboardTableFooter (ISheet dashboard, IFont font, int rowIndex, string footer)
		{
			var style = dashboard.Workbook.CreateCellStyle ();
			var row = dashboard.CreateRow (rowIndex);
			ICell cell;

			style.Alignment = HorizontalAlignment.Left;
			style.BorderBottom = BorderStyle.Thin;
			style.BorderRight = BorderStyle.Thin;
			style.BorderLeft = BorderStyle.Thin;
			style.BorderTop = BorderStyle.Thin;
			style.SetFont (font);

			cell = row.CreateCell ((int) TableColumn.Fund, CellType.String);
			cell.SetCellValue (footer);
			cell.CellStyle = style;

			cell = row.CreateCell ((int) TableColumn.Position, CellType.String);
			cell.CellStyle = style;

			cell = row.CreateCell ((int) TableColumn.Variance, CellType.String);
			cell.CellStyle = style;

			var region = new CellRangeAddress (rowIndex, rowIndex, (int) TableColumn.Fund, (int) TableColumn.Variance);
			dashboard.AddMergedRegion (region);
		}

		static void CreateDashboardTable (ISheet dashboard, IFont small, IFont bold, IFont font, int rowIndex, int numSymbols, string name)
		{
			var footer = string.Format ("* Percent above/below the {0}", name);
			var header = string.Format ("Ivy Portfolio {0} Signals", name);

			CreateDashboardTableHeader (dashboard, bold, font, rowIndex, header);
			CreateDashboardTableFooter (dashboard, small, rowIndex + numSymbols + 2, footer);
		}

		static void CreateDashboardTableRow (ISheet dashboard, IFont bold, IFont font, int rowIndex, string symbol, DataColumn column)
		{
			var boldStyle = dashboard.Workbook.CreateCellStyle ();
			var style = dashboard.Workbook.CreateCellStyle ();
			var row = dashboard.CreateRow (rowIndex);
			ICell cell;

			boldStyle.Alignment = HorizontalAlignment.Center;
			boldStyle.BorderBottom = BorderStyle.Thin;
			boldStyle.BorderRight = BorderStyle.Thin;
			boldStyle.BorderLeft = BorderStyle.Thin;
			boldStyle.BorderTop = BorderStyle.Thin;
			boldStyle.SetFont (bold);

			style.Alignment = HorizontalAlignment.Center;
			style.BorderBottom = BorderStyle.Thin;
			style.BorderRight = BorderStyle.Thin;
			style.BorderLeft = BorderStyle.Thin;
			style.BorderTop = BorderStyle.Thin;
			style.SetFont (font);

			// Create the cell for the symbol name
			cell = row.CreateCell ((int) TableColumn.Fund, CellType.String);
			cell.SetCellValue (symbol);
			cell.CellStyle = style;

			// Create the cell for the buy/sell position
			cell = row.CreateCell ((int) TableColumn.Position, CellType.Formula);
			cell.SetCellFormula (string.Format ("IF({0}{1} > 0, \"Invested\", \"Cash\")",
			                                    (char) ('A' + TableColumn.Variance),
			                                    rowIndex + 1));
			cell.CellStyle = boldStyle;

			// Create the cell for the variance
			var variance = string.Format ("ROUND(({0}!{1}2 - {0}!{2}2) / {0}!{2}2 * 100, 2)", symbol, 
			                             (char) ('A' + DataColumn.AdjClose), (char) ('A' + column));
			cell = row.CreateCell ((int) TableColumn.Variance, CellType.Formula);
			cell.SetCellFormula (variance);
			cell.CellStyle = style;
		}

		static void CreateDashboardLegendHeader (ISheet dashboard, IFont bold, int rowIndex, string title)
		{
			var style = dashboard.Workbook.CreateCellStyle ();
			var row = dashboard.GetRow (rowIndex);
			ICell cell;

			if (style is XSSFCellStyle)
				((XSSFCellStyle) style).SetFillForegroundColor (CustomLightBlue);
			else
				style.FillForegroundColor = LightBlue;
			style.FillPattern = FillPattern.SolidForeground;
			style.Alignment = HorizontalAlignment.Center;
			style.BorderBottom = BorderStyle.Thin;
			style.BorderRight = BorderStyle.Thin;
			style.BorderLeft = BorderStyle.Thin;
			style.BorderTop = BorderStyle.Thin;
			style.SetFont (bold);

			cell = row.CreateCell ((int) LegendColumn.Fund, CellType.String);
			cell.SetCellValue (title);
			cell.CellStyle = style;

			cell = row.CreateCell ((int) LegendColumn.Name, CellType.String);
			cell.CellStyle = style;

			cell = row.CreateCell ((int) LegendColumn.Name + 1, CellType.String);
			cell.CellStyle = style;

			var region = new CellRangeAddress (rowIndex, rowIndex, (int) LegendColumn.Fund, (int) LegendColumn.Name + 1);
			dashboard.AddMergedRegion (region);

			style = dashboard.Workbook.CreateCellStyle ();
			style.FillForegroundColor = LightGrey;
			style.FillPattern = FillPattern.SolidForeground;
			style.Alignment = HorizontalAlignment.Center;
			style.BorderBottom = BorderStyle.Thin;
			style.BorderRight = BorderStyle.Thin;
			style.BorderLeft = BorderStyle.Thin;
			style.BorderTop = BorderStyle.Thin;
			style.SetFont (bold);

			row = dashboard.GetRow (rowIndex + 1);

			foreach (LegendColumn column in Enum.GetValues (typeof (LegendColumn))) {
				cell = row.CreateCell ((int) column, CellType.String);
				cell.SetCellValue (column.ToString ());
				cell.CellStyle = style;
			}

			cell = row.CreateCell ((int) LegendColumn.Name + 1, CellType.String);
			cell.CellStyle = style;

			region = new CellRangeAddress (rowIndex + 1, rowIndex + 1, (int) LegendColumn.Name, (int) LegendColumn.Name + 1);
			dashboard.AddMergedRegion (region);
		}

		static void CreateDashboardLegend (ISheet dashboard, IFont bold, IFont font, int rowIndex, string title, string[] symbols, IDictionary<string, string> descriptions)
		{
			var fundStyle = dashboard.Workbook.CreateCellStyle ();
			var nameStyle = dashboard.Workbook.CreateCellStyle ();

			CreateDashboardLegendHeader (dashboard, bold, rowIndex, title);

			fundStyle.FillForegroundColor = LightGrey;
			fundStyle.FillPattern = FillPattern.SolidForeground;
			fundStyle.Alignment = HorizontalAlignment.Center;
			fundStyle.BorderBottom = BorderStyle.Thin;
			fundStyle.BorderRight = BorderStyle.Thin;
			fundStyle.BorderLeft = BorderStyle.Thin;
			fundStyle.BorderTop = BorderStyle.Thin;
			fundStyle.SetFont (font);

			nameStyle.Alignment = HorizontalAlignment.Left;
			nameStyle.BorderBottom = BorderStyle.Thin;
			nameStyle.BorderRight = BorderStyle.Thin;
			nameStyle.BorderLeft = BorderStyle.Thin;
			nameStyle.BorderTop = BorderStyle.Thin;
			nameStyle.SetFont (font);

			for (int i = 0; i < symbols.Length; i++) {
				var row = dashboard.GetRow (rowIndex + 2 + i);
				string name;
				ICell cell;

				cell = row.CreateCell ((int) LegendColumn.Fund, CellType.String);
				cell.SetCellValue (symbols[i]);
				cell.CellStyle = fundStyle;

				cell = row.CreateCell ((int) LegendColumn.Name, CellType.String);
				if (!descriptions.TryGetValue (symbols[i], out name))
					cell.SetCellValue (string.Empty);
				else
					cell.SetCellValue (name);
				cell.CellStyle = nameStyle;

				cell = row.CreateCell ((int) LegendColumn.Name + 1, CellType.String);
				cell.CellStyle = nameStyle;

				var region = new CellRangeAddress (row.RowNum, row.RowNum, (int) LegendColumn.Name, (int) LegendColumn.Name + 1);
				dashboard.AddMergedRegion (region);
			}
		}

		static async Task<string> GetStockDescriptionAsync (HttpClient client, string symbol)
		{
			const string format = "https://finance.yahoo.com/quote/{0}?p={0}";
			var requestUri = string.Format (format, symbol);

			var html = await client.GetStringAsync (requestUri);
			int startIndex, endIndex;

			// extract the cookie crumb
			var crumbs = CrumbRegex.Matches (html);

			if (crumbs.Count > 0) {
				crumb = crumbs[0].Groups["crumb"].Value;
				crumb = crumb.Replace ("\\u002F", "/");
			} else {
				crumb = "xxxxxxxxxxx";
			}

			if ((endIndex = html.IndexOf ("(" + symbol + ")", StringComparison.Ordinal)) <= 0) {
				Console.WriteLine ("Failed to locate \"({0})\" in:", symbol);
				Console.WriteLine ("{0}", html);
				Console.WriteLine ();
				return string.Empty;
			}

			if (html[endIndex - 1] == ' ')
				endIndex--;

			startIndex = endIndex;

			while (startIndex > 0 && html[startIndex - 1] != '>')
				startIndex--;

			return html.Substring (startIndex, endIndex - startIndex).Replace ("&amp;", "&");
		}

		static async Task<string> GetStockDataAsync (HttpClient client, string symbol, DateTime start, DateTime end)
		{
			const string format = "https://query1.finance.yahoo.com/v7/finance/download/{0}?period1={1}&period2={2}&interval=1d&events=history&crumb={3}";
			var requestUri = string.Format (format, symbol, (start - UnixEpoch).TotalSeconds, (end - UnixEpoch).TotalSeconds, crumb);

			var data = await client.GetStringAsync (requestUri);

			return data;
		}

		static bool IsNull (string[] data)
		{
			for (int i = 1; i < data.Length; i++) {
				if (data[i] == "null")
					return true;
			}

			return false;
		}

		static async Task<ISheet> CreateSheetAsync (HttpClient client, IWorkbook workbook, IFont font, string symbol, DateTime start, DateTime end)
		{
			var csv = await GetStockDataAsync (client, symbol, start, end);
			var sheet = workbook.CreateSheet (symbol);
			var hstyle = workbook.CreateCellStyle ();
			var style = workbook.CreateCellStyle ();

			hstyle.Alignment = HorizontalAlignment.Center;
			hstyle.FillPattern = FillPattern.SolidForeground;
			hstyle.FillForegroundColor = LightGrey;
			hstyle.ShrinkToFit = false;
			hstyle.SetFont (font);

			style.Alignment = HorizontalAlignment.Center;
			style.ShrinkToFit = false;
			style.SetFont (font);

			sheet.DefaultColumnWidth = 12;

			using (var reader = new StringReader (csv)) {
				var columnNames = reader.ReadLine ().Split (',');
				var endOfMonthRows = new List<int> ();
				var lines = new List<string> ();
				int previousMonth = -1;
				var columnIndex = 0;
				var rowIndex = 1;
				string line;
				ICell cell;
				IRow row;

				for (int i = sheet.LastRowNum; i > 0; i--) {
					row = sheet.GetRow (i - 1);
					sheet.RemoveRow (row);
				}

				// Add the Titles for the data columns
				row = sheet.CreateRow (0);
				while (columnIndex < columnNames.Length) {
					cell = row.CreateCell (columnIndex, CellType.String);
					cell.SetCellValue (columnNames[columnIndex]);
					cell.CellStyle = hstyle;
					columnIndex++;
				}

				// Add the Titles for the formula columns
				//sheet.SetDefaultColumnStyle ((int) DataColumn.SMA200Day, style);
				cell = row.CreateCell ((int) DataColumn.SMA200Day, CellType.String);
				cell.SetCellValue ("200-Day SMA");
				cell.CellStyle = hstyle;

				//sheet.SetDefaultColumnStyle ((int) DataColumn.SMA10Month, style);
				cell = row.CreateCell ((int) DataColumn.SMA10Month, CellType.String);
				cell.SetCellValue ("10 Month SMA");
				cell.CellStyle = hstyle;

				//sheet.SetDefaultColumnStyle ((int) DataColumn.SMA12Month, style);
				cell = row.CreateCell ((int) DataColumn.SMA12Month, CellType.String);
				cell.SetCellValue ("12 Month SMA");
				cell.CellStyle = hstyle;

				// The data is sorted such that the latest data is at the end
				while ((line = reader.ReadLine ()) != null)
					lines.Add (line);

				// Read the data in reverse
				for (int i = lines.Count - 1; i >= 0; i--) {
					var data = lines[i].Split (',');
					DateTime date;

					if (IsNull (data))
						continue;

					row = sheet.CreateRow (rowIndex++);
					columnIndex = 0;

					// Note: the first column is a DateTime value, all other values are stock price values
					date = DateTime.Parse (data[0], CultureInfo.InvariantCulture);
					cell = row.CreateCell (columnIndex++, CellType.String);
					cell.SetCellValue (data[0]);
					cell.CellStyle = style;

					if (date.Month != previousMonth) {
						endOfMonthRows.Add (rowIndex - 1);
						previousMonth = date.Month;
					}

					while (columnIndex < data.Length) {
						cell = row.CreateCell (columnIndex, CellType.Numeric);
						cell.SetCellValue (double.Parse (data[columnIndex], CultureInfo.InvariantCulture));
						cell.CellStyle = style;
						columnIndex++;
					}
				}

				// Set the formula for the 200-Day SMA cells
				for (int i = 1; i < rowIndex - 200; i++) {
					row = sheet.GetRow (i);

					cell = row.CreateCell ((int) DataColumn.SMA200Day, CellType.Formula);
					cell.SetCellFormula (string.Format ("AVERAGE({0}{1}:{0}{2})", (char) ('A' + DataColumn.AdjClose), i + 1, i + 201));
					cell.CellStyle = style;
				}

				// Set the formulas for the 10-Month and 12-Month SMA cells
				for (int i = 0; i < endOfMonthRows.Count - 12; i++) {
					var items = new List<string> ();

					row = sheet.GetRow (endOfMonthRows[i]);

					cell = row.CreateCell ((int) DataColumn.SMA10Month, CellType.Formula);
					for (int k = 0; k < 10; k++)
						items.Add (string.Format ("{0}{1}", (char) ('A' + DataColumn.AdjClose), endOfMonthRows[i + k] + 1));
					cell.SetCellFormula (string.Format ("AVERAGE({0})", string.Join (", ", items)));
					cell.CellStyle = style;

					cell = row.CreateCell ((int) DataColumn.SMA12Month, CellType.Formula);
					for (int k = 0; k < 2; k++)
						items.Add (string.Format ("{0}{1}", (char) ('A' + DataColumn.AdjClose), endOfMonthRows[i + 10 + k] + 1));
					cell.SetCellFormula (string.Format ("AVERAGE({0})", string.Join (", ", items)));
					cell.CellStyle = style;
				}
			}

			return sheet;
		}
	}
}

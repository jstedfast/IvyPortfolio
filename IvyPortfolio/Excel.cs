//
// Excel.cs
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
using System.Threading;
using System.Globalization;
using System.Threading.Tasks;
using System.Collections.Generic;

using NPOI.SS.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace IvyPortfolio
{
	public static class Excel
	{
		class MovingAverageColumn
		{
			public MovingAverage MovingAverage { get; private set; }
			public int DataColumn { get; private set; }
			public string Title { get; private set; }
			public int RowIndex { get; set; }

			public MovingAverageColumn (MovingAverage movingAverage, int dataColumn, int rowIndex)
			{
				MovingAverage = movingAverage;
				Title = movingAverage.Title;
				DataColumn = dataColumn;
				RowIndex = rowIndex;

				if (string.IsNullOrEmpty (Title))
					Title = string.Format (CultureInfo.InvariantCulture, "{0}-{1} SMA", movingAverage.Period, movingAverage.PeriodType);
			}
		}

		public static async Task<IWorkbook> CreateSpreadsheetAsync (IFinancialService client, Document document, CancellationToken cancellationToken)
		{
			var movingAverageColumns = GetMovingAverageColumns (document, out var monthly);
			var descriptions = new Dictionary<string, string> ();
			var workbook = new XSSFWorkbook ();
			var symbols = document.Symbols;
			DateTime start, end;

			GetDateRange (monthly, out start, out end);

			Console.WriteLine ("Generating '{0}' based on data from {1} to {2}", document.FileName, start, end);

			var positionRegions = new CellRangeAddress[movingAverageColumns.Count];
			var varianceRegions = new CellRangeAddress[movingAverageColumns.Count];
			var dashboard = workbook.CreateSheet ("Dashboard");
			var charts = workbook.CreateSheet ("Charts");
			var small = workbook.CreateFont ();
			var bold = workbook.CreateFont ();
			var font = workbook.CreateFont ();

			small.FontName = bold.FontName = font.FontName = "Arial";
			small.FontHeightInPoints = 8;
			font.FontHeightInPoints = 11;
			bold.FontHeightInPoints = 11;
			small.IsItalic = true;
			bold.IsBold = true;

			dashboard.DefaultColumnWidth = 18;

			for (int i = 0; i < movingAverageColumns.Count; i++) {
				var column = movingAverageColumns[i];

				CreateDashboardTable (dashboard, small, bold, font, column.RowIndex, symbols.Length, column.Title);
				column.RowIndex += 2;

				int lastRowIndex = column.RowIndex + symbols.Length;

				positionRegions[i] = new CellRangeAddress (column.RowIndex, lastRowIndex, (int) TableColumn.Position, (int) TableColumn.Position);
				varianceRegions[i] = new CellRangeAddress (column.RowIndex, lastRowIndex, (int) TableColumn.Variance, (int) TableColumn.Variance);
			}

			foreach (var symbol in symbols) {
				var description = await client.GetStockDescriptionAsync (symbol, cancellationToken).ConfigureAwait (false);

				descriptions.Add (symbol, description);

				try {
					await CreateSheetAsync (client, workbook, font, movingAverageColumns, symbol, start, end, cancellationToken).ConfigureAwait (false);
				} catch (Exception ex) {
					Console.WriteLine ("\tFailed to get stock data for {0}: {1}\n{2}", symbol, ex.Message, ex);
					throw;
				}

				foreach (var column in movingAverageColumns) {
					CreateDashboardTableRow (dashboard, bold, font, column.RowIndex, symbol, column.DataColumn);
					column.RowIndex++;
				}
			}

			ApplyConditionalPositionFormatting (dashboard, positionRegions);
			ApplyConditionalVarianceFormatting (dashboard, varianceRegions);

			CreateDashboardLegend (dashboard, bold, font, 0, "Funds", symbols, descriptions);

			using (var stream = File.Create (document.FileName))
				workbook.Write (stream);

			return workbook;
		}

		static List<MovingAverageColumn> GetMovingAverageColumns (Document document, out bool monthly)
		{
			var movingAverageColumns = new List<MovingAverageColumn> ();
			int dataColumn = (int) DataColumn.MovingAverage;
			int rowIndex = 0;

			monthly = false;

			foreach (var movingAverage in document.MovingAverages) {
				if (movingAverage.PeriodType == MovingAveragePeriodType.Month)
					monthly = true;

				movingAverageColumns.Add (new MovingAverageColumn (movingAverage, dataColumn, rowIndex));
				rowIndex += document.Symbols.Length + 4;
				dataColumn++;
			}

			return movingAverageColumns;
		}

		static void GetDateRange (bool monthly, out DateTime start, out DateTime end)
		{
			var today = DateTime.Today.ToUniversalTime ();

			if (monthly && today.Day < DateTime.DaysInMonth (today.Year, today.Month))
				end = today.AddDays (-1 * today.Day);
			else
				end = today.AddDays (-1);

			start = end.AddYears (-4);
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
			pattern.FillBackgroundColor = Style.Green;

			rules[index++] = rule;

			// Create the "Cash" formatting rule
			rule = formatting.CreateConditionalFormattingRule (ComparisonOperator.Equal, "\"Cash\"");
			pattern = rule.CreatePatternFormatting ();
			pattern.FillPattern = FillPattern.SolidForeground;
			pattern.FillBackgroundColor = Style.Red;

			font = rule.CreateFontFormatting ();
			font.FontColorIndex = Style.White;

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
			pattern.FillBackgroundColor = Style.LightGreen;

			rules[index++] = rule;

			// Add the "Neutral" formatting
			rule = formatting.CreateConditionalFormattingRule (ComparisonOperator.Between, "-2", "2");
			pattern = rule.CreatePatternFormatting ();
			pattern.FillPattern = FillPattern.SolidForeground;
			pattern.FillBackgroundColor = Style.LightYellow;

			rules[index++] = rule;

			// Add the "Sell" formatting
			rule = formatting.CreateConditionalFormattingRule (ComparisonOperator.LessThan, "-2");
			pattern = rule.CreatePatternFormatting ();
			pattern.FillPattern = FillPattern.SolidForeground;
			pattern.FillBackgroundColor = Style.LightRed;

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
				((XSSFCellStyle) style).SetFillForegroundColor (Style.CustomLightBlue);
			else
				style.FillForegroundColor = Style.LightBlue;
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
			style.FillForegroundColor = Style.LightGrey;
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

		static void CreateDashboardTableRow (ISheet dashboard, IFont bold, IFont font, int rowIndex, string symbol, int column)
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
				((XSSFCellStyle) style).SetFillForegroundColor (Style.CustomLightBlue);
			else
				style.FillForegroundColor = Style.LightBlue;
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
			style.FillForegroundColor = Style.LightGrey;
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

			fundStyle.FillForegroundColor = Style.LightGrey;
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

		static bool IsNull (IStockData stockData, int row)
		{
			for (int i = 1; i < stockData.Columns; i++) {
				if (stockData.GetValue (row, i) == null)
					return true;
			}

			return false;
		}

		static async Task<ISheet> CreateSheetAsync (IFinancialService client, IWorkbook workbook, IFont font, List<MovingAverageColumn>movingAverageColumns, string symbol, DateTime start, DateTime end, CancellationToken cancellationToken)
		{
			var stockData = await client.GetStockDataAsync (symbol, start, end, cancellationToken).ConfigureAwait (false);
			var sheet = workbook.CreateSheet (symbol);
			var hstyle = workbook.CreateCellStyle ();
			var style = workbook.CreateCellStyle ();

			hstyle.Alignment = HorizontalAlignment.Center;
			hstyle.FillPattern = FillPattern.SolidForeground;
			hstyle.FillForegroundColor = Style.LightGrey;
			hstyle.ShrinkToFit = false;
			hstyle.SetFont (font);

			style.Alignment = HorizontalAlignment.Center;
			style.ShrinkToFit = false;
			style.SetFont (font);

			sheet.DefaultColumnWidth = 12;

			var endOfMonthRows = new List<int> ();
			int previousMonth = -1;
			var columnIndex = 0;
			var rowIndex = 1;
			ICell cell;
			IRow row;

			for (int i = sheet.LastRowNum; i > 0; i--) {
				row = sheet.GetRow (i - 1);
				sheet.RemoveRow (row);
			}

			// Add the Titles for the data columns
			row = sheet.CreateRow (0);
			while (columnIndex < stockData.Columns) {
				cell = row.CreateCell (columnIndex, CellType.String);
				cell.SetCellValue (stockData.GetHeader (columnIndex));
				cell.CellStyle = hstyle;
				columnIndex++;
			}

			// Add the Titles for the Moving Average columns
			foreach (var movingAverageColumn in movingAverageColumns) {
				//sheet.SetDefaultColumnStyle (movingAverageColumn.DataColumn, style);
				cell = row.CreateCell (movingAverageColumn.DataColumn, CellType.String);
				cell.SetCellValue (movingAverageColumn.Title);
				cell.CellStyle = hstyle;
			}

			// Read the data in reverse
			for (int i = stockData.Rows - 1; i >= 0; i--) {
				object value;

				if (IsNull (stockData, i))
					continue;

				value = stockData.GetValue (i, 0);
				if (!(value is DateTime date))
					continue;

				row = sheet.CreateRow (rowIndex++);
				columnIndex = 0;

				// Note: the first column is a DateTime value, all other values are stock price values
				cell = row.CreateCell (columnIndex++, CellType.String);
				cell.SetCellValue (date.ToString ("yyyy-MM-dd", CultureInfo.InvariantCulture));
				cell.CellStyle = style;

				if (date.Month != previousMonth) {
					endOfMonthRows.Add (rowIndex - 1);
					previousMonth = date.Month;
				}

				while (columnIndex < stockData.Columns) {
					cell = row.CreateCell (columnIndex, CellType.Numeric);
					value = stockData.GetValue (i, columnIndex);
					if (value is double number)
						cell.SetCellValue (number);
					cell.CellStyle = style;
					columnIndex++;
				}
			}

			// Populate the formulas for the Moving Average columns
			foreach (var movingAverageColumn in movingAverageColumns) {
				var movingAverage = movingAverageColumn.MovingAverage;

				switch (movingAverage.PeriodType) {
				case MovingAveragePeriodType.Day:
					SetSimpleDayMovingAverageFormulas (sheet, style, movingAverageColumn.DataColumn, rowIndex, movingAverage.Period);
					break;
				case MovingAveragePeriodType.Month:
					SetSimpleMonthMovingAverageFormuas (sheet, style, movingAverageColumn.DataColumn, endOfMonthRows, movingAverage.Period);
					break;
				}
			}

			return sheet;
		}

		static void SetSimpleDayMovingAverageFormulas (ISheet sheet, ICellStyle style, int dataColumn, int maxRowIndex, int days)
		{
			for (int i = 1; i < maxRowIndex - days; i++) {
				var row = sheet.GetRow (i);

				var cell = row.CreateCell (dataColumn, CellType.Formula);
				cell.SetCellFormula (string.Format ("AVERAGE({0}{1}:{0}{2})", (char) ('A' + DataColumn.AdjClose), i + 1, i + days + 1));
				cell.CellStyle = style;
			}
		}

		static void SetSimpleMonthMovingAverageFormuas (ISheet sheet, ICellStyle style, int dataColumn, List<int> endOfMonthRows, int months)
		{
			for (int i = 0; i < endOfMonthRows.Count - months; i++) {
				var items = new List<string> ();

				var row = sheet.GetRow (endOfMonthRows[i]);

				var cell = row.CreateCell (dataColumn, CellType.Formula);
				for (int month = 0; month < months; month++)
					items.Add (string.Format ("{0}{1}", (char) ('A' + DataColumn.AdjClose), endOfMonthRows[i + month] + 1));
				cell.SetCellFormula (string.Format ("AVERAGE({0})", string.Join (", ", items)));
				cell.CellStyle = style;
			}
		}
	}
}

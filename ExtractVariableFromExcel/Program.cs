using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExtractVariableFromExcel
{
	class Program
	{
		private static readonly List<string> TargetExtensions = new() { ".xlsx", ".xlsm", ".xltx", ".xltm" };

		static void Main(string[] args)
		{
			var files = args.ToList();

			// ファイルが指定されていない
			if (files.Count == 0)
			{
				Console.WriteLine("[エラー] ファイルの指定がありません。EXE ファイルに Excelファイルををドロップしてください。");
				Console.ReadLine();
				return;
			}

			// 検索対象の文字は何か？
			Console.WriteLine("部分一致検索を行います。検索文字を入力してください。");
			var targetStr = Console.ReadLine();

			// 検索対象文字がない
			if (string.IsNullOrWhiteSpace(targetStr))
			{
				Console.WriteLine("[エラー] 空白以外の文字列を入力してください。");
				Console.ReadLine();
				return;
			}

			var outputBook = new XLWorkbook();
			outputBook.AddWorksheet("Result");
			var outputCurrentRow = 1;

			// Excel 検索
			foreach (var file in files)
			{
				// ファイルが、Excel拡張子ではない場合はスキップ
				if (!TargetExtensions.Contains(Path.GetExtension(file))) continue;

				var results = new List<Result>();
				var book = new XLWorkbook(@$"{file}");

				// シートを指定
				foreach (var sheet in book.Worksheets)
				{
					// 行を指定
					for (var i = 1; i <= sheet.LastRowUsed().RowNumber(); i++)
					{
						// 列を指定
						for (var j = 1; j <= sheet.LastColumnUsed().ColumnNumber(); j++)
						{
							// 検索
							var cellContent = sheet.Cell(i, j).Value.ToString();
							if (!string.IsNullOrWhiteSpace(cellContent) && cellContent.Contains(targetStr))
							{
								var result = new Result
								{
									Value = cellContent,
									Horizontal = HorizontalToStr(sheet.Cell(i, j).Style.Alignment.Horizontal)
								};
								results.Add(result);
							}
						}
					}

					// 出力
					outputBook.Worksheet(1).Cell(outputCurrentRow, 1).SetValue($"ファイル名：{Path.GetFileName(file)}　　シート名:{sheet.Name}");
					outputCurrentRow++;

					foreach (var result in results)
					{
						outputBook.Worksheet(1).Cell(outputCurrentRow, 1).SetValue(result.Value).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center); ;
						outputBook.Worksheet(1).Cell(outputCurrentRow, 2).SetValue("ラベル").Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
						if (result.Value.Contains("DATE")) outputBook.Worksheet(1).Cell(outputCurrentRow, 3).SetValue("年月日_西暦・和暦_年月日").Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center); ;
						if (result.Value.Contains("SEIKYUKIKAN")) outputBook.Worksheet(1).Cell(outputCurrentRow, 3).SetValue(@"終了日が存在：年月日_西暦・和暦_年月日 ～ 年月日_西暦・和暦_年月日
終了日が不存在：年月日_西暦・和暦_年月日").Style.Alignment.SetVertical(XLAlignmentVerticalValues.Top);
						outputBook.Worksheet(1).Cell(outputCurrentRow, 4).SetValue(result.Horizontal).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center); ;
						outputCurrentRow++;
					}

					// 次のシートの結果書き込み時は、一行開ける
					outputCurrentRow++;
				}
			}

			// 保存
			outputBook.SaveAs(Path.Combine($@"{Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)}", "Search.xlsx"));
		}

		private static string HorizontalToStr(XLAlignmentHorizontalValues arg)
		{
			if (arg == XLAlignmentHorizontalValues.Center) return "中央揃え";
			if (arg == XLAlignmentHorizontalValues.Left) return "左揃え";
			if (arg == XLAlignmentHorizontalValues.Right) return "右揃え";

			return "左揃え";
		}
	}

	class Result
	{
		public string Value { get; set; }

		public string Horizontal { get; set; }
	}
}

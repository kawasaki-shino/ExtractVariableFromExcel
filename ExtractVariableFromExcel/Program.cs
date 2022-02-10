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

				var result = new List<string>();
				var book = new XLWorkbook(@$"{file}");

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
								result.Add(cellContent);
							}
						}
					}

					// 出力
					outputBook.Worksheet(1).Cell(outputCurrentRow, 1).SetValue($"ファイル名：{Path.GetFileName(file)}　　シート名:{sheet.Name}");
					outputCurrentRow++;

					foreach (var str in result)
					{
						outputBook.Worksheet(1).Cell(outputCurrentRow, 1).SetValue(str);
						outputCurrentRow++;
					}

					// 次のシートの結果書き込み時は、一行開ける
					outputCurrentRow++;
				}
			}

			// 保存
			outputBook.SaveAs(Path.Combine($@"{Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)}", "Search.xlsx"));
		}
	}
}

using System;
using System.Linq;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.IO;
using OfficeOpenXml;

/// <summary>
/// vbファイルのステップ数をExcelファイルに出力
/// </summary>
namespace VBCodeStepChecker {
    class Program {

        /// <summary>
        /// 保存Excelファイル名
        /// </summary>
        private const string OutputFileName = "result.xlsx";

        /// <summary>
        /// Sheet名
        /// </summary>
        private const string SheetName = "result";

        /// <summary>
        /// 列名
        /// </summary>
        private struct ExcelColumnName {
            public const string FileName = "ファイル名",
                                StepNumber = "ステップ数";
        }

        /// <summary>
        /// 値を設定する列番号
        /// </summary>
        private class ColumnNumber {
            public const int FileName = 1,
                             StepNumber = 2;
        }


        static void Main(string[] args) {
            //起動時のメッセージ(使い方)
            WriteExplanationMessage();

            try {
                while (true) {
                    var input = Console.ReadLine().Trim().Trim('"').ToLower();

                    if (!Path.IsPathRooted(input)) {
                        Console.WriteLine("絶対パスを入力してください");
                        continue;
                    }

                    if (Path.GetExtension(input) == ".vb") {
                        //ファイル単体の場合
                        var t = CreateVBFileValidationResultFile(input);
                        t.Wait();
                        WriteProcessEndMessage();
                    } else {
                        //ディレクトリ指定の場合
                        //末尾にバックスラッシュが付いていない場合は付与
                        if (!input.EndsWith(@"\")) { input += @"\"; }
                        var t = CreateDirectoryAllValidationResultFile(input);
                        t.Wait();

                        if (!t.Result) {
                            Console.WriteLine("検証ディレクトリに.vbファイルが存在しません");
                        }else {
                            WriteProcessEndMessage();
                        }
                    }

                    //アプリケーションを続行するかの選択
                    if (!IsProceed()) { return; }
                }
            } catch {
                WriteExceptionMessage();
            }
        }

        /// <summary>
        /// .vbファイルステップ数を取得し､Excelファイルに出力
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        static async Task CreateVBFileValidationResultFile(string input) {
            //ファイル名の取得
            var fileName = Path.GetFileName(input);
            //ステップ数取得
            var value = File.ReadAllLines(input);
            var stepNumber = value.Count(e => IsCountTarget(e));

            //既存ファイル削除
            if (File.Exists(OutputFileName)) { File.Delete(OutputFileName); }

            await Task.Run(() => {

                //Excelファイル作成
                var outputFile = new FileInfo(OutputFileName);

                using (var book = new ExcelPackage(outputFile))
                using (var sheet = book.Workbook.Worksheets.Add(SheetName)) {
                    //列名設定
                    SetColumnName(sheet.Cells);

                    //検証ファイル名設定
                    sheet.Cells[2, ColumnNumber.FileName].Value = fileName;
                    //ステップ数設定
                    sheet.Cells[2, ColumnNumber.StepNumber].Value = stepNumber;

                    //列幅自動調整
                    ColumnsAutoFit(sheet);

                    book.Save();
                }
            });
        }

        /// <summary>
        /// 指定ディレクトリ内全ての.vbファイルステップ数を取得し､Excelファイルに出力
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        static async Task<bool> CreateDirectoryAllValidationResultFile(string input) {
            //ディレクトリ内の拡張子｢.vb｣のパスを取得(.designer.vbは除外する)
            var files = Directory.GetFiles(input, "*.vb").Where(e => !e.EndsWith(".designer.vb"));

            //ファイルが取得できない場合はfalseを返す
            if (files.Count() == 0) { return false; }

            //既存ファイル削除
            if (File.Exists(OutputFileName)) { File.Delete(OutputFileName); }

            await Task.Run(() => {

                //Excelファイル作成
                var outputFile = new FileInfo(OutputFileName);

                using (var book = new ExcelPackage(outputFile))
                using (var sheet = book.Workbook.Worksheets.Add(SheetName)) {
                    //列名設定
                    SetColumnName(sheet.Cells);

                    foreach (var f in files.Select((s, i) => new { FilePath = s, RowNumber = i + 2 })) {
                        //ファイル名の取得
                        var fileName = Path.GetFileName(f.FilePath);
                        //ステップ数取得
                        var value = File.ReadAllLines(f.FilePath);
                        var stepNumber = value.Count(e => IsCountTarget(e));

                        //検証ファイル名設定
                        sheet.Cells[f.RowNumber, ColumnNumber.FileName].Value = fileName;
                        //ステップ数設定
                        sheet.Cells[f.RowNumber, ColumnNumber.StepNumber].Value = stepNumber;
                    }

                    //列幅自動調整
                    ColumnsAutoFit(sheet);

                    book.Save();
                }
            });

            return true;
        }

        /// <summary>
        /// 列名設定
        /// </summary>
        /// <param name="cells"></param>
        static void SetColumnName(ExcelRange cells) {
            cells[1, ColumnNumber.FileName].Value = ExcelColumnName.FileName;
            cells[1, ColumnNumber.StepNumber].Value = ExcelColumnName.StepNumber;
        }

        /// <summary>
        /// 値が行数カウントの対象か検証
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        static bool IsCountTarget(string value) {
            return (!Regex.IsMatch(value.Trim(), @"^[\s\'#]") && !String.IsNullOrWhiteSpace(value.Trim()));
        }

        /// <summary>
        /// 列幅自動調整
        /// </summary>
        /// <param name="sheet"></param>
        static void ColumnsAutoFit(ExcelWorksheet sheet) {
            sheet.Column(ColumnNumber.FileName).AutoFit();
            sheet.Column(ColumnNumber.StepNumber).AutoFit();
        }
        
        /// <summary>
        /// 処理が終了した場合､続行するか選択
        /// </summary>
        /// <returns></returns>
        static bool IsProceed() {
            Console.WriteLine("続けて処理を行いますか? Y/N");
            while (true) {
                var response = Console.ReadLine().ToUpper();
                switch (response) {
                    case "Y":
                        return true;
                    case "N":
                        return false;
                }
            }
        }

        #region メッセージ関数

        /// <summary>
        /// 使い方
        /// </summary>
        static void WriteExplanationMessage() {
            Console.WriteLine("検証を行うvbファイル､ディレクトリをドラッグアンドドロップしてください");
        }

        /// <summary>
        /// ファイル出力完了メッセージ
        /// </summary>
        static void WriteProcessEndMessage() {
            Console.WriteLine("ファイルの出力が完了しました");
        }

        /// <summary>
        /// 例外発生時のメッセージ
        /// </summary>
        static void WriteExceptionMessage() {
            Console.WriteLine("例外が発生しました{0}アプリケーションを終了します", Environment.NewLine);
            Console.ReadLine();
        }

        #endregion
    }
}

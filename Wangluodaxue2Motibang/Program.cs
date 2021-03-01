
using ClosedXML.Excel;
using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;

namespace Wangluodaxue2Motibang {
    class Program {
        static readonly bool IsAll = true;
        static readonly string PatternTypeSingleSelect = "单选题";
        static readonly string PatternTypeMutipleSelect = "多选题";
        static readonly string PatternTypeJudgement = "判断题";
        static string PatternLevel { get; set; }
        static bool IsPatternLevel(string level) {
            if (!string.IsNullOrEmpty(level)) {
                return level.Equals(PatternLevel);
            } else {
                return false;
            }
        }
        static bool IsPatternTypeSelect(string type) {
            if (!string.IsNullOrEmpty(type)) {
                return type.Equals(PatternTypeSingleSelect) || type.Equals(PatternTypeMutipleSelect);
            } else {
                return false;
            }
        }
        static bool IsPatternTypeJudgement(string type) {
            if (!string.IsNullOrEmpty(type)) {
                return type.Equals(PatternTypeJudgement);
            } else {
                return false;
            }
        }
        static void Main(string[] args) {
            string Help = @"
备份：
w2m 全部/初级工/中级工/高级工/技师 d:/the/path/to/网络大学题库.xlsx d:/the/path/to/磨题帮题库.xlsx
";
            if (args.Length != 3) {
                Console.WriteLine(Help);
                return;
            }
            PatternLevel = args[0];
            var InputFile = args[1];
            var OutputFile = args[2];
            
            string FileName = $"{OutputFile}_{PatternLevel}";
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (var workbook = new XLWorkbook()) {
                var worksheet = workbook.Worksheets.Add("题库");
                var currentRow = 1;
                worksheet.Cell(currentRow, 1).Value = "标题";
                worksheet.Range("B1:H1").Row(1).Merge();
                worksheet.Cell(currentRow, 2).Value = FileName;
                currentRow++;
                worksheet.Cell(currentRow, 1).Value = "描述";
                worksheet.Cell(currentRow, 2).Value = FileName;
                worksheet.Range("B2:H2").Row(1).Merge();
                currentRow++;
                worksheet.Cell(currentRow, 1).Value = "用时";
                worksheet.Cell(currentRow, 2).Value = "1000";
                currentRow++;
                worksheet.Cell(currentRow, 1).Value = "题干";
                worksheet.Cell(currentRow, 2).Value = "题型";
                worksheet.Cell(currentRow, 3).Value = "选择项1";
                worksheet.Cell(currentRow, 4).Value = "选择项2";
                worksheet.Cell(currentRow, 5).Value = "选择项3";
                worksheet.Cell(currentRow, 6).Value = "选择项4";
                worksheet.Cell(currentRow, 7).Value = "解析";
                worksheet.Cell(currentRow, 8).Value = "答案";
                worksheet.Cell(currentRow, 9).Value = "得分";
                using (var stream = File.Open(InputFile, FileMode.Open, FileAccess.Read)) {
                    using (var reader = ExcelReaderFactory.CreateReader(stream)) {
                        do {
                            while (reader.Read()) {
                                try {
                                    if (IsAll || IsPatternLevel(reader.GetString(4))) {
                                        currentRow++;
                                        worksheet.Cell(currentRow, 1).Value = reader.GetString(7);
                                        //worksheet.Cell(currentRow, 2).Value = reader.GetString(6);
                                        worksheet.Cell(currentRow, 7).Value = $"{reader.GetString(10)}\n{reader.GetString(11)}";
                                        
                                        worksheet.Cell(currentRow, 9).Value = "1";
                                        if (IsPatternTypeSelect(reader.GetString(6))) {
                                            
                                            var Options = reader.GetString(8).Split("$;$");
                                            for (var i = 0; i < Options.Length; i++) {
                                                if (Options[i].Contains("/")) {
                                                    //worksheet.Cell(currentRow, i + 3).Value = string.Join("除以", Options[i].Split("/"));
                                                    worksheet.Cell(currentRow, i + 3).Value = $"\'{Options[i]}";
                                                } else {
                                                    worksheet.Cell(currentRow, i + 3).Value = Options[i];
                                                }
                                            }
                                            worksheet.Cell(currentRow, 8).Value = reader.GetString(9);
                                        } else if (IsPatternTypeJudgement(reader.GetString(6))) {
                                            worksheet.Cell(currentRow, 8).Value = reader.GetString(9).Equals("A") ? "对" : "错";
                                        }
                                    }

                                } catch (IndexOutOfRangeException e) {
                                    Console.WriteLine(e);
                                }
                            }
                        } while (reader.NextResult());
                    }
                    using (var memstream = new MemoryStream()) {
                        workbook.SaveAs(memstream);
                        var content = memstream.ToArray();
                        File.WriteAllBytes(OutputFile, content);

                    }
                }
            }
            
        }
    }
}

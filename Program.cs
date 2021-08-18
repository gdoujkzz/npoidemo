using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;

namespace 导出数据到Excel模板
{
    public class AwardInfo
    {
        public string Country { get; set; }

        public int Num { get; set; }

        public int Rank { get; set; }

        public int LastRank { get; set; }
    }
    class Program
    {
        static void Main(string[] args)
        {

            //准备数据
            var time = DateTime.Now.ToString("yyyy-MM-dd");
            var address = "中国-北京";

            var awardInfos = new List<AwardInfo>();
            awardInfos.Add(new AwardInfo() { Country = "中国", Num = 100, Rank = 1, LastRank = 1 });
            awardInfos.Add(new AwardInfo() { Country = "美国", Num = 80, Rank = 2, LastRank = 2 });
            awardInfos.Add(new AwardInfo() { Country = "俄罗斯", Num = 60, Rank = 3, LastRank = 4 });
            awardInfos.Add(new AwardInfo() { Country = "加拿大", Num = 40, Rank = 4, LastRank = 3 });
            awardInfos.Add(new AwardInfo() { Country = "巴基斯坦", Num = 20, Rank = 5, LastRank = 5 });
            awardInfos.Add(new AwardInfo() { Country = "中国台湾", Num = 10, Rank = 6, LastRank = 7 });
            awardInfos.Add(new AwardInfo() { Country = "中国香港", Num = 9, Rank = 7, LastRank = 6 });
            awardInfos.Add(new AwardInfo() { Country = "阿富汗", Num = 8, Rank = 8, LastRank = 9 });

            //把数据写入Excel
            string path = $"{System.AppDomain.CurrentDomain.BaseDirectory}\\templates\\简单的模板.xlsx";
            Console.WriteLine(path);
            try
            {
                IWorkbook workbook = null;

                using (var fs = new FileStream(path, FileMode.Open, FileAccess.ReadWrite))
                {
                    if (path.IndexOf(".xlsx") > 0) // 2007
                        workbook = new XSSFWorkbook(fs);
                    else if (path.IndexOf(".xls") > 0) // 2003
                        workbook = new HSSFWorkbook(fs);
                    if (workbook != null)
                    {
                        ISheet sheet = workbook.GetSheet("金牌统计");

                        //内容
                        ICellStyle style = workbook.CreateCellStyle();
                        style.BorderBottom = BorderStyle.Thin;
                        style.BorderLeft = BorderStyle.Thin;
                        style.BorderRight = BorderStyle.Thin;
                        style.BorderTop = BorderStyle.Thin;
                        var font = workbook.CreateFont();
                        font.FontName = "宋体";
                        font.FontHeightInPoints = 9;
                        style.SetFont(font);



                        sheet.GetRow(1).GetCell(1).SetCellValue(time);
                        sheet.GetRow(1).GetCell(3).SetCellValue(address);


                        //第四行开始创建新行
                        for (int i = 0; i < awardInfos.Count; i++)
                        {
                            var rowIndex = i + 3; //跳过首部三行
                            //现在想要的结果：要把数据插入到现在固定的表格模板中
                            //思路：每次插入数据，看看当前模板是否能够容纳得下，
                            //如果可以就直接赋值。不可用就先把当前行到最后一行整体往后移动一行。
                            if (rowIndex >= 7)
                            {
                                //向下移动一行。
                                var row4Style = sheet.GetRow(rowIndex);   //获取第四行是因为你创建的新行，
                                                                          //要进行赋值
                                sheet.ShiftRows(rowIndex, sheet.LastRowNum, 1, true, false); //从当前行到最后一行，整体后移动。
                                var rowt = sheet.CreateRow(rowIndex);
                                for (int t = 0; t < 4; t++)
                                {
                                    var tcell = rowt.CreateCell(t);
                                    tcell.CellStyle = row4Style.GetCell(0).CellStyle;
                                }
                            }

                            var row5 = sheet.GetRow(rowIndex);
                            var cell50 = row5.GetCell(0);
                            cell50.CellStyle = style;
                            cell50.SetCellValue(awardInfos[i].Country.ToString());
                            var cell51 = row5.GetCell(1);
                            cell51.CellStyle = style;
                            cell51.SetCellValue(awardInfos[i].Num.ToString());
                            var cell52 = row5.GetCell(2);
                            cell52.CellStyle = style;
                            cell52.SetCellValue(awardInfos[i].Rank.ToString());
                            var cell53 = row5.GetCell(3);
                            cell53.CellStyle = style;
                            cell53.SetCellValue(awardInfos[i].LastRank
                          .ToString());
                        }
                    }
                }
                using (FileStream fs = new FileStream("test.xlsx", FileMode.OpenOrCreate))
                {
                    workbook.Write(fs);
                    workbook.Close();
                    Console.Write("成功");
                    //base64string = new byte[ms.Length];
                    //ms.Position = 0;
                    //ms.Read(base64string, 0, base64string.Length);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }
    }
}

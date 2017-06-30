using DocumentFormat.OpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using Word;

namespace Test
{
    class Program
    {
        public enum eInfluence
        {
            Bad,
            None,
            Good
        }
        static void Main(string[] args)
        {
            Docments doc = new Docments(@"C:\Users\Zhang\Documents\test_word.docx");
            // doc.AddStyle();
            // doc.AddStylesPartToPackage();
            //  doc.CreateParagraphStyle(ParagraphStyle.Heading1.ToString(), ParagraphStyle.Heading1.ToString());
            //Word.ParaAndFontFormat formatCollection = new Word.ParaAndFontFormat();
            //formatCollection.SetFontFormat(20, color: Color.Blue, italic: true);
            //doc.AddParagraph("Hello World", formatCollection.formatCollection);
            //formatCollection.SetFontFormat(50, true, highlightColor: HighlightColor.Blue);
            //formatCollection.SetParaFormat(paragraphStyle: ParagraphStyle.Heading2);
            //doc.AddParagraph("this is a test para", formatCollection.formatCollection);
            //doc.AddParagraph();
            //formatCollection.SetFontFormat(20);
            //formatCollection.SetParaFormat(justificationValues: JustificationValues.Center);
            //formatCollection.SetFontFormat(30, true, true, color: Color.AliceBlue);
            ////  formatCollection.SetParaFormat(paragraphStyle: ParagraphStyle.Heading1);
            //doc.AddParagraph("this is a new para", formatCollection.formatCollection);

            //doc.AddText("123");

            //var table = doc.AddTable(3, 3);
            ////table.MergeCell(1, 2, 1, 3);
            ////table.MergeCell(2, 1, 2, 2);
            //table.MergeCell(2, 1, 3, 1);
            //table.CellText(1, 1, "test", JustificationValues.Center);

            //table.SetCellStyle(1, 1, Color.Azure);

            //table.SetRowStyle(3, Color.Cyan);



            // Word.ParaAndFontFormat fc = new Word.ParaAndFontFormat();
            // fc.SetFontFormat(fontSize: 22, fontName: "华文中宋", color: Color.FromArgb(68, 84, 106));
            // fc.SetParaFormat(justificationValues: JustificationValues.Center);
            // doc.AddBlankLine(3);
            // doc.AddParagraph("中小学生学业诊断分析系统 ", fc.formatCollection);
            // doc.AddParagraph("学业支持子系统 ", fc.formatCollection);
            // doc.AddParagraph("个体测评报告", fc.formatCollection);
            // doc.AddBlankLine(22);
            // fc.SetFontFormat(fontSize: 16, fontName: "中宋", underlineValues: UnderlineValues.Single);
            // doc.AddParagraph($"学校： 	1232435", fc.formatCollection);
            // doc.AddBlankLine(2);
            // doc.AddParagraph($"姓名： 	1232435", fc.formatCollection);
            // doc.AddBlackPage();
            // doc.AddParagraph("this is a new page");

            //// doc.PasteFrom(@"D:\bspublish\App_Data\Measurements\1302\师生关系量表描述文件.docx");
            // List<ChartSubArea> chartAreas = new List<ChartSubArea>();
            // chartAreas.Add(new ChartSubArea() { Color = SchemeColorValues.Accent1, Label = "1st Qtr", Value = "8.2" });
            // chartAreas.Add(new ChartSubArea() { Color = SchemeColorValues.Accent2, Label = "2st Qtr", Value = "3.2" });
            // chartAreas.Add(new ChartSubArea() { Color = SchemeColorValues.Accent3, Label = "3st Qtr", Value = "1.4" });
            // chartAreas.Add(new ChartSubArea() { Color = SchemeColorValues.Accent4, Label = "4st Qtr", Value = "1.2" });
            // //var chart1=doc.AddChart("chart1","rId12");
            // //chart1.AddNewBarAndLineChart(chartAreas,10,5);
            // doc.AddBlackPage();

            // doc.PasteFrom(@"D:\bspublish\App_Data\Tasks\2017060301001\20170601系统描述文件.docx");

            //量表<维度<评估，登记，影响>>
            var detaildata = new Dictionary<string, Dictionary<string,  ValueTuple<string, string, eInfluence>>>();
            var tem11 = new Dictionary<string, ValueTuple<string, string, eInfluence>>();
            tem11.Add("冲突性", ("123", "差", eInfluence.Bad));
            var tem2 = new Dictionary<string, ValueTuple<string, string, eInfluence>>();
            tem2.Add("回避性", ("123", "差", eInfluence.Bad));
            //var tem3 = new Dictionary<string, ValueTuple<string, string, eInfluence>>();
            //tem3.Add("亲密性", ("123", "差", eInfluence.Bad));
            //var tem4 = new Dictionary<string, ValueTuple<string, string, eInfluence>>();
            //tem4.Add("依恋性", ("123", "差", eInfluence.Bad));
            detaildata.Add("test1", tem11);
            detaildata.Add("test2", tem2);
            //detaildata.Add("test3", tem3);
            //detaildata.Add("test4", tem4);

            var DetailData = detaildata;

            int rowNum = 1, colNum = 12;
            foreach (var tem in DetailData)
            {
                rowNum = rowNum + tem.Value.Values.Count();
            }
            int maxRowNum = (rowNum / 2) + (rowNum % 2);
            var table = doc.AddTable(maxRowNum, colNum);

           // var table = doc.AddTable(3, 12);
            //   table.AddTableBorder(Color.Red, Color.Red);
            table.SetRowStyle(1, System.Drawing.Color.PaleVioletRed);
            table.MergeCell(1, 7, 1, 10);
            table.MergeCell(1, 1, 1, 4);
            table.SetCellStyle(1, 4, bold: true);
            table.SetCellStyle(1, 1, bold: true);
            table.CellText(1, 1, "维度名称", JustificationValues.Center);
            table.CellText(1, 5, "状态", JustificationValues.Center);
            table.CellText(1, 6, "是否需要改善", JustificationValues.Center);
            table.CellText(1, 7, "维度名称", JustificationValues.Center);
            table.CellText(1, 11, "状态", JustificationValues.Center);
            table.CellText(1, 12, "是否需要改善", JustificationValues.Center);
            int temRowNum = 2, temColNum = 1;
            bool isSecendCol = false;
            foreach (var tem in DetailData)
            {
                again: if (!isSecendCol)
                {
                    int mergeRow = tem.Value.Count();
                    table.MergeCell(temRowNum, 1, mergeRow, 1);
                    //  table.SetCellStyle(temRowNum, 1);
                    table.CellText(temRowNum, 1, tem.Key);

                    foreach (var tem1 in tem.Value)
                    {
                        table.MergeCell(temRowNum, 2, temRowNum, 4);
                        table.CellText(temRowNum, 2, tem1.Key);
                        table.CellText(temRowNum, 5, tem1.Value.Item2);
                        if (tem1.Value.Item3 == eInfluence.Bad)
                        {
                            table.CellText(temRowNum, 6, "😭");
                        }
                        temRowNum++;
                        if (temRowNum > maxRowNum)
                        {
                            isSecendCol = true;
                            temRowNum = 2;
                            goto again;
                        }

                    }
                }
                else
                {
                    int mergeRow = tem.Value.Count();
                    int tem_num = 1;
                    table.MergeCell(temRowNum, 7, mergeRow, 7);
                    table.SetCellStyle(temRowNum, 1);
                    table.CellText(temRowNum, 7, tem.Key);
                    foreach (var tem1 in tem.Value)
                    {
                        if (tem_num < maxRowNum)
                        {
                            tem_num++;
                            continue;
                        }
                        table.MergeCell(temRowNum, 8, temRowNum, 10);
                        table.CellText(temRowNum, 8, tem1.Key);
                        table.CellText(temRowNum, 11, tem1.Value.Item2);
                        if (tem1.Value.Item3 == eInfluence.Bad)
                        {
                            table.CellText(temRowNum, 12, "😭", JustificationValues.Center);
                        }
                        temRowNum++;
                        if (temRowNum > maxRowNum)
                        {
                            // isSecendCol = true;
                            break;
                        }
                    }
                }
            }
            doc.Close();





            Console.WriteLine();
            System.Diagnostics.Process.Start(@"C:\Users\Zhang\Documents\test_word.docx");
            Console.Read();
        }
      
    }
}

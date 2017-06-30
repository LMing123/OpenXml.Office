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

            var table = doc.AddTable(3, 3);
            //table.MergeCell(1, 2, 1, 3);
            //table.MergeCell(2, 1, 2, 2);
            table.MergeCell(2, 1, 3, 1);
            table.CellText(1, 1, "test", JustificationValues.Center);

            table.SetCellStyle(1, 1, Color.Azure);

            table.SetRowStyle(3, Color.Cyan);



            Word.ParaAndFontFormat fc = new Word.ParaAndFontFormat();
            fc.SetFontFormat(fontSize: 22, fontName: "华文中宋", color: Color.FromArgb(68, 84, 106));
            fc.SetParaFormat(justificationValues: JustificationValues.Center);
            doc.AddBlankLine(3);
            doc.AddParagraph("中小学生学业诊断分析系统 ", fc.formatCollection);
            doc.AddParagraph("学业支持子系统 ", fc.formatCollection);
            doc.AddParagraph("个体测评报告", fc.formatCollection);
            doc.AddBlankLine(22);
            fc.SetFontFormat(fontSize: 16, fontName: "中宋", underlineValues: UnderlineValues.Single);
            doc.AddParagraph($"学校： 	1232435", fc.formatCollection);
            doc.AddBlankLine(2);
            doc.AddParagraph($"姓名： 	1232435", fc.formatCollection);
            doc.AddBlackPage();
            doc.AddParagraph("this is a new page");

           // doc.PasteFrom(@"D:\bspublish\App_Data\Measurements\1302\师生关系量表描述文件.docx");
            List<ChartSubArea> chartAreas = new List<ChartSubArea>();
            chartAreas.Add(new ChartSubArea() { Color = SchemeColorValues.Accent1, Label = "1st Qtr", Value = "8.2" });
            chartAreas.Add(new ChartSubArea() { Color = SchemeColorValues.Accent2, Label = "2st Qtr", Value = "3.2" });
            chartAreas.Add(new ChartSubArea() { Color = SchemeColorValues.Accent3, Label = "3st Qtr", Value = "1.4" });
            chartAreas.Add(new ChartSubArea() { Color = SchemeColorValues.Accent4, Label = "4st Qtr", Value = "1.2" });
            //var chart1=doc.AddChart("chart1","rId12");
            //chart1.AddNewBarAndLineChart(chartAreas,10,5);
            doc.AddBlackPage();

           // doc.PasteFrom(@"D:\bspublish\App_Data\Tasks\2017060301001\20170601系统描述文件.docx");
            doc.Close();





            Console.WriteLine();
            System.Diagnostics.Process.Start(@"C:\Users\Zhang\Documents\test_word.docx");
            Console.Read();
        }
      
    }
}

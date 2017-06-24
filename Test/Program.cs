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
            Docments doc = new Docments(@"C:\Users\Zhang\Documents\ ","test_word.docx");

            doc.AddStylesPartToPackage();
            doc.CreateParagraphStyle(ParagraphStyle.Heading1.ToString(), ParagraphStyle.Heading1.ToString());

            Word.ParaAndFontFormat formatCollection = new Word.ParaAndFontFormat();
            formatCollection.SetFontFormat(20, color: Color.Blue, italic: true);
            doc.AddParagraph("Hello World", formatCollection.formatCollection);
            formatCollection.SetFontFormat(50, true, highlightColor: HighlightColor.Blue);
            formatCollection.SetParaFormat(paragraphStyle: ParagraphStyle.Heading2);
            doc.AddParagraph("this is a test para", formatCollection.formatCollection);
            doc.AddParagraph();
            formatCollection.SetFontFormat(30, true, true, color: Color.AliceBlue);
          //  formatCollection.SetParaFormat(paragraphStyle: ParagraphStyle.Heading1);
            doc.AddParagraph("this is a new para",formatCollection.formatCollection);
            
            doc.AddText("123");

           var table= doc.AddTable(3, 3);
            //table.MergeCell(1, 2, 1, 3);
            //table.MergeCell(2, 1, 2, 2);
            table.MergeCell(1, 1, 2, 3);


            doc.Close();





            Console.WriteLine();
            System.Diagnostics.Process.Start(@"C:\Users\Zhang\Documents\test_word.docx");
            Console.Read();
        }
      
    }
}

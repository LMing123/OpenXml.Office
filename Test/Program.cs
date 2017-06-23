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
            Docments doc = new Docments(@"d:\", "test_word.docx");

            doc.AddStylesPartToPackage();
            doc.CreateParagraphStyle(ParagraphStyle.Heading1.ToString(), ParagraphStyle.Heading1.ToString());

            Word.ParaAndFontFormat formatCollection = new Word.ParaAndFontFormat();
            formatCollection.SetFontFormat(20, color: Color.Blue, italic: true);
            doc.AddParagraph("Hello World", formatCollection.formatCollection);
            formatCollection.SetFontFormat(50, true, highlightColor: HighlightColor.Blue);
            formatCollection.SetParaFormat(paragraphStyle: ParagraphStyle.Heading2);
            doc.AddParagraph("this is a test para", formatCollection.formatCollection);
            doc.AddParagraph();
            
            doc.AddText("123");

           var table= doc.AddTable(2, 3);
            table.MergeCell(1, 1, 1, 1);
            

            doc.Close();





            Console.WriteLine();
            System.Diagnostics.Process.Start(@"d:\test_word.docx");
            Console.Read();
        }
      
    }
}

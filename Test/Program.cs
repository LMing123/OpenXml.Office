using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            Docments doc = new Docments(@"d:\", "test_word.docx");
            Word.ParaAndFontFormat fc = new Word.ParaAndFontFormat();
            fc.SetFontSize("20");
            doc.AddParagraph("Hello World", fc.fc);
            fc.SetFontSize("50");
            doc.AddParagraph("this is a test para",fc.fc);
            doc.AddParagraph();
            doc.AddText("123");

            doc.Close();
            Console.WriteLine();

            Console.Read();
        }
    }
}

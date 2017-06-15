using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            Docments doc = new Docments(@"d:\", "test_word.docx");
            Word.ParaAndFontFormat fc = new Word.ParaAndFontFormat();
            fc.SetFontFormat(20,color: Color.Blue,italic:true);
            doc.AddParagraph("Hello World", fc.fc);
            fc.SetFontFormat(50,true,highlightColor: HighlightColor.Blue);
            doc.AddParagraph("this is a test para",fc.fc);
            doc.AddParagraph();
            doc.AddText("123");
            doc.Close();

           

            Console.WriteLine();
            System.Diagnostics.Process.Start(@"d:\test_word.docx");
            Console.Read();
        }
    }
}

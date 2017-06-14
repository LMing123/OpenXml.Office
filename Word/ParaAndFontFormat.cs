using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Threading.Tasks;

namespace Word
{
    public class FormatCollection
    {
       public ParagraphProperties ParagraphProperties { get; set; }
       public RunProperties RunProperties { set; get; }
    }
    public class ParaAndFontFormat
    {
       public FormatCollection fc;
        public ParaAndFontFormat()
        {
            fc = new FormatCollection();
        }
        public void SetFontSize(string fontSize)
        {           
            fc.RunProperties = new RunProperties();
            fc.RunProperties.FontSize = new FontSize() { Val = fontSize };
        }
    }
}

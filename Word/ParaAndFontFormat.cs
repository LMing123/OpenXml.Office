using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Threading.Tasks;
using System.Resources;
using System.Media;

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
        public void SetFontFormat(int fontSize, bool? bold = null, bool? italic = null, System.Drawing.Color? color = null, bool? shadow = null, HighlightColor? highlightColor = null)
        {
            fc.RunProperties = new RunProperties();
            fc.RunProperties.FontSize = new FontSize() { Val = (fontSize * 2).ToString() };
            if (bold != null)
            {
                fc.RunProperties.Bold = new Bold() { Val = new DocumentFormat.OpenXml.OnOffValue(bold) };
            }
            if (color != null)
            {
                fc.RunProperties.Color = new Color() { Val = String.Format("{0:X6}", color.Value.R << 16 | color.Value.G << 8 | color.Value.B) };
            }
            if (shadow != null)
            {
                fc.RunProperties.Shadow = new Shadow() { Val = new DocumentFormat.OpenXml.OnOffValue(shadow) };
            }
            if (highlightColor != null)
            {
                fc.RunProperties.Highlight = new Highlight() { Val = (HighlightColorValues)highlightColor };
            }
            if (italic != null)
            {
                fc.RunProperties.Italic = new Italic() { Val = new DocumentFormat.OpenXml.OnOffValue(italic) };
            }
        }
        public void SetParaFormat(int? firstLineChars = null, JustificationValues? justificationValues = null, ParagraphStyle? paragraphStyle = null)
        {
            fc.ParagraphProperties = new ParagraphProperties();
            if (firstLineChars != null)
            {
                fc.ParagraphProperties.Indentation = new Indentation() { FirstLineChars = firstLineChars * 100 };//首行缩进
            }
            if (justificationValues != null)
            {
                fc.ParagraphProperties.Justification = new Justification() { Val = (DocumentFormat.OpenXml.Wordprocessing.JustificationValues)justificationValues };
            }
            if (paragraphStyle != null)
            {
                
                fc.ParagraphProperties.ParagraphStyleId = new ParagraphStyleId() { Val = paragraphStyle.ToString() };
                fc.ParagraphProperties.OutlineLevel = new OutlineLevel() { Val = (int)paragraphStyle.Value-1 };
            }
        }
    }
}

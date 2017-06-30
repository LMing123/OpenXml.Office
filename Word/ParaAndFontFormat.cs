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
        public FormatCollection formatCollection;
        public ParaAndFontFormat()
        {
            formatCollection = new FormatCollection();
        }
        public void SetFontFormat(int fontSize,string fontName=null, bool? bold = null, bool? italic = null, System.Drawing.Color? color = null, bool? shadow = null, HighlightColor? highlightColor = null,UnderlineValues?underlineValues=null)
        {
            formatCollection.RunProperties = new RunProperties();
            formatCollection.RunProperties.FontSize = new FontSize() { Val = (fontSize * 2).ToString() };
            if(fontName!=null)
            {
                formatCollection.RunProperties.RunFonts = new RunFonts() { EastAsia = fontName };
            }
            if (bold != null)
            {
                formatCollection.RunProperties.Bold = new Bold() { Val = new DocumentFormat.OpenXml.OnOffValue(bold) };
            }
            if (color != null)
            {
                formatCollection.RunProperties.Color = new Color() { Val = String.Format("{0:X6}", color.Value.R << 16 | color.Value.G << 8 | color.Value.B) };
            }
            if (shadow != null)
            {
                formatCollection.RunProperties.Shadow = new Shadow() { Val = new DocumentFormat.OpenXml.OnOffValue(shadow) };
            }
            if (highlightColor != null)
            {
                formatCollection.RunProperties.Highlight = new Highlight() { Val = (HighlightColorValues)highlightColor };
            }
            if (italic != null)
            {
                formatCollection.RunProperties.Italic = new Italic() { Val = new DocumentFormat.OpenXml.OnOffValue(italic) };
            }
            if(underlineValues!=null)
            {
                formatCollection.RunProperties.Underline = new Underline() { Val =(DocumentFormat.OpenXml.Wordprocessing.UnderlineValues) underlineValues };

            }

        }
        /// <summary>
        /// 设置段落格式
        /// </summary>
        /// <param name="firstLineChars">首行缩进</param>
        /// <param name="justificationValues">对齐方式</param>
        /// <param name="outlineLevel">大纲级别</param>        /// 
        /// <param name="paragraphStyle">段落风格</param>
        public void SetParaFormat(int? firstLineChars = null, JustificationValues? justificationValues = null, int? outlineLevel = null, ParagraphStyle? paragraphStyle = null)
        {
            formatCollection.ParagraphProperties = new ParagraphProperties();
            if (firstLineChars != null)
            {
                formatCollection.ParagraphProperties.Indentation = new Indentation() { FirstLineChars = firstLineChars * 100 };//首行缩进
            }
            if (justificationValues != null)
            {
                formatCollection.ParagraphProperties.Justification = new Justification() { Val = (DocumentFormat.OpenXml.Wordprocessing.JustificationValues)justificationValues };
            }
            if (paragraphStyle != null)
            {

                formatCollection.ParagraphProperties.ParagraphStyleId = new ParagraphStyleId() { Val = paragraphStyle.ToString() };
                formatCollection.ParagraphProperties.OutlineLevel = new OutlineLevel() { Val = (int)paragraphStyle.Value - 1 };
            }
            if (outlineLevel != null)
            {
                formatCollection.ParagraphProperties.OutlineLevel = new OutlineLevel() { Val = outlineLevel };

            }
          //  formatCollection.ParagraphProperties.Shading=new Shading() { }

        }
    }
}

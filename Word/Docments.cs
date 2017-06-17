using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat;
using DocumentFormat.OpenXml;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Xml.Linq;
using System.Xml;

namespace Word
{
    public class Docments
    {
        WordprocessingDocument doc;
        MainDocumentPart mainPart;
        Body body;

        public Docments(string path, string docName)
        {
            if (!File.Exists(path))
            {
                doc = WordprocessingDocument.Create(path + @"\" + docName, WordprocessingDocumentType.Document);
                mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document();
                body = mainPart.Document.AppendChild(new Body());
            }
            else
            {
                doc = WordprocessingDocument.Open(path + @"\" + docName, true);
                if (doc.MainDocumentPart == null)
                {
                    mainPart = doc.AddMainDocumentPart();
                }
                else
                {
                    mainPart = doc.MainDocumentPart;
                }
                if (mainPart.Document == null)
                {
                    mainPart.Document = new Document();
                }
                else
                {
                    mainPart.Document = new Document();
                }
                if (mainPart.Document.Body == null)
                {
                    body = mainPart.Document.AppendChild(new Body());
                }
                else
                {
                    body = mainPart.Document.Body;
                }
            }

        }
        public void AddParagraph()
        {

            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());

            run.AppendChild(new Text("\n"));

        }

        public void AddParagraph(string text, FormatCollection fc = null)
        {

            Paragraph para = body.AppendChild(new Paragraph());
            para.ParagraphProperties = new ParagraphProperties();
            Run run = para.AppendChild(new Run());
            if (fc != null)
            {
                para.ParagraphProperties = fc.ParagraphProperties;
                run.RunProperties = fc.RunProperties;
            }


            run.AppendChild(new Text(text));

        }

        public void AddText(string text)
        {
            var m = body.ChildElements.Where(i => i.LocalName == "p");
            m.Last().AppendChild(new Run(new Text(text)));
        }

        public void CreateParagraphStyle(string styleID, string styleName, string aliases = "")
        {
            CreateAndAddParagraphStyle.CreateParagraphStyle(mainPart.StyleDefinitionsPart, styleID, styleName, aliases);
        }

        public StyleDefinitionsPart AddStylesPartToPackage(WordprocessingDocument doc = null)
        {
            doc = this.doc;
            return CreateAndAddParagraphStyle.AddStylesPartToPackage(doc);
        }
        public void Close()
        {
            doc.Close();
        }


    }
}

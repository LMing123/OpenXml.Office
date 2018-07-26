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
using d = DocumentFormat.OpenXml.Drawing;
using dc = DocumentFormat.OpenXml.Drawing.Charts;
using dw = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Word.tableModel;
using Word.Enum;
using Lsj.Util.Collections;

namespace Word
{
    public class Docments
    {
        WordprocessingDocument doc;
        MainDocumentPart mainPart;

        Body body;

        public Docments(string path)
        {
            if (!File.Exists(path))
            {
                doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
                mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document();
                body = mainPart.Document.AppendChild(new Body());
            }
            else
            {
                doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);

                //doc = WordprocessingDocument.Open(path, true);
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

        public void DocReadOnly(bool readonlyFlag)
        {
            if (readonlyFlag)
            {
                if (doc.MainDocumentPart.DocumentSettingsPart == null)
                {
                    doc.MainDocumentPart.AddNewPart<DocumentSettingsPart>();

                }
                var setting = doc.MainDocumentPart.DocumentSettingsPart;
                setting.Settings = new Settings();
                setting.Settings.WriteProtection = new WriteProtection() { Hash = new Base64BinaryValue() { Value = "9oN7nWkCAyEZib1RomSJTjmPpCY=" } };
            }

        }

        public void AddParagraph()
        {

            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text("\n"));

        }
        public void AddBlankLine(int row)
        {

            for (int i = 0; i < row; i++)
            {
                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());
                run.AppendChild(new Text("\n"));
            }
        }

        public void AddParagraph(string text, FormatCollection formatCollection = null)
        {

            Paragraph para = body.AppendChild(new Paragraph());
            para.ParagraphProperties = new ParagraphProperties();
            Run run = para.AppendChild(new Run());
            if (formatCollection != null)
            {
                if (formatCollection.ParagraphProperties != null)
                {
                    para.ParagraphProperties = (ParagraphProperties)formatCollection.ParagraphProperties.Clone();

                }
                if (formatCollection.RunProperties != null)
                {
                    run.RunProperties = (RunProperties)formatCollection.RunProperties.Clone();

                }
            }


            run.AppendChild(new Text(text));

        }

        public void AddText(string text)
        {
            var m = body.ChildElements.Where(i => i.LocalName == "p");
            m.Last().AppendChild(new Run(new Text(text)));
        }

        public void AddBlackPage()
        {
            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run(new Break() { Type = BreakValues.Page }));
        }

        public void PasteFrom(string path)
        {
            if (!File.Exists(path))
            {
                throw new Exception("文件不存在");
            }
            else
            {
                using (var copy_doc = WordprocessingDocument.Open(path, true))
                {
                    var content = copy_doc.MainDocumentPart.Document.Body.CloneNode(true);
                    if (doc.MainDocumentPart.StyleDefinitionsPart == null)
                    {
                        doc.MainDocumentPart.AddPart<StyleDefinitionsPart>(copy_doc.MainDocumentPart.StyleDefinitionsPart);
                    }
                    else
                    {
                        var stylePart = doc.MainDocumentPart.StyleDefinitionsPart;
                        stylePart.Styles.Append(copy_doc.MainDocumentPart.StyleDefinitionsPart.Styles.CloneNode(true));
                    }


                    //复制图片
                    foreach (var item in copy_doc.MainDocumentPart.ImageParts)
                    {
                        var tem = item.HyperlinkRelationships;
                        var type = item.ContentType;
                        var id = copy_doc.MainDocumentPart.GetIdOfPart(item);
                        var pic = doc.MainDocumentPart.AddImagePart(type, id);
                        pic.FeedData(item.GetStream());
                    }

                    ///不知道为嘛冲突反正去了sectPr就好了~~
                    ///下面那个是不取name为sectPr的
                    for(int i=0;i<content.Count();i++)
                    {
                        if (content.ElementAt(i).LocalName == "p")
                        {
                            for (int j = 0; j < content.ElementAt(i).Count(); j++)
                            {
                                if (content.ElementAt(i).ElementAt(j).LocalName == "bookmarkStart" || content.ElementAt(i).ElementAt(j).LocalName == "bookmarkEnd")
                                {
                                    content.ElementAt(i).ElementAt(j).Remove();
                                }
                            }
                        }
                    }
                    for (int i = 0; i < content.Count(); i++)
                    {
                        if (content.ElementAt(i).LocalName == "sectPr" || content.ElementAt(i).LocalName == "bookmarkStart" || content.ElementAt(i).LocalName == "bookmarkEnd")
                        {
                            content.ElementAt(i).Remove();
                        }

                    }
                    // var openXmlElement = content.Where(n => n.LocalName == "p");


                    body.Append(content.Select(t => t.CloneNode(true)));
                    copy_doc.Close();
                }


            }
        }

        public Table AddTable(int row, int column)
        {
            DocumentFormat.OpenXml.Wordprocessing.Table table = new DocumentFormat.OpenXml.Wordprocessing.Table();
            body.AppendChild(table);
            return new Table(table, row, column);
        }

        public Chart AddChart(string chartName, string rid)
        {

            ChartPart chartPart = mainPart.AddNewPart<ChartPart>(rid);
            //EmbeddedPackagePart embeddedPackagePart1 = chartPart.AddNewPart<EmbeddedPackagePart>("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "rId3");
            //GenerateEmbeddedPackagePart1Content(embeddedPackagePart1);
            Paragraph paragraph = new Paragraph() { RsidParagraphAddition = "00C75AEB", RsidRunAdditionDefault = "000F3EFF" };

            // Create a new run that has an inline drawing object
            Run run = new Run();
            Drawing drawing = new Drawing();

            dw.Inline inline = new dw.Inline();
            inline.Append(new dw.Extent() { Cx = 5274310L, Cy = 3076575L });
            inline.Append(new dw.EffectExtent() { LeftEdge = 0, TopEdge = 0, RightEdge = 2540, BottomEdge = 9525 });
            dw.DocProperties docPros = new dw.DocProperties() { Id = 6666, Name = chartName };
            inline.Append(docPros);
            inline.Append(new dw.NonVisualGraphicFrameDrawingProperties());

            d.Graphic g = new d.Graphic();
            d.GraphicData graphicData = new d.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" };
            dc.ChartReference chartReference = new dc.ChartReference() { Id = rid };
            graphicData.Append(chartReference);

            g.Append(graphicData);
            inline.Append(g);
            drawing.Append(inline);
            run.Append(drawing);
            paragraph.Append(run);

            body.AppendChild(paragraph);
            return new Chart(chartPart);
        }

        public void AddSummaryTable(SafeDictionary<string, SafeDictionary<string, (string, string, Word.Enum.eInfluence, double)>> content)
        {
            GeneratedClass gc = new GeneratedClass();

            body.AppendChild(gc.GenerateSummaryTable(content));
        }

        public void Addtable2(string title, string evaluate, SafeDictionary<string, (string, string, eInfluence, double)> content)
        {
            GeneratedClass gc = new GeneratedClass();

            body.AppendChild(gc.GenerateTable2(title, evaluate, content));
        }
        public void Addtable3(string title, string evaluate, SafeDictionary<string, (string, string, eInfluence, double)> content)
        {
            GeneratedClass gc = new GeneratedClass();

            body.AppendChild(gc.GenerateTable3(title, evaluate, content));
        }
        public void Addtable4(string title, string evaluate, SafeDictionary<string, (string, string, eInfluence, double)> content)
        {
            GeneratedClass gc = new GeneratedClass();
            body.AppendChild(gc.GenerateTable4(title, evaluate, content));
        }

        public void Addtable5(string title, string evaluate, SafeDictionary<string, (string, string, eInfluence, double)> content)
        {
            GeneratedClass gc = new GeneratedClass();

            body.AppendChild(gc.GenerateTable5(title, evaluate, content));
        }

        public void AddStyle()
        {
            StyleDefinitionsPart styleDefinitionsPart1 = mainPart.AddNewPart<StyleDefinitionsPart>("rId1");
            CreateAndAddParagraphStyle.AddStyle(styleDefinitionsPart1);

        }

        #region Binary Data
        private string embeddedPackagePart1Data = "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";

        private System.IO.Stream GetBinaryDataStream(string base64String)
        {
            return new System.IO.MemoryStream(System.Convert.FromBase64String(base64String));
        }
        private void GenerateEmbeddedPackagePart1Content(EmbeddedPackagePart embeddedPackagePart1)
        {
            System.IO.Stream data = GetBinaryDataStream(embeddedPackagePart1Data);
            embeddedPackagePart1.FeedData(data);
            data.Close();
        }


        #endregion



        /// <summary>
        /// 暂未完工版本
        /// </summary>
        /// <param name="styleID"></param>
        /// <param name="styleName"></param>
        /// <param name="aliases"></param>
        //TODO 创建自定义段落风格未完成
        public void CreateParagraphStyle(string styleID, string styleName, string aliases = "")
        {
            CreateAndAddParagraphStyle.CreateParagraphStyle(mainPart.StyleDefinitionsPart, styleID, styleName, aliases);
        }
        /// <summary>
        /// 暂未完工版本
        /// </summary>
        /// <param name="doc"></param>
        /// <returns>StyleDefinitionsPart</returns>
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

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Word.tableModel
{
    using DocumentFormat.OpenXml.Wordprocessing;
    using DocumentFormat.OpenXml;
    using Word.Enum;
    public partial class GeneratedClass
    {
        public Table GenerateTable3(Dictionary<string, Dictionary<string, ValueTuple<string, string, eInfluence>>> content)
        {
            Table table1 = new Table();

            TableProperties tableProperties1 = new TableProperties();
            TableWidth tableWidth1 = new TableWidth() { Width = "10589", Type = TableWidthUnitValues.Dxa };
            TableIndentation tableIndentation1 = new TableIndentation() { Width = -1310, Type = TableWidthUnitValues.Dxa };

            TableBorders tableBorders1 = new TableBorders();
            TopBorder topBorder1 = new TopBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder1 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder1 = new InsideVerticalBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableBorders1.Append(topBorder1);
            tableBorders1.Append(leftBorder1);
            tableBorders1.Append(bottomBorder1);
            tableBorders1.Append(rightBorder1);
            tableBorders1.Append(insideHorizontalBorder1);
            tableBorders1.Append(insideVerticalBorder1);
            TableLayout tableLayout1 = new TableLayout() { Type = TableLayoutValues.Fixed };

            TableCellMarginDefault tableCellMarginDefault1 = new TableCellMarginDefault();
            TableCellLeftMargin tableCellLeftMargin1 = new TableCellLeftMargin() { Width = 0, Type = TableWidthValues.Dxa };
            TableCellRightMargin tableCellRightMargin1 = new TableCellRightMargin() { Width = 0, Type = TableWidthValues.Dxa };

            tableCellMarginDefault1.Append(tableCellLeftMargin1);
            tableCellMarginDefault1.Append(tableCellRightMargin1);
            TableLook tableLook1 = new TableLook() { Val = "04A0", FirstRow = true, LastRow = false, FirstColumn = true, LastColumn = false, NoHorizontalBand = false, NoVerticalBand = true };

            tableProperties1.Append(tableWidth1);
            tableProperties1.Append(tableIndentation1);
            tableProperties1.Append(tableBorders1);
            tableProperties1.Append(tableLayout1);
            tableProperties1.Append(tableCellMarginDefault1);
            tableProperties1.Append(tableLook1);

            TableGrid tableGrid1 = new TableGrid();
            GridColumn gridColumn1 = new GridColumn() { Width = "1437" };
            GridColumn gridColumn2 = new GridColumn() { Width = "709" };
            GridColumn gridColumn3 = new GridColumn() { Width = "850" };
            GridColumn gridColumn4 = new GridColumn() { Width = "1560" };
            GridColumn gridColumn5 = new GridColumn() { Width = "6033" };

            tableGrid1.Append(gridColumn1);
            tableGrid1.Append(gridColumn2);
            tableGrid1.Append(gridColumn3);
            tableGrid1.Append(gridColumn4);
            tableGrid1.Append(gridColumn5);

            TableRow tableRow1 = new TableRow() { RsidTableRowMarkRevision = "007642DC", RsidTableRowAddition = "00B97AF7", RsidTableRowProperties = "009E6323" };

            TableRowProperties tableRowProperties1 = new TableRowProperties();
            TableRowHeight tableRowHeight1 = new TableRowHeight() { Val = (UInt32Value)226U };

            tableRowProperties1.Append(tableRowHeight1);

            TableCell tableCell1 = new TableCell();

            TableCellProperties tableCellProperties1 = new TableCellProperties();
            TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = "10589", Type = TableWidthUnitValues.Dxa };
            GridSpan gridSpan1 = new GridSpan() { Val = 5 };
            Shading shading1 = new Shading() { Val = ShadingPatternValues.Clear, Color = "000000", Fill = "D99795" };
            TableCellVerticalAlignment tableCellVerticalAlignment1 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties1.Append(tableCellWidth1);
            tableCellProperties1.Append(gridSpan1);
            tableCellProperties1.Append(shading1);
            tableCellProperties1.Append(tableCellVerticalAlignment1);

            Paragraph paragraph1 = new Paragraph() { RsidParagraphMarkRevision = "007642DC", RsidParagraphAddition = "00B97AF7", RsidParagraphProperties = "009E6323", RsidRunAdditionDefault = "00B97AF7" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            Indentation indentation1 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification1 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
            RunFonts runFonts1 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Bold bold1 = new Bold();
            Kern kern1 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties1.Append(runFonts1);
            paragraphMarkRunProperties1.Append(bold1);
            paragraphMarkRunProperties1.Append(kern1);
            paragraphMarkRunProperties1.Append(fontSize1);
            paragraphMarkRunProperties1.Append(fontSizeComplexScript1);

            paragraphProperties1.Append(indentation1);
            paragraphProperties1.Append(justification1);
            paragraphProperties1.Append(paragraphMarkRunProperties1);

            Run run1 = new Run();

            RunProperties runProperties1 = new RunProperties();
            RunFonts runFonts2 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Bold bold2 = new Bold();
            Kern kern2 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize2 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "24" };

            runProperties1.Append(runFonts2);
            runProperties1.Append(bold2);
            runProperties1.Append(kern2);
            runProperties1.Append(fontSize2);
            runProperties1.Append(fontSizeComplexScript2);
            Text text1 = new Text();
            text1.Text = "班级环境量表";

            run1.Append(runProperties1);
            run1.Append(text1);

           

          

            Run run4 = new Run() { RsidRunProperties = "00E54493", RsidRunAddition = "00E54493" };

            RunProperties runProperties4 = new RunProperties();
            RunFonts runFonts5 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Bold bold5 = new Bold();
            Color color1 = new Color() { Val = "FF0000" };
            Kern kern5 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize5 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "24" };

            runProperties4.Append(runFonts5);
            runProperties4.Append(bold5);
            runProperties4.Append(color1);
            runProperties4.Append(kern5);
            runProperties4.Append(fontSize5);
            runProperties4.Append(fontSizeComplexScript5);
            Text text4 = new Text();
            text4.Text = "【量表名称】";

            run4.Append(runProperties4);
            run4.Append(text4);

        
            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);       
            paragraph1.Append(run4);


            tableCell1.Append(tableCellProperties1);
            tableCell1.Append(paragraph1);

            tableRow1.Append(tableRowProperties1);
            tableRow1.Append(tableCell1);

            TableRow tableRow2 = new TableRow() { RsidTableRowMarkRevision = "007642DC", RsidTableRowAddition = "00B97AF7", RsidTableRowProperties = "009E6323" };

            TableRowProperties tableRowProperties2 = new TableRowProperties();
            TableRowHeight tableRowHeight2 = new TableRowHeight() { Val = (UInt32Value)226U };

            tableRowProperties2.Append(tableRowHeight2);

            TableCell tableCell2 = new TableCell();

            TableCellProperties tableCellProperties2 = new TableCellProperties();
            TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = "1437", Type = TableWidthUnitValues.Dxa };
            Shading shading2 = new Shading() { Val = ShadingPatternValues.Clear, Color = "000000", Fill = "D99795" };
            TableCellVerticalAlignment tableCellVerticalAlignment2 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties2.Append(tableCellWidth2);
            tableCellProperties2.Append(shading2);
            tableCellProperties2.Append(tableCellVerticalAlignment2);

            Paragraph paragraph2 = new Paragraph() { RsidParagraphMarkRevision = "007642DC", RsidParagraphAddition = "00B97AF7", RsidParagraphProperties = "00B97AF7", RsidRunAdditionDefault = "00B97AF7" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            Indentation indentation2 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification2 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
            RunFonts runFonts7 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Bold bold7 = new Bold();
            Kern kern7 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize7 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript7 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties2.Append(runFonts7);
            paragraphMarkRunProperties2.Append(bold7);
            paragraphMarkRunProperties2.Append(kern7);
            paragraphMarkRunProperties2.Append(fontSize7);
            paragraphMarkRunProperties2.Append(fontSizeComplexScript7);

            paragraphProperties2.Append(indentation2);
            paragraphProperties2.Append(justification2);
            paragraphProperties2.Append(paragraphMarkRunProperties2);

            Run run6 = new Run() { RsidRunProperties = "007642DC" };

            RunProperties runProperties6 = new RunProperties();
            RunFonts runFonts8 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Bold bold8 = new Bold();
            Kern kern8 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize8 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript8 = new FontSizeComplexScript() { Val = "24" };

            runProperties6.Append(runFonts8);
            runProperties6.Append(bold8);
            runProperties6.Append(kern8);
            runProperties6.Append(fontSize8);
            runProperties6.Append(fontSizeComplexScript8);
            Text text6 = new Text();
            text6.Text = "维度名称";

            run6.Append(runProperties6);
            run6.Append(text6);

            paragraph2.Append(paragraphProperties2);
            paragraph2.Append(run6);

            tableCell2.Append(tableCellProperties2);
            tableCell2.Append(paragraph2);

            TableCell tableCell3 = new TableCell();

            TableCellProperties tableCellProperties3 = new TableCellProperties();
            TableCellWidth tableCellWidth3 = new TableCellWidth() { Width = "709", Type = TableWidthUnitValues.Dxa };
            Shading shading3 = new Shading() { Val = ShadingPatternValues.Clear, Color = "000000", Fill = "D99795" };

            tableCellProperties3.Append(tableCellWidth3);
            tableCellProperties3.Append(shading3);

            Paragraph paragraph3 = new Paragraph() { RsidParagraphMarkRevision = "007642DC", RsidParagraphAddition = "00B97AF7", RsidParagraphProperties = "00B97AF7", RsidRunAdditionDefault = "00B97AF7" };

            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
            Indentation indentation3 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification3 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
            RunFonts runFonts9 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Bold bold9 = new Bold();
            Kern kern9 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize9 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript9 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties3.Append(runFonts9);
            paragraphMarkRunProperties3.Append(bold9);
            paragraphMarkRunProperties3.Append(kern9);
            paragraphMarkRunProperties3.Append(fontSize9);
            paragraphMarkRunProperties3.Append(fontSizeComplexScript9);

            paragraphProperties3.Append(indentation3);
            paragraphProperties3.Append(justification3);
            paragraphProperties3.Append(paragraphMarkRunProperties3);

            Run run7 = new Run();

            RunProperties runProperties7 = new RunProperties();
            RunFonts runFonts10 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Bold bold10 = new Bold();
            Kern kern10 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize10 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript10 = new FontSizeComplexScript() { Val = "24" };

            runProperties7.Append(runFonts10);
            runProperties7.Append(bold10);
            runProperties7.Append(kern10);
            runProperties7.Append(fontSize10);
            runProperties7.Append(fontSizeComplexScript10);
            Text text7 = new Text();
            text7.Text = "分数";

            run7.Append(runProperties7);
            run7.Append(text7);

            paragraph3.Append(paragraphProperties3);
            paragraph3.Append(run7);

            tableCell3.Append(tableCellProperties3);
            tableCell3.Append(paragraph3);

            TableCell tableCell4 = new TableCell();

            TableCellProperties tableCellProperties4 = new TableCellProperties();
            TableCellWidth tableCellWidth4 = new TableCellWidth() { Width = "850", Type = TableWidthUnitValues.Dxa };
            Shading shading4 = new Shading() { Val = ShadingPatternValues.Clear, Color = "000000", Fill = "D99795" };
            TableCellVerticalAlignment tableCellVerticalAlignment3 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties4.Append(tableCellWidth4);
            tableCellProperties4.Append(shading4);
            tableCellProperties4.Append(tableCellVerticalAlignment3);

            Paragraph paragraph4 = new Paragraph() { RsidParagraphMarkRevision = "007642DC", RsidParagraphAddition = "00B97AF7", RsidParagraphProperties = "00B97AF7", RsidRunAdditionDefault = "00B97AF7" };

            ParagraphProperties paragraphProperties4 = new ParagraphProperties();
            Indentation indentation4 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification4 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties4 = new ParagraphMarkRunProperties();
            RunFonts runFonts11 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Bold bold11 = new Bold();
            Kern kern11 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize11 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript11 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties4.Append(runFonts11);
            paragraphMarkRunProperties4.Append(bold11);
            paragraphMarkRunProperties4.Append(kern11);
            paragraphMarkRunProperties4.Append(fontSize11);
            paragraphMarkRunProperties4.Append(fontSizeComplexScript11);

            paragraphProperties4.Append(indentation4);
            paragraphProperties4.Append(justification4);
            paragraphProperties4.Append(paragraphMarkRunProperties4);

            Run run8 = new Run();

            RunProperties runProperties8 = new RunProperties();
            RunFonts runFonts12 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Bold bold12 = new Bold();
            Kern kern12 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize12 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript12 = new FontSizeComplexScript() { Val = "24" };

            runProperties8.Append(runFonts12);
            runProperties8.Append(bold12);
            runProperties8.Append(kern12);
            runProperties8.Append(fontSize12);
            runProperties8.Append(fontSizeComplexScript12);
            Text text8 = new Text();
            text8.Text = "水平";

            run8.Append(runProperties8);
            run8.Append(text8);

            paragraph4.Append(paragraphProperties4);
            paragraph4.Append(run8);

            tableCell4.Append(tableCellProperties4);
            tableCell4.Append(paragraph4);

            TableCell tableCell5 = new TableCell();

            TableCellProperties tableCellProperties5 = new TableCellProperties();
            TableCellWidth tableCellWidth5 = new TableCellWidth() { Width = "1560", Type = TableWidthUnitValues.Dxa };
            Shading shading5 = new Shading() { Val = ShadingPatternValues.Clear, Color = "000000", Fill = "D99795" };

            tableCellProperties5.Append(tableCellWidth5);
            tableCellProperties5.Append(shading5);

            Paragraph paragraph5 = new Paragraph() { RsidParagraphAddition = "00B97AF7", RsidParagraphProperties = "00B97AF7", RsidRunAdditionDefault = "00B97AF7" };

            ParagraphProperties paragraphProperties5 = new ParagraphProperties();
            Indentation indentation5 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification5 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties5 = new ParagraphMarkRunProperties();
            RunFonts runFonts13 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Bold bold13 = new Bold();
            Kern kern13 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize13 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript13 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties5.Append(runFonts13);
            paragraphMarkRunProperties5.Append(bold13);
            paragraphMarkRunProperties5.Append(kern13);
            paragraphMarkRunProperties5.Append(fontSize13);
            paragraphMarkRunProperties5.Append(fontSizeComplexScript13);

            paragraphProperties5.Append(indentation5);
            paragraphProperties5.Append(justification5);
            paragraphProperties5.Append(paragraphMarkRunProperties5);

            Run run9 = new Run();

            RunProperties runProperties9 = new RunProperties();
            RunFonts runFonts14 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Bold bold14 = new Bold();
            Kern kern14 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize14 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript14 = new FontSizeComplexScript() { Val = "24" };

            runProperties9.Append(runFonts14);
            runProperties9.Append(bold14);
            runProperties9.Append(kern14);
            runProperties9.Append(fontSize14);
            runProperties9.Append(fontSizeComplexScript14);
            Text text9 = new Text();
            text9.Text = "对学习的影响";

            run9.Append(runProperties9);
            run9.Append(text9);

            paragraph5.Append(paragraphProperties5);
            paragraph5.Append(run9);

            tableCell5.Append(tableCellProperties5);
            tableCell5.Append(paragraph5);

            TableCell tableCell6 = new TableCell();

            TableCellProperties tableCellProperties6 = new TableCellProperties();
            TableCellWidth tableCellWidth6 = new TableCellWidth() { Width = "6033", Type = TableWidthUnitValues.Dxa };
            Shading shading6 = new Shading() { Val = ShadingPatternValues.Clear, Color = "000000", Fill = "D99795" };
            TableCellVerticalAlignment tableCellVerticalAlignment4 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties6.Append(tableCellWidth6);
            tableCellProperties6.Append(shading6);
            tableCellProperties6.Append(tableCellVerticalAlignment4);

            Paragraph paragraph6 = new Paragraph() { RsidParagraphMarkRevision = "007642DC", RsidParagraphAddition = "00B97AF7", RsidParagraphProperties = "00B97AF7", RsidRunAdditionDefault = "00B97AF7" };

            ParagraphProperties paragraphProperties6 = new ParagraphProperties();
            Indentation indentation6 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification6 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties6 = new ParagraphMarkRunProperties();
            RunFonts runFonts16 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Bold bold16 = new Bold();
            Kern kern16 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize16 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript16 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties6.Append(runFonts16);
            paragraphMarkRunProperties6.Append(bold16);
            paragraphMarkRunProperties6.Append(kern16);
            paragraphMarkRunProperties6.Append(fontSize16);
            paragraphMarkRunProperties6.Append(fontSizeComplexScript16);

            paragraphProperties6.Append(indentation6);
            paragraphProperties6.Append(justification6);
            paragraphProperties6.Append(paragraphMarkRunProperties6);

            Run run11 = new Run();

            RunProperties runProperties11 = new RunProperties();
            RunFonts runFonts17 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Bold bold17 = new Bold();
            Kern kern17 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize17 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript17 = new FontSizeComplexScript() { Val = "24" };

            runProperties11.Append(runFonts17);
            runProperties11.Append(bold17);
            runProperties11.Append(kern17);
            runProperties11.Append(fontSize17);
            runProperties11.Append(fontSizeComplexScript17);
            Text text11 = new Text();
            text11.Text = "班级环境类型";

            run11.Append(runProperties11);
            run11.Append(text11);

            paragraph6.Append(paragraphProperties6);
            paragraph6.Append(run11);

            tableCell6.Append(tableCellProperties6);
            tableCell6.Append(paragraph6);

            tableRow2.Append(tableRowProperties2);
            tableRow2.Append(tableCell2);
            tableRow2.Append(tableCell3);
            tableRow2.Append(tableCell4);
            tableRow2.Append(tableCell5);
            tableRow2.Append(tableCell6);

            table1.Append(tableProperties1);
            table1.Append(tableGrid1);
            table1.Append(tableRow1);
            table1.Append(tableRow2);

            int count = 1;
            for(int i=1;i<3;i++)
            {
                TableRow tableRow3 = new TableRow() { RsidTableRowMarkRevision = "007642DC", RsidTableRowAddition = "00B97AF7", RsidTableRowProperties = "00E9225D" };

                TableRowProperties tableRowProperties3 = new TableRowProperties();
                TableRowHeight tableRowHeight3 = new TableRowHeight() { Val = (UInt32Value)1325U };

                tableRowProperties3.Append(tableRowHeight3);

                TableCell tableCell7 = new TableCell();

                TableCellProperties tableCellProperties7 = new TableCellProperties();
                TableCellWidth tableCellWidth7 = new TableCellWidth() { Width = "1437", Type = TableWidthUnitValues.Dxa };
                Shading shading7 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
                TableCellVerticalAlignment tableCellVerticalAlignment5 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

                tableCellProperties7.Append(tableCellWidth7);
                tableCellProperties7.Append(shading7);
                tableCellProperties7.Append(tableCellVerticalAlignment5);

                Paragraph paragraph7 = new Paragraph() { RsidParagraphMarkRevision = "007642DC", RsidParagraphAddition = "00B97AF7", RsidParagraphProperties = "00B97AF7", RsidRunAdditionDefault = "00B97AF7" };

                ParagraphProperties paragraphProperties7 = new ParagraphProperties();
                Indentation indentation7 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
                Justification justification7 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties7 = new ParagraphMarkRunProperties();
                RunFonts runFonts19 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
                Bold bold19 = new Bold();
                Color color3 = new Color() { Val = "000000" };
                Kern kern19 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize19 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript19 = new FontSizeComplexScript() { Val = "24" };

                paragraphMarkRunProperties7.Append(runFonts19);
                paragraphMarkRunProperties7.Append(bold19);
                paragraphMarkRunProperties7.Append(color3);
                paragraphMarkRunProperties7.Append(kern19);
                paragraphMarkRunProperties7.Append(fontSize19);
                paragraphMarkRunProperties7.Append(fontSizeComplexScript19);

                paragraphProperties7.Append(indentation7);
                paragraphProperties7.Append(justification7);
                paragraphProperties7.Append(paragraphMarkRunProperties7);

                Run run13 = new Run();

                RunProperties runProperties13 = new RunProperties();
                RunFonts runFonts20 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
                Bold bold20 = new Bold();
                Color color4 = new Color() { Val = "000000" };
                Kern kern20 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize20 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript20 = new FontSizeComplexScript() { Val = "24" };

                runProperties13.Append(runFonts20);
                runProperties13.Append(bold20);
                runProperties13.Append(color4);
                runProperties13.Append(kern20);
                runProperties13.Append(fontSize20);
                runProperties13.Append(fontSizeComplexScript20);
                Text text13 = new Text();
                text13.Text = "师生关系";

                run13.Append(runProperties13);
                run13.Append(text13);

                paragraph7.Append(paragraphProperties7);
                paragraph7.Append(run13);

                tableCell7.Append(tableCellProperties7);
                tableCell7.Append(paragraph7);

                TableCell tableCell8 = new TableCell();

                TableCellProperties tableCellProperties8 = new TableCellProperties();
                TableCellWidth tableCellWidth8 = new TableCellWidth() { Width = "709", Type = TableWidthUnitValues.Dxa };
                TableCellVerticalAlignment tableCellVerticalAlignment6 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

                tableCellProperties8.Append(tableCellWidth8);
                tableCellProperties8.Append(tableCellVerticalAlignment6);

                Paragraph paragraph8 = new Paragraph() { RsidParagraphAddition = "00B97AF7", RsidParagraphProperties = "00B97AF7", RsidRunAdditionDefault = "00B97AF7" };

                ParagraphProperties paragraphProperties8 = new ParagraphProperties();
                Indentation indentation8 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };

                paragraphProperties8.Append(indentation8);

                Run run15 = new Run();

                RunProperties runProperties15 = new RunProperties();
                RunFonts runFonts22 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };

                runProperties15.Append(runFonts22);
                Text text15 = new Text();
                text15.Text = "34.44";

                run15.Append(runProperties15);
                run15.Append(text15);

                paragraph8.Append(paragraphProperties8);
                paragraph8.Append(run15);

                tableCell8.Append(tableCellProperties8);
                tableCell8.Append(paragraph8);

                TableCell tableCell9 = new TableCell();

                TableCellProperties tableCellProperties9 = new TableCellProperties();
                TableCellWidth tableCellWidth9 = new TableCellWidth() { Width = "850", Type = TableWidthUnitValues.Dxa };
                Shading shading8 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
                TableCellVerticalAlignment tableCellVerticalAlignment7 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

                tableCellProperties9.Append(tableCellWidth9);
                tableCellProperties9.Append(shading8);
                tableCellProperties9.Append(tableCellVerticalAlignment7);

                Paragraph paragraph10 = new Paragraph() { RsidParagraphMarkRevision = "001F26F7", RsidParagraphAddition = "00B97AF7", RsidParagraphProperties = "00B97AF7", RsidRunAdditionDefault = "00B97AF7" };

                ParagraphProperties paragraphProperties10 = new ParagraphProperties();
                SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Auto };
                Indentation indentation10 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
                Justification justification8 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties9 = new ParagraphMarkRunProperties();
                RunFonts runFonts50 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                Kern kern23 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize23 = new FontSize() { Val = "20" };
                FontSizeComplexScript fontSizeComplexScript23 = new FontSizeComplexScript() { Val = "20" };

                paragraphMarkRunProperties9.Append(runFonts50);
                paragraphMarkRunProperties9.Append(kern23);
                paragraphMarkRunProperties9.Append(fontSize23);
                paragraphMarkRunProperties9.Append(fontSizeComplexScript23);

                paragraphProperties10.Append(spacingBetweenLines1);
                paragraphProperties10.Append(indentation10);
                paragraphProperties10.Append(justification8);
                paragraphProperties10.Append(paragraphMarkRunProperties9);

                Run run45 = new Run();

                RunProperties runProperties43 = new RunProperties();
                RunFonts runFonts51 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Segoe UI Symbol", HighAnsi = "Segoe UI Symbol", ComplexScript = "Segoe UI Symbol" };
                Kern kern24 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize24 = new FontSize() { Val = "20" };
                FontSizeComplexScript fontSizeComplexScript24 = new FontSizeComplexScript() { Val = "20" };

                runProperties43.Append(runFonts51);
                runProperties43.Append(kern24);
                runProperties43.Append(fontSize24);
                runProperties43.Append(fontSizeComplexScript24);
                Text text16 = new Text();
                text16.Text = "低水平";

                run45.Append(runProperties43);
                run45.Append(text16);

                paragraph10.Append(paragraphProperties10);
                paragraph10.Append(run45);

                Paragraph paragraph11 = new Paragraph() { RsidParagraphMarkRevision = "001B313A", RsidParagraphAddition = "00B97AF7", RsidParagraphProperties = "00B97AF7", RsidRunAdditionDefault = "00F82954" };

                ParagraphProperties paragraphProperties11 = new ParagraphProperties();
                Indentation indentation11 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };

                ParagraphMarkRunProperties paragraphMarkRunProperties10 = new ParagraphMarkRunProperties();
                RunFonts runFonts52 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
                Kern kern25 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize25 = new FontSize() { Val = "2" };
                FontSizeComplexScript fontSizeComplexScript25 = new FontSizeComplexScript() { Val = "24" };

                paragraphMarkRunProperties10.Append(runFonts52);
                paragraphMarkRunProperties10.Append(kern25);
                paragraphMarkRunProperties10.Append(fontSize25);
                paragraphMarkRunProperties10.Append(fontSizeComplexScript25);

                paragraphProperties11.Append(indentation11);
                paragraphProperties11.Append(paragraphMarkRunProperties10);

                Run run46 = new Run();

                RunProperties runProperties44 = new RunProperties();
                RunFonts runFonts53 = new RunFonts() { Ascii = "Segoe UI Symbol", HighAnsi = "Segoe UI Symbol", ComplexScript = "Segoe UI Symbol" };
                Kern kern26 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize26 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript26 = new FontSizeComplexScript() { Val = "24" };

                runProperties44.Append(runFonts53);
                runProperties44.Append(kern26);
                runProperties44.Append(fontSize26);
                runProperties44.Append(fontSizeComplexScript26);
                FieldChar fieldChar3 = new FieldChar() { FieldCharType = FieldCharValues.End };

                run46.Append(runProperties44);
                run46.Append(fieldChar3);

                paragraph11.Append(paragraphProperties11);
                paragraph11.Append(run46);

                tableCell9.Append(tableCellProperties9);
                tableCell9.Append(paragraph10);
                tableCell9.Append(paragraph11);

                TableCell tableCell10 = new TableCell();

                TableCellProperties tableCellProperties10 = new TableCellProperties();
                TableCellWidth tableCellWidth10 = new TableCellWidth() { Width = "1560", Type = TableWidthUnitValues.Dxa };

                tableCellProperties10.Append(tableCellWidth10);

                Paragraph paragraph13 = new Paragraph() { RsidParagraphMarkRevision = "001F26F7", RsidParagraphAddition = "00B97AF7", RsidParagraphProperties = "00B97AF7", RsidRunAdditionDefault = "00B97AF7" };

                ParagraphProperties paragraphProperties13 = new ParagraphProperties();
                SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Auto };
                Indentation indentation13 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
                Justification justification10 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties12 = new ParagraphMarkRunProperties();
                RunFonts runFonts81 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                Kern kern28 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize28 = new FontSize() { Val = "20" };
                FontSizeComplexScript fontSizeComplexScript28 = new FontSizeComplexScript() { Val = "20" };

                paragraphMarkRunProperties12.Append(runFonts81);
                paragraphMarkRunProperties12.Append(kern28);
                paragraphMarkRunProperties12.Append(fontSize28);
                paragraphMarkRunProperties12.Append(fontSizeComplexScript28);

                paragraphProperties13.Append(spacingBetweenLines2);
                paragraphProperties13.Append(indentation13);
                paragraphProperties13.Append(justification10);
                paragraphProperties13.Append(paragraphMarkRunProperties12);

                Run run76 = new Run() { RsidRunProperties = "001F26F7" };

                RunProperties runProperties72 = new RunProperties();
                RunFonts runFonts82 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                Kern kern29 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize29 = new FontSize() { Val = "20" };
                FontSizeComplexScript fontSizeComplexScript29 = new FontSizeComplexScript() { Val = "20" };

                runProperties72.Append(runFonts82);
                runProperties72.Append(kern29);
                runProperties72.Append(fontSize29);
                runProperties72.Append(fontSizeComplexScript29);
                Text text17 = new Text();
                text17.Text = "消极影响";

                run76.Append(runProperties72);
                run76.Append(text17);

                paragraph13.Append(paragraphProperties13);
                paragraph13.Append(run76);

                Paragraph paragraph14 = new Paragraph() { RsidParagraphMarkRevision = "007642DC", RsidParagraphAddition = "00B97AF7", RsidParagraphProperties = "00B97AF7", RsidRunAdditionDefault = "00F82954" };

                ParagraphProperties paragraphProperties14 = new ParagraphProperties();
                Indentation indentation14 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
                Justification justification11 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties13 = new ParagraphMarkRunProperties();
                RunFonts runFonts83 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
                Color color6 = new Color() { Val = "000000" };
                Kern kern30 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize30 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript30 = new FontSizeComplexScript() { Val = "24" };

                paragraphMarkRunProperties13.Append(runFonts83);
                paragraphMarkRunProperties13.Append(color6);
                paragraphMarkRunProperties13.Append(kern30);
                paragraphMarkRunProperties13.Append(fontSize30);
                paragraphMarkRunProperties13.Append(fontSizeComplexScript30);

                paragraphProperties14.Append(indentation14);
                paragraphProperties14.Append(justification11);
                paragraphProperties14.Append(paragraphMarkRunProperties13);

                Run run77 = new Run();

                RunProperties runProperties73 = new RunProperties();
                RunFonts runFonts84 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
                Color color7 = new Color() { Val = "000000" };
                Kern kern31 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize31 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript31 = new FontSizeComplexScript() { Val = "24" };

                runProperties73.Append(runFonts84);
                runProperties73.Append(color7);
                runProperties73.Append(kern31);
                runProperties73.Append(fontSize31);
                runProperties73.Append(fontSizeComplexScript31);
                FieldChar fieldChar6 = new FieldChar() { FieldCharType = FieldCharValues.End };

                run77.Append(runProperties73);
                run77.Append(fieldChar6);

                paragraph14.Append(paragraphProperties14);
                paragraph14.Append(run77);

                tableCell10.Append(tableCellProperties10);
                tableCell10.Append(paragraph13);
                tableCell10.Append(paragraph14);

                TableCell tableCell11 = new TableCell();

                TableCellProperties tableCellProperties11 = new TableCellProperties();
                TableCellWidth tableCellWidth11 = new TableCellWidth() { Width = "6033", Type = TableWidthUnitValues.Dxa };
                VerticalMerge verticalMerge1 = new VerticalMerge();
                if (count==1)
                {
                    verticalMerge1.Val = MergedCellValues.Restart;
                }
                else
                {
                    verticalMerge1.Val = MergedCellValues.Continue;
                }
                Shading shading9 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
                TableCellVerticalAlignment tableCellVerticalAlignment8 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

                tableCellProperties11.Append(tableCellWidth11);
                tableCellProperties11.Append(verticalMerge1);
                tableCellProperties11.Append(shading9);
                tableCellProperties11.Append(tableCellVerticalAlignment8);

                Paragraph paragraph15 = new Paragraph() { RsidParagraphMarkRevision = "00B97AF7", RsidParagraphAddition = "00B97AF7", RsidParagraphProperties = "00B97AF7", RsidRunAdditionDefault = "00B97AF7" };

                ParagraphProperties paragraphProperties15 = new ParagraphProperties();
                Indentation indentation15 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
                Justification justification12 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties14 = new ParagraphMarkRunProperties();
                Bold bold22 = new Bold();
                Kern kern32 = new Kern() { Val = (UInt32Value)0U };
                FontSizeComplexScript fontSizeComplexScript32 = new FontSizeComplexScript() { Val = "20" };

                paragraphMarkRunProperties14.Append(bold22);
                paragraphMarkRunProperties14.Append(kern32);
                paragraphMarkRunProperties14.Append(fontSizeComplexScript32);

                paragraphProperties15.Append(indentation15);
                paragraphProperties15.Append(justification12);
                paragraphProperties15.Append(paragraphMarkRunProperties14);

                Run run78 = new Run() { RsidRunProperties = "00B97AF7" };

                RunProperties runProperties74 = new RunProperties();
                RunFonts runFonts85 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Times New Roman" };
                Bold bold23 = new Bold();
                Color color8 = new Color() { Val = "FF0000" };

                runProperties74.Append(runFonts85);
                runProperties74.Append(bold23);
                runProperties74.Append(color8);
                Text text18 = new Text();
                text18.Text = "【量表评价】";

                run78.Append(runProperties74);
                run78.Append(text18);

                paragraph15.Append(paragraphProperties15);
                paragraph15.Append(run78);


                Paragraph paragraph16 = new Paragraph() { RsidParagraphMarkRevision = "001F26F7", RsidParagraphAddition = "00B97AF7", RsidParagraphProperties = "00B97AF7", RsidRunAdditionDefault = "00B97AF7" };

                ParagraphProperties paragraphProperties16 = new ParagraphProperties();
                SpacingBetweenLines spacingBetweenLines3 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Auto };
                Indentation indentation16 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };

                ParagraphMarkRunProperties paragraphMarkRunProperties15 = new ParagraphMarkRunProperties();
                RunFonts runFonts114 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                Kern kern33 = new Kern() { Val = (UInt32Value)0U };
                FontSizeComplexScript fontSizeComplexScript33 = new FontSizeComplexScript() { Val = "20" };

                paragraphMarkRunProperties15.Append(runFonts114);
                paragraphMarkRunProperties15.Append(kern33);
                paragraphMarkRunProperties15.Append(fontSizeComplexScript33);

                paragraphProperties16.Append(spacingBetweenLines3);
                paragraphProperties16.Append(indentation16);
                paragraphProperties16.Append(paragraphMarkRunProperties15);

                Run run109 = new Run() { RsidRunProperties = "001F26F7" };

                RunProperties runProperties105 = new RunProperties();
                RunFonts runFonts115 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                Kern kern34 = new Kern() { Val = (UInt32Value)0U };
                FontSizeComplexScript fontSizeComplexScript34 = new FontSizeComplexScript() { Val = "20" };

                runProperties105.Append(runFonts115);
                runProperties105.Append(kern34);
                runProperties105.Append(fontSizeComplexScript34);
                Text text20 = new Text();
                text20.Text = "该生认为其所在班级的类型是一般型，心理学研究发现：该类型班级学生的学业成绩明显高于其他两类班级的学生。该类型的班级人际关系较好，师生和同学之间有一定的支持与信任、宽容与互动；课堂秩序性较好；该生觉得教师布置的作业适当，适度的课业负担转化为了该生学习的动力。";

                run109.Append(runProperties105);
                run109.Append(text20);

                paragraph16.Append(paragraphProperties16);
                paragraph16.Append(run109);

                Paragraph paragraph17 = new Paragraph() { RsidParagraphMarkRevision = "00C840C5", RsidParagraphAddition = "00B97AF7", RsidParagraphProperties = "00B97AF7", RsidRunAdditionDefault = "00F82954" };

                ParagraphProperties paragraphProperties17 = new ParagraphProperties();
                Indentation indentation17 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };

                ParagraphMarkRunProperties paragraphMarkRunProperties16 = new ParagraphMarkRunProperties();
                RunFonts runFonts116 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
                Color color10 = new Color() { Val = "000000" };
                Kern kern35 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize32 = new FontSize() { Val = "2" };
                FontSizeComplexScript fontSizeComplexScript35 = new FontSizeComplexScript() { Val = "24" };

                paragraphMarkRunProperties16.Append(runFonts116);
                paragraphMarkRunProperties16.Append(color10);
                paragraphMarkRunProperties16.Append(kern35);
                paragraphMarkRunProperties16.Append(fontSize32);
                paragraphMarkRunProperties16.Append(fontSizeComplexScript35);

                paragraphProperties17.Append(indentation17);
                paragraphProperties17.Append(paragraphMarkRunProperties16);

                Run run110 = new Run() { RsidRunProperties = "006605C4" };

                RunProperties runProperties106 = new RunProperties();
                RunFonts runFonts117 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
                Color color11 = new Color() { Val = "000000" };
                Kern kern36 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize33 = new FontSize() { Val = "40" };
                FontSizeComplexScript fontSizeComplexScript36 = new FontSizeComplexScript() { Val = "24" };

                runProperties106.Append(runFonts117);
                runProperties106.Append(color11);
                runProperties106.Append(kern36);
                runProperties106.Append(fontSize33);
                runProperties106.Append(fontSizeComplexScript36);
                FieldChar fieldChar9 = new FieldChar() { FieldCharType = FieldCharValues.End };

                run110.Append(runProperties106);
                run110.Append(fieldChar9);

                paragraph17.Append(paragraphProperties17);
                paragraph17.Append(run110);

                tableCell11.Append(tableCellProperties11);
                tableCell11.Append(paragraph15);
                tableCell11.Append(paragraph16);
                tableCell11.Append(paragraph17);

                tableRow3.Append(tableRowProperties3);
                tableRow3.Append(tableCell7);
                tableRow3.Append(tableCell8);
                tableRow3.Append(tableCell9);
                tableRow3.Append(tableCell10);
                tableRow3.Append(tableCell11);

                table1.Append(tableRow3);

                count++;
            }

            return table1;

        }
    }
}

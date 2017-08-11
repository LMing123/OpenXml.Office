using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Word.tableModel
{
    using DocumentFormat.OpenXml.Wordprocessing;
    using DocumentFormat.OpenXml;
    public partial class GeneratedClass
    {
        public Table GenerateTable4()
        {
            Table table1 = new Table();

            TableProperties tableProperties1 = new TableProperties();
            TablePositionProperties tablePositionProperties1 = new TablePositionProperties() { LeftFromText = 180, RightFromText = 180, VerticalAnchor = VerticalAnchorValues.Text, HorizontalAnchor = HorizontalAnchorValues.Margin, TablePositionXAlignment = HorizontalAlignmentValues.Center, TablePositionY = -1 };
            TableWidth tableWidth1 = new TableWidth() { Width = "10758", Type = TableWidthUnitValues.Dxa };

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

            tableProperties1.Append(tablePositionProperties1);
            tableProperties1.Append(tableWidth1);
            tableProperties1.Append(tableBorders1);
            tableProperties1.Append(tableLayout1);
            tableProperties1.Append(tableCellMarginDefault1);
            tableProperties1.Append(tableLook1);

            TableGrid tableGrid1 = new TableGrid();
            GridColumn gridColumn1 = new GridColumn() { Width = "2702" };
            GridColumn gridColumn2 = new GridColumn() { Width = "685" };
            GridColumn gridColumn3 = new GridColumn() { Width = "1087" };
            GridColumn gridColumn4 = new GridColumn() { Width = "886" };
            GridColumn gridColumn5 = new GridColumn() { Width = "5398" };

            tableGrid1.Append(gridColumn1);
            tableGrid1.Append(gridColumn2);
            tableGrid1.Append(gridColumn3);
            tableGrid1.Append(gridColumn4);
            tableGrid1.Append(gridColumn5);

            TableRow tableRow1 = new TableRow() { RsidTableRowMarkRevision = "007642DC", RsidTableRowAddition = "001B2170", RsidTableRowProperties = "001B2170" };

            TableRowProperties tableRowProperties1 = new TableRowProperties();
            TableRowHeight tableRowHeight1 = new TableRowHeight() { Val = (UInt32Value)530U };

            tableRowProperties1.Append(tableRowHeight1);

            TableCell tableCell1 = new TableCell();

            TableCellProperties tableCellProperties1 = new TableCellProperties();
            TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = "10758", Type = TableWidthUnitValues.Dxa };
            GridSpan gridSpan1 = new GridSpan() { Val = 5 };
            Shading shading1 = new Shading() { Val = ShadingPatternValues.Clear, Color = "000000", Fill = "D99795" };

            tableCellProperties1.Append(tableCellWidth1);
            tableCellProperties1.Append(gridSpan1);
            tableCellProperties1.Append(shading1);

            Paragraph paragraph1 = new Paragraph() { RsidParagraphMarkRevision = "007642DC", RsidParagraphAddition = "001B2170", RsidParagraphProperties = "009E6323", RsidRunAdditionDefault = "001B2170" };

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
            text1.Text = "学业归因量表";

            run1.Append(runProperties1);
            run1.Append(text1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            tableCell1.Append(tableCellProperties1);
            tableCell1.Append(paragraph1);

            tableRow1.Append(tableRowProperties1);
            tableRow1.Append(tableCell1);

            TableRow tableRow2 = new TableRow() { RsidTableRowMarkRevision = "007642DC", RsidTableRowAddition = "001B2170", RsidTableRowProperties = "001B2170" };

            TableRowProperties tableRowProperties2 = new TableRowProperties();
            TableRowHeight tableRowHeight2 = new TableRowHeight() { Val = (UInt32Value)530U };

            tableRowProperties2.Append(tableRowHeight2);

            TableCell tableCell2 = new TableCell();

            TableCellProperties tableCellProperties2 = new TableCellProperties();
            TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = "2702", Type = TableWidthUnitValues.Dxa };
            Shading shading2 = new Shading() { Val = ShadingPatternValues.Clear, Color = "000000", Fill = "D99795" };
            TableCellVerticalAlignment tableCellVerticalAlignment1 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties2.Append(tableCellWidth2);
            tableCellProperties2.Append(shading2);
            tableCellProperties2.Append(tableCellVerticalAlignment1);

            Paragraph paragraph2 = new Paragraph() { RsidParagraphMarkRevision = "007642DC", RsidParagraphAddition = "001B2170", RsidParagraphProperties = "001B2170", RsidRunAdditionDefault = "001B2170" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            Indentation indentation2 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification2 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
            RunFonts runFonts6 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Bold bold6 = new Bold();
            Kern kern6 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize6 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties2.Append(runFonts6);
            paragraphMarkRunProperties2.Append(bold6);
            paragraphMarkRunProperties2.Append(kern6);
            paragraphMarkRunProperties2.Append(fontSize6);
            paragraphMarkRunProperties2.Append(fontSizeComplexScript6);

            paragraphProperties2.Append(indentation2);
            paragraphProperties2.Append(justification2);
            paragraphProperties2.Append(paragraphMarkRunProperties2);

            Run run5 = new Run() { RsidRunProperties = "007642DC" };

            RunProperties runProperties5 = new RunProperties();
            RunFonts runFonts7 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Bold bold7 = new Bold();
            Kern kern7 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize7 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript7 = new FontSizeComplexScript() { Val = "24" };

            runProperties5.Append(runFonts7);
            runProperties5.Append(bold7);
            runProperties5.Append(kern7);
            runProperties5.Append(fontSize7);
            runProperties5.Append(fontSizeComplexScript7);
            Text text5 = new Text();
            text5.Text = "维度名称";

            run5.Append(runProperties5);
            run5.Append(text5);

            paragraph2.Append(paragraphProperties2);
            paragraph2.Append(run5);

            tableCell2.Append(tableCellProperties2);
            tableCell2.Append(paragraph2);

            TableCell tableCell3 = new TableCell();

            TableCellProperties tableCellProperties3 = new TableCellProperties();
            TableCellWidth tableCellWidth3 = new TableCellWidth() { Width = "685", Type = TableWidthUnitValues.Dxa };
            Shading shading3 = new Shading() { Val = ShadingPatternValues.Clear, Color = "000000", Fill = "D99795" };
            TableCellVerticalAlignment tableCellVerticalAlignment2 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties3.Append(tableCellWidth3);
            tableCellProperties3.Append(shading3);
            tableCellProperties3.Append(tableCellVerticalAlignment2);

            Paragraph paragraph3 = new Paragraph() { RsidParagraphMarkRevision = "007642DC", RsidParagraphAddition = "001B2170", RsidParagraphProperties = "001B2170", RsidRunAdditionDefault = "001B2170" };

            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
            Indentation indentation3 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification3 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
            RunFonts runFonts8 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Bold bold8 = new Bold();
            Kern kern8 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize8 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript8 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties3.Append(runFonts8);
            paragraphMarkRunProperties3.Append(bold8);
            paragraphMarkRunProperties3.Append(kern8);
            paragraphMarkRunProperties3.Append(fontSize8);
            paragraphMarkRunProperties3.Append(fontSizeComplexScript8);

            paragraphProperties3.Append(indentation3);
            paragraphProperties3.Append(justification3);
            paragraphProperties3.Append(paragraphMarkRunProperties3);

            Run run6 = new Run() { RsidRunProperties = "007642DC" };

            RunProperties runProperties6 = new RunProperties();
            RunFonts runFonts9 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Bold bold9 = new Bold();
            Kern kern9 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize9 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript9 = new FontSizeComplexScript() { Val = "24" };

            runProperties6.Append(runFonts9);
            runProperties6.Append(bold9);
            runProperties6.Append(kern9);
            runProperties6.Append(fontSize9);
            runProperties6.Append(fontSizeComplexScript9);
            Text text6 = new Text();
            text6.Text = "得分";

            run6.Append(runProperties6);
            run6.Append(text6);

            paragraph3.Append(paragraphProperties3);
            paragraph3.Append(run6);

            tableCell3.Append(tableCellProperties3);
            tableCell3.Append(paragraph3);

            TableCell tableCell4 = new TableCell();

            TableCellProperties tableCellProperties4 = new TableCellProperties();
            TableCellWidth tableCellWidth4 = new TableCellWidth() { Width = "1087", Type = TableWidthUnitValues.Dxa };
            Shading shading4 = new Shading() { Val = ShadingPatternValues.Clear, Color = "000000", Fill = "D99795" };
            TableCellVerticalAlignment tableCellVerticalAlignment3 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties4.Append(tableCellWidth4);
            tableCellProperties4.Append(shading4);
            tableCellProperties4.Append(tableCellVerticalAlignment3);

            Paragraph paragraph4 = new Paragraph() { RsidParagraphMarkRevision = "007642DC", RsidParagraphAddition = "001B2170", RsidParagraphProperties = "001B2170", RsidRunAdditionDefault = "001B2170" };

            ParagraphProperties paragraphProperties4 = new ParagraphProperties();
            Indentation indentation4 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification4 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties4 = new ParagraphMarkRunProperties();
            RunFonts runFonts10 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Bold bold10 = new Bold();
            Kern kern10 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize10 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript10 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties4.Append(runFonts10);
            paragraphMarkRunProperties4.Append(bold10);
            paragraphMarkRunProperties4.Append(kern10);
            paragraphMarkRunProperties4.Append(fontSize10);
            paragraphMarkRunProperties4.Append(fontSizeComplexScript10);

            paragraphProperties4.Append(indentation4);
            paragraphProperties4.Append(justification4);
            paragraphProperties4.Append(paragraphMarkRunProperties4);

            Run run7 = new Run();

            RunProperties runProperties7 = new RunProperties();
            RunFonts runFonts11 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Bold bold11 = new Bold();
            Kern kern11 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize11 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript11 = new FontSizeComplexScript() { Val = "24" };

            runProperties7.Append(runFonts11);
            runProperties7.Append(bold11);
            runProperties7.Append(kern11);
            runProperties7.Append(fontSize11);
            runProperties7.Append(fontSizeComplexScript11);
            Text text7 = new Text();
            text7.Text = "水平";

            run7.Append(runProperties7);
            run7.Append(text7);

            paragraph4.Append(paragraphProperties4);
            paragraph4.Append(run7);

            tableCell4.Append(tableCellProperties4);
            tableCell4.Append(paragraph4);

            TableCell tableCell5 = new TableCell();

            TableCellProperties tableCellProperties5 = new TableCellProperties();
            TableCellWidth tableCellWidth5 = new TableCellWidth() { Width = "886", Type = TableWidthUnitValues.Dxa };
            Shading shading5 = new Shading() { Val = ShadingPatternValues.Clear, Color = "000000", Fill = "D99795" };
            TableCellVerticalAlignment tableCellVerticalAlignment4 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties5.Append(tableCellWidth5);
            tableCellProperties5.Append(shading5);
            tableCellProperties5.Append(tableCellVerticalAlignment4);

            Paragraph paragraph5 = new Paragraph() { RsidParagraphMarkRevision = "007642DC", RsidParagraphAddition = "001B2170", RsidParagraphProperties = "001B2170", RsidRunAdditionDefault = "001B2170" };

            ParagraphProperties paragraphProperties5 = new ParagraphProperties();
            Indentation indentation5 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification5 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties5 = new ParagraphMarkRunProperties();
            RunFonts runFonts12 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Bold bold12 = new Bold();
            Kern kern12 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize12 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript12 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties5.Append(runFonts12);
            paragraphMarkRunProperties5.Append(bold12);
            paragraphMarkRunProperties5.Append(kern12);
            paragraphMarkRunProperties5.Append(fontSize12);
            paragraphMarkRunProperties5.Append(fontSizeComplexScript12);

            paragraphProperties5.Append(indentation5);
            paragraphProperties5.Append(justification5);
            paragraphProperties5.Append(paragraphMarkRunProperties5);

            Run run8 = new Run();

            RunProperties runProperties8 = new RunProperties();
            RunFonts runFonts13 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Bold bold13 = new Bold();
            Kern kern13 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize13 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript13 = new FontSizeComplexScript() { Val = "24" };

            runProperties8.Append(runFonts13);
            runProperties8.Append(bold13);
            runProperties8.Append(kern13);
            runProperties8.Append(fontSize13);
            runProperties8.Append(fontSizeComplexScript13);
            Text text8 = new Text();
            text8.Text = "对";

            run8.Append(runProperties8);
            run8.Append(text8);

            Run run9 = new Run();

            RunProperties runProperties9 = new RunProperties();
            RunFonts runFonts14 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
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
            text9.Text = "学习的影响";

            run9.Append(runProperties9);
            run9.Append(text9);

            paragraph5.Append(paragraphProperties5);
            paragraph5.Append(run8);
            paragraph5.Append(run9);

            tableCell5.Append(tableCellProperties5);
            tableCell5.Append(paragraph5);

            TableCell tableCell6 = new TableCell();

            TableCellProperties tableCellProperties6 = new TableCellProperties();
            TableCellWidth tableCellWidth6 = new TableCellWidth() { Width = "5398", Type = TableWidthUnitValues.Dxa };
            Shading shading6 = new Shading() { Val = ShadingPatternValues.Clear, Color = "000000", Fill = "D99795" };
            TableCellVerticalAlignment tableCellVerticalAlignment5 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties6.Append(tableCellWidth6);
            tableCellProperties6.Append(shading6);
            tableCellProperties6.Append(tableCellVerticalAlignment5);

            Paragraph paragraph6 = new Paragraph() { RsidParagraphMarkRevision = "007642DC", RsidParagraphAddition = "001B2170", RsidParagraphProperties = "001B2170", RsidRunAdditionDefault = "001B2170" };

            ParagraphProperties paragraphProperties6 = new ParagraphProperties();
            Indentation indentation6 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification6 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties6 = new ParagraphMarkRunProperties();
            RunFonts runFonts15 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Bold bold15 = new Bold();
            Kern kern15 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize15 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript15 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties6.Append(runFonts15);
            paragraphMarkRunProperties6.Append(bold15);
            paragraphMarkRunProperties6.Append(kern15);
            paragraphMarkRunProperties6.Append(fontSize15);
            paragraphMarkRunProperties6.Append(fontSizeComplexScript15);

            paragraphProperties6.Append(indentation6);
            paragraphProperties6.Append(justification6);
            paragraphProperties6.Append(paragraphMarkRunProperties6);

            Run run10 = new Run();

            RunProperties runProperties10 = new RunProperties();
            RunFonts runFonts16 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Bold bold16 = new Bold();
            Kern kern16 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize16 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript16 = new FontSizeComplexScript() { Val = "24" };

            runProperties10.Append(runFonts16);
            runProperties10.Append(bold16);
            runProperties10.Append(kern16);
            runProperties10.Append(fontSize16);
            runProperties10.Append(fontSizeComplexScript16);
            Text text10 = new Text();
            text10.Text = "归因方式";

            run10.Append(runProperties10);
            run10.Append(text10);

            paragraph6.Append(paragraphProperties6);
            paragraph6.Append(run10);

            tableCell6.Append(tableCellProperties6);
            tableCell6.Append(paragraph6);

            tableRow2.Append(tableRowProperties2);
            tableRow2.Append(tableCell2);
            tableRow2.Append(tableCell3);
            tableRow2.Append(tableCell4);
            tableRow2.Append(tableCell5);
            tableRow2.Append(tableCell6);

            TableRow tableRow3 = new TableRow() { RsidTableRowMarkRevision = "007642DC", RsidTableRowAddition = "001B2170", RsidTableRowProperties = "001B2170" };

            TableRowProperties tableRowProperties3 = new TableRowProperties();
            TableRowHeight tableRowHeight3 = new TableRowHeight() { Val = (UInt32Value)1846U };

            tableRowProperties3.Append(tableRowHeight3);

            TableCell tableCell7 = new TableCell();

            TableCellProperties tableCellProperties7 = new TableCellProperties();
            TableCellWidth tableCellWidth7 = new TableCellWidth() { Width = "2702", Type = TableWidthUnitValues.Dxa };
            Shading shading7 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment6 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties7.Append(tableCellWidth7);
            tableCellProperties7.Append(shading7);
            tableCellProperties7.Append(tableCellVerticalAlignment6);

            Paragraph paragraph7 = new Paragraph() { RsidParagraphMarkRevision = "004F3E28", RsidParagraphAddition = "001B2170", RsidParagraphProperties = "001B2170", RsidRunAdditionDefault = "001B2170" };

            ParagraphProperties paragraphProperties7 = new ParagraphProperties();
            Indentation indentation7 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification7 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties7 = new ParagraphMarkRunProperties();
            RunFonts runFonts17 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Bold bold17 = new Bold();
            Color color3 = new Color() { Val = "000000" };
            Kern kern17 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize17 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript17 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties7.Append(runFonts17);
            paragraphMarkRunProperties7.Append(bold17);
            paragraphMarkRunProperties7.Append(color3);
            paragraphMarkRunProperties7.Append(kern17);
            paragraphMarkRunProperties7.Append(fontSize17);
            paragraphMarkRunProperties7.Append(fontSizeComplexScript17);

            paragraphProperties7.Append(indentation7);
            paragraphProperties7.Append(justification7);
            paragraphProperties7.Append(paragraphMarkRunProperties7);

            Run run11 = new Run();

            RunProperties runProperties11 = new RunProperties();
            RunFonts runFonts18 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Bold bold18 = new Bold();
            Color color4 = new Color() { Val = "000000" };
            Kern kern18 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize18 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript18 = new FontSizeComplexScript() { Val = "24" };

            runProperties11.Append(runFonts18);
            runProperties11.Append(bold18);
            runProperties11.Append(color4);
            runProperties11.Append(kern18);
            runProperties11.Append(fontSize18);
            runProperties11.Append(fontSizeComplexScript18);
            Text text11 = new Text();
            text11.Text = "学业成功-内归因";

            run11.Append(runProperties11);
            run11.Append(text11);

            paragraph7.Append(paragraphProperties7);
            paragraph7.Append(run11);
            tableCell7.Append(tableCellProperties7);
            tableCell7.Append(paragraph7);

            TableCell tableCell8 = new TableCell();

            TableCellProperties tableCellProperties8 = new TableCellProperties();
            TableCellWidth tableCellWidth8 = new TableCellWidth() { Width = "685", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment7 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties8.Append(tableCellWidth8);
            tableCellProperties8.Append(tableCellVerticalAlignment7);

            Paragraph paragraph8 = new Paragraph() { RsidParagraphMarkRevision = "001B2170", RsidParagraphAddition = "001B2170", RsidParagraphProperties = "001B2170", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties8 = new ParagraphProperties();
            Indentation indentation8 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification8 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties8 = new ParagraphMarkRunProperties();
            FontSize fontSize21 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript21 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties8.Append(fontSize21);
            paragraphMarkRunProperties8.Append(fontSizeComplexScript21);

            paragraphProperties8.Append(indentation8);
            paragraphProperties8.Append(justification8);
            paragraphProperties8.Append(paragraphMarkRunProperties8);

            paragraph8.Append(paragraphProperties8);

            Paragraph paragraph9 = new Paragraph() { RsidParagraphMarkRevision = "001B2170", RsidParagraphAddition = "001B2170", RsidParagraphProperties = "001B2170", RsidRunAdditionDefault = "001B2170" };

            ParagraphProperties paragraphProperties9 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation9 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification9 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties9 = new ParagraphMarkRunProperties();
            RunFonts runFonts40 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "宋体" };
            Color color7 = new Color() { Val = "000000" };
            Kern kern21 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize43 = new FontSize() { Val = "22" };

            paragraphMarkRunProperties9.Append(runFonts40);
            paragraphMarkRunProperties9.Append(color7);
            paragraphMarkRunProperties9.Append(kern21);
            paragraphMarkRunProperties9.Append(fontSize43);

            paragraphProperties9.Append(spacingBetweenLines1);
            paragraphProperties9.Append(indentation9);
            paragraphProperties9.Append(justification9);
            paragraphProperties9.Append(paragraphMarkRunProperties9);

            Run run35 = new Run() { RsidRunProperties = "001B2170" };

            RunProperties runProperties35 = new RunProperties();
            RunFonts runFonts41 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "宋体" };
            Color color8 = new Color() { Val = "000000" };
            Kern kern22 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize44 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript22 = new FontSizeComplexScript() { Val = "24" };

            runProperties35.Append(runFonts41);
            runProperties35.Append(color8);
            runProperties35.Append(kern22);
            runProperties35.Append(fontSize44);
            runProperties35.Append(fontSizeComplexScript22);
            Text text14 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text14.Text = "61.40 ";

            run35.Append(runProperties35);
            run35.Append(text14);

            paragraph9.Append(paragraphProperties9);
            paragraph9.Append(run35);

            Paragraph paragraph10 = new Paragraph() { RsidParagraphMarkRevision = "001B2170", RsidParagraphAddition = "001B2170", RsidParagraphProperties = "001B2170", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties10 = new ParagraphProperties();
            Indentation indentation10 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification10 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties10 = new ParagraphMarkRunProperties();
            RunFonts runFonts42 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Bold bold21 = new Bold();
            Color color9 = new Color() { Val = "000000" };
            Kern kern23 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize45 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript23 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties10.Append(runFonts42);
            paragraphMarkRunProperties10.Append(bold21);
            paragraphMarkRunProperties10.Append(color9);
            paragraphMarkRunProperties10.Append(kern23);
            paragraphMarkRunProperties10.Append(fontSize45);
            paragraphMarkRunProperties10.Append(fontSizeComplexScript23);

            paragraphProperties10.Append(indentation10);
            paragraphProperties10.Append(justification10);
            paragraphProperties10.Append(paragraphMarkRunProperties10);

            Run run36 = new Run() { RsidRunProperties = "001B2170" };

            RunProperties runProperties36 = new RunProperties();
            RunFonts runFonts43 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Bold bold22 = new Bold();
            Color color10 = new Color() { Val = "000000" };
            Kern kern24 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize46 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript24 = new FontSizeComplexScript() { Val = "24" };

            runProperties36.Append(runFonts43);
            runProperties36.Append(bold22);
            runProperties36.Append(color10);
            runProperties36.Append(kern24);
            runProperties36.Append(fontSize46);
            runProperties36.Append(fontSizeComplexScript24);
            FieldChar fieldChar3 = new FieldChar() { FieldCharType = FieldCharValues.End };

            run36.Append(runProperties36);
            run36.Append(fieldChar3);

            paragraph10.Append(paragraphProperties10);
            paragraph10.Append(run36);

            tableCell8.Append(tableCellProperties8);
            tableCell8.Append(paragraph8);
            tableCell8.Append(paragraph9);
            tableCell8.Append(paragraph10);

            TableCell tableCell9 = new TableCell();

            TableCellProperties tableCellProperties9 = new TableCellProperties();
            TableCellWidth tableCellWidth9 = new TableCellWidth() { Width = "1087", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment8 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties9.Append(tableCellWidth9);
            tableCellProperties9.Append(tableCellVerticalAlignment8);

            Paragraph paragraph11 = new Paragraph() { RsidParagraphMarkRevision = "001B2170", RsidParagraphAddition = "001B2170", RsidParagraphProperties = "001B2170", RsidRunAdditionDefault = "001B2170" };

            ParagraphProperties paragraphProperties11 = new ParagraphProperties();
            Indentation indentation11 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification11 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties11 = new ParagraphMarkRunProperties();
            RunFonts runFonts44 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Bold bold23 = new Bold();
            Color color11 = new Color() { Val = "000000" };
            Kern kern25 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize47 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript25 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties11.Append(runFonts44);
            paragraphMarkRunProperties11.Append(bold23);
            paragraphMarkRunProperties11.Append(color11);
            paragraphMarkRunProperties11.Append(kern25);
            paragraphMarkRunProperties11.Append(fontSize47);
            paragraphMarkRunProperties11.Append(fontSizeComplexScript25);

            paragraphProperties11.Append(indentation11);
            paragraphProperties11.Append(justification11);
            paragraphProperties11.Append(paragraphMarkRunProperties11);

            Run run37 = new Run() { RsidRunProperties = "001B2170" };

            RunProperties runProperties37 = new RunProperties();
            RunFonts runFonts45 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Bold bold24 = new Bold();
            Color color12 = new Color() { Val = "000000" };
            Kern kern26 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize48 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript26 = new FontSizeComplexScript() { Val = "24" };

            runProperties37.Append(runFonts45);
            runProperties37.Append(bold24);
            runProperties37.Append(color12);
            runProperties37.Append(kern26);
            runProperties37.Append(fontSize48);
            runProperties37.Append(fontSizeComplexScript26);
            Text text15 = new Text();
            text15.Text = "高";

            run37.Append(runProperties37);
            run37.Append(text15);

            Run run38 = new Run() { RsidRunProperties = "001B2170" };

            RunProperties runProperties38 = new RunProperties();
            RunFonts runFonts46 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Bold bold25 = new Bold();
            Color color13 = new Color() { Val = "000000" };
            Kern kern27 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize49 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript27 = new FontSizeComplexScript() { Val = "24" };

            runProperties38.Append(runFonts46);
            runProperties38.Append(bold25);
            runProperties38.Append(color13);
            runProperties38.Append(kern27);
            runProperties38.Append(fontSize49);
            runProperties38.Append(fontSizeComplexScript27);
            Text text16 = new Text();
            text16.Text = "水平";

            run38.Append(runProperties38);
            run38.Append(text16);

            paragraph11.Append(paragraphProperties11);
            paragraph11.Append(run37);
            paragraph11.Append(run38);

            tableCell9.Append(tableCellProperties9);
            tableCell9.Append(paragraph11);

            TableCell tableCell10 = new TableCell();

            TableCellProperties tableCellProperties10 = new TableCellProperties();
            TableCellWidth tableCellWidth10 = new TableCellWidth() { Width = "886", Type = TableWidthUnitValues.Dxa };
            Shading shading8 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment9 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties10.Append(tableCellWidth10);
            tableCellProperties10.Append(shading8);
            tableCellProperties10.Append(tableCellVerticalAlignment9);

            Paragraph paragraph12 = new Paragraph() { RsidParagraphMarkRevision = "001B2170", RsidParagraphAddition = "001B2170", RsidParagraphProperties = "001B2170", RsidRunAdditionDefault = "001B2170" };

            ParagraphProperties paragraphProperties12 = new ParagraphProperties();
            Indentation indentation12 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification12 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties12 = new ParagraphMarkRunProperties();
            RunFonts runFonts47 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Bold bold26 = new Bold();
            Color color14 = new Color() { Val = "000000" };
            Kern kern28 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize50 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript28 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties12.Append(runFonts47);
            paragraphMarkRunProperties12.Append(bold26);
            paragraphMarkRunProperties12.Append(color14);
            paragraphMarkRunProperties12.Append(kern28);
            paragraphMarkRunProperties12.Append(fontSize50);
            paragraphMarkRunProperties12.Append(fontSizeComplexScript28);

            paragraphProperties12.Append(indentation12);
            paragraphProperties12.Append(justification12);
            paragraphProperties12.Append(paragraphMarkRunProperties12);

            Run run39 = new Run() { RsidRunProperties = "001B2170" };

            RunProperties runProperties39 = new RunProperties();
            RunFonts runFonts48 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Color color15 = new Color() { Val = "000000" };
            Kern kern29 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize51 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript29 = new FontSizeComplexScript() { Val = "24" };

            runProperties39.Append(runFonts48);
            runProperties39.Append(color15);
            runProperties39.Append(kern29);
            runProperties39.Append(fontSize51);
            runProperties39.Append(fontSizeComplexScript29);
            Text text17 = new Text();
            text17.Text = "-";

            run39.Append(runProperties39);
            run39.Append(text17);

            Run run40 = new Run() { RsidRunProperties = "001B2170" };

            RunProperties runProperties40 = new RunProperties();
            RunFonts runFonts49 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Color color16 = new Color() { Val = "000000" };
            Kern kern30 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize52 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript30 = new FontSizeComplexScript() { Val = "24" };

            runProperties40.Append(runFonts49);
            runProperties40.Append(color16);
            runProperties40.Append(kern30);
            runProperties40.Append(fontSize52);
            runProperties40.Append(fontSizeComplexScript30);
            Text text18 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text18.Text = " ";

            run40.Append(runProperties40);
            run40.Append(text18);

            Run run41 = new Run() { RsidRunProperties = "001B2170" };

            RunProperties runProperties41 = new RunProperties();
            RunFonts runFonts50 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Color color17 = new Color() { Val = "000000" };
            Kern kern31 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize53 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript31 = new FontSizeComplexScript() { Val = "24" };

            runProperties41.Append(runFonts50);
            runProperties41.Append(color17);
            runProperties41.Append(kern31);
            runProperties41.Append(fontSize53);
            runProperties41.Append(fontSizeComplexScript31);
            Text text19 = new Text();
            text19.Text = "-";

            run41.Append(runProperties41);
            run41.Append(text19);

            Run run42 = new Run() { RsidRunProperties = "001B2170" };

            RunProperties runProperties42 = new RunProperties();
            RunFonts runFonts51 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Color color18 = new Color() { Val = "000000" };
            Kern kern32 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize54 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript32 = new FontSizeComplexScript() { Val = "24" };

            runProperties42.Append(runFonts51);
            runProperties42.Append(color18);
            runProperties42.Append(kern32);
            runProperties42.Append(fontSize54);
            runProperties42.Append(fontSizeComplexScript32);
            Text text20 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text20.Text = " -";

            run42.Append(runProperties42);
            run42.Append(text20);

            paragraph12.Append(paragraphProperties12);
            paragraph12.Append(run39);
            paragraph12.Append(run40);
            paragraph12.Append(run41);
            paragraph12.Append(run42);

            tableCell10.Append(tableCellProperties10);
            tableCell10.Append(paragraph12);

            TableCell tableCell11 = new TableCell();

            TableCellProperties tableCellProperties11 = new TableCellProperties();
            TableCellWidth tableCellWidth11 = new TableCellWidth() { Width = "5398", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge1 = new VerticalMerge() { Val = MergedCellValues.Restart };
            Shading shading9 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment10 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties11.Append(tableCellWidth11);
            tableCellProperties11.Append(verticalMerge1);
            tableCellProperties11.Append(shading9);
            tableCellProperties11.Append(tableCellVerticalAlignment10);

            Paragraph paragraph13 = new Paragraph() { RsidParagraphMarkRevision = "00C16BA1", RsidParagraphAddition = "001B2170", RsidParagraphProperties = "001B2170", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties13 = new ParagraphProperties();
            Indentation indentation13 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };

            ParagraphMarkRunProperties paragraphMarkRunProperties13 = new ParagraphMarkRunProperties();
            FontSize fontSize55 = new FontSize() { Val = "4" };
            FontSizeComplexScript fontSizeComplexScript33 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties13.Append(fontSize55);
            paragraphMarkRunProperties13.Append(fontSizeComplexScript33);

            paragraphProperties13.Append(indentation13);
            paragraphProperties13.Append(paragraphMarkRunProperties13);

          

            paragraph13.Append(paragraphProperties13);
          

            Paragraph paragraph14 = new Paragraph() { RsidParagraphMarkRevision = "00C16BA1", RsidParagraphAddition = "001B2170", RsidParagraphProperties = "001B2170", RsidRunAdditionDefault = "001B2170" };

            ParagraphProperties paragraphProperties14 = new ParagraphProperties();
            Indentation indentation14 = new Indentation() { FirstLine = "480" };

            ParagraphMarkRunProperties paragraphMarkRunProperties14 = new ParagraphMarkRunProperties();
            RunFonts runFonts71 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "宋体" };
            Color color19 = new Color() { Val = "000000" };
            Kern kern33 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize56 = new FontSize() { Val = "24" };

            paragraphMarkRunProperties14.Append(runFonts71);
            paragraphMarkRunProperties14.Append(color19);
            paragraphMarkRunProperties14.Append(kern33);
            paragraphMarkRunProperties14.Append(fontSize56);

            paragraphProperties14.Append(indentation14);
            paragraphProperties14.Append(paragraphMarkRunProperties14);

            Run run64 = new Run() { RsidRunProperties = "00C16BA1" };

            RunProperties runProperties62 = new RunProperties();
            RunFonts runFonts72 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "宋体" };
            Color color20 = new Color() { Val = "000000" };
            Kern kern34 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize57 = new FontSize() { Val = "24" };

            runProperties62.Append(runFonts72);
            runProperties62.Append(color20);
            runProperties62.Append(kern34);
            runProperties62.Append(fontSize57);
            Text text21 = new Text();
            text21.Text = "学业成功时的归因方式是内归因，是一种比较积极的归因方式。倾向于把学业成功归因于自身学习努力和学习能力这样的可控因素，是对自己能力的一种肯定，" +
                "他们对于学习能够更有掌控感和成就感，从而增强了对于学习的控制和对成功的渴望，那么他们就会积极地争取成功，会感到满意、自豪等积极情绪，有助于维持和激发其随后的学习动机，他们就更加愿意去学习。在这种情况下，更容易完成我们所设的任务，而不是选择逃避。";

            run64.Append(runProperties62);
            run64.Append(text21);

            paragraph14.Append(paragraphProperties14);
            paragraph14.Append(run64);

            Paragraph paragraph15 = new Paragraph() { RsidParagraphMarkRevision = "007D1A31", RsidParagraphAddition = "001B2170", RsidParagraphProperties = "001B2170", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties15 = new ParagraphProperties();
            Indentation indentation15 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };

            ParagraphMarkRunProperties paragraphMarkRunProperties15 = new ParagraphMarkRunProperties();
            RunFonts runFonts75 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Color color23 = new Color() { Val = "000000" };
            Kern kern37 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize60 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript34 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties15.Append(runFonts75);
            paragraphMarkRunProperties15.Append(color23);
            paragraphMarkRunProperties15.Append(kern37);
            paragraphMarkRunProperties15.Append(fontSize60);
            paragraphMarkRunProperties15.Append(fontSizeComplexScript34);

            paragraphProperties15.Append(indentation15);
            paragraphProperties15.Append(paragraphMarkRunProperties15);

            Run run67 = new Run();

            RunProperties runProperties65 = new RunProperties();
            RunFonts runFonts76 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Color color24 = new Color() { Val = "000000" };
            Kern kern38 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize61 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript35 = new FontSizeComplexScript() { Val = "24" };

            runProperties65.Append(runFonts76);
            runProperties65.Append(color24);
            runProperties65.Append(kern38);
            runProperties65.Append(fontSize61);
            runProperties65.Append(fontSizeComplexScript35);
            FieldChar fieldChar6 = new FieldChar() { FieldCharType = FieldCharValues.End };

            run67.Append(runProperties65);
            run67.Append(fieldChar6);

            paragraph15.Append(paragraphProperties15);
            paragraph15.Append(run67);

            tableCell11.Append(tableCellProperties11);
            tableCell11.Append(paragraph13);
            tableCell11.Append(paragraph14);
            tableCell11.Append(paragraph15);

            tableRow3.Append(tableRowProperties3);
            tableRow3.Append(tableCell7);
            tableRow3.Append(tableCell8);
            tableRow3.Append(tableCell9);
            tableRow3.Append(tableCell10);
            tableRow3.Append(tableCell11);

            TableRow tableRow4 = new TableRow() { RsidTableRowMarkRevision = "007642DC", RsidTableRowAddition = "001B2170", RsidTableRowProperties = "001B2170" };

            TableRowProperties tableRowProperties4 = new TableRowProperties();
            TableRowHeight tableRowHeight4 = new TableRowHeight() { Val = (UInt32Value)1097U };

            tableRowProperties4.Append(tableRowHeight4);

            TableCell tableCell12 = new TableCell();

            TableCellProperties tableCellProperties12 = new TableCellProperties();
            TableCellWidth tableCellWidth12 = new TableCellWidth() { Width = "2702", Type = TableWidthUnitValues.Dxa };
            Shading shading10 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment11 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties12.Append(tableCellWidth12);
            tableCellProperties12.Append(shading10);
            tableCellProperties12.Append(tableCellVerticalAlignment11);

            Paragraph paragraph16 = new Paragraph() { RsidParagraphMarkRevision = "00973E52", RsidParagraphAddition = "001B2170", RsidParagraphProperties = "001B2170", RsidRunAdditionDefault = "001B2170" };

            ParagraphProperties paragraphProperties16 = new ParagraphProperties();
            Indentation indentation16 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification13 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties16 = new ParagraphMarkRunProperties();
            RunFonts runFonts77 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Kern kern39 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize62 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript36 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties16.Append(runFonts77);
            paragraphMarkRunProperties16.Append(kern39);
            paragraphMarkRunProperties16.Append(fontSize62);
            paragraphMarkRunProperties16.Append(fontSizeComplexScript36);

            paragraphProperties16.Append(indentation16);
            paragraphProperties16.Append(justification13);
            paragraphProperties16.Append(paragraphMarkRunProperties16);

            Run run68 = new Run();

            RunProperties runProperties66 = new RunProperties();
            RunFonts runFonts78 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Kern kern40 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize63 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript37 = new FontSizeComplexScript() { Val = "24" };

            runProperties66.Append(runFonts78);
            runProperties66.Append(kern40);
            runProperties66.Append(fontSize63);
            runProperties66.Append(fontSizeComplexScript37);
            Text text24 = new Text();
            text24.Text = "学业成功--外归因";

            run68.Append(runProperties66);
            run68.Append(text24);

            paragraph16.Append(paragraphProperties16);
            paragraph16.Append(run68);
          

            tableCell12.Append(tableCellProperties12);
            tableCell12.Append(paragraph16);

            TableCell tableCell13 = new TableCell();

            TableCellProperties tableCellProperties13 = new TableCellProperties();
            TableCellWidth tableCellWidth13 = new TableCellWidth() { Width = "685", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment12 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties13.Append(tableCellWidth13);
            tableCellProperties13.Append(tableCellVerticalAlignment12);

            Paragraph paragraph17 = new Paragraph() { RsidParagraphMarkRevision = "001B2170", RsidParagraphAddition = "001B2170", RsidParagraphProperties = "001B2170", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties17 = new ParagraphProperties();
            Indentation indentation17 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification14 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties17 = new ParagraphMarkRunProperties();
            FontSize fontSize68 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript42 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties17.Append(fontSize68);
            paragraphMarkRunProperties17.Append(fontSizeComplexScript42);

            paragraphProperties17.Append(indentation17);
            paragraphProperties17.Append(justification14);
            paragraphProperties17.Append(paragraphMarkRunProperties17);

            paragraph17.Append(paragraphProperties17);

            Paragraph paragraph18 = new Paragraph() { RsidParagraphMarkRevision = "001B2170", RsidParagraphAddition = "001B2170", RsidParagraphProperties = "001B2170", RsidRunAdditionDefault = "001B2170" };

            ParagraphProperties paragraphProperties18 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation18 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification15 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties18 = new ParagraphMarkRunProperties();
            RunFonts runFonts102 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "宋体" };
            Color color25 = new Color() { Val = "000000" };
            Kern kern45 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize90 = new FontSize() { Val = "22" };

            paragraphMarkRunProperties18.Append(runFonts102);
            paragraphMarkRunProperties18.Append(color25);
            paragraphMarkRunProperties18.Append(kern45);
            paragraphMarkRunProperties18.Append(fontSize90);

            paragraphProperties18.Append(spacingBetweenLines2);
            paragraphProperties18.Append(indentation18);
            paragraphProperties18.Append(justification15);
            paragraphProperties18.Append(paragraphMarkRunProperties18);

            Run run94 = new Run() { RsidRunProperties = "001B2170" };

            RunProperties runProperties92 = new RunProperties();
            RunFonts runFonts103 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "宋体" };
            Color color26 = new Color() { Val = "000000" };
            Kern kern46 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize91 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript43 = new FontSizeComplexScript() { Val = "24" };

            runProperties92.Append(runFonts103);
            runProperties92.Append(color26);
            runProperties92.Append(kern46);
            runProperties92.Append(fontSize91);
            runProperties92.Append(fontSizeComplexScript43);
            Text text29 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text29.Text = "45.00 ";

            run94.Append(runProperties92);
            run94.Append(text29);

            paragraph18.Append(paragraphProperties18);
            paragraph18.Append(run94);

            Paragraph paragraph19 = new Paragraph() { RsidParagraphMarkRevision = "001B2170", RsidParagraphAddition = "001B2170", RsidParagraphProperties = "001B2170", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties19 = new ParagraphProperties();
            Indentation indentation19 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification16 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties19 = new ParagraphMarkRunProperties();
            RunFonts runFonts104 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Bold bold27 = new Bold();
            Color color27 = new Color() { Val = "000000" };
            Kern kern47 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize92 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript44 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties19.Append(runFonts104);
            paragraphMarkRunProperties19.Append(bold27);
            paragraphMarkRunProperties19.Append(color27);
            paragraphMarkRunProperties19.Append(kern47);
            paragraphMarkRunProperties19.Append(fontSize92);
            paragraphMarkRunProperties19.Append(fontSizeComplexScript44);

            paragraphProperties19.Append(indentation19);
            paragraphProperties19.Append(justification16);
            paragraphProperties19.Append(paragraphMarkRunProperties19);

            Run run95 = new Run() { RsidRunProperties = "001B2170" };

            RunProperties runProperties93 = new RunProperties();
            RunFonts runFonts105 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Bold bold28 = new Bold();
            Color color28 = new Color() { Val = "000000" };
            Kern kern48 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize93 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript45 = new FontSizeComplexScript() { Val = "24" };

            runProperties93.Append(runFonts105);
            runProperties93.Append(bold28);
            runProperties93.Append(color28);
            runProperties93.Append(kern48);
            runProperties93.Append(fontSize93);
            runProperties93.Append(fontSizeComplexScript45);
            FieldChar fieldChar9 = new FieldChar() { FieldCharType = FieldCharValues.End };

            run95.Append(runProperties93);
            run95.Append(fieldChar9);

            paragraph19.Append(paragraphProperties19);
            paragraph19.Append(run95);

            tableCell13.Append(tableCellProperties13);
            tableCell13.Append(paragraph17);
            tableCell13.Append(paragraph18);
            tableCell13.Append(paragraph19);

            TableCell tableCell14 = new TableCell();

            TableCellProperties tableCellProperties14 = new TableCellProperties();
            TableCellWidth tableCellWidth14 = new TableCellWidth() { Width = "1087", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment13 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties14.Append(tableCellWidth14);
            tableCellProperties14.Append(tableCellVerticalAlignment13);

            Paragraph paragraph20 = new Paragraph() { RsidParagraphMarkRevision = "001B2170", RsidParagraphAddition = "001B2170", RsidParagraphProperties = "001B2170", RsidRunAdditionDefault = "001B2170" };

            ParagraphProperties paragraphProperties20 = new ParagraphProperties();
            Indentation indentation20 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification17 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties20 = new ParagraphMarkRunProperties();
            RunFonts runFonts106 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Bold bold29 = new Bold();
            Color color29 = new Color() { Val = "000000" };
            Kern kern49 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize94 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript46 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties20.Append(runFonts106);
            paragraphMarkRunProperties20.Append(bold29);
            paragraphMarkRunProperties20.Append(color29);
            paragraphMarkRunProperties20.Append(kern49);
            paragraphMarkRunProperties20.Append(fontSize94);
            paragraphMarkRunProperties20.Append(fontSizeComplexScript46);

            paragraphProperties20.Append(indentation20);
            paragraphProperties20.Append(justification17);
            paragraphProperties20.Append(paragraphMarkRunProperties20);

            Run run96 = new Run() { RsidRunProperties = "001B2170" };

            RunProperties runProperties94 = new RunProperties();
            RunFonts runFonts107 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Bold bold30 = new Bold();
            Color color30 = new Color() { Val = "000000" };
            Kern kern50 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize95 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript47 = new FontSizeComplexScript() { Val = "24" };

            runProperties94.Append(runFonts107);
            runProperties94.Append(bold30);
            runProperties94.Append(color30);
            runProperties94.Append(kern50);
            runProperties94.Append(fontSize95);
            runProperties94.Append(fontSizeComplexScript47);
            Text text30 = new Text();
            text30.Text = "中等水平";

            run96.Append(runProperties94);
            run96.Append(text30);

            paragraph20.Append(paragraphProperties20);
            paragraph20.Append(run96);

            tableCell14.Append(tableCellProperties14);
            tableCell14.Append(paragraph20);

            TableCell tableCell15 = new TableCell();

            TableCellProperties tableCellProperties15 = new TableCellProperties();
            TableCellWidth tableCellWidth15 = new TableCellWidth() { Width = "886", Type = TableWidthUnitValues.Dxa };
            Shading shading11 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment14 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties15.Append(tableCellWidth15);
            tableCellProperties15.Append(shading11);
            tableCellProperties15.Append(tableCellVerticalAlignment14);

            Paragraph paragraph21 = new Paragraph() { RsidParagraphMarkRevision = "001B2170", RsidParagraphAddition = "001B2170", RsidParagraphProperties = "001B2170", RsidRunAdditionDefault = "001B2170" };

            ParagraphProperties paragraphProperties21 = new ParagraphProperties();
            Indentation indentation21 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification18 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties21 = new ParagraphMarkRunProperties();
            RunFonts runFonts109 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Bold bold32 = new Bold();
            Color color32 = new Color() { Val = "000000" };
            Kern kern52 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize97 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript49 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties21.Append(runFonts109);
            paragraphMarkRunProperties21.Append(bold32);
            paragraphMarkRunProperties21.Append(color32);
            paragraphMarkRunProperties21.Append(kern52);
            paragraphMarkRunProperties21.Append(fontSize97);
            paragraphMarkRunProperties21.Append(fontSizeComplexScript49);

            paragraphProperties21.Append(indentation21);
            paragraphProperties21.Append(justification18);
            paragraphProperties21.Append(paragraphMarkRunProperties21);

            Run run98 = new Run() { RsidRunProperties = "001B2170" };

            RunProperties runProperties96 = new RunProperties();
            RunFonts runFonts110 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Color color33 = new Color() { Val = "000000" };
            Kern kern53 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize98 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript50 = new FontSizeComplexScript() { Val = "24" };

            runProperties96.Append(runFonts110);
            runProperties96.Append(color33);
            runProperties96.Append(kern53);
            runProperties96.Append(fontSize98);
            runProperties96.Append(fontSizeComplexScript50);
            Text text32 = new Text();
            text32.Text = "- - -";

            run98.Append(runProperties96);
            run98.Append(text32);

            paragraph21.Append(paragraphProperties21);
            paragraph21.Append(run98);

            tableCell15.Append(tableCellProperties15);
            tableCell15.Append(paragraph21);

            TableCell tableCell16 = new TableCell();

            TableCellProperties tableCellProperties16 = new TableCellProperties();
            TableCellWidth tableCellWidth16 = new TableCellWidth() { Width = "5398", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge2 = new VerticalMerge();
            Shading shading12 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment15 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties16.Append(tableCellWidth16);
            tableCellProperties16.Append(verticalMerge2);
            tableCellProperties16.Append(shading12);
            tableCellProperties16.Append(tableCellVerticalAlignment15);

            Paragraph paragraph22 = new Paragraph() { RsidParagraphMarkRevision = "002A093C", RsidParagraphAddition = "001B2170", RsidParagraphProperties = "001B2170", RsidRunAdditionDefault = "001B2170" };

            ParagraphProperties paragraphProperties22 = new ParagraphProperties();
            Indentation indentation22 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };

            ParagraphMarkRunProperties paragraphMarkRunProperties22 = new ParagraphMarkRunProperties();
            RunFonts runFonts114 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Color color37 = new Color() { Val = "000000" };
            Kern kern57 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize102 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript54 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties22.Append(runFonts114);
            paragraphMarkRunProperties22.Append(color37);
            paragraphMarkRunProperties22.Append(kern57);
            paragraphMarkRunProperties22.Append(fontSize102);
            paragraphMarkRunProperties22.Append(fontSizeComplexScript54);

            paragraphProperties22.Append(indentation22);
            paragraphProperties22.Append(paragraphMarkRunProperties22);

            paragraph22.Append(paragraphProperties22);

            tableCell16.Append(tableCellProperties16);
            tableCell16.Append(paragraph22);

            tableRow4.Append(tableRowProperties4);
            tableRow4.Append(tableCell12);
            tableRow4.Append(tableCell13);
            tableRow4.Append(tableCell14);
            tableRow4.Append(tableCell15);
            tableRow4.Append(tableCell16);

            TableRow tableRow5 = new TableRow() { RsidTableRowMarkRevision = "007642DC", RsidTableRowAddition = "001B2170", RsidTableRowProperties = "001B2170" };

            TableRowProperties tableRowProperties5 = new TableRowProperties();
            TableRowHeight tableRowHeight5 = new TableRowHeight() { Val = (UInt32Value)1650U };

            tableRowProperties5.Append(tableRowHeight5);

            TableCell tableCell17 = new TableCell();

            TableCellProperties tableCellProperties17 = new TableCellProperties();
            TableCellWidth tableCellWidth17 = new TableCellWidth() { Width = "2702", Type = TableWidthUnitValues.Dxa };
            Shading shading13 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment16 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties17.Append(tableCellWidth17);
            tableCellProperties17.Append(shading13);
            tableCellProperties17.Append(tableCellVerticalAlignment16);

            Paragraph paragraph23 = new Paragraph() { RsidParagraphMarkRevision = "00973E52", RsidParagraphAddition = "001B2170", RsidParagraphProperties = "001B2170", RsidRunAdditionDefault = "001B2170" };

            ParagraphProperties paragraphProperties23 = new ParagraphProperties();
            Indentation indentation23 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification19 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties23 = new ParagraphMarkRunProperties();
            RunFonts runFonts115 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Kern kern58 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize103 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript55 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties23.Append(runFonts115);
            paragraphMarkRunProperties23.Append(kern58);
            paragraphMarkRunProperties23.Append(fontSize103);
            paragraphMarkRunProperties23.Append(fontSizeComplexScript55);

            paragraphProperties23.Append(indentation23);
            paragraphProperties23.Append(justification19);
            paragraphProperties23.Append(paragraphMarkRunProperties23);

            Run run102 = new Run();

            RunProperties runProperties100 = new RunProperties();
            RunFonts runFonts116 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Bold bold33 = new Bold();
            Color color38 = new Color() { Val = "000000" };
            Kern kern59 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize104 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript56 = new FontSizeComplexScript() { Val = "24" };

            runProperties100.Append(runFonts116);
            runProperties100.Append(bold33);
            runProperties100.Append(color38);
            runProperties100.Append(kern59);
            runProperties100.Append(fontSize104);
            runProperties100.Append(fontSizeComplexScript56);
            Text text36 = new Text();
            text36.Text = "学业失败-内归因";

            run102.Append(runProperties100);
            run102.Append(text36);

            paragraph23.Append(paragraphProperties23);
            paragraph23.Append(run102);

            tableCell17.Append(tableCellProperties17);
            tableCell17.Append(paragraph23);

            TableCell tableCell18 = new TableCell();

            TableCellProperties tableCellProperties18 = new TableCellProperties();
            TableCellWidth tableCellWidth18 = new TableCellWidth() { Width = "685", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment17 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties18.Append(tableCellWidth18);
            tableCellProperties18.Append(tableCellVerticalAlignment17);

            Paragraph paragraph24 = new Paragraph() { RsidParagraphMarkRevision = "001B2170", RsidParagraphAddition = "001B2170", RsidParagraphProperties = "001B2170", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties24 = new ParagraphProperties();
            Indentation indentation24 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification20 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties24 = new ParagraphMarkRunProperties();
            FontSize fontSize108 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript60 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties24.Append(fontSize108);
            paragraphMarkRunProperties24.Append(fontSizeComplexScript60);

            paragraphProperties24.Append(indentation24);
            paragraphProperties24.Append(justification20);
            paragraphProperties24.Append(paragraphMarkRunProperties24);

            paragraph24.Append(paragraphProperties24);

            Paragraph paragraph25 = new Paragraph() { RsidParagraphMarkRevision = "001B2170", RsidParagraphAddition = "001B2170", RsidParagraphProperties = "001B2170", RsidRunAdditionDefault = "001B2170" };

            ParagraphProperties paragraphProperties25 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines3 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation25 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification21 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties25 = new ParagraphMarkRunProperties();
            RunFonts runFonts139 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "宋体" };
            Color color42 = new Color() { Val = "000000" };
            Kern kern63 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize130 = new FontSize() { Val = "22" };

            paragraphMarkRunProperties25.Append(runFonts139);
            paragraphMarkRunProperties25.Append(color42);
            paragraphMarkRunProperties25.Append(kern63);
            paragraphMarkRunProperties25.Append(fontSize130);

            paragraphProperties25.Append(spacingBetweenLines3);
            paragraphProperties25.Append(indentation25);
            paragraphProperties25.Append(justification21);
            paragraphProperties25.Append(paragraphMarkRunProperties25);

            Run run127 = new Run() { RsidRunProperties = "001B2170" };

            RunProperties runProperties125 = new RunProperties();
            RunFonts runFonts140 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "宋体" };
            Color color43 = new Color() { Val = "000000" };
            Kern kern64 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize131 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript61 = new FontSizeComplexScript() { Val = "24" };

            runProperties125.Append(runFonts140);
            runProperties125.Append(color43);
            runProperties125.Append(kern64);
            runProperties125.Append(fontSize131);
            runProperties125.Append(fontSizeComplexScript61);
            Text text40 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text40.Text = "32.15 ";

            run127.Append(runProperties125);
            run127.Append(text40);

            paragraph25.Append(paragraphProperties25);
            paragraph25.Append(run127);

            Paragraph paragraph26 = new Paragraph() { RsidParagraphMarkRevision = "001B2170", RsidParagraphAddition = "001B2170", RsidParagraphProperties = "001B2170", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties26 = new ParagraphProperties();
            Indentation indentation26 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification22 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties26 = new ParagraphMarkRunProperties();
            RunFonts runFonts141 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Bold bold37 = new Bold();
            Color color44 = new Color() { Val = "000000" };
            Kern kern65 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize132 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript62 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties26.Append(runFonts141);
            paragraphMarkRunProperties26.Append(bold37);
            paragraphMarkRunProperties26.Append(color44);
            paragraphMarkRunProperties26.Append(kern65);
            paragraphMarkRunProperties26.Append(fontSize132);
            paragraphMarkRunProperties26.Append(fontSizeComplexScript62);

            paragraphProperties26.Append(indentation26);
            paragraphProperties26.Append(justification22);
            paragraphProperties26.Append(paragraphMarkRunProperties26);

            Run run128 = new Run() { RsidRunProperties = "001B2170" };

            RunProperties runProperties126 = new RunProperties();
            RunFonts runFonts142 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Bold bold38 = new Bold();
            Color color45 = new Color() { Val = "000000" };
            Kern kern66 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize133 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript63 = new FontSizeComplexScript() { Val = "24" };

            runProperties126.Append(runFonts142);
            runProperties126.Append(bold38);
            runProperties126.Append(color45);
            runProperties126.Append(kern66);
            runProperties126.Append(fontSize133);
            runProperties126.Append(fontSizeComplexScript63);
            FieldChar fieldChar12 = new FieldChar() { FieldCharType = FieldCharValues.End };

            run128.Append(runProperties126);
            run128.Append(fieldChar12);

            paragraph26.Append(paragraphProperties26);
            paragraph26.Append(run128);

            tableCell18.Append(tableCellProperties18);
            tableCell18.Append(paragraph24);
            tableCell18.Append(paragraph25);
            tableCell18.Append(paragraph26);

            TableCell tableCell19 = new TableCell();

            TableCellProperties tableCellProperties19 = new TableCellProperties();
            TableCellWidth tableCellWidth19 = new TableCellWidth() { Width = "1087", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment18 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties19.Append(tableCellWidth19);
            tableCellProperties19.Append(tableCellVerticalAlignment18);

            Paragraph paragraph27 = new Paragraph() { RsidParagraphMarkRevision = "001B2170", RsidParagraphAddition = "001B2170", RsidParagraphProperties = "001B2170", RsidRunAdditionDefault = "001B2170" };

            ParagraphProperties paragraphProperties27 = new ParagraphProperties();
            Indentation indentation27 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification23 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties27 = new ParagraphMarkRunProperties();
            RunFonts runFonts143 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Bold bold39 = new Bold();
            Color color46 = new Color() { Val = "000000" };
            Kern kern67 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize134 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript64 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties27.Append(runFonts143);
            paragraphMarkRunProperties27.Append(bold39);
            paragraphMarkRunProperties27.Append(color46);
            paragraphMarkRunProperties27.Append(kern67);
            paragraphMarkRunProperties27.Append(fontSize134);
            paragraphMarkRunProperties27.Append(fontSizeComplexScript64);

            paragraphProperties27.Append(indentation27);
            paragraphProperties27.Append(justification23);
            paragraphProperties27.Append(paragraphMarkRunProperties27);

            Run run129 = new Run() { RsidRunProperties = "001B2170" };

            RunProperties runProperties127 = new RunProperties();
            RunFonts runFonts144 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Bold bold40 = new Bold();
            Color color47 = new Color() { Val = "000000" };
            Kern kern68 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize135 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript65 = new FontSizeComplexScript() { Val = "24" };

            runProperties127.Append(runFonts144);
            runProperties127.Append(bold40);
            runProperties127.Append(color47);
            runProperties127.Append(kern68);
            runProperties127.Append(fontSize135);
            runProperties127.Append(fontSizeComplexScript65);
            Text text41 = new Text();
            text41.Text = "低水平";

            run129.Append(runProperties127);
            run129.Append(text41);       

            paragraph27.Append(paragraphProperties27);
            paragraph27.Append(run129);

            tableCell19.Append(tableCellProperties19);
            tableCell19.Append(paragraph27);

            TableCell tableCell20 = new TableCell();

            TableCellProperties tableCellProperties20 = new TableCellProperties();
            TableCellWidth tableCellWidth20 = new TableCellWidth() { Width = "886", Type = TableWidthUnitValues.Dxa };
            Shading shading14 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment19 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties20.Append(tableCellWidth20);
            tableCellProperties20.Append(shading14);
            tableCellProperties20.Append(tableCellVerticalAlignment19);

            Paragraph paragraph28 = new Paragraph() { RsidParagraphMarkRevision = "001B2170", RsidParagraphAddition = "001B2170", RsidParagraphProperties = "001B2170", RsidRunAdditionDefault = "001B2170" };

            ParagraphProperties paragraphProperties28 = new ParagraphProperties();
            Indentation indentation28 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification24 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties28 = new ParagraphMarkRunProperties();
            RunFonts runFonts146 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Bold bold42 = new Bold();
            Color color49 = new Color() { Val = "000000" };
            Kern kern70 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize137 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript67 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties28.Append(runFonts146);
            paragraphMarkRunProperties28.Append(bold42);
            paragraphMarkRunProperties28.Append(color49);
            paragraphMarkRunProperties28.Append(kern70);
            paragraphMarkRunProperties28.Append(fontSize137);
            paragraphMarkRunProperties28.Append(fontSizeComplexScript67);

            paragraphProperties28.Append(indentation28);
            paragraphProperties28.Append(justification24);
            paragraphProperties28.Append(paragraphMarkRunProperties28);

            Run run131 = new Run() { RsidRunProperties = "001B2170" };

            RunProperties runProperties129 = new RunProperties();
            RunFonts runFonts147 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Color color50 = new Color() { Val = "000000" };
            Kern kern71 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize138 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript68 = new FontSizeComplexScript() { Val = "24" };

            runProperties129.Append(runFonts147);
            runProperties129.Append(color50);
            runProperties129.Append(kern71);
            runProperties129.Append(fontSize138);
            runProperties129.Append(fontSizeComplexScript68);
            Text text43 = new Text();
            text43.Text = "- - -";

            run131.Append(runProperties129);
            run131.Append(text43);


            paragraph28.Append(paragraphProperties28);
            paragraph28.Append(run131);
 
            tableCell20.Append(tableCellProperties20);
            tableCell20.Append(paragraph28);

            TableCell tableCell21 = new TableCell();

            TableCellProperties tableCellProperties21 = new TableCellProperties();
            TableCellWidth tableCellWidth21 = new TableCellWidth() { Width = "5398", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge3 = new VerticalMerge() { Val = MergedCellValues.Restart };
            Shading shading15 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment20 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties21.Append(tableCellWidth21);
            tableCellProperties21.Append(verticalMerge3);
            tableCellProperties21.Append(shading15);
            tableCellProperties21.Append(tableCellVerticalAlignment20);

            Paragraph paragraph29 = new Paragraph() { RsidParagraphMarkRevision = "00C16BA1", RsidParagraphAddition = "001B2170", RsidParagraphProperties = "001B2170", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties29 = new ParagraphProperties();
            Indentation indentation29 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };

            ParagraphMarkRunProperties paragraphMarkRunProperties29 = new ParagraphMarkRunProperties();
            FontSize fontSize142 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript72 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties29.Append(fontSize142);
            paragraphMarkRunProperties29.Append(fontSizeComplexScript72);

            paragraphProperties29.Append(indentation29);
            paragraphProperties29.Append(paragraphMarkRunProperties29);  

            paragraph29.Append(paragraphProperties29);

            Paragraph paragraph30 = new Paragraph() { RsidParagraphMarkRevision = "00C16BA1", RsidParagraphAddition = "001B2170", RsidParagraphProperties = "001B2170", RsidRunAdditionDefault = "001B2170" };

            ParagraphProperties paragraphProperties30 = new ParagraphProperties();
            Indentation indentation30 = new Indentation() { FirstLine = "480" };

            ParagraphMarkRunProperties paragraphMarkRunProperties30 = new ParagraphMarkRunProperties();
            RunFonts runFonts170 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "宋体" };
            Color color54 = new Color() { Val = "000000" };
            Kern kern75 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize143 = new FontSize() { Val = "22" };

            paragraphMarkRunProperties30.Append(runFonts170);
            paragraphMarkRunProperties30.Append(color54);
            paragraphMarkRunProperties30.Append(kern75);
            paragraphMarkRunProperties30.Append(fontSize143);

            paragraphProperties30.Append(indentation30);
            paragraphProperties30.Append(paragraphMarkRunProperties30);

            Run run156 = new Run() { RsidRunProperties = "00C16BA1" };

            RunProperties runProperties152 = new RunProperties();
            RunFonts runFonts171 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "宋体" };
            Color color55 = new Color() { Val = "000000" };
            Kern kern76 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize144 = new FontSize() { Val = "24" };

            runProperties152.Append(runFonts171);
            runProperties152.Append(color55);
            runProperties152.Append(kern76);
            runProperties152.Append(fontSize144);
            Text text47 = new Text();
            text47.Text = "学业失败时的归因方式是外归因，是一种相对积极的归因方式。学习失败时归因于运气不佳，题目难度太大等外部因素时，" +
                "学生就不会产生内疚或羞愧的心理。即使考试不理想，也不会否认自己的学习能力，不会打击继续努力信心。这种归因方式对于学业成绩的提高有一定的推动作用。";

            run156.Append(runProperties152);
            run156.Append(text47);

            paragraph30.Append(paragraphProperties30);
            paragraph30.Append(run156);

            Paragraph paragraph31 = new Paragraph() { RsidParagraphMarkRevision = "007D1A31", RsidParagraphAddition = "001B2170", RsidParagraphProperties = "001B2170", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties31 = new ParagraphProperties();
            Indentation indentation31 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };

            ParagraphMarkRunProperties paragraphMarkRunProperties31 = new ParagraphMarkRunProperties();
            RunFonts runFonts172 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Color color56 = new Color() { Val = "000000" };
            Kern kern77 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize145 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript73 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties31.Append(runFonts172);
            paragraphMarkRunProperties31.Append(color56);
            paragraphMarkRunProperties31.Append(kern77);
            paragraphMarkRunProperties31.Append(fontSize145);
            paragraphMarkRunProperties31.Append(fontSizeComplexScript73);

            paragraphProperties31.Append(indentation31);
            paragraphProperties31.Append(paragraphMarkRunProperties31);

            Run run157 = new Run();

            RunProperties runProperties153 = new RunProperties();
            RunFonts runFonts173 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Color color57 = new Color() { Val = "000000" };
            Kern kern78 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize146 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript74 = new FontSizeComplexScript() { Val = "24" };

            runProperties153.Append(runFonts173);
            runProperties153.Append(color57);
            runProperties153.Append(kern78);
            runProperties153.Append(fontSize146);
            runProperties153.Append(fontSizeComplexScript74);
            FieldChar fieldChar15 = new FieldChar() { FieldCharType = FieldCharValues.End };

            run157.Append(runProperties153);
            run157.Append(fieldChar15);

            paragraph31.Append(paragraphProperties31);
            paragraph31.Append(run157);

            tableCell21.Append(tableCellProperties21);
            tableCell21.Append(paragraph29);
            tableCell21.Append(paragraph30);
            tableCell21.Append(paragraph31);

            tableRow5.Append(tableRowProperties5);
            tableRow5.Append(tableCell17);
            tableRow5.Append(tableCell18);
            tableRow5.Append(tableCell19);
            tableRow5.Append(tableCell20);
            tableRow5.Append(tableCell21);

            TableRow tableRow6 = new TableRow() { RsidTableRowMarkRevision = "007642DC", RsidTableRowAddition = "001B2170", RsidTableRowProperties = "001B2170" };

            TableRowProperties tableRowProperties6 = new TableRowProperties();
            TableRowHeight tableRowHeight6 = new TableRowHeight() { Val = (UInt32Value)1097U };

            tableRowProperties6.Append(tableRowHeight6);

            TableCell tableCell22 = new TableCell();

            TableCellProperties tableCellProperties22 = new TableCellProperties();
            TableCellWidth tableCellWidth22 = new TableCellWidth() { Width = "2702", Type = TableWidthUnitValues.Dxa };
            Shading shading16 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment21 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties22.Append(tableCellWidth22);
            tableCellProperties22.Append(shading16);
            tableCellProperties22.Append(tableCellVerticalAlignment21);

            Paragraph paragraph32 = new Paragraph() { RsidParagraphMarkRevision = "00973E52", RsidParagraphAddition = "001B2170", RsidParagraphProperties = "001B2170", RsidRunAdditionDefault = "001B2170" };

            ParagraphProperties paragraphProperties32 = new ParagraphProperties();
            Indentation indentation32 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification25 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties32 = new ParagraphMarkRunProperties();
            RunFonts runFonts174 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Kern kern79 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize147 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript75 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties32.Append(runFonts174);
            paragraphMarkRunProperties32.Append(kern79);
            paragraphMarkRunProperties32.Append(fontSize147);
            paragraphMarkRunProperties32.Append(fontSizeComplexScript75);

            paragraphProperties32.Append(indentation32);
            paragraphProperties32.Append(justification25);
            paragraphProperties32.Append(paragraphMarkRunProperties32);

            Run run158 = new Run();

            RunProperties runProperties154 = new RunProperties();
            RunFonts runFonts175 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Kern kern80 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize148 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript76 = new FontSizeComplexScript() { Val = "24" };

            runProperties154.Append(runFonts175);
            runProperties154.Append(kern80);
            runProperties154.Append(fontSize148);
            runProperties154.Append(fontSizeComplexScript76);
            Text text48 = new Text();
            text48.Text = "学业失败-外归因";

            run158.Append(runProperties154);
            run158.Append(text48);

            paragraph32.Append(paragraphProperties32);
            paragraph32.Append(run158);

            tableCell22.Append(tableCellProperties22);
            tableCell22.Append(paragraph32);

            TableCell tableCell23 = new TableCell();

            TableCellProperties tableCellProperties23 = new TableCellProperties();
            TableCellWidth tableCellWidth23 = new TableCellWidth() { Width = "685", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment22 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties23.Append(tableCellWidth23);
            tableCellProperties23.Append(tableCellVerticalAlignment22);

            Paragraph paragraph33 = new Paragraph() { RsidParagraphMarkRevision = "001B2170", RsidParagraphAddition = "001B2170", RsidParagraphProperties = "001B2170", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties33 = new ParagraphProperties();
            Indentation indentation33 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification26 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties33 = new ParagraphMarkRunProperties();
            FontSize fontSize153 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript81 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties33.Append(fontSize153);
            paragraphMarkRunProperties33.Append(fontSizeComplexScript81);

            paragraphProperties33.Append(indentation33);
            paragraphProperties33.Append(justification26);
            paragraphProperties33.Append(paragraphMarkRunProperties33);

            paragraph33.Append(paragraphProperties33);

            Paragraph paragraph34 = new Paragraph() { RsidParagraphMarkRevision = "001B2170", RsidParagraphAddition = "001B2170", RsidParagraphProperties = "001B2170", RsidRunAdditionDefault = "001B2170" };

            ParagraphProperties paragraphProperties34 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines4 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation34 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification27 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties34 = new ParagraphMarkRunProperties();
            RunFonts runFonts199 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "宋体" };
            Color color58 = new Color() { Val = "000000" };
            Kern kern85 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize175 = new FontSize() { Val = "22" };

            paragraphMarkRunProperties34.Append(runFonts199);
            paragraphMarkRunProperties34.Append(color58);
            paragraphMarkRunProperties34.Append(kern85);
            paragraphMarkRunProperties34.Append(fontSize175);

            paragraphProperties34.Append(spacingBetweenLines4);
            paragraphProperties34.Append(indentation34);
            paragraphProperties34.Append(justification27);
            paragraphProperties34.Append(paragraphMarkRunProperties34);

            Run run184 = new Run() { RsidRunProperties = "001B2170" };

            RunProperties runProperties180 = new RunProperties();
            RunFonts runFonts200 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "宋体" };
            Color color59 = new Color() { Val = "000000" };
            Kern kern86 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize176 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript82 = new FontSizeComplexScript() { Val = "24" };

            runProperties180.Append(runFonts200);
            runProperties180.Append(color59);
            runProperties180.Append(kern86);
            runProperties180.Append(fontSize176);
            runProperties180.Append(fontSizeComplexScript82);
            Text text53 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text53.Text = "36.76 ";

            run184.Append(runProperties180);
            run184.Append(text53);

            paragraph34.Append(paragraphProperties34);
            paragraph34.Append(run184);

            Paragraph paragraph35 = new Paragraph() { RsidParagraphMarkRevision = "001B2170", RsidParagraphAddition = "001B2170", RsidParagraphProperties = "001B2170", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties35 = new ParagraphProperties();
            Indentation indentation35 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification28 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties35 = new ParagraphMarkRunProperties();
            RunFonts runFonts201 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Bold bold43 = new Bold();
            Color color60 = new Color() { Val = "000000" };
            Kern kern87 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize177 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript83 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties35.Append(runFonts201);
            paragraphMarkRunProperties35.Append(bold43);
            paragraphMarkRunProperties35.Append(color60);
            paragraphMarkRunProperties35.Append(kern87);
            paragraphMarkRunProperties35.Append(fontSize177);
            paragraphMarkRunProperties35.Append(fontSizeComplexScript83);

            paragraphProperties35.Append(indentation35);
            paragraphProperties35.Append(justification28);
            paragraphProperties35.Append(paragraphMarkRunProperties35);

            Run run185 = new Run() { RsidRunProperties = "001B2170" };

            RunProperties runProperties181 = new RunProperties();
            RunFonts runFonts202 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Bold bold44 = new Bold();
            Color color61 = new Color() { Val = "000000" };
            Kern kern88 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize178 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript84 = new FontSizeComplexScript() { Val = "24" };

            runProperties181.Append(runFonts202);
            runProperties181.Append(bold44);
            runProperties181.Append(color61);
            runProperties181.Append(kern88);
            runProperties181.Append(fontSize178);
            runProperties181.Append(fontSizeComplexScript84);
            FieldChar fieldChar18 = new FieldChar() { FieldCharType = FieldCharValues.End };

            run185.Append(runProperties181);
            run185.Append(fieldChar18);

            paragraph35.Append(paragraphProperties35);
            paragraph35.Append(run185);

            tableCell23.Append(tableCellProperties23);
            tableCell23.Append(paragraph33);
            tableCell23.Append(paragraph34);
            tableCell23.Append(paragraph35);

            TableCell tableCell24 = new TableCell();

            TableCellProperties tableCellProperties24 = new TableCellProperties();
            TableCellWidth tableCellWidth24 = new TableCellWidth() { Width = "1087", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment23 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties24.Append(tableCellWidth24);
            tableCellProperties24.Append(tableCellVerticalAlignment23);

            Paragraph paragraph36 = new Paragraph() { RsidParagraphMarkRevision = "001B2170", RsidParagraphAddition = "001B2170", RsidParagraphProperties = "001B2170", RsidRunAdditionDefault = "001B2170" };

            ParagraphProperties paragraphProperties36 = new ParagraphProperties();
            Indentation indentation36 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification29 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties36 = new ParagraphMarkRunProperties();
            RunFonts runFonts203 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Bold bold45 = new Bold();
            Color color62 = new Color() { Val = "000000" };
            Kern kern89 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize179 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript85 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties36.Append(runFonts203);
            paragraphMarkRunProperties36.Append(bold45);
            paragraphMarkRunProperties36.Append(color62);
            paragraphMarkRunProperties36.Append(kern89);
            paragraphMarkRunProperties36.Append(fontSize179);
            paragraphMarkRunProperties36.Append(fontSizeComplexScript85);

            paragraphProperties36.Append(indentation36);
            paragraphProperties36.Append(justification29);
            paragraphProperties36.Append(paragraphMarkRunProperties36);

            Run run186 = new Run() { RsidRunProperties = "001B2170" };

            RunProperties runProperties182 = new RunProperties();
            RunFonts runFonts204 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Bold bold46 = new Bold();
            Color color63 = new Color() { Val = "000000" };
            Kern kern90 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize180 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript86 = new FontSizeComplexScript() { Val = "24" };

            runProperties182.Append(runFonts204);
            runProperties182.Append(bold46);
            runProperties182.Append(color63);
            runProperties182.Append(kern90);
            runProperties182.Append(fontSize180);
            runProperties182.Append(fontSizeComplexScript86);
            Text text54 = new Text();
            text54.Text = "低水平";

            run186.Append(runProperties182);
            run186.Append(text54);    

            paragraph36.Append(paragraphProperties36);
            paragraph36.Append(run186);

            tableCell24.Append(tableCellProperties24);
            tableCell24.Append(paragraph36);

            TableCell tableCell25 = new TableCell();

            TableCellProperties tableCellProperties25 = new TableCellProperties();
            TableCellWidth tableCellWidth25 = new TableCellWidth() { Width = "886", Type = TableWidthUnitValues.Dxa };
            Shading shading17 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment24 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties25.Append(tableCellWidth25);
            tableCellProperties25.Append(shading17);
            tableCellProperties25.Append(tableCellVerticalAlignment24);

            Paragraph paragraph37 = new Paragraph() { RsidParagraphMarkRevision = "001B2170", RsidParagraphAddition = "001B2170", RsidParagraphProperties = "001B2170", RsidRunAdditionDefault = "001B2170" };

            ParagraphProperties paragraphProperties37 = new ParagraphProperties();
            Indentation indentation37 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification30 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties37 = new ParagraphMarkRunProperties();
            RunFonts runFonts206 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Bold bold48 = new Bold();
            Color color65 = new Color() { Val = "000000" };
            Kern kern92 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize182 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript88 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties37.Append(runFonts206);
            paragraphMarkRunProperties37.Append(bold48);
            paragraphMarkRunProperties37.Append(color65);
            paragraphMarkRunProperties37.Append(kern92);
            paragraphMarkRunProperties37.Append(fontSize182);
            paragraphMarkRunProperties37.Append(fontSizeComplexScript88);

            paragraphProperties37.Append(indentation37);
            paragraphProperties37.Append(justification30);
            paragraphProperties37.Append(paragraphMarkRunProperties37);

            Run run188 = new Run() { RsidRunProperties = "001B2170" };

            RunProperties runProperties184 = new RunProperties();
            RunFonts runFonts207 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Color color66 = new Color() { Val = "000000" };
            Kern kern93 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize183 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript89 = new FontSizeComplexScript() { Val = "24" };

            runProperties184.Append(runFonts207);
            runProperties184.Append(color66);
            runProperties184.Append(kern93);
            runProperties184.Append(fontSize183);
            runProperties184.Append(fontSizeComplexScript89);
            Text text56 = new Text();
            text56.Text = "- - -";

            run188.Append(runProperties184);
            run188.Append(text56);

            paragraph37.Append(paragraphProperties37);
            paragraph37.Append(run188);

            tableCell25.Append(tableCellProperties25);
            tableCell25.Append(paragraph37);

            TableCell tableCell26 = new TableCell();

            TableCellProperties tableCellProperties26 = new TableCellProperties();
            TableCellWidth tableCellWidth26 = new TableCellWidth() { Width = "5398", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge4 = new VerticalMerge();
            Shading shading18 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment25 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties26.Append(tableCellWidth26);
            tableCellProperties26.Append(verticalMerge4);
            tableCellProperties26.Append(shading18);
            tableCellProperties26.Append(tableCellVerticalAlignment25);

            Paragraph paragraph38 = new Paragraph() { RsidParagraphMarkRevision = "002A093C", RsidParagraphAddition = "001B2170", RsidParagraphProperties = "001B2170", RsidRunAdditionDefault = "001B2170" };

            ParagraphProperties paragraphProperties38 = new ParagraphProperties();
            Indentation indentation38 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };

            ParagraphMarkRunProperties paragraphMarkRunProperties38 = new ParagraphMarkRunProperties();
            RunFonts runFonts211 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Color color70 = new Color() { Val = "000000" };
            Kern kern97 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize187 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript93 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties38.Append(runFonts211);
            paragraphMarkRunProperties38.Append(color70);
            paragraphMarkRunProperties38.Append(kern97);
            paragraphMarkRunProperties38.Append(fontSize187);
            paragraphMarkRunProperties38.Append(fontSizeComplexScript93);

            paragraphProperties38.Append(indentation38);
            paragraphProperties38.Append(paragraphMarkRunProperties38);

            paragraph38.Append(paragraphProperties38);

            tableCell26.Append(tableCellProperties26);
            tableCell26.Append(paragraph38);

            tableRow6.Append(tableRowProperties6);
            tableRow6.Append(tableCell22);
            tableRow6.Append(tableCell23);
            tableRow6.Append(tableCell24);
            tableRow6.Append(tableCell25);
            tableRow6.Append(tableCell26);

            table1.Append(tableProperties1);
            table1.Append(tableGrid1);
            table1.Append(tableRow1);
            table1.Append(tableRow2);
            table1.Append(tableRow3);
            table1.Append(tableRow4);
            table1.Append(tableRow5);
            table1.Append(tableRow6);
            return table1;

        }
    }
}

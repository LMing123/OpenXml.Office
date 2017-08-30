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
    using Lsj.Util.Collections;

    public partial class GeneratedClass
    {
        public Table GenerateTable4(string title, string evaluate, SafeDictionary<string, (string, string, eInfluence, double)> contents)
        {


            Table table1 = new Table();

            TableProperties tableProperties1 = new TableProperties();
            TablePositionProperties tablePositionProperties1 = new TablePositionProperties() { LeftFromText = 180, RightFromText = 180, VerticalAnchor = VerticalAnchorValues.Text, HorizontalAnchor = HorizontalAnchorValues.Margin, TablePositionXAlignment = HorizontalAlignmentValues.Center, TablePositionY = -1 };
            TableWidth tableWidth1 = new TableWidth() { Width = "10713", Type = TableWidthUnitValues.Dxa };

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
            GridColumn gridColumn1 = new GridColumn() { Width = "2690" };
            GridColumn gridColumn2 = new GridColumn() { Width = "682" };
            GridColumn gridColumn3 = new GridColumn() { Width = "1082" };
            GridColumn gridColumn4 = new GridColumn() { Width = "882" };
            GridColumn gridColumn5 = new GridColumn() { Width = "5377" };

            tableGrid1.Append(gridColumn1);
            tableGrid1.Append(gridColumn2);
            tableGrid1.Append(gridColumn3);
            tableGrid1.Append(gridColumn4);
            tableGrid1.Append(gridColumn5);

            TableRow tableRow1 = new TableRow() { RsidTableRowMarkRevision = "007642DC", RsidTableRowAddition = "001D4055", RsidTableRowProperties = "006A0186" };

            TablePropertyExceptions tablePropertyExceptions1 = new TablePropertyExceptions();

            TableCellMarginDefault tableCellMarginDefault2 = new TableCellMarginDefault();
            TopMargin topMargin1 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            BottomMargin bottomMargin1 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };

            tableCellMarginDefault2.Append(topMargin1);
            tableCellMarginDefault2.Append(bottomMargin1);

            tablePropertyExceptions1.Append(tableCellMarginDefault2);

            TableRowProperties tableRowProperties1 = new TableRowProperties();
            TableRowHeight tableRowHeight1 = new TableRowHeight() { Val = (UInt32Value)505U };

            tableRowProperties1.Append(tableRowHeight1);

            TableCell tableCell1 = new TableCell();

            TableCellProperties tableCellProperties1 = new TableCellProperties();
            TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = "10713", Type = TableWidthUnitValues.Dxa };
            GridSpan gridSpan1 = new GridSpan() { Val = 5 };
            Shading shading1 = new Shading() { Val = ShadingPatternValues.Clear, Color = "000000", Fill = "D99795" };

            tableCellProperties1.Append(tableCellWidth1);
            tableCellProperties1.Append(gridSpan1);
            tableCellProperties1.Append(shading1);

            Paragraph paragraph1 = new Paragraph() { RsidParagraphMarkRevision = "007642DC", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "000C649A" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
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
            LastRenderedPageBreak lastRenderedPageBreak1 = new LastRenderedPageBreak();
            Text text1 = new Text();
            text1.Text = "学业归因量表";

            run1.Append(runProperties1);
            run1.Append(lastRenderedPageBreak1);
            run1.Append(text1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            tableCell1.Append(tableCellProperties1);
            tableCell1.Append(paragraph1);

            tableRow1.Append(tablePropertyExceptions1);
            tableRow1.Append(tableRowProperties1);
            tableRow1.Append(tableCell1);

            TableRow tableRow2 = new TableRow() { RsidTableRowMarkRevision = "007642DC", RsidTableRowAddition = "001D4055", RsidTableRowProperties = "006A0186" };

            TablePropertyExceptions tablePropertyExceptions2 = new TablePropertyExceptions();

            TableCellMarginDefault tableCellMarginDefault3 = new TableCellMarginDefault();
            TopMargin topMargin2 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            BottomMargin bottomMargin2 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };

            tableCellMarginDefault3.Append(topMargin2);
            tableCellMarginDefault3.Append(bottomMargin2);

            tablePropertyExceptions2.Append(tableCellMarginDefault3);

            TableRowProperties tableRowProperties2 = new TableRowProperties();
            TableRowHeight tableRowHeight2 = new TableRowHeight() { Val = (UInt32Value)505U };

            tableRowProperties2.Append(tableRowHeight2);

            TableCell tableCell2 = new TableCell();

            TableCellProperties tableCellProperties2 = new TableCellProperties();
            TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = "2690", Type = TableWidthUnitValues.Dxa };
            Shading shading2 = new Shading() { Val = ShadingPatternValues.Clear, Color = "000000", Fill = "D99795" };
            TableCellVerticalAlignment tableCellVerticalAlignment1 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties2.Append(tableCellWidth2);
            tableCellProperties2.Append(shading2);
            tableCellProperties2.Append(tableCellVerticalAlignment1);

            Paragraph paragraph2 = new Paragraph() { RsidParagraphMarkRevision = "007642DC", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "000C649A" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            Justification justification2 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
            RunFonts runFonts3 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Bold bold3 = new Bold();
            Kern kern3 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize3 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties2.Append(runFonts3);
            paragraphMarkRunProperties2.Append(bold3);
            paragraphMarkRunProperties2.Append(kern3);
            paragraphMarkRunProperties2.Append(fontSize3);
            paragraphMarkRunProperties2.Append(fontSizeComplexScript3);

            paragraphProperties2.Append(justification2);
            paragraphProperties2.Append(paragraphMarkRunProperties2);

            Run run2 = new Run() { RsidRunProperties = "007642DC" };

            RunProperties runProperties2 = new RunProperties();
            RunFonts runFonts4 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Bold bold4 = new Bold();
            Kern kern4 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize4 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "24" };

            runProperties2.Append(runFonts4);
            runProperties2.Append(bold4);
            runProperties2.Append(kern4);
            runProperties2.Append(fontSize4);
            runProperties2.Append(fontSizeComplexScript4);
            Text text2 = new Text();
            text2.Text = "维度名称";

            run2.Append(runProperties2);
            run2.Append(text2);

            paragraph2.Append(paragraphProperties2);
            paragraph2.Append(run2);

            tableCell2.Append(tableCellProperties2);
            tableCell2.Append(paragraph2);

            TableCell tableCell3 = new TableCell();

            TableCellProperties tableCellProperties3 = new TableCellProperties();
            TableCellWidth tableCellWidth3 = new TableCellWidth() { Width = "682", Type = TableWidthUnitValues.Dxa };
            Shading shading3 = new Shading() { Val = ShadingPatternValues.Clear, Color = "000000", Fill = "D99795" };
            TableCellVerticalAlignment tableCellVerticalAlignment2 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties3.Append(tableCellWidth3);
            tableCellProperties3.Append(shading3);
            tableCellProperties3.Append(tableCellVerticalAlignment2);

            Paragraph paragraph3 = new Paragraph() { RsidParagraphMarkRevision = "007642DC", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "000C649A" };

            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
            Justification justification3 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
            RunFonts runFonts5 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Bold bold5 = new Bold();
            Kern kern5 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize5 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties3.Append(runFonts5);
            paragraphMarkRunProperties3.Append(bold5);
            paragraphMarkRunProperties3.Append(kern5);
            paragraphMarkRunProperties3.Append(fontSize5);
            paragraphMarkRunProperties3.Append(fontSizeComplexScript5);

            paragraphProperties3.Append(justification3);
            paragraphProperties3.Append(paragraphMarkRunProperties3);

            Run run3 = new Run() { RsidRunProperties = "007642DC" };

            RunProperties runProperties3 = new RunProperties();
            RunFonts runFonts6 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Bold bold6 = new Bold();
            Kern kern6 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize6 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "24" };

            runProperties3.Append(runFonts6);
            runProperties3.Append(bold6);
            runProperties3.Append(kern6);
            runProperties3.Append(fontSize6);
            runProperties3.Append(fontSizeComplexScript6);
            Text text3 = new Text();
            text3.Text = "得分";

            run3.Append(runProperties3);
            run3.Append(text3);

            paragraph3.Append(paragraphProperties3);
            paragraph3.Append(run3);

            tableCell3.Append(tableCellProperties3);
            tableCell3.Append(paragraph3);

            TableCell tableCell4 = new TableCell();

            TableCellProperties tableCellProperties4 = new TableCellProperties();
            TableCellWidth tableCellWidth4 = new TableCellWidth() { Width = "1082", Type = TableWidthUnitValues.Dxa };
            Shading shading4 = new Shading() { Val = ShadingPatternValues.Clear, Color = "000000", Fill = "D99795" };
            TableCellVerticalAlignment tableCellVerticalAlignment3 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties4.Append(tableCellWidth4);
            tableCellProperties4.Append(shading4);
            tableCellProperties4.Append(tableCellVerticalAlignment3);

            Paragraph paragraph4 = new Paragraph() { RsidParagraphMarkRevision = "007642DC", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "000C649A" };

            ParagraphProperties paragraphProperties4 = new ParagraphProperties();
            Justification justification4 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties4 = new ParagraphMarkRunProperties();
            RunFonts runFonts7 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Bold bold7 = new Bold();
            Kern kern7 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize7 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript7 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties4.Append(runFonts7);
            paragraphMarkRunProperties4.Append(bold7);
            paragraphMarkRunProperties4.Append(kern7);
            paragraphMarkRunProperties4.Append(fontSize7);
            paragraphMarkRunProperties4.Append(fontSizeComplexScript7);

            paragraphProperties4.Append(justification4);
            paragraphProperties4.Append(paragraphMarkRunProperties4);

            Run run4 = new Run();

            RunProperties runProperties4 = new RunProperties();
            RunFonts runFonts8 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Bold bold8 = new Bold();
            Kern kern8 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize8 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript8 = new FontSizeComplexScript() { Val = "24" };

            runProperties4.Append(runFonts8);
            runProperties4.Append(bold8);
            runProperties4.Append(kern8);
            runProperties4.Append(fontSize8);
            runProperties4.Append(fontSizeComplexScript8);
            Text text4 = new Text();
            text4.Text = "水平";

            run4.Append(runProperties4);
            run4.Append(text4);

            paragraph4.Append(paragraphProperties4);
            paragraph4.Append(run4);

            tableCell4.Append(tableCellProperties4);
            tableCell4.Append(paragraph4);

            TableCell tableCell5 = new TableCell();

            TableCellProperties tableCellProperties5 = new TableCellProperties();
            TableCellWidth tableCellWidth5 = new TableCellWidth() { Width = "882", Type = TableWidthUnitValues.Dxa };
            Shading shading5 = new Shading() { Val = ShadingPatternValues.Clear, Color = "000000", Fill = "D99795" };
            TableCellVerticalAlignment tableCellVerticalAlignment4 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties5.Append(tableCellWidth5);
            tableCellProperties5.Append(shading5);
            tableCellProperties5.Append(tableCellVerticalAlignment4);

            Paragraph paragraph5 = new Paragraph() { RsidParagraphMarkRevision = "007642DC", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "000C649A" };

            ParagraphProperties paragraphProperties5 = new ParagraphProperties();
            Justification justification5 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties5 = new ParagraphMarkRunProperties();
            RunFonts runFonts9 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Bold bold9 = new Bold();
            Kern kern9 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize9 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript9 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties5.Append(runFonts9);
            paragraphMarkRunProperties5.Append(bold9);
            paragraphMarkRunProperties5.Append(kern9);
            paragraphMarkRunProperties5.Append(fontSize9);
            paragraphMarkRunProperties5.Append(fontSizeComplexScript9);

            paragraphProperties5.Append(justification5);
            paragraphProperties5.Append(paragraphMarkRunProperties5);

            Run run5 = new Run();

            RunProperties runProperties5 = new RunProperties();
            RunFonts runFonts10 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Bold bold10 = new Bold();
            Kern kern10 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize10 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript10 = new FontSizeComplexScript() { Val = "24" };

            runProperties5.Append(runFonts10);
            runProperties5.Append(bold10);
            runProperties5.Append(kern10);
            runProperties5.Append(fontSize10);
            runProperties5.Append(fontSizeComplexScript10);
            Text text5 = new Text();
            text5.Text = "对学习的影响";

            run5.Append(runProperties5);
            run5.Append(text5);

            paragraph5.Append(paragraphProperties5);
            paragraph5.Append(run5);

            tableCell5.Append(tableCellProperties5);
            tableCell5.Append(paragraph5);

            TableCell tableCell6 = new TableCell();

            TableCellProperties tableCellProperties6 = new TableCellProperties();
            TableCellWidth tableCellWidth6 = new TableCellWidth() { Width = "5375", Type = TableWidthUnitValues.Dxa };
            Shading shading6 = new Shading() { Val = ShadingPatternValues.Clear, Color = "000000", Fill = "D99795" };
            TableCellVerticalAlignment tableCellVerticalAlignment5 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties6.Append(tableCellWidth6);
            tableCellProperties6.Append(shading6);
            tableCellProperties6.Append(tableCellVerticalAlignment5);

            Paragraph paragraph6 = new Paragraph() { RsidParagraphMarkRevision = "007642DC", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "000C649A" };

            ParagraphProperties paragraphProperties6 = new ParagraphProperties();
            Justification justification6 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties6 = new ParagraphMarkRunProperties();
            RunFonts runFonts11 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Bold bold11 = new Bold();
            Kern kern11 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize11 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript11 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties6.Append(runFonts11);
            paragraphMarkRunProperties6.Append(bold11);
            paragraphMarkRunProperties6.Append(kern11);
            paragraphMarkRunProperties6.Append(fontSize11);
            paragraphMarkRunProperties6.Append(fontSizeComplexScript11);

            paragraphProperties6.Append(justification6);
            paragraphProperties6.Append(paragraphMarkRunProperties6);

            Run run6 = new Run();

            RunProperties runProperties6 = new RunProperties();
            RunFonts runFonts12 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Bold bold12 = new Bold();
            Kern kern12 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize12 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript12 = new FontSizeComplexScript() { Val = "24" };

            runProperties6.Append(runFonts12);
            runProperties6.Append(bold12);
            runProperties6.Append(kern12);
            runProperties6.Append(fontSize12);
            runProperties6.Append(fontSizeComplexScript12);
            Text text6 = new Text();
            text6.Text = "归因方式";

            run6.Append(runProperties6);
            run6.Append(text6);

            paragraph6.Append(paragraphProperties6);
            paragraph6.Append(run6);

            tableCell6.Append(tableCellProperties6);
            tableCell6.Append(paragraph6);

            tableRow2.Append(tablePropertyExceptions2);
            tableRow2.Append(tableRowProperties2);
            tableRow2.Append(tableCell2);
            tableRow2.Append(tableCell3);
            tableRow2.Append(tableCell4);
            tableRow2.Append(tableCell5);
            tableRow2.Append(tableCell6);

            TableRow tableRow3 = new TableRow() { RsidTableRowMarkRevision = "007642DC", RsidTableRowAddition = "001D4055", RsidTableRowProperties = "00500C13" };

            TablePropertyExceptions tablePropertyExceptions3 = new TablePropertyExceptions();

            TableCellMarginDefault tableCellMarginDefault4 = new TableCellMarginDefault();
            TopMargin topMargin3 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            BottomMargin bottomMargin3 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };

            tableCellMarginDefault4.Append(topMargin3);
            tableCellMarginDefault4.Append(bottomMargin3);

            tablePropertyExceptions3.Append(tableCellMarginDefault4);

            TableRowProperties tableRowProperties3 = new TableRowProperties();
            TableRowHeight tableRowHeight3 = new TableRowHeight() { Val = (UInt32Value)2806U };

            tableRowProperties3.Append(tableRowHeight3);

            TableCell tableCell7 = new TableCell();

            TableCellProperties tableCellProperties7 = new TableCellProperties();
            TableCellWidth tableCellWidth7 = new TableCellWidth() { Width = "2690", Type = TableWidthUnitValues.Dxa };
            Shading shading7 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment6 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties7.Append(tableCellWidth7);
            tableCellProperties7.Append(shading7);
            tableCellProperties7.Append(tableCellVerticalAlignment6);

            Paragraph paragraph7 = new Paragraph() { RsidParagraphMarkRevision = "004F3E28", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "000C649A" };

            ParagraphProperties paragraphProperties7 = new ParagraphProperties();
            Justification justification7 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties7 = new ParagraphMarkRunProperties();
            RunFonts runFonts13 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Bold bold13 = new Bold();
            Color color1 = new Color() { Val = "000000" };
            Kern kern13 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize13 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript13 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties7.Append(runFonts13);
            paragraphMarkRunProperties7.Append(bold13);
            paragraphMarkRunProperties7.Append(color1);
            paragraphMarkRunProperties7.Append(kern13);
            paragraphMarkRunProperties7.Append(fontSize13);
            paragraphMarkRunProperties7.Append(fontSizeComplexScript13);

            paragraphProperties7.Append(justification7);
            paragraphProperties7.Append(paragraphMarkRunProperties7);

            Run run7 = new Run();

            RunProperties runProperties7 = new RunProperties();
            RunFonts runFonts14 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Bold bold14 = new Bold();
            Color color2 = new Color() { Val = "000000" };
            Kern kern14 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize14 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript14 = new FontSizeComplexScript() { Val = "24" };

            runProperties7.Append(runFonts14);
            runProperties7.Append(bold14);
            runProperties7.Append(color2);
            runProperties7.Append(kern14);
            runProperties7.Append(fontSize14);
            runProperties7.Append(fontSizeComplexScript14);
            Text text7 = new Text();
            text7.Text = "学业成功";

            run7.Append(runProperties7);
            run7.Append(text7);

            Run run8 = new Run();

            RunProperties runProperties8 = new RunProperties();
            RunFonts runFonts15 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Bold bold15 = new Bold();
            Color color3 = new Color() { Val = "000000" };
            Kern kern15 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize15 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript15 = new FontSizeComplexScript() { Val = "24" };

            runProperties8.Append(runFonts15);
            runProperties8.Append(bold15);
            runProperties8.Append(color3);
            runProperties8.Append(kern15);
            runProperties8.Append(fontSize15);
            runProperties8.Append(fontSizeComplexScript15);
            Text text8 = new Text();
            text8.Text = "-";

            run8.Append(runProperties8);
            run8.Append(text8);

            Run run9 = new Run();

            RunProperties runProperties9 = new RunProperties();
            RunFonts runFonts16 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Bold bold16 = new Bold();
            Color color4 = new Color() { Val = "000000" };
            Kern kern16 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize16 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript16 = new FontSizeComplexScript() { Val = "24" };

            runProperties9.Append(runFonts16);
            runProperties9.Append(bold16);
            runProperties9.Append(color4);
            runProperties9.Append(kern16);
            runProperties9.Append(fontSize16);
            runProperties9.Append(fontSizeComplexScript16);
            Text text9 = new Text();
            text9.Text = "内归因";

            run9.Append(runProperties9);
            run9.Append(text9);

            paragraph7.Append(paragraphProperties7);
            paragraph7.Append(run7);
            paragraph7.Append(run8);
            paragraph7.Append(run9);

            tableCell7.Append(tableCellProperties7);
            tableCell7.Append(paragraph7);

            TableCell tableCell8 = new TableCell();

            TableCellProperties tableCellProperties8 = new TableCellProperties();
            TableCellWidth tableCellWidth8 = new TableCellWidth() { Width = "682", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment7 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties8.Append(tableCellWidth8);
            tableCellProperties8.Append(tableCellVerticalAlignment7);

            Paragraph paragraph8 = new Paragraph() { RsidParagraphMarkRevision = "001B2170", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "000C649A" };

            ParagraphProperties paragraphProperties8 = new ParagraphProperties();
            Justification justification8 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties8 = new ParagraphMarkRunProperties();
            FontSize fontSize17 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript17 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties8.Append(fontSize17);
            paragraphMarkRunProperties8.Append(fontSizeComplexScript17);

            paragraphProperties8.Append(justification8);
            paragraphProperties8.Append(paragraphMarkRunProperties8);

            paragraph8.Append(paragraphProperties8);

            Paragraph paragraph9 = new Paragraph() { RsidParagraphMarkRevision = "001B2170", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "000C649A" };

            ParagraphProperties paragraphProperties9 = new ParagraphProperties();
            Justification justification9 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties9 = new ParagraphMarkRunProperties();
            RunFonts runFonts17 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "宋体" };
            Color color5 = new Color() { Val = "000000" };
            Kern kern17 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize18 = new FontSize() { Val = "22" };

            paragraphMarkRunProperties9.Append(runFonts17);
            paragraphMarkRunProperties9.Append(color5);
            paragraphMarkRunProperties9.Append(kern17);
            paragraphMarkRunProperties9.Append(fontSize18);

            paragraphProperties9.Append(justification9);
            paragraphProperties9.Append(paragraphMarkRunProperties9);

            Run run10 = new Run() { RsidRunProperties = "001B2170" };

            RunProperties runProperties10 = new RunProperties();
            RunFonts runFonts18 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "宋体" };
            Color color6 = new Color() { Val = "000000" };
            Kern kern18 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize19 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript18 = new FontSizeComplexScript() { Val = "24" };

            runProperties10.Append(runFonts18);
            runProperties10.Append(color6);
            runProperties10.Append(kern18);
            runProperties10.Append(fontSize19);
            runProperties10.Append(fontSizeComplexScript18);
            Text text10 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text10.Text = "61.40 ";

            run10.Append(runProperties10);
            run10.Append(text10);

            paragraph9.Append(paragraphProperties9);
            paragraph9.Append(run10);

            Paragraph paragraph10 = new Paragraph() { RsidParagraphMarkRevision = "001B2170", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "000C649A" };

            ParagraphProperties paragraphProperties10 = new ParagraphProperties();
            Justification justification10 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties10 = new ParagraphMarkRunProperties();
            RunFonts runFonts19 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Bold bold17 = new Bold();
            Color color7 = new Color() { Val = "000000" };
            Kern kern19 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize20 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript19 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties10.Append(runFonts19);
            paragraphMarkRunProperties10.Append(bold17);
            paragraphMarkRunProperties10.Append(color7);
            paragraphMarkRunProperties10.Append(kern19);
            paragraphMarkRunProperties10.Append(fontSize20);
            paragraphMarkRunProperties10.Append(fontSizeComplexScript19);

            paragraphProperties10.Append(justification10);
            paragraphProperties10.Append(paragraphMarkRunProperties10);

            paragraph10.Append(paragraphProperties10);

            tableCell8.Append(tableCellProperties8);
            tableCell8.Append(paragraph8);
            tableCell8.Append(paragraph9);
            tableCell8.Append(paragraph10);

            TableCell tableCell9 = new TableCell();

            TableCellProperties tableCellProperties9 = new TableCellProperties();
            TableCellWidth tableCellWidth9 = new TableCellWidth() { Width = "1082", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment8 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties9.Append(tableCellWidth9);
            tableCellProperties9.Append(tableCellVerticalAlignment8);

            Paragraph paragraph11 = new Paragraph() { RsidParagraphMarkRevision = "001B2170", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "000C649A" };

            ParagraphProperties paragraphProperties11 = new ParagraphProperties();
            Justification justification11 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties11 = new ParagraphMarkRunProperties();
            RunFonts runFonts20 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Bold bold18 = new Bold();
            Color color8 = new Color() { Val = "000000" };
            Kern kern20 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize21 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript20 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties11.Append(runFonts20);
            paragraphMarkRunProperties11.Append(bold18);
            paragraphMarkRunProperties11.Append(color8);
            paragraphMarkRunProperties11.Append(kern20);
            paragraphMarkRunProperties11.Append(fontSize21);
            paragraphMarkRunProperties11.Append(fontSizeComplexScript20);

            paragraphProperties11.Append(justification11);
            paragraphProperties11.Append(paragraphMarkRunProperties11);

            Run run11 = new Run() { RsidRunProperties = "001B2170" };

            RunProperties runProperties11 = new RunProperties();
            RunFonts runFonts21 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Bold bold19 = new Bold();
            Color color9 = new Color() { Val = "000000" };
            Kern kern21 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize22 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript21 = new FontSizeComplexScript() { Val = "24" };

            runProperties11.Append(runFonts21);
            runProperties11.Append(bold19);
            runProperties11.Append(color9);
            runProperties11.Append(kern21);
            runProperties11.Append(fontSize22);
            runProperties11.Append(fontSizeComplexScript21);
            Text text11 = new Text();
            text11.Text = "高水平";

            run11.Append(runProperties11);
            run11.Append(text11);

            paragraph11.Append(paragraphProperties11);
            paragraph11.Append(run11);

            tableCell9.Append(tableCellProperties9);
            tableCell9.Append(paragraph11);

            TableCell tableCell10 = new TableCell();

            TableCellProperties tableCellProperties10 = new TableCellProperties();
            TableCellWidth tableCellWidth10 = new TableCellWidth() { Width = "882", Type = TableWidthUnitValues.Dxa };
            Shading shading8 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment9 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties10.Append(tableCellWidth10);
            tableCellProperties10.Append(shading8);
            tableCellProperties10.Append(tableCellVerticalAlignment9);

            Paragraph paragraph12 = new Paragraph() { RsidParagraphMarkRevision = "001B2170", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "000C649A" };

            ParagraphProperties paragraphProperties12 = new ParagraphProperties();
            Justification justification12 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties12 = new ParagraphMarkRunProperties();
            RunFonts runFonts22 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Bold bold20 = new Bold();
            Color color10 = new Color() { Val = "000000" };
            Kern kern22 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize23 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript22 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties12.Append(runFonts22);
            paragraphMarkRunProperties12.Append(bold20);
            paragraphMarkRunProperties12.Append(color10);
            paragraphMarkRunProperties12.Append(kern22);
            paragraphMarkRunProperties12.Append(fontSize23);
            paragraphMarkRunProperties12.Append(fontSizeComplexScript22);

            paragraphProperties12.Append(justification12);
            paragraphProperties12.Append(paragraphMarkRunProperties12);

            Run run12 = new Run() { RsidRunProperties = "001B2170" };

            RunProperties runProperties12 = new RunProperties();
            RunFonts runFonts23 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Color color11 = new Color() { Val = "000000" };
            Kern kern23 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize24 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript23 = new FontSizeComplexScript() { Val = "24" };

            runProperties12.Append(runFonts23);
            runProperties12.Append(color11);
            runProperties12.Append(kern23);
            runProperties12.Append(fontSize24);
            runProperties12.Append(fontSizeComplexScript23);
            Text text12 = new Text();
            text12.Text = "- - -";

            run12.Append(runProperties12);
            run12.Append(text12);

            paragraph12.Append(paragraphProperties12);
            paragraph12.Append(run12);

            tableCell10.Append(tableCellProperties10);
            tableCell10.Append(paragraph12);

            TableCell tableCell11 = new TableCell();

            TableCellProperties tableCellProperties11 = new TableCellProperties();
            TableCellWidth tableCellWidth11 = new TableCellWidth() { Width = "5375", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge1 = new VerticalMerge() { Val = MergedCellValues.Restart };
            Shading shading9 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment10 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties11.Append(tableCellWidth11);
            tableCellProperties11.Append(verticalMerge1);
            tableCellProperties11.Append(shading9);
            tableCellProperties11.Append(tableCellVerticalAlignment10);

            Paragraph paragraph13 = new Paragraph() { RsidParagraphMarkRevision = "00C16BA1", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "000C649A" };

            ParagraphProperties paragraphProperties13 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties13 = new ParagraphMarkRunProperties();
            FontSize fontSize25 = new FontSize() { Val = "4" };
            FontSizeComplexScript fontSizeComplexScript24 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties13.Append(fontSize25);
            paragraphMarkRunProperties13.Append(fontSizeComplexScript24);

            paragraphProperties13.Append(paragraphMarkRunProperties13);

            paragraph13.Append(paragraphProperties13);

            Paragraph paragraph14 = new Paragraph() { RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "000C649A" };

            ParagraphProperties paragraphProperties14 = new ParagraphProperties();
            Indentation indentation1 = new Indentation() { FirstLine = "480" };

            ParagraphMarkRunProperties paragraphMarkRunProperties14 = new ParagraphMarkRunProperties();
            RunFonts runFonts24 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "宋体" };
            Color color12 = new Color() { Val = "000000" };
            Kern kern24 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize26 = new FontSize() { Val = "24" };

            paragraphMarkRunProperties14.Append(runFonts24);
            paragraphMarkRunProperties14.Append(color12);
            paragraphMarkRunProperties14.Append(kern24);
            paragraphMarkRunProperties14.Append(fontSize26);

            paragraphProperties14.Append(indentation1);
            paragraphProperties14.Append(paragraphMarkRunProperties14);

            Run run13 = new Run() { RsidRunProperties = "00C16BA1" };

            RunProperties runProperties13 = new RunProperties();
            RunFonts runFonts25 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "宋体" };
            Color color13 = new Color() { Val = "000000" };
            Kern kern25 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize27 = new FontSize() { Val = "24" };

            runProperties13.Append(runFonts25);
            runProperties13.Append(color13);
            runProperties13.Append(kern25);
            runProperties13.Append(fontSize27);
            Text text13 = new Text();
            text13.Text = "学业成功时的归因方式是内归因，是一种比较积极的归因方式。倾向于把学业成功归因于自身学习努力和学习能力这样的可控因素，是对自己能力的一种肯定，他们对于学习能够更有掌控感和成就感，从而增强了对于学习的控制和对成功的渴望，那么他们就会积极地争取成功，会感到满意、自豪等积极情绪，有助于维持和激发其随后的学习动机，他们就更加愿意去学习。在这种情况下，更容易完成我们所设的任务，而不是选择逃避。";

            run13.Append(runProperties13);
            run13.Append(text13);

            paragraph14.Append(paragraphProperties14);
            paragraph14.Append(run13);

            Paragraph paragraph15 = new Paragraph() { RsidParagraphMarkRevision = "00C16BA1", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "000C649A" };

            ParagraphProperties paragraphProperties15 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties15 = new ParagraphMarkRunProperties();
            FontSize fontSize28 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript25 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties15.Append(fontSize28);
            paragraphMarkRunProperties15.Append(fontSizeComplexScript25);

            paragraphProperties15.Append(paragraphMarkRunProperties15);

            paragraph15.Append(paragraphProperties15);

            Paragraph paragraph16 = new Paragraph() { RsidParagraphMarkRevision = "00C16BA1", RsidParagraphAddition = "00500C13", RsidParagraphProperties = "00500C13", RsidRunAdditionDefault = "00500C13" };

            ParagraphProperties paragraphProperties16 = new ParagraphProperties();
            Indentation indentation2 = new Indentation() { FirstLine = "480" };

            ParagraphMarkRunProperties paragraphMarkRunProperties16 = new ParagraphMarkRunProperties();
            RunFonts runFonts26 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "宋体" };
            Color color14 = new Color() { Val = "000000" };
            Kern kern26 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize29 = new FontSize() { Val = "22" };

            paragraphMarkRunProperties16.Append(runFonts26);
            paragraphMarkRunProperties16.Append(color14);
            paragraphMarkRunProperties16.Append(kern26);
            paragraphMarkRunProperties16.Append(fontSize29);

            paragraphProperties16.Append(indentation2);
            paragraphProperties16.Append(paragraphMarkRunProperties16);

            Run run14 = new Run() { RsidRunProperties = "00C16BA1" };

            RunProperties runProperties14 = new RunProperties();
            RunFonts runFonts27 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "宋体" };
            Color color15 = new Color() { Val = "000000" };
            Kern kern27 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize30 = new FontSize() { Val = "24" };

            runProperties14.Append(runFonts27);
            runProperties14.Append(color15);
            runProperties14.Append(kern27);
            runProperties14.Append(fontSize30);
            Text text14 = new Text();
            text14.Text = "学业失败时的归因方式是外归因，是一种相对积极的归因方式。学习失败时归因于运气不佳，题目难度太大等外部因素时，学生就不会产生内疚或羞愧的心理。即使考试不理想，也不会否认自己的学习能力，不会打击继续努力信心。这种归因方式对于学业成绩的提高有一定的推动作用。";

            run14.Append(runProperties14);
            run14.Append(text14);

            paragraph16.Append(paragraphProperties16);
            paragraph16.Append(run14);

            Paragraph paragraph17 = new Paragraph() { RsidParagraphMarkRevision = "00C16BA1", RsidParagraphAddition = "00500C13", RsidParagraphProperties = "00500C13", RsidRunAdditionDefault = "00500C13" };

            ParagraphProperties paragraphProperties17 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties17 = new ParagraphMarkRunProperties();
            RunFonts runFonts28 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "宋体" };
            Color color16 = new Color() { Val = "000000" };
            Kern kern28 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize31 = new FontSize() { Val = "24" };

            paragraphMarkRunProperties17.Append(runFonts28);
            paragraphMarkRunProperties17.Append(color16);
            paragraphMarkRunProperties17.Append(kern28);
            paragraphMarkRunProperties17.Append(fontSize31);

            paragraphProperties17.Append(paragraphMarkRunProperties17);

            paragraph17.Append(paragraphProperties17);

            Paragraph paragraph18 = new Paragraph() { RsidParagraphMarkRevision = "00500C13", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "00500C13", RsidRunAdditionDefault = "000C649A" };

            ParagraphProperties paragraphProperties18 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties18 = new ParagraphMarkRunProperties();
            RunFonts runFonts29 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Color color17 = new Color() { Val = "000000" };
            Kern kern29 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize32 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript26 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties18.Append(runFonts29);
            paragraphMarkRunProperties18.Append(color17);
            paragraphMarkRunProperties18.Append(kern29);
            paragraphMarkRunProperties18.Append(fontSize32);
            paragraphMarkRunProperties18.Append(fontSizeComplexScript26);

            paragraphProperties18.Append(paragraphMarkRunProperties18);
            BookmarkStart bookmarkStart1 = new BookmarkStart() { Name = "_GoBack", Id = "0" };
            BookmarkEnd bookmarkEnd1 = new BookmarkEnd() { Id = "0" };

            paragraph18.Append(paragraphProperties18);
            paragraph18.Append(bookmarkStart1);
            paragraph18.Append(bookmarkEnd1);

            tableCell11.Append(tableCellProperties11);
            tableCell11.Append(paragraph13);
            tableCell11.Append(paragraph14);
            tableCell11.Append(paragraph15);
            tableCell11.Append(paragraph16);
            tableCell11.Append(paragraph17);
            tableCell11.Append(paragraph18);

            tableRow3.Append(tablePropertyExceptions3);
            tableRow3.Append(tableRowProperties3);
            tableRow3.Append(tableCell7);
            tableRow3.Append(tableCell8);
            tableRow3.Append(tableCell9);
            tableRow3.Append(tableCell10);
            tableRow3.Append(tableCell11);

            TableRow tableRow4 = new TableRow() { RsidTableRowMarkRevision = "007642DC", RsidTableRowAddition = "001D4055", RsidTableRowProperties = "006A0186" };

            TablePropertyExceptions tablePropertyExceptions4 = new TablePropertyExceptions();

            TableCellMarginDefault tableCellMarginDefault5 = new TableCellMarginDefault();
            TopMargin topMargin4 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            BottomMargin bottomMargin4 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };

            tableCellMarginDefault5.Append(topMargin4);
            tableCellMarginDefault5.Append(bottomMargin4);

            tablePropertyExceptions4.Append(tableCellMarginDefault5);

            TableRowProperties tableRowProperties4 = new TableRowProperties();
            TableRowHeight tableRowHeight4 = new TableRowHeight() { Val = (UInt32Value)1046U };

            tableRowProperties4.Append(tableRowHeight4);

            TableCell tableCell12 = new TableCell();

            TableCellProperties tableCellProperties12 = new TableCellProperties();
            TableCellWidth tableCellWidth12 = new TableCellWidth() { Width = "2690", Type = TableWidthUnitValues.Dxa };
            Shading shading10 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment11 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties12.Append(tableCellWidth12);
            tableCellProperties12.Append(shading10);
            tableCellProperties12.Append(tableCellVerticalAlignment11);

            Paragraph paragraph19 = new Paragraph() { RsidParagraphMarkRevision = "00973E52", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "000C649A" };

            ParagraphProperties paragraphProperties19 = new ParagraphProperties();
            Justification justification13 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties19 = new ParagraphMarkRunProperties();
            RunFonts runFonts30 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Kern kern30 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize33 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript27 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties19.Append(runFonts30);
            paragraphMarkRunProperties19.Append(kern30);
            paragraphMarkRunProperties19.Append(fontSize33);
            paragraphMarkRunProperties19.Append(fontSizeComplexScript27);

            paragraphProperties19.Append(justification13);
            paragraphProperties19.Append(paragraphMarkRunProperties19);

            Run run15 = new Run();

            RunProperties runProperties15 = new RunProperties();
            RunFonts runFonts31 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Kern kern31 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize34 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript28 = new FontSizeComplexScript() { Val = "24" };

            runProperties15.Append(runFonts31);
            runProperties15.Append(kern31);
            runProperties15.Append(fontSize34);
            runProperties15.Append(fontSizeComplexScript28);
            Text text15 = new Text();
            text15.Text = "学业成功";

            run15.Append(runProperties15);
            run15.Append(text15);

            Run run16 = new Run();

            RunProperties runProperties16 = new RunProperties();
            RunFonts runFonts32 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Kern kern32 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize35 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript29 = new FontSizeComplexScript() { Val = "24" };

            runProperties16.Append(runFonts32);
            runProperties16.Append(kern32);
            runProperties16.Append(fontSize35);
            runProperties16.Append(fontSizeComplexScript29);
            Text text16 = new Text();
            text16.Text = "-";

            run16.Append(runProperties16);
            run16.Append(text16);

            Run run17 = new Run() { RsidRunProperties = "00973E52" };

            RunProperties runProperties17 = new RunProperties();
            RunFonts runFonts33 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Kern kern33 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize36 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript30 = new FontSizeComplexScript() { Val = "24" };

            runProperties17.Append(runFonts33);
            runProperties17.Append(kern33);
            runProperties17.Append(fontSize36);
            runProperties17.Append(fontSizeComplexScript30);
            Text text17 = new Text();
            text17.Text = "外归因";

            run17.Append(runProperties17);
            run17.Append(text17);

            paragraph19.Append(paragraphProperties19);
            paragraph19.Append(run15);
            paragraph19.Append(run16);
            paragraph19.Append(run17);

            tableCell12.Append(tableCellProperties12);
            tableCell12.Append(paragraph19);

            TableCell tableCell13 = new TableCell();

            TableCellProperties tableCellProperties13 = new TableCellProperties();
            TableCellWidth tableCellWidth13 = new TableCellWidth() { Width = "682", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment12 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties13.Append(tableCellWidth13);
            tableCellProperties13.Append(tableCellVerticalAlignment12);

            Paragraph paragraph20 = new Paragraph() { RsidParagraphMarkRevision = "001B2170", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "000C649A" };

            ParagraphProperties paragraphProperties20 = new ParagraphProperties();
            Justification justification14 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties20 = new ParagraphMarkRunProperties();
            FontSize fontSize37 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript31 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties20.Append(fontSize37);
            paragraphMarkRunProperties20.Append(fontSizeComplexScript31);

            paragraphProperties20.Append(justification14);
            paragraphProperties20.Append(paragraphMarkRunProperties20);

            paragraph20.Append(paragraphProperties20);

            Paragraph paragraph21 = new Paragraph() { RsidParagraphMarkRevision = "001B2170", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "000C649A" };

            ParagraphProperties paragraphProperties21 = new ParagraphProperties();
            Justification justification15 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties21 = new ParagraphMarkRunProperties();
            RunFonts runFonts34 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "宋体" };
            Color color18 = new Color() { Val = "000000" };
            Kern kern34 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize38 = new FontSize() { Val = "22" };

            paragraphMarkRunProperties21.Append(runFonts34);
            paragraphMarkRunProperties21.Append(color18);
            paragraphMarkRunProperties21.Append(kern34);
            paragraphMarkRunProperties21.Append(fontSize38);

            paragraphProperties21.Append(justification15);
            paragraphProperties21.Append(paragraphMarkRunProperties21);

            Run run18 = new Run() { RsidRunProperties = "001B2170" };

            RunProperties runProperties18 = new RunProperties();
            RunFonts runFonts35 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "宋体" };
            Color color19 = new Color() { Val = "000000" };
            Kern kern35 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize39 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript32 = new FontSizeComplexScript() { Val = "24" };

            runProperties18.Append(runFonts35);
            runProperties18.Append(color19);
            runProperties18.Append(kern35);
            runProperties18.Append(fontSize39);
            runProperties18.Append(fontSizeComplexScript32);
            Text text18 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text18.Text = "45.00 ";

            run18.Append(runProperties18);
            run18.Append(text18);

            paragraph21.Append(paragraphProperties21);
            paragraph21.Append(run18);

            Paragraph paragraph22 = new Paragraph() { RsidParagraphMarkRevision = "001B2170", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "000C649A" };

            ParagraphProperties paragraphProperties22 = new ParagraphProperties();
            Justification justification16 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties22 = new ParagraphMarkRunProperties();
            RunFonts runFonts36 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Bold bold21 = new Bold();
            Color color20 = new Color() { Val = "000000" };
            Kern kern36 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize40 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript33 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties22.Append(runFonts36);
            paragraphMarkRunProperties22.Append(bold21);
            paragraphMarkRunProperties22.Append(color20);
            paragraphMarkRunProperties22.Append(kern36);
            paragraphMarkRunProperties22.Append(fontSize40);
            paragraphMarkRunProperties22.Append(fontSizeComplexScript33);

            paragraphProperties22.Append(justification16);
            paragraphProperties22.Append(paragraphMarkRunProperties22);

            paragraph22.Append(paragraphProperties22);

            tableCell13.Append(tableCellProperties13);
            tableCell13.Append(paragraph20);
            tableCell13.Append(paragraph21);
            tableCell13.Append(paragraph22);

            TableCell tableCell14 = new TableCell();

            TableCellProperties tableCellProperties14 = new TableCellProperties();
            TableCellWidth tableCellWidth14 = new TableCellWidth() { Width = "1082", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment13 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties14.Append(tableCellWidth14);
            tableCellProperties14.Append(tableCellVerticalAlignment13);

            Paragraph paragraph23 = new Paragraph() { RsidParagraphMarkRevision = "001B2170", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "000C649A" };

            ParagraphProperties paragraphProperties23 = new ParagraphProperties();
            Justification justification17 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties23 = new ParagraphMarkRunProperties();
            RunFonts runFonts37 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Bold bold22 = new Bold();
            Color color21 = new Color() { Val = "000000" };
            Kern kern37 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize41 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript34 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties23.Append(runFonts37);
            paragraphMarkRunProperties23.Append(bold22);
            paragraphMarkRunProperties23.Append(color21);
            paragraphMarkRunProperties23.Append(kern37);
            paragraphMarkRunProperties23.Append(fontSize41);
            paragraphMarkRunProperties23.Append(fontSizeComplexScript34);

            paragraphProperties23.Append(justification17);
            paragraphProperties23.Append(paragraphMarkRunProperties23);

            Run run19 = new Run() { RsidRunProperties = "001B2170" };

            RunProperties runProperties19 = new RunProperties();
            RunFonts runFonts38 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Bold bold23 = new Bold();
            Color color22 = new Color() { Val = "000000" };
            Kern kern38 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize42 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript35 = new FontSizeComplexScript() { Val = "24" };

            runProperties19.Append(runFonts38);
            runProperties19.Append(bold23);
            runProperties19.Append(color22);
            runProperties19.Append(kern38);
            runProperties19.Append(fontSize42);
            runProperties19.Append(fontSizeComplexScript35);
            Text text19 = new Text();
            text19.Text = "中等";

            run19.Append(runProperties19);
            run19.Append(text19);

            Run run20 = new Run() { RsidRunProperties = "001B2170" };

            RunProperties runProperties20 = new RunProperties();
            RunFonts runFonts39 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Bold bold24 = new Bold();
            Color color23 = new Color() { Val = "000000" };
            Kern kern39 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize43 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript36 = new FontSizeComplexScript() { Val = "24" };

            runProperties20.Append(runFonts39);
            runProperties20.Append(bold24);
            runProperties20.Append(color23);
            runProperties20.Append(kern39);
            runProperties20.Append(fontSize43);
            runProperties20.Append(fontSizeComplexScript36);
            Text text20 = new Text();
            text20.Text = "水平";

            run20.Append(runProperties20);
            run20.Append(text20);

            paragraph23.Append(paragraphProperties23);
            paragraph23.Append(run19);
            paragraph23.Append(run20);

            tableCell14.Append(tableCellProperties14);
            tableCell14.Append(paragraph23);

            TableCell tableCell15 = new TableCell();

            TableCellProperties tableCellProperties15 = new TableCellProperties();
            TableCellWidth tableCellWidth15 = new TableCellWidth() { Width = "882", Type = TableWidthUnitValues.Dxa };
            Shading shading11 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment14 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties15.Append(tableCellWidth15);
            tableCellProperties15.Append(shading11);
            tableCellProperties15.Append(tableCellVerticalAlignment14);

            Paragraph paragraph24 = new Paragraph() { RsidParagraphMarkRevision = "001B2170", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "000C649A" };

            ParagraphProperties paragraphProperties24 = new ParagraphProperties();
            Justification justification18 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties24 = new ParagraphMarkRunProperties();
            RunFonts runFonts40 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Bold bold25 = new Bold();
            Color color24 = new Color() { Val = "000000" };
            Kern kern40 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize44 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript37 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties24.Append(runFonts40);
            paragraphMarkRunProperties24.Append(bold25);
            paragraphMarkRunProperties24.Append(color24);
            paragraphMarkRunProperties24.Append(kern40);
            paragraphMarkRunProperties24.Append(fontSize44);
            paragraphMarkRunProperties24.Append(fontSizeComplexScript37);

            paragraphProperties24.Append(justification18);
            paragraphProperties24.Append(paragraphMarkRunProperties24);

            Run run21 = new Run() { RsidRunProperties = "001B2170" };

            RunProperties runProperties21 = new RunProperties();
            RunFonts runFonts41 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Color color25 = new Color() { Val = "000000" };
            Kern kern41 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize45 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript38 = new FontSizeComplexScript() { Val = "24" };

            runProperties21.Append(runFonts41);
            runProperties21.Append(color25);
            runProperties21.Append(kern41);
            runProperties21.Append(fontSize45);
            runProperties21.Append(fontSizeComplexScript38);
            Text text21 = new Text();
            text21.Text = "-";

            run21.Append(runProperties21);
            run21.Append(text21);

            Run run22 = new Run() { RsidRunProperties = "001B2170" };

            RunProperties runProperties22 = new RunProperties();
            RunFonts runFonts42 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Color color26 = new Color() { Val = "000000" };
            Kern kern42 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize46 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript39 = new FontSizeComplexScript() { Val = "24" };

            runProperties22.Append(runFonts42);
            runProperties22.Append(color26);
            runProperties22.Append(kern42);
            runProperties22.Append(fontSize46);
            runProperties22.Append(fontSizeComplexScript39);
            Text text22 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text22.Text = " ";

            run22.Append(runProperties22);
            run22.Append(text22);

            Run run23 = new Run() { RsidRunProperties = "001B2170" };

            RunProperties runProperties23 = new RunProperties();
            RunFonts runFonts43 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Color color27 = new Color() { Val = "000000" };
            Kern kern43 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize47 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript40 = new FontSizeComplexScript() { Val = "24" };

            runProperties23.Append(runFonts43);
            runProperties23.Append(color27);
            runProperties23.Append(kern43);
            runProperties23.Append(fontSize47);
            runProperties23.Append(fontSizeComplexScript40);
            Text text23 = new Text();
            text23.Text = "-";

            run23.Append(runProperties23);
            run23.Append(text23);

            Run run24 = new Run() { RsidRunProperties = "001B2170" };

            RunProperties runProperties24 = new RunProperties();
            RunFonts runFonts44 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Color color28 = new Color() { Val = "000000" };
            Kern kern44 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize48 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript41 = new FontSizeComplexScript() { Val = "24" };

            runProperties24.Append(runFonts44);
            runProperties24.Append(color28);
            runProperties24.Append(kern44);
            runProperties24.Append(fontSize48);
            runProperties24.Append(fontSizeComplexScript41);
            Text text24 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text24.Text = " -";

            run24.Append(runProperties24);
            run24.Append(text24);

            paragraph24.Append(paragraphProperties24);
            paragraph24.Append(run21);
            paragraph24.Append(run22);
            paragraph24.Append(run23);
            paragraph24.Append(run24);

            tableCell15.Append(tableCellProperties15);
            tableCell15.Append(paragraph24);

            TableCell tableCell16 = new TableCell();

            TableCellProperties tableCellProperties16 = new TableCellProperties();
            TableCellWidth tableCellWidth16 = new TableCellWidth() { Width = "5375", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge2 = new VerticalMerge();
            Shading shading12 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment15 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties16.Append(tableCellWidth16);
            tableCellProperties16.Append(verticalMerge2);
            tableCellProperties16.Append(shading12);
            tableCellProperties16.Append(tableCellVerticalAlignment15);

            Paragraph paragraph25 = new Paragraph() { RsidParagraphMarkRevision = "002A093C", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "000C649A" };

            ParagraphProperties paragraphProperties25 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties25 = new ParagraphMarkRunProperties();
            RunFonts runFonts45 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Color color29 = new Color() { Val = "000000" };
            Kern kern45 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize49 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript42 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties25.Append(runFonts45);
            paragraphMarkRunProperties25.Append(color29);
            paragraphMarkRunProperties25.Append(kern45);
            paragraphMarkRunProperties25.Append(fontSize49);
            paragraphMarkRunProperties25.Append(fontSizeComplexScript42);

            paragraphProperties25.Append(paragraphMarkRunProperties25);

            paragraph25.Append(paragraphProperties25);

            tableCell16.Append(tableCellProperties16);
            tableCell16.Append(paragraph25);

            tableRow4.Append(tablePropertyExceptions4);
            tableRow4.Append(tableRowProperties4);
            tableRow4.Append(tableCell12);
            tableRow4.Append(tableCell13);
            tableRow4.Append(tableCell14);
            tableRow4.Append(tableCell15);
            tableRow4.Append(tableCell16);

            table1.Append(tableProperties1);
            table1.Append(tableGrid1);
            table1.Append(tableRow1);
            table1.Append(tableRow2);
            table1.Append(tableRow3);
            table1.Append(tableRow4);
            return table1;

        }
    }
}

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
        // Creates an Table instance and adds its children.
        //量表，维度，评价 等级 影响 得分
        public Table GenerateSummaryTable(SafeDictionary<string, SafeDictionary<string, (string, string, Word.Enum.eInfluence, double)>> detaildate)
        {

            Table table1 = new Table();

            TableProperties tableProperties1 = new TableProperties();
            TableWidth tableWidth1 = new TableWidth() { Width = "8285", Type = TableWidthUnitValues.Dxa };
            TableJustification tableJustification1 = new TableJustification() { Val = TableRowAlignmentValues.Center };

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
            tableProperties1.Append(tableJustification1);
            tableProperties1.Append(tableBorders1);
            tableProperties1.Append(tableLayout1);
            tableProperties1.Append(tableCellMarginDefault1);
            tableProperties1.Append(tableLook1);

            table1.Append(tableProperties1);

            TableGrid tableGrid1 = new TableGrid();
            GridColumn gridColumn1 = new GridColumn() { Width = "1986" };
            GridColumn gridColumn2 = new GridColumn() { Width = "3186" };
            GridColumn gridColumn3 = new GridColumn() { Width = "935" };
            GridColumn gridColumn4 = new GridColumn() { Width = "935" };
            GridColumn gridColumn5 = new GridColumn() { Width = "1243" };

            tableGrid1.Append(gridColumn1);
            tableGrid1.Append(gridColumn2);
            tableGrid1.Append(gridColumn3);
            tableGrid1.Append(gridColumn4);
            tableGrid1.Append(gridColumn5);
            table1.Append(tableGrid1);

            TableRow tableRow1 = new TableRow() { RsidTableRowMarkRevision = "00885DA2", RsidTableRowAddition = "00A844B0", RsidTableRowProperties = "00522EB6" };

            TableRowProperties tableRowProperties1 = new TableRowProperties();
            TableRowHeight tableRowHeight1 = new TableRowHeight() { Val = (UInt32Value)595U };
            TableJustification tableJustification2 = new TableJustification() { Val = TableRowAlignmentValues.Center };

            tableRowProperties1.Append(tableRowHeight1);
            tableRowProperties1.Append(tableJustification2);

            TableCell tableCell1 = new TableCell();

            TableCellProperties tableCellProperties1 = new TableCellProperties();
            TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = "5172", Type = TableWidthUnitValues.Dxa };
            GridSpan gridSpan1 = new GridSpan() { Val = 2 };
            Shading shading1 = new Shading() { Val = ShadingPatternValues.Clear, Color = "000000", Fill = "D99795" };
            TableCellVerticalAlignment tableCellVerticalAlignment1 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };
            HideMark hideMark1 = new HideMark();

            tableCellProperties1.Append(tableCellWidth1);
            tableCellProperties1.Append(gridSpan1);
            tableCellProperties1.Append(shading1);
            tableCellProperties1.Append(tableCellVerticalAlignment1);
            tableCellProperties1.Append(hideMark1);

            Paragraph paragraph1 = new Paragraph() { RsidParagraphMarkRevision = "00885DA2", RsidParagraphAddition = "00A844B0", RsidParagraphProperties = "00885DA2", RsidRunAdditionDefault = "00A844B0" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            WidowControl widowControl1 = new WidowControl();
            AdjustRightIndent adjustRightIndent1 = new AdjustRightIndent() { Val = false };
            SnapToGrid snapToGrid1 = new SnapToGrid() { Val = false };
            Indentation indentation1 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification1 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
            RunFonts runFonts1 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold1 = new Bold();
            Kern kern1 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties1.Append(runFonts1);
            paragraphMarkRunProperties1.Append(bold1);
            paragraphMarkRunProperties1.Append(kern1);
            paragraphMarkRunProperties1.Append(fontSize1);
            paragraphMarkRunProperties1.Append(fontSizeComplexScript1);

            paragraphProperties1.Append(widowControl1);
            paragraphProperties1.Append(adjustRightIndent1);
            paragraphProperties1.Append(snapToGrid1);
            paragraphProperties1.Append(indentation1);
            paragraphProperties1.Append(justification1);
            paragraphProperties1.Append(paragraphMarkRunProperties1);

            Run run1 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties1 = new RunProperties();
            RunFonts runFonts2 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体", ComplexScript = "Tahoma" };
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
            text1.Text = "维度名称";

            run1.Append(runProperties1);
            run1.Append(text1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            tableCell1.Append(tableCellProperties1);
            tableCell1.Append(paragraph1);

            TableCell tableCell2 = new TableCell();

            TableCellProperties tableCellProperties2 = new TableCellProperties();
            TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = "935", Type = TableWidthUnitValues.Dxa };
            Shading shading2 = new Shading() { Val = ShadingPatternValues.Clear, Color = "000000", Fill = "D99795" };
            TableCellVerticalAlignment tableCellVerticalAlignment2 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties2.Append(tableCellWidth2);
            tableCellProperties2.Append(shading2);
            tableCellProperties2.Append(tableCellVerticalAlignment2);

            Paragraph paragraph2 = new Paragraph() { RsidParagraphMarkRevision = "00885DA2", RsidParagraphAddition = "00A844B0", RsidParagraphProperties = "00885DA2", RsidRunAdditionDefault = "00A844B0" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            WidowControl widowControl2 = new WidowControl();
            AdjustRightIndent adjustRightIndent2 = new AdjustRightIndent() { Val = false };
            SnapToGrid snapToGrid2 = new SnapToGrid() { Val = false };
            Indentation indentation2 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification2 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
            RunFonts runFonts3 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold3 = new Bold();
            Kern kern3 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize3 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties2.Append(runFonts3);
            paragraphMarkRunProperties2.Append(bold3);
            paragraphMarkRunProperties2.Append(kern3);
            paragraphMarkRunProperties2.Append(fontSize3);
            paragraphMarkRunProperties2.Append(fontSizeComplexScript3);

            paragraphProperties2.Append(widowControl2);
            paragraphProperties2.Append(adjustRightIndent2);
            paragraphProperties2.Append(snapToGrid2);
            paragraphProperties2.Append(indentation2);
            paragraphProperties2.Append(justification2);
            paragraphProperties2.Append(paragraphMarkRunProperties2);

            Run run2 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties2 = new RunProperties();
            RunFonts runFonts4 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体", ComplexScript = "Tahoma" };
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
            text2.Text = "得分";

            run2.Append(runProperties2);
            run2.Append(text2);

            paragraph2.Append(paragraphProperties2);
            paragraph2.Append(run2);

            tableCell2.Append(tableCellProperties2);
            tableCell2.Append(paragraph2);

            TableCell tableCell3 = new TableCell();

            TableCellProperties tableCellProperties3 = new TableCellProperties();
            TableCellWidth tableCellWidth3 = new TableCellWidth() { Width = "935", Type = TableWidthUnitValues.Dxa };
            Shading shading3 = new Shading() { Val = ShadingPatternValues.Clear, Color = "000000", Fill = "D99795" };
            TableCellVerticalAlignment tableCellVerticalAlignment3 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };
            HideMark hideMark2 = new HideMark();

            tableCellProperties3.Append(tableCellWidth3);
            tableCellProperties3.Append(shading3);
            tableCellProperties3.Append(tableCellVerticalAlignment3);
            tableCellProperties3.Append(hideMark2);

            Paragraph paragraph3 = new Paragraph() { RsidParagraphMarkRevision = "00885DA2", RsidParagraphAddition = "00A844B0", RsidParagraphProperties = "00885DA2", RsidRunAdditionDefault = "00A844B0" };

            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
            WidowControl widowControl3 = new WidowControl();
            AdjustRightIndent adjustRightIndent3 = new AdjustRightIndent() { Val = false };
            SnapToGrid snapToGrid3 = new SnapToGrid() { Val = false };
            Indentation indentation3 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification3 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
            RunFonts runFonts5 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold5 = new Bold();
            Kern kern5 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize5 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties3.Append(runFonts5);
            paragraphMarkRunProperties3.Append(bold5);
            paragraphMarkRunProperties3.Append(kern5);
            paragraphMarkRunProperties3.Append(fontSize5);
            paragraphMarkRunProperties3.Append(fontSizeComplexScript5);

            paragraphProperties3.Append(widowControl3);
            paragraphProperties3.Append(adjustRightIndent3);
            paragraphProperties3.Append(snapToGrid3);
            paragraphProperties3.Append(indentation3);
            paragraphProperties3.Append(justification3);
            paragraphProperties3.Append(paragraphMarkRunProperties3);

            Run run3 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties3 = new RunProperties();
            RunFonts runFonts6 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体", ComplexScript = "Tahoma" };
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
            text3.Text = "水平";

            run3.Append(runProperties3);
            run3.Append(text3);

            paragraph3.Append(paragraphProperties3);
            paragraph3.Append(run3);

            tableCell3.Append(tableCellProperties3);
            tableCell3.Append(paragraph3);

            TableCell tableCell4 = new TableCell();

            TableCellProperties tableCellProperties4 = new TableCellProperties();
            TableCellWidth tableCellWidth4 = new TableCellWidth() { Width = "1243", Type = TableWidthUnitValues.Dxa };
            Shading shading4 = new Shading() { Val = ShadingPatternValues.Clear, Color = "000000", Fill = "D99795" };
            TableCellVerticalAlignment tableCellVerticalAlignment4 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties4.Append(tableCellWidth4);
            tableCellProperties4.Append(shading4);
            tableCellProperties4.Append(tableCellVerticalAlignment4);

            Paragraph paragraph4 = new Paragraph() { RsidParagraphMarkRevision = "00885DA2", RsidParagraphAddition = "00A844B0", RsidParagraphProperties = "00885DA2", RsidRunAdditionDefault = "00A844B0" };

            ParagraphProperties paragraphProperties4 = new ParagraphProperties();
            WidowControl widowControl4 = new WidowControl();
            AdjustRightIndent adjustRightIndent4 = new AdjustRightIndent() { Val = false };
            SnapToGrid snapToGrid4 = new SnapToGrid() { Val = false };
            Indentation indentation4 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification4 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties4 = new ParagraphMarkRunProperties();
            RunFonts runFonts7 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold7 = new Bold();
            Kern kern7 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize7 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript7 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties4.Append(runFonts7);
            paragraphMarkRunProperties4.Append(bold7);
            paragraphMarkRunProperties4.Append(kern7);
            paragraphMarkRunProperties4.Append(fontSize7);
            paragraphMarkRunProperties4.Append(fontSizeComplexScript7);

            paragraphProperties4.Append(widowControl4);
            paragraphProperties4.Append(adjustRightIndent4);
            paragraphProperties4.Append(snapToGrid4);
            paragraphProperties4.Append(indentation4);
            paragraphProperties4.Append(justification4);
            paragraphProperties4.Append(paragraphMarkRunProperties4);

            Run run4 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties4 = new RunProperties();
            RunFonts runFonts8 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体", ComplexScript = "Tahoma" };
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
            text4.Text = "是否需要改善";

            run4.Append(runProperties4);
            run4.Append(text4);

            paragraph4.Append(paragraphProperties4);
            paragraph4.Append(run4);


            tableCell4.Append(tableCellProperties4);
            tableCell4.Append(paragraph4);

            tableRow1.Append(tableRowProperties1);
            tableRow1.Append(tableCell1);
            tableRow1.Append(tableCell2);
            tableRow1.Append(tableCell3);
            tableRow1.Append(tableCell4);
            table1.Append(tableRow1);

            ////正文开始

            foreach (var measurement in detaildate)
            {
                TableRow tableRow2 = new TableRow() { RsidTableRowMarkRevision = "00885DA2", RsidTableRowAddition = "00522EB6", RsidTableRowProperties = "00522EB6" };

                TableRowProperties tableRowProperties2 = new TableRowProperties();
                TableRowHeight tableRowHeight2 = new TableRowHeight() { Val = (UInt32Value)367U };
                TableJustification tableJustification3 = new TableJustification() { Val = TableRowAlignmentValues.Center };

                tableRowProperties2.Append(tableRowHeight2);
                tableRowProperties2.Append(tableJustification3);

                TableCell tableCell5 = new TableCell();

                TableCellProperties tableCellProperties5 = new TableCellProperties();
                TableCellWidth tableCellWidth5 = new TableCellWidth() { Width = "1986", Type = TableWidthUnitValues.Dxa };
                VerticalMerge verticalMerge1 = new VerticalMerge() { Val = MergedCellValues.Restart };
                Shading shading5 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
                TableCellVerticalAlignment tableCellVerticalAlignment5 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

                tableCellProperties5.Append(tableCellWidth5);
                tableCellProperties5.Append(verticalMerge1);
                tableCellProperties5.Append(shading5);
                tableCellProperties5.Append(tableCellVerticalAlignment5);

                Paragraph paragraph5 = new Paragraph() { RsidParagraphMarkRevision = "00885DA2", RsidParagraphAddition = "00522EB6", RsidParagraphProperties = "001F4E3B", RsidRunAdditionDefault = "00522EB6" };

                ParagraphProperties paragraphProperties5 = new ParagraphProperties();
                WidowControl widowControl5 = new WidowControl();
                AdjustRightIndent adjustRightIndent5 = new AdjustRightIndent() { Val = false };
                SnapToGrid snapToGrid5 = new SnapToGrid() { Val = false };
                Indentation indentation5 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
                Justification justification5 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties5 = new ParagraphMarkRunProperties();
                RunFonts runFonts10 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体", ComplexScript = "Tahoma" };
                Bold bold10 = new Bold();
                Color color1 = new Color() { Val = "000000" };
                Kern kern10 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize10 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript10 = new FontSizeComplexScript() { Val = "24" };

                paragraphMarkRunProperties5.Append(runFonts10);
                paragraphMarkRunProperties5.Append(bold10);
                paragraphMarkRunProperties5.Append(color1);
                paragraphMarkRunProperties5.Append(kern10);
                paragraphMarkRunProperties5.Append(fontSize10);
                paragraphMarkRunProperties5.Append(fontSizeComplexScript10);

                paragraphProperties5.Append(widowControl5);
                paragraphProperties5.Append(adjustRightIndent5);
                paragraphProperties5.Append(snapToGrid5);
                paragraphProperties5.Append(indentation5);
                paragraphProperties5.Append(justification5);
                paragraphProperties5.Append(paragraphMarkRunProperties5);

                Run run6 = new Run() { RsidRunProperties = "00885DA2" };

                RunProperties runProperties6 = new RunProperties();
                RunFonts runFonts11 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体", ComplexScript = "Tahoma" };
                Bold bold11 = new Bold();
                Color color2 = new Color() { Val = "000000" };
                Kern kern11 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize11 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript11 = new FontSizeComplexScript() { Val = "24" };

                runProperties6.Append(runFonts11);
                runProperties6.Append(bold11);
                runProperties6.Append(color2);
                runProperties6.Append(kern11);
                runProperties6.Append(fontSize11);
                runProperties6.Append(fontSizeComplexScript11);
                Text text6 = new Text();
                text6.Text =measurement.Key;//TODO 量表名

                run6.Append(runProperties6);
                run6.Append(text6);

                paragraph5.Append(paragraphProperties5);
                paragraph5.Append(run6);

                tableCell5.Append(tableCellProperties5);
                tableCell5.Append(paragraph5);

                TableCell tableCell6 = new TableCell();

                TableCellProperties tableCellProperties6 = new TableCellProperties();
                TableCellWidth tableCellWidth6 = new TableCellWidth() { Width = "3186", Type = TableWidthUnitValues.Dxa };

                TableCellBorders tableCellBorders1 = new TableCellBorders();
                TopBorder topBorder2 = new TopBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
                LeftBorder leftBorder2 = new LeftBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
                BottomBorder bottomBorder2 = new BottomBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
                RightBorder rightBorder2 = new RightBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

                tableCellBorders1.Append(topBorder2);
                tableCellBorders1.Append(leftBorder2);
                tableCellBorders1.Append(bottomBorder2);
                tableCellBorders1.Append(rightBorder2);
                Shading shading6 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
                TableCellVerticalAlignment tableCellVerticalAlignment6 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

                tableCellProperties6.Append(tableCellWidth6);
                tableCellProperties6.Append(tableCellBorders1);
                tableCellProperties6.Append(shading6);
                tableCellProperties6.Append(tableCellVerticalAlignment6);

                Paragraph paragraph6 = new Paragraph() { RsidParagraphMarkRevision = "00885DA2", RsidParagraphAddition = "00522EB6", RsidParagraphProperties = "006F5D65", RsidRunAdditionDefault = "00522EB6" };

                ParagraphProperties paragraphProperties6 = new ParagraphProperties();
                WidowControl widowControl6 = new WidowControl();
                AdjustRightIndent adjustRightIndent6 = new AdjustRightIndent() { Val = false };
                SnapToGrid snapToGrid6 = new SnapToGrid() { Val = false };
                Indentation indentation6 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
                Justification justification6 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties6 = new ParagraphMarkRunProperties();
                RunFonts runFonts13 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体", ComplexScript = "Tahoma" };
                Bold bold13 = new Bold();
                Color color4 = new Color() { Val = "000000" };
                Kern kern13 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize13 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript13 = new FontSizeComplexScript() { Val = "24" };

                paragraphMarkRunProperties6.Append(runFonts13);
                paragraphMarkRunProperties6.Append(bold13);
                paragraphMarkRunProperties6.Append(color4);
                paragraphMarkRunProperties6.Append(kern13);
                paragraphMarkRunProperties6.Append(fontSize13);
                paragraphMarkRunProperties6.Append(fontSizeComplexScript13);

                paragraphProperties6.Append(widowControl6);
                paragraphProperties6.Append(adjustRightIndent6);
                paragraphProperties6.Append(snapToGrid6);
                paragraphProperties6.Append(indentation6);
                paragraphProperties6.Append(justification6);
                paragraphProperties6.Append(paragraphMarkRunProperties6);

                Run run8 = new Run() { RsidRunProperties = "00885DA2" };

                RunProperties runProperties8 = new RunProperties();
                RunFonts runFonts14 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体", ComplexScript = "Tahoma" };
                Bold bold14 = new Bold();
                Color color5 = new Color() { Val = "000000" };
                Kern kern14 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize14 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript14 = new FontSizeComplexScript() { Val = "24" };

                runProperties8.Append(runFonts14);
                runProperties8.Append(bold14);
                runProperties8.Append(color5);
                runProperties8.Append(kern14);
                runProperties8.Append(fontSize14);
                runProperties8.Append(fontSizeComplexScript14);
                Text text8 = new Text();
                text8.Text = measurement.Value.Keys.FirstOrDefault();//TODO第一行维度名称

                run8.Append(runProperties8);
                run8.Append(text8);

                paragraph6.Append(paragraphProperties6);
                paragraph6.Append(run8);

                tableCell6.Append(tableCellProperties6);
                tableCell6.Append(paragraph6);

                TableCell tableCell7 = new TableCell();

                TableCellProperties tableCellProperties7 = new TableCellProperties();
                TableCellWidth tableCellWidth7 = new TableCellWidth() { Width = "935", Type = TableWidthUnitValues.Dxa };

                TableCellBorders tableCellBorders2 = new TableCellBorders();
                TopBorder topBorder3 = new TopBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
                LeftBorder leftBorder3 = new LeftBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
                BottomBorder bottomBorder3 = new BottomBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
                RightBorder rightBorder3 = new RightBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

                tableCellBorders2.Append(topBorder3);
                tableCellBorders2.Append(leftBorder3);
                tableCellBorders2.Append(bottomBorder3);
                tableCellBorders2.Append(rightBorder3);

                tableCellProperties7.Append(tableCellWidth7);
                tableCellProperties7.Append(tableCellBorders2);

                Paragraph paragraph7 = new Paragraph() { RsidParagraphMarkRevision = "00A844B0", RsidParagraphAddition = "00522EB6", RsidParagraphProperties = "001F4E3B", RsidRunAdditionDefault = "006B6F55" };

                ParagraphProperties paragraphProperties7 = new ParagraphProperties();
                WidowControl widowControl7 = new WidowControl();
                AdjustRightIndent adjustRightIndent7 = new AdjustRightIndent() { Val = false };
                SnapToGrid snapToGrid7 = new SnapToGrid() { Val = false };
                Indentation indentation7 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
                Justification justification7 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties7 = new ParagraphMarkRunProperties();
                RunFonts runFonts16 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体", ComplexScript = "Tahoma" };
                Bold bold16 = new Bold();
                Color color7 = new Color() { Val = "000000" };
                Kern kern16 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize16 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript16 = new FontSizeComplexScript() { Val = "24" };

                paragraphMarkRunProperties7.Append(runFonts16);
                paragraphMarkRunProperties7.Append(bold16);
                paragraphMarkRunProperties7.Append(color7);
                paragraphMarkRunProperties7.Append(kern16);
                paragraphMarkRunProperties7.Append(fontSize16);
                paragraphMarkRunProperties7.Append(fontSizeComplexScript16);

                paragraphProperties7.Append(widowControl7);
                paragraphProperties7.Append(adjustRightIndent7);
                paragraphProperties7.Append(snapToGrid7);
                paragraphProperties7.Append(indentation7);
                paragraphProperties7.Append(justification7);
                paragraphProperties7.Append(paragraphMarkRunProperties7);

                Run run10 = new Run();

                RunProperties runProperties10 = new RunProperties();
                RunFonts runFonts17 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", EastAsia = "宋体", ComplexScript = "Times New Roman" };

                runProperties10.Append(runFonts17);
                Text text10 = new Text();
                text10.Text = measurement.Value.FirstOrDefault().Value.Item4.ToString();//TODO第一行得分

                run10.Append(runProperties10);
                run10.Append(text10);

                paragraph7.Append(paragraphProperties7);
                paragraph7.Append(run10);

                tableCell7.Append(tableCellProperties7);
                tableCell7.Append(paragraph7);

                TableCell tableCell8 = new TableCell();

                TableCellProperties tableCellProperties8 = new TableCellProperties();
                TableCellWidth tableCellWidth8 = new TableCellWidth() { Width = "935", Type = TableWidthUnitValues.Dxa };

                TableCellBorders tableCellBorders3 = new TableCellBorders();
                TopBorder topBorder4 = new TopBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
                LeftBorder leftBorder4 = new LeftBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
                BottomBorder bottomBorder4 = new BottomBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
                RightBorder rightBorder4 = new RightBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

                tableCellBorders3.Append(topBorder4);
                tableCellBorders3.Append(leftBorder4);
                tableCellBorders3.Append(bottomBorder4);
                tableCellBorders3.Append(rightBorder4);
                Shading shading7 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
                TableCellVerticalAlignment tableCellVerticalAlignment7 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

                tableCellProperties8.Append(tableCellWidth8);
                tableCellProperties8.Append(tableCellBorders3);
                tableCellProperties8.Append(shading7);
                tableCellProperties8.Append(tableCellVerticalAlignment7);

                Paragraph paragraph8 = new Paragraph() { RsidParagraphMarkRevision = "00A844B0", RsidParagraphAddition = "00522EB6", RsidParagraphProperties = "001F4E3B", RsidRunAdditionDefault = "006B6F55" };

                ParagraphProperties paragraphProperties8 = new ParagraphProperties();
                WidowControl widowControl8 = new WidowControl();
                AdjustRightIndent adjustRightIndent8 = new AdjustRightIndent() { Val = false };
                SnapToGrid snapToGrid8 = new SnapToGrid() { Val = false };
                Indentation indentation8 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
                Justification justification8 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties8 = new ParagraphMarkRunProperties();
                RunFonts runFonts18 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", EastAsia = "宋体", ComplexScript = "Times New Roman" };

                paragraphMarkRunProperties8.Append(runFonts18);

                paragraphProperties8.Append(widowControl8);
                paragraphProperties8.Append(adjustRightIndent8);
                paragraphProperties8.Append(snapToGrid8);
                paragraphProperties8.Append(indentation8);
                paragraphProperties8.Append(justification8);
                paragraphProperties8.Append(paragraphMarkRunProperties8);

                Run run11 = new Run();

                RunProperties runProperties11 = new RunProperties();
                RunFonts runFonts19 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", EastAsia = "宋体", ComplexScript = "Times New Roman" };

                runProperties11.Append(runFonts19);
                Text text11 = new Text();
                text11.Text = measurement.Value.FirstOrDefault().Value.Item2;   //TODO第一行水平

                run11.Append(runProperties11);
                run11.Append(text11);

                paragraph8.Append(paragraphProperties8);
                paragraph8.Append(run11);

                tableCell8.Append(tableCellProperties8);
                tableCell8.Append(paragraph8);

                TableCell tableCell9 = new TableCell();

                TableCellProperties tableCellProperties9 = new TableCellProperties();
                TableCellWidth tableCellWidth9 = new TableCellWidth() { Width = "1243", Type = TableWidthUnitValues.Dxa };

                TableCellBorders tableCellBorders4 = new TableCellBorders();
                TopBorder topBorder5 = new TopBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
                LeftBorder leftBorder5 = new LeftBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
                BottomBorder bottomBorder5 = new BottomBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
                RightBorder rightBorder5 = new RightBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

                tableCellBorders4.Append(topBorder5);
                tableCellBorders4.Append(leftBorder5);
                tableCellBorders4.Append(bottomBorder5);
                tableCellBorders4.Append(rightBorder5);
                TableCellVerticalAlignment tableCellVerticalAlignment8 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

                tableCellProperties9.Append(tableCellWidth9);
                tableCellProperties9.Append(tableCellBorders4);
                tableCellProperties9.Append(tableCellVerticalAlignment8);

                Paragraph paragraph9 = new Paragraph() { RsidParagraphMarkRevision = "00885DA2", RsidParagraphAddition = "00522EB6", RsidParagraphProperties = "001F4E3B", RsidRunAdditionDefault = "00F82954" };

                ParagraphProperties paragraphProperties9 = new ParagraphProperties();
                WidowControl widowControl9 = new WidowControl();
                AdjustRightIndent adjustRightIndent9 = new AdjustRightIndent() { Val = false };
                SnapToGrid snapToGrid9 = new SnapToGrid() { Val = false };
                Indentation indentation9 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
                Justification justification9 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties9 = new ParagraphMarkRunProperties();
                RunFonts runFonts20 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", EastAsia = "宋体", ComplexScript = "Times New Roman" };
                Kern kern17 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize17 = new FontSize() { Val = "2" };
                FontSizeComplexScript fontSizeComplexScript17 = new FontSizeComplexScript() { Val = "2" };

                paragraphMarkRunProperties9.Append(runFonts20);
                paragraphMarkRunProperties9.Append(kern17);
                paragraphMarkRunProperties9.Append(fontSize17);
                paragraphMarkRunProperties9.Append(fontSizeComplexScript17);

                paragraphProperties9.Append(widowControl9);
                paragraphProperties9.Append(adjustRightIndent9);
                paragraphProperties9.Append(snapToGrid9);
                paragraphProperties9.Append(indentation9);
                paragraphProperties9.Append(justification9);
                paragraphProperties9.Append(paragraphMarkRunProperties9);


                paragraph9.Append(paragraphProperties9);

                Paragraph paragraph10 = new Paragraph() { RsidParagraphMarkRevision = "00885DA2", RsidParagraphAddition = "00522EB6", RsidParagraphProperties = "001F4E3B", RsidRunAdditionDefault = "0018090C" };

                ParagraphProperties paragraphProperties10 = new ParagraphProperties();
                WidowControl widowControl10 = new WidowControl();
                AdjustRightIndent adjustRightIndent10 = new AdjustRightIndent() { Val = false };
                SnapToGrid snapToGrid10 = new SnapToGrid() { Val = false };
                Indentation indentation10 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
                Justification justification10 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties10 = new ParagraphMarkRunProperties();
                RunFonts runFonts50 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体", ComplexScript = "Tahoma" };
                Color color37 = new Color() { Val = "000000" };
                Kern kern47 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize47 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript47 = new FontSizeComplexScript() { Val = "24" };

                paragraphMarkRunProperties10.Append(runFonts50);
                paragraphMarkRunProperties10.Append(color37);
                paragraphMarkRunProperties10.Append(kern47);
                paragraphMarkRunProperties10.Append(fontSize47);
                paragraphMarkRunProperties10.Append(fontSizeComplexScript47);

                paragraphProperties10.Append(widowControl10);
                paragraphProperties10.Append(adjustRightIndent10);
                paragraphProperties10.Append(snapToGrid10);
                paragraphProperties10.Append(indentation10);
                paragraphProperties10.Append(justification10);
                paragraphProperties10.Append(paragraphMarkRunProperties10);
                BookmarkStart bookmarkStart1 = new BookmarkStart() { Name = "_GoBack", Id = "0" };
                BookmarkEnd bookmarkEnd1 = new BookmarkEnd() { Id = "0" };

                Run run41 = new Run() { RsidRunProperties = "00885DA2" };

                RunProperties runProperties41 = new RunProperties();
                RunFonts runFonts51 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体", ComplexScript = "Tahoma" };
                Color color38 = new Color() { Val = "FF0000" };
                Kern kern48 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize48 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript48 = new FontSizeComplexScript() { Val = "24" };

                runProperties41.Append(runFonts51);
                runProperties41.Append(color38);
                runProperties41.Append(kern48);
                runProperties41.Append(fontSize48);
                runProperties41.Append(fontSizeComplexScript48);
                Text text12 = new Text();
                text12.Text = measurement.Value.FirstOrDefault().Value.Item3== eInfluence.消极影响?"■":"";//TODO第一行是否需要改善

                run41.Append(runProperties41);
                run41.Append(text12);

                paragraph10.Append(paragraphProperties10);
                paragraph10.Append(bookmarkStart1);
                paragraph10.Append(bookmarkEnd1);
                paragraph10.Append(run41);

                Paragraph paragraph11 = new Paragraph() { RsidParagraphMarkRevision = "00885DA2", RsidParagraphAddition = "00522EB6", RsidParagraphProperties = "001F4E3B", RsidRunAdditionDefault = "00F82954" };

                ParagraphProperties paragraphProperties11 = new ParagraphProperties();
                WidowControl widowControl11 = new WidowControl();
                AdjustRightIndent adjustRightIndent11 = new AdjustRightIndent() { Val = false };
                SnapToGrid snapToGrid11 = new SnapToGrid() { Val = false };
                Indentation indentation11 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
                Justification justification11 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties11 = new ParagraphMarkRunProperties();
                RunFonts runFonts52 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
                Bold bold46 = new Bold();
                Color color39 = new Color() { Val = "000000" };
                Kern kern49 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize49 = new FontSize() { Val = "2" };
                FontSizeComplexScript fontSizeComplexScript49 = new FontSizeComplexScript() { Val = "2" };

                paragraphMarkRunProperties11.Append(runFonts52);
                paragraphMarkRunProperties11.Append(bold46);
                paragraphMarkRunProperties11.Append(color39);
                paragraphMarkRunProperties11.Append(kern49);
                paragraphMarkRunProperties11.Append(fontSize49);
                paragraphMarkRunProperties11.Append(fontSizeComplexScript49);

                paragraphProperties11.Append(widowControl11);
                paragraphProperties11.Append(adjustRightIndent11);
                paragraphProperties11.Append(snapToGrid11);
                paragraphProperties11.Append(indentation11);
                paragraphProperties11.Append(justification11);
                paragraphProperties11.Append(paragraphMarkRunProperties11);
                paragraph11.Append(paragraphProperties11);


                tableCell9.Append(tableCellProperties9);
                tableCell9.Append(paragraph9);
                tableCell9.Append(paragraph10);
                tableCell9.Append(paragraph11);

                tableRow2.Append(tableRowProperties2);
                tableRow2.Append(tableCell5);
                tableRow2.Append(tableCell6);
                tableRow2.Append(tableCell7);
                tableRow2.Append(tableCell8);
                tableRow2.Append(tableCell9);

                table1.Append(tableRow2);
                bool isFirst = true;
                foreach (var demension in measurement.Value)
                {
                    if(isFirst)
                    {
                        isFirst = false;
                        continue;
                    }
                    TableRow tableRow3 = new TableRow() { RsidTableRowMarkRevision = "00885DA2", RsidTableRowAddition = "00522EB6", RsidTableRowProperties = "00522EB6" };

                    TableRowProperties tableRowProperties3 = new TableRowProperties();
                    TableRowHeight tableRowHeight3 = new TableRowHeight() { Val = (UInt32Value)367U };
                    TableJustification tableJustification4 = new TableJustification() { Val = TableRowAlignmentValues.Center };

                    tableRowProperties3.Append(tableRowHeight3);
                    tableRowProperties3.Append(tableJustification4);

                    TableCell tableCell10 = new TableCell();

                    TableCellProperties tableCellProperties10 = new TableCellProperties();
                    TableCellWidth tableCellWidth10 = new TableCellWidth() { Width = "1986", Type = TableWidthUnitValues.Dxa };
                    VerticalMerge verticalMerge2 = new VerticalMerge();
                    Shading shading8 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
                    TableCellVerticalAlignment tableCellVerticalAlignment9 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

                    tableCellProperties10.Append(tableCellWidth10);
                    tableCellProperties10.Append(verticalMerge2);
                    tableCellProperties10.Append(shading8);
                    tableCellProperties10.Append(tableCellVerticalAlignment9);

                    Paragraph paragraph12 = new Paragraph() { RsidParagraphMarkRevision = "00885DA2", RsidParagraphAddition = "00522EB6", RsidParagraphProperties = "001F4E3B", RsidRunAdditionDefault = "00522EB6" };

                    ParagraphProperties paragraphProperties12 = new ParagraphProperties();
                    WidowControl widowControl12 = new WidowControl();
                    AdjustRightIndent adjustRightIndent12 = new AdjustRightIndent() { Val = false };
                    SnapToGrid snapToGrid12 = new SnapToGrid() { Val = false };
                    Indentation indentation12 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
                    Justification justification12 = new Justification() { Val = JustificationValues.Center };

                    ParagraphMarkRunProperties paragraphMarkRunProperties12 = new ParagraphMarkRunProperties();
                    RunFonts runFonts54 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体", ComplexScript = "Tahoma" };
                    Bold bold48 = new Bold();
                    Color color41 = new Color() { Val = "000000" };
                    Kern kern51 = new Kern() { Val = (UInt32Value)0U };
                    FontSize fontSize51 = new FontSize() { Val = "24" };
                    FontSizeComplexScript fontSizeComplexScript51 = new FontSizeComplexScript() { Val = "24" };

                    paragraphMarkRunProperties12.Append(runFonts54);
                    paragraphMarkRunProperties12.Append(bold48);
                    paragraphMarkRunProperties12.Append(color41);
                    paragraphMarkRunProperties12.Append(kern51);
                    paragraphMarkRunProperties12.Append(fontSize51);
                    paragraphMarkRunProperties12.Append(fontSizeComplexScript51);

                    paragraphProperties12.Append(widowControl12);
                    paragraphProperties12.Append(adjustRightIndent12);
                    paragraphProperties12.Append(snapToGrid12);
                    paragraphProperties12.Append(indentation12);
                    paragraphProperties12.Append(justification12);
                    paragraphProperties12.Append(paragraphMarkRunProperties12);

                    paragraph12.Append(paragraphProperties12);

                    tableCell10.Append(tableCellProperties10);
                    tableCell10.Append(paragraph12);

                    TableCell tableCell11 = new TableCell();

                    TableCellProperties tableCellProperties11 = new TableCellProperties();
                    TableCellWidth tableCellWidth11 = new TableCellWidth() { Width = "3186", Type = TableWidthUnitValues.Dxa };

                    TableCellBorders tableCellBorders5 = new TableCellBorders();
                    TopBorder topBorder6 = new TopBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
                    LeftBorder leftBorder6 = new LeftBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
                    BottomBorder bottomBorder6 = new BottomBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
                    RightBorder rightBorder6 = new RightBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

                    tableCellBorders5.Append(topBorder6);
                    tableCellBorders5.Append(leftBorder6);
                    tableCellBorders5.Append(bottomBorder6);
                    tableCellBorders5.Append(rightBorder6);
                    Shading shading9 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
                    TableCellVerticalAlignment tableCellVerticalAlignment10 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

                    tableCellProperties11.Append(tableCellWidth11);
                    tableCellProperties11.Append(tableCellBorders5);
                    tableCellProperties11.Append(shading9);
                    tableCellProperties11.Append(tableCellVerticalAlignment10);

                    Paragraph paragraph13 = new Paragraph() { RsidParagraphMarkRevision = "00A844B0", RsidParagraphAddition = "00522EB6", RsidParagraphProperties = "006F5D65", RsidRunAdditionDefault = "00522EB6" };

                    ParagraphProperties paragraphProperties13 = new ParagraphProperties();
                    WidowControl widowControl13 = new WidowControl();
                    AdjustRightIndent adjustRightIndent13 = new AdjustRightIndent() { Val = false };
                    SnapToGrid snapToGrid13 = new SnapToGrid() { Val = false };
                    Indentation indentation13 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
                    Justification justification13 = new Justification() { Val = JustificationValues.Center };

                    ParagraphMarkRunProperties paragraphMarkRunProperties13 = new ParagraphMarkRunProperties();
                    RunFonts runFonts55 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体", ComplexScript = "Tahoma" };
                    Bold bold49 = new Bold();
                    Color color42 = new Color() { Val = "000000" };
                    Kern kern52 = new Kern() { Val = (UInt32Value)0U };
                    FontSize fontSize52 = new FontSize() { Val = "24" };
                    FontSizeComplexScript fontSizeComplexScript52 = new FontSizeComplexScript() { Val = "24" };

                    paragraphMarkRunProperties13.Append(runFonts55);
                    paragraphMarkRunProperties13.Append(bold49);
                    paragraphMarkRunProperties13.Append(color42);
                    paragraphMarkRunProperties13.Append(kern52);
                    paragraphMarkRunProperties13.Append(fontSize52);
                    paragraphMarkRunProperties13.Append(fontSizeComplexScript52);

                    paragraphProperties13.Append(widowControl13);
                    paragraphProperties13.Append(adjustRightIndent13);
                    paragraphProperties13.Append(snapToGrid13);
                    paragraphProperties13.Append(indentation13);
                    paragraphProperties13.Append(justification13);
                    paragraphProperties13.Append(paragraphMarkRunProperties13);

                    Run run43 = new Run() { RsidRunProperties = "00A844B0" };

                    RunProperties runProperties43 = new RunProperties();
                    RunFonts runFonts56 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体", ComplexScript = "Tahoma" };
                    Bold bold50 = new Bold();
                    Color color43 = new Color() { Val = "000000" };
                    Kern kern53 = new Kern() { Val = (UInt32Value)0U };
                    FontSize fontSize53 = new FontSize() { Val = "24" };
                    FontSizeComplexScript fontSizeComplexScript53 = new FontSizeComplexScript() { Val = "24" };

                    runProperties43.Append(runFonts56);
                    runProperties43.Append(bold50);
                    runProperties43.Append(color43);
                    runProperties43.Append(kern53);
                    runProperties43.Append(fontSize53);
                    runProperties43.Append(fontSizeComplexScript53);
                    Text text13 = new Text();
                    text13.Text = demension.Key;//TODO其他维度名

                    run43.Append(runProperties43);
                    run43.Append(text13);

                    paragraph13.Append(paragraphProperties13);
                    paragraph13.Append(run43);

                    tableCell11.Append(tableCellProperties11);
                    tableCell11.Append(paragraph13);

                    TableCell tableCell12 = new TableCell();

                    TableCellProperties tableCellProperties12 = new TableCellProperties();
                    TableCellWidth tableCellWidth12 = new TableCellWidth() { Width = "935", Type = TableWidthUnitValues.Dxa };

                    TableCellBorders tableCellBorders6 = new TableCellBorders();
                    TopBorder topBorder7 = new TopBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
                    LeftBorder leftBorder7 = new LeftBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
                    BottomBorder bottomBorder7 = new BottomBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
                    RightBorder rightBorder7 = new RightBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

                    tableCellBorders6.Append(topBorder7);
                    tableCellBorders6.Append(leftBorder7);
                    tableCellBorders6.Append(bottomBorder7);
                    tableCellBorders6.Append(rightBorder7);

                    tableCellProperties12.Append(tableCellWidth12);
                    tableCellProperties12.Append(tableCellBorders6);

                    Paragraph paragraph14 = new Paragraph() { RsidParagraphMarkRevision = "00A844B0", RsidParagraphAddition = "00522EB6", RsidParagraphProperties = "001F4E3B", RsidRunAdditionDefault = "007B3C50" };

                    ParagraphProperties paragraphProperties14 = new ParagraphProperties();
                    WidowControl widowControl14 = new WidowControl();
                    AdjustRightIndent adjustRightIndent14 = new AdjustRightIndent() { Val = false };
                    SnapToGrid snapToGrid14 = new SnapToGrid() { Val = false };
                    Indentation indentation14 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
                    Justification justification14 = new Justification() { Val = JustificationValues.Center };

                    ParagraphMarkRunProperties paragraphMarkRunProperties14 = new ParagraphMarkRunProperties();
                    RunFonts runFonts58 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体", ComplexScript = "Tahoma" };
                    Bold bold52 = new Bold();
                    Color color45 = new Color() { Val = "000000" };
                    Kern kern55 = new Kern() { Val = (UInt32Value)0U };
                    FontSize fontSize55 = new FontSize() { Val = "24" };
                    FontSizeComplexScript fontSizeComplexScript55 = new FontSizeComplexScript() { Val = "24" };

                    paragraphMarkRunProperties14.Append(runFonts58);
                    paragraphMarkRunProperties14.Append(bold52);
                    paragraphMarkRunProperties14.Append(color45);
                    paragraphMarkRunProperties14.Append(kern55);
                    paragraphMarkRunProperties14.Append(fontSize55);
                    paragraphMarkRunProperties14.Append(fontSizeComplexScript55);

                    paragraphProperties14.Append(widowControl14);
                    paragraphProperties14.Append(adjustRightIndent14);
                    paragraphProperties14.Append(snapToGrid14);
                    paragraphProperties14.Append(indentation14);
                    paragraphProperties14.Append(justification14);
                    paragraphProperties14.Append(paragraphMarkRunProperties14);

                    Run run45 = new Run();

                    RunProperties runProperties45 = new RunProperties();
                    RunFonts runFonts59 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", EastAsia = "宋体", ComplexScript = "Times New Roman" };

                    runProperties45.Append(runFonts59);
                    Text text15 = new Text();
                    text15.Text = demension.Value.Item4.ToString();//TODO其他得分

                    run45.Append(runProperties45);
                    run45.Append(text15);

                    paragraph14.Append(paragraphProperties14);
                    paragraph14.Append(run45);

                    tableCell12.Append(tableCellProperties12);
                    tableCell12.Append(paragraph14);

                    TableCell tableCell13 = new TableCell();

                    TableCellProperties tableCellProperties13 = new TableCellProperties();
                    TableCellWidth tableCellWidth13 = new TableCellWidth() { Width = "935", Type = TableWidthUnitValues.Dxa };

                    TableCellBorders tableCellBorders7 = new TableCellBorders();
                    TopBorder topBorder8 = new TopBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
                    LeftBorder leftBorder8 = new LeftBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
                    BottomBorder bottomBorder8 = new BottomBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
                    RightBorder rightBorder8 = new RightBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

                    tableCellBorders7.Append(topBorder8);
                    tableCellBorders7.Append(leftBorder8);
                    tableCellBorders7.Append(bottomBorder8);
                    tableCellBorders7.Append(rightBorder8);
                    Shading shading10 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
                    TableCellVerticalAlignment tableCellVerticalAlignment11 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

                    tableCellProperties13.Append(tableCellWidth13);
                    tableCellProperties13.Append(tableCellBorders7);
                    tableCellProperties13.Append(shading10);
                    tableCellProperties13.Append(tableCellVerticalAlignment11);

                    Paragraph paragraph15 = new Paragraph() { RsidParagraphMarkRevision = "00A844B0", RsidParagraphAddition = "00522EB6", RsidParagraphProperties = "00A844B0", RsidRunAdditionDefault = "007B3C50" };

                    ParagraphProperties paragraphProperties15 = new ParagraphProperties();
                    WidowControl widowControl15 = new WidowControl();
                    AdjustRightIndent adjustRightIndent15 = new AdjustRightIndent() { Val = false };
                    SnapToGrid snapToGrid15 = new SnapToGrid() { Val = false };
                    Indentation indentation15 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
                    Justification justification15 = new Justification() { Val = JustificationValues.Center };

                    ParagraphMarkRunProperties paragraphMarkRunProperties15 = new ParagraphMarkRunProperties();
                    RunFonts runFonts60 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", EastAsia = "宋体", ComplexScript = "Times New Roman" };

                    paragraphMarkRunProperties15.Append(runFonts60);

                    paragraphProperties15.Append(widowControl15);
                    paragraphProperties15.Append(adjustRightIndent15);
                    paragraphProperties15.Append(snapToGrid15);
                    paragraphProperties15.Append(indentation15);
                    paragraphProperties15.Append(justification15);
                    paragraphProperties15.Append(paragraphMarkRunProperties15);

                    Run run46 = new Run();

                    RunProperties runProperties46 = new RunProperties();
                    RunFonts runFonts61 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", EastAsia = "宋体", ComplexScript = "Times New Roman" };

                    runProperties46.Append(runFonts61);
                    Text text16 = new Text();
                    text16.Text = demension.Value.Item2;

                    run46.Append(runProperties46);
                    run46.Append(text16);

                    paragraph15.Append(paragraphProperties15);
                    paragraph15.Append(run46);

                    tableCell13.Append(tableCellProperties13);
                    tableCell13.Append(paragraph15);

                    TableCell tableCell14 = new TableCell();

                    TableCellProperties tableCellProperties14 = new TableCellProperties();
                    TableCellWidth tableCellWidth14 = new TableCellWidth() { Width = "1243", Type = TableWidthUnitValues.Dxa };

                    TableCellBorders tableCellBorders8 = new TableCellBorders();
                    TopBorder topBorder9 = new TopBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
                    LeftBorder leftBorder9 = new LeftBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
                    BottomBorder bottomBorder9 = new BottomBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
                    RightBorder rightBorder9 = new RightBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

                    tableCellBorders8.Append(topBorder9);
                    tableCellBorders8.Append(leftBorder9);
                    tableCellBorders8.Append(bottomBorder9);
                    tableCellBorders8.Append(rightBorder9);
                    TableCellVerticalAlignment tableCellVerticalAlignment12 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

                    tableCellProperties14.Append(tableCellWidth14);
                    tableCellProperties14.Append(tableCellBorders8);
                    tableCellProperties14.Append(tableCellVerticalAlignment12);

                    Paragraph paragraph16 = new Paragraph() { RsidParagraphMarkRevision = "00885DA2", RsidParagraphAddition = "00522EB6", RsidParagraphProperties = "001F4E3B", RsidRunAdditionDefault = "00F82954" };

                    ParagraphProperties paragraphProperties16 = new ParagraphProperties();
                    WidowControl widowControl16 = new WidowControl();
                    AdjustRightIndent adjustRightIndent16 = new AdjustRightIndent() { Val = false };
                    SnapToGrid snapToGrid16 = new SnapToGrid() { Val = false };
                    Indentation indentation16 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
                    Justification justification16 = new Justification() { Val = JustificationValues.Center };

                    ParagraphMarkRunProperties paragraphMarkRunProperties16 = new ParagraphMarkRunProperties();
                    RunFonts runFonts62 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", EastAsia = "宋体", ComplexScript = "Times New Roman" };
                    Color color46 = new Color() { Val = "FF0000" };
                    Kern kern56 = new Kern() { Val = (UInt32Value)0U };
                    FontSize fontSize56 = new FontSize() { Val = "2" };
                    FontSizeComplexScript fontSizeComplexScript56 = new FontSizeComplexScript() { Val = "2" };

                    paragraphMarkRunProperties16.Append(runFonts62);
                    paragraphMarkRunProperties16.Append(color46);
                    paragraphMarkRunProperties16.Append(kern56);
                    paragraphMarkRunProperties16.Append(fontSize56);
                    paragraphMarkRunProperties16.Append(fontSizeComplexScript56);

                    paragraphProperties16.Append(widowControl16);
                    paragraphProperties16.Append(adjustRightIndent16);
                    paragraphProperties16.Append(snapToGrid16);
                    paragraphProperties16.Append(indentation16);
                    paragraphProperties16.Append(justification16);
                    paragraphProperties16.Append(paragraphMarkRunProperties16);

                    paragraph16.Append(paragraphProperties16);

                    Paragraph paragraph17 = new Paragraph() { RsidParagraphMarkRevision = "00885DA2", RsidParagraphAddition = "00522EB6", RsidParagraphProperties = "001F4E3B", RsidRunAdditionDefault = "00522EB6" };

                    ParagraphProperties paragraphProperties17 = new ParagraphProperties();
                    WidowControl widowControl17 = new WidowControl();
                    AdjustRightIndent adjustRightIndent17 = new AdjustRightIndent() { Val = false };
                    SnapToGrid snapToGrid17 = new SnapToGrid() { Val = false };
                    Indentation indentation17 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
                    Justification justification17 = new Justification() { Val = JustificationValues.Center };

                    ParagraphMarkRunProperties paragraphMarkRunProperties17 = new ParagraphMarkRunProperties();
                    RunFonts runFonts92 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体", ComplexScript = "Tahoma" };
                    Color color76 = new Color() { Val = "FF0000" };
                    Kern kern86 = new Kern() { Val = (UInt32Value)0U };
                    FontSize fontSize86 = new FontSize() { Val = "24" };
                    FontSizeComplexScript fontSizeComplexScript86 = new FontSizeComplexScript() { Val = "24" };

                    paragraphMarkRunProperties17.Append(runFonts92);
                    paragraphMarkRunProperties17.Append(color76);
                    paragraphMarkRunProperties17.Append(kern86);
                    paragraphMarkRunProperties17.Append(fontSize86);
                    paragraphMarkRunProperties17.Append(fontSizeComplexScript86);

                    paragraphProperties17.Append(widowControl17);
                    paragraphProperties17.Append(adjustRightIndent17);
                    paragraphProperties17.Append(snapToGrid17);
                    paragraphProperties17.Append(indentation17);
                    paragraphProperties17.Append(justification17);
                    paragraphProperties17.Append(paragraphMarkRunProperties17);

                    Run run76 = new Run() { RsidRunProperties = "00885DA2" };

                    RunProperties runProperties76 = new RunProperties();
                    RunFonts runFonts93 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体", ComplexScript = "Tahoma" };
                    Color color77 = new Color() { Val = "FF0000" };
                    Kern kern87 = new Kern() { Val = (UInt32Value)0U };
                    FontSize fontSize87 = new FontSize() { Val = "24" };
                    FontSizeComplexScript fontSizeComplexScript87 = new FontSizeComplexScript() { Val = "24" };

                    runProperties76.Append(runFonts93);
                    runProperties76.Append(color77);
                    runProperties76.Append(kern87);
                    runProperties76.Append(fontSize87);
                    runProperties76.Append(fontSizeComplexScript87);
                    Text text17 = new Text();
                    text17.Text = demension.Value.Item3== eInfluence.消极影响?"■":"";

                    run76.Append(runProperties76);
                    run76.Append(text17);

                    paragraph17.Append(paragraphProperties17);
                    paragraph17.Append(run76);

                    Paragraph paragraph18 = new Paragraph() { RsidParagraphMarkRevision = "00885DA2", RsidParagraphAddition = "00522EB6", RsidParagraphProperties = "001F4E3B", RsidRunAdditionDefault = "00F82954" };

                    ParagraphProperties paragraphProperties18 = new ParagraphProperties();
                    WidowControl widowControl18 = new WidowControl();
                    AdjustRightIndent adjustRightIndent18 = new AdjustRightIndent() { Val = false };
                    SnapToGrid snapToGrid18 = new SnapToGrid() { Val = false };
                    Indentation indentation18 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
                    Justification justification18 = new Justification() { Val = JustificationValues.Center };

                    ParagraphMarkRunProperties paragraphMarkRunProperties18 = new ParagraphMarkRunProperties();
                    RunFonts runFonts94 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
                    Bold bold82 = new Bold();
                    Color color78 = new Color() { Val = "FF0000" };
                    Kern kern88 = new Kern() { Val = (UInt32Value)0U };
                    FontSize fontSize88 = new FontSize() { Val = "2" };
                    FontSizeComplexScript fontSizeComplexScript88 = new FontSizeComplexScript() { Val = "2" };

                    paragraphMarkRunProperties18.Append(runFonts94);
                    paragraphMarkRunProperties18.Append(bold82);
                    paragraphMarkRunProperties18.Append(color78);
                    paragraphMarkRunProperties18.Append(kern88);
                    paragraphMarkRunProperties18.Append(fontSize88);
                    paragraphMarkRunProperties18.Append(fontSizeComplexScript88);

                    paragraphProperties18.Append(widowControl18);
                    paragraphProperties18.Append(adjustRightIndent18);
                    paragraphProperties18.Append(snapToGrid18);
                    paragraphProperties18.Append(indentation18);
                    paragraphProperties18.Append(justification18);
                    paragraphProperties18.Append(paragraphMarkRunProperties18);

                    paragraph18.Append(paragraphProperties18);

                    tableCell14.Append(tableCellProperties14);
                    tableCell14.Append(paragraph16);
                    tableCell14.Append(paragraph17);
                    tableCell14.Append(paragraph18);

                    tableRow3.Append(tableRowProperties3);
                    tableRow3.Append(tableCell10);
                    tableRow3.Append(tableCell11);
                    tableRow3.Append(tableCell12);
                    tableRow3.Append(tableCell13);
                    tableRow3.Append(tableCell14);

                    table1.Append(tableRow3);
                }
              
                
            }




#region testaera
           /* TableRow tableRow4 = new TableRow() { RsidTableRowMarkRevision = "007B3C50", RsidTableRowAddition = "007B3C50", RsidTableRowProperties = "007B3C50" };

            TableRowProperties tableRowProperties4 = new TableRowProperties();
            TableRowHeight tableRowHeight4 = new TableRowHeight() { Val = (UInt32Value)647U };
            TableJustification tableJustification5 = new TableJustification() { Val = TableRowAlignmentValues.Center };

            tableRowProperties4.Append(tableRowHeight4);
            tableRowProperties4.Append(tableJustification5);

            TableCell tableCell15 = new TableCell();

            TableCellProperties tableCellProperties15 = new TableCellProperties();
            TableCellWidth tableCellWidth15 = new TableCellWidth() { Width = "1986", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge3 = new VerticalMerge() { Val = MergedCellValues.Restart };
            Shading shading11 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment13 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties15.Append(tableCellWidth15);
            tableCellProperties15.Append(verticalMerge3);
            tableCellProperties15.Append(shading11);
            tableCellProperties15.Append(tableCellVerticalAlignment13);

            Paragraph paragraph19 = new Paragraph() { RsidParagraphMarkRevision = "00885DA2", RsidParagraphAddition = "007B3C50", RsidParagraphProperties = "001F4E3B", RsidRunAdditionDefault = "007B3C50" };

            ParagraphProperties paragraphProperties19 = new ParagraphProperties();
            WidowControl widowControl19 = new WidowControl();
            AdjustRightIndent adjustRightIndent19 = new AdjustRightIndent() { Val = false };
            SnapToGrid snapToGrid19 = new SnapToGrid() { Val = false };
            Indentation indentation19 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification19 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties19 = new ParagraphMarkRunProperties();
            RunFonts runFonts96 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold84 = new Bold();
            Color color80 = new Color() { Val = "000000" };
            Kern kern90 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize90 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript90 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties19.Append(runFonts96);
            paragraphMarkRunProperties19.Append(bold84);
            paragraphMarkRunProperties19.Append(color80);
            paragraphMarkRunProperties19.Append(kern90);
            paragraphMarkRunProperties19.Append(fontSize90);
            paragraphMarkRunProperties19.Append(fontSizeComplexScript90);

            paragraphProperties19.Append(widowControl19);
            paragraphProperties19.Append(adjustRightIndent19);
            paragraphProperties19.Append(snapToGrid19);
            paragraphProperties19.Append(indentation19);
            paragraphProperties19.Append(justification19);
            paragraphProperties19.Append(paragraphMarkRunProperties19);

            Run run78 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties78 = new RunProperties();
            RunFonts runFonts97 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold85 = new Bold();
            Color color81 = new Color() { Val = "000000" };
            Kern kern91 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize91 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript91 = new FontSizeComplexScript() { Val = "24" };

            runProperties78.Append(runFonts97);
            runProperties78.Append(bold85);
            runProperties78.Append(color81);
            runProperties78.Append(kern91);
            runProperties78.Append(fontSize91);
            runProperties78.Append(fontSizeComplexScript91);
            Text text18 = new Text();
            text18.Text = "师生";

            run78.Append(runProperties78);
            run78.Append(text18);

            Run run79 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties79 = new RunProperties();
            RunFonts runFonts98 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold86 = new Bold();
            Color color82 = new Color() { Val = "000000" };
            Kern kern92 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize92 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript92 = new FontSizeComplexScript() { Val = "24" };

            runProperties79.Append(runFonts98);
            runProperties79.Append(bold86);
            runProperties79.Append(color82);
            runProperties79.Append(kern92);
            runProperties79.Append(fontSize92);
            runProperties79.Append(fontSizeComplexScript92);
            Text text19 = new Text();
            text19.Text = "关系";

            run79.Append(runProperties79);
            run79.Append(text19);

            paragraph19.Append(paragraphProperties19);
            paragraph19.Append(run78);
            paragraph19.Append(run79);

            Paragraph paragraph20 = new Paragraph() { RsidParagraphMarkRevision = "00885DA2", RsidParagraphAddition = "007B3C50", RsidParagraphProperties = "001F4E3B", RsidRunAdditionDefault = "007B3C50" };

            ParagraphProperties paragraphProperties20 = new ParagraphProperties();
            WidowControl widowControl20 = new WidowControl();
            AdjustRightIndent adjustRightIndent20 = new AdjustRightIndent() { Val = false };
            SnapToGrid snapToGrid20 = new SnapToGrid() { Val = false };
            Indentation indentation20 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification20 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties20 = new ParagraphMarkRunProperties();
            RunFonts runFonts99 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", EastAsia = "宋体", ComplexScript = "Times New Roman" };
            Kern kern93 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize93 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript93 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties20.Append(runFonts99);
            paragraphMarkRunProperties20.Append(kern93);
            paragraphMarkRunProperties20.Append(fontSize93);
            paragraphMarkRunProperties20.Append(fontSizeComplexScript93);

            paragraphProperties20.Append(widowControl20);
            paragraphProperties20.Append(adjustRightIndent20);
            paragraphProperties20.Append(snapToGrid20);
            paragraphProperties20.Append(indentation20);
            paragraphProperties20.Append(justification20);
            paragraphProperties20.Append(paragraphMarkRunProperties20);

            paragraph20.Append(paragraphProperties20);

            tableCell15.Append(tableCellProperties15);
            tableCell15.Append(paragraph19);
            tableCell15.Append(paragraph20);

            TableCell tableCell16 = new TableCell();

            TableCellProperties tableCellProperties16 = new TableCellProperties();
            TableCellWidth tableCellWidth16 = new TableCellWidth() { Width = "3186", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders9 = new TableCellBorders();
            TopBorder topBorder10 = new TopBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder10 = new LeftBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder10 = new BottomBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder10 = new RightBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders9.Append(topBorder10);
            tableCellBorders9.Append(leftBorder10);
            tableCellBorders9.Append(bottomBorder10);
            tableCellBorders9.Append(rightBorder10);
            Shading shading12 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment14 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties16.Append(tableCellWidth16);
            tableCellProperties16.Append(tableCellBorders9);
            tableCellProperties16.Append(shading12);
            tableCellProperties16.Append(tableCellVerticalAlignment14);

            Paragraph paragraph21 = new Paragraph() { RsidParagraphMarkRevision = "007B3C50", RsidParagraphAddition = "007B3C50", RsidParagraphProperties = "006F5D65", RsidRunAdditionDefault = "007B3C50" };

            ParagraphProperties paragraphProperties21 = new ParagraphProperties();
            WidowControl widowControl21 = new WidowControl();
            AdjustRightIndent adjustRightIndent21 = new AdjustRightIndent() { Val = false };
            SnapToGrid snapToGrid21 = new SnapToGrid() { Val = false };
            Indentation indentation21 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification21 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties21 = new ParagraphMarkRunProperties();
            RunFonts runFonts100 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", EastAsia = "宋体", ComplexScript = "Times New Roman" };

            paragraphMarkRunProperties21.Append(runFonts100);

            paragraphProperties21.Append(widowControl21);
            paragraphProperties21.Append(adjustRightIndent21);
            paragraphProperties21.Append(snapToGrid21);
            paragraphProperties21.Append(indentation21);
            paragraphProperties21.Append(justification21);
            paragraphProperties21.Append(paragraphMarkRunProperties21);

            Run run80 = new Run() { RsidRunProperties = "00522EB6" };

            RunProperties runProperties80 = new RunProperties();
            RunFonts runFonts101 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold87 = new Bold();
            Color color83 = new Color() { Val = "000000" };
            Kern kern94 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize94 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript94 = new FontSizeComplexScript() { Val = "24" };

            runProperties80.Append(runFonts101);
            runProperties80.Append(bold87);
            runProperties80.Append(color83);
            runProperties80.Append(kern94);
            runProperties80.Append(fontSize94);
            runProperties80.Append(fontSizeComplexScript94);
            Text text20 = new Text();
            text20.Text = "回避";

            run80.Append(runProperties80);
            run80.Append(text20);

            Run run81 = new Run() { RsidRunProperties = "00522EB6" };

            RunProperties runProperties81 = new RunProperties();
            RunFonts runFonts102 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold88 = new Bold();
            Color color84 = new Color() { Val = "000000" };
            Kern kern95 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize95 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript95 = new FontSizeComplexScript() { Val = "24" };

            runProperties81.Append(runFonts102);
            runProperties81.Append(bold88);
            runProperties81.Append(color84);
            runProperties81.Append(kern95);
            runProperties81.Append(fontSize95);
            runProperties81.Append(fontSizeComplexScript95);
            Text text21 = new Text();
            text21.Text = "性";

            run81.Append(runProperties81);
            run81.Append(text21);

            paragraph21.Append(paragraphProperties21);
            paragraph21.Append(run80);
            paragraph21.Append(run81);

            tableCell16.Append(tableCellProperties16);
            tableCell16.Append(paragraph21);

            TableCell tableCell17 = new TableCell();

            TableCellProperties tableCellProperties17 = new TableCellProperties();
            TableCellWidth tableCellWidth17 = new TableCellWidth() { Width = "935", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders10 = new TableCellBorders();
            TopBorder topBorder11 = new TopBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder11 = new LeftBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder11 = new BottomBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder11 = new RightBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders10.Append(topBorder11);
            tableCellBorders10.Append(leftBorder11);
            tableCellBorders10.Append(bottomBorder11);
            tableCellBorders10.Append(rightBorder11);

            tableCellProperties17.Append(tableCellWidth17);
            tableCellProperties17.Append(tableCellBorders10);

            Paragraph paragraph22 = new Paragraph() { RsidParagraphMarkRevision = "007B3C50", RsidParagraphAddition = "007B3C50", RsidParagraphProperties = "007B3C50", RsidRunAdditionDefault = "007B3C50" };

            ParagraphProperties paragraphProperties22 = new ParagraphProperties();
            WidowControl widowControl22 = new WidowControl();
            AdjustRightIndent adjustRightIndent22 = new AdjustRightIndent() { Val = false };
            SnapToGrid snapToGrid22 = new SnapToGrid() { Val = false };
            Indentation indentation22 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification22 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties22 = new ParagraphMarkRunProperties();
            RunFonts runFonts103 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", EastAsia = "宋体", ComplexScript = "Times New Roman" };

            paragraphMarkRunProperties22.Append(runFonts103);

            paragraphProperties22.Append(widowControl22);
            paragraphProperties22.Append(adjustRightIndent22);
            paragraphProperties22.Append(snapToGrid22);
            paragraphProperties22.Append(indentation22);
            paragraphProperties22.Append(justification22);
            paragraphProperties22.Append(paragraphMarkRunProperties22);

            Run run82 = new Run();

            RunProperties runProperties82 = new RunProperties();
            RunFonts runFonts104 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", EastAsia = "宋体", ComplexScript = "Times New Roman" };

            runProperties82.Append(runFonts104);
            Text text22 = new Text();
            text22.Text = "123";

            run82.Append(runProperties82);
            run82.Append(text22);

            paragraph22.Append(paragraphProperties22);
            paragraph22.Append(run82);

            tableCell17.Append(tableCellProperties17);
            tableCell17.Append(paragraph22);

            TableCell tableCell18 = new TableCell();

            TableCellProperties tableCellProperties18 = new TableCellProperties();
            TableCellWidth tableCellWidth18 = new TableCellWidth() { Width = "935", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders11 = new TableCellBorders();
            TopBorder topBorder12 = new TopBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder12 = new LeftBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder12 = new BottomBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder12 = new RightBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders11.Append(topBorder12);
            tableCellBorders11.Append(leftBorder12);
            tableCellBorders11.Append(bottomBorder12);
            tableCellBorders11.Append(rightBorder12);
            Shading shading13 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment15 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties18.Append(tableCellWidth18);
            tableCellProperties18.Append(tableCellBorders11);
            tableCellProperties18.Append(shading13);
            tableCellProperties18.Append(tableCellVerticalAlignment15);

            Paragraph paragraph23 = new Paragraph() { RsidParagraphMarkRevision = "00522EB6", RsidParagraphAddition = "007B3C50", RsidParagraphProperties = "001F4E3B", RsidRunAdditionDefault = "007B3C50" };

            ParagraphProperties paragraphProperties23 = new ParagraphProperties();
            WidowControl widowControl23 = new WidowControl();
            AdjustRightIndent adjustRightIndent23 = new AdjustRightIndent() { Val = false };
            SnapToGrid snapToGrid23 = new SnapToGrid() { Val = false };
            Indentation indentation23 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification23 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties23 = new ParagraphMarkRunProperties();
            RunFonts runFonts105 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", EastAsia = "宋体", ComplexScript = "Times New Roman" };

            paragraphMarkRunProperties23.Append(runFonts105);

            paragraphProperties23.Append(widowControl23);
            paragraphProperties23.Append(adjustRightIndent23);
            paragraphProperties23.Append(snapToGrid23);
            paragraphProperties23.Append(indentation23);
            paragraphProperties23.Append(justification23);
            paragraphProperties23.Append(paragraphMarkRunProperties23);

            Run run83 = new Run();

            RunProperties runProperties83 = new RunProperties();
            RunFonts runFonts106 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", EastAsia = "宋体", ComplexScript = "Times New Roman" };

            runProperties83.Append(runFonts106);
            Text text23 = new Text();
            text23.Text = "123";

            run83.Append(runProperties83);
            run83.Append(text23);

            paragraph23.Append(paragraphProperties23);
            paragraph23.Append(run83);

            tableCell18.Append(tableCellProperties18);
            tableCell18.Append(paragraph23);

            TableCell tableCell19 = new TableCell();

            TableCellProperties tableCellProperties19 = new TableCellProperties();
            TableCellWidth tableCellWidth19 = new TableCellWidth() { Width = "1243", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders12 = new TableCellBorders();
            TopBorder topBorder13 = new TopBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder13 = new LeftBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder13 = new BottomBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder13 = new RightBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders12.Append(topBorder13);
            tableCellBorders12.Append(leftBorder13);
            tableCellBorders12.Append(bottomBorder13);
            tableCellBorders12.Append(rightBorder13);
            TableCellVerticalAlignment tableCellVerticalAlignment16 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties19.Append(tableCellWidth19);
            tableCellProperties19.Append(tableCellBorders12);
            tableCellProperties19.Append(tableCellVerticalAlignment16);

            Paragraph paragraph24 = new Paragraph() { RsidParagraphMarkRevision = "007B3C50", RsidParagraphAddition = "007B3C50", RsidParagraphProperties = "001F4E3B", RsidRunAdditionDefault = "0018090C" };

            ParagraphProperties paragraphProperties24 = new ParagraphProperties();
            WidowControl widowControl24 = new WidowControl();
            AdjustRightIndent adjustRightIndent24 = new AdjustRightIndent() { Val = false };
            SnapToGrid snapToGrid24 = new SnapToGrid() { Val = false };
            Indentation indentation24 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification24 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties24 = new ParagraphMarkRunProperties();
            RunFonts runFonts107 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", EastAsia = "宋体", ComplexScript = "Times New Roman" };

            paragraphMarkRunProperties24.Append(runFonts107);

            paragraphProperties24.Append(widowControl24);
            paragraphProperties24.Append(adjustRightIndent24);
            paragraphProperties24.Append(snapToGrid24);
            paragraphProperties24.Append(indentation24);
            paragraphProperties24.Append(justification24);
            paragraphProperties24.Append(paragraphMarkRunProperties24);

            Run run84 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties84 = new RunProperties();
            RunFonts runFonts108 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Color color85 = new Color() { Val = "FF0000" };
            Kern kern96 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize96 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript96 = new FontSizeComplexScript() { Val = "24" };

            runProperties84.Append(runFonts108);
            runProperties84.Append(color85);
            runProperties84.Append(kern96);
            runProperties84.Append(fontSize96);
            runProperties84.Append(fontSizeComplexScript96);
            Text text24 = new Text();
            text24.Text = "■";

            run84.Append(runProperties84);
            run84.Append(text24);

            paragraph24.Append(paragraphProperties24);
            paragraph24.Append(run84);

            tableCell19.Append(tableCellProperties19);
            tableCell19.Append(paragraph24);

            tableRow4.Append(tableRowProperties4);
            tableRow4.Append(tableCell15);
            tableRow4.Append(tableCell16);
            tableRow4.Append(tableCell17);
            tableRow4.Append(tableCell18);
            tableRow4.Append(tableCell19);

            TableRow tableRow5 = new TableRow() { RsidTableRowMarkRevision = "00885DA2", RsidTableRowAddition = "007B3C50", RsidTableRowProperties = "00522EB6" };

            TableRowProperties tableRowProperties5 = new TableRowProperties();
            TableRowHeight tableRowHeight5 = new TableRowHeight() { Val = (UInt32Value)367U };
            TableJustification tableJustification6 = new TableJustification() { Val = TableRowAlignmentValues.Center };

            tableRowProperties5.Append(tableRowHeight5);
            tableRowProperties5.Append(tableJustification6);

            TableCell tableCell20 = new TableCell();

            TableCellProperties tableCellProperties20 = new TableCellProperties();
            TableCellWidth tableCellWidth20 = new TableCellWidth() { Width = "1986", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge4 = new VerticalMerge();
            Shading shading14 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment17 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties20.Append(tableCellWidth20);
            tableCellProperties20.Append(verticalMerge4);
            tableCellProperties20.Append(shading14);
            tableCellProperties20.Append(tableCellVerticalAlignment17);

            Paragraph paragraph25 = new Paragraph() { RsidParagraphMarkRevision = "00885DA2", RsidParagraphAddition = "007B3C50", RsidParagraphProperties = "001F4E3B", RsidRunAdditionDefault = "007B3C50" };

            ParagraphProperties paragraphProperties25 = new ParagraphProperties();
            WidowControl widowControl25 = new WidowControl();
            AdjustRightIndent adjustRightIndent25 = new AdjustRightIndent() { Val = false };
            SnapToGrid snapToGrid25 = new SnapToGrid() { Val = false };
            Indentation indentation25 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification25 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties25 = new ParagraphMarkRunProperties();
            RunFonts runFonts109 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold89 = new Bold();
            Color color86 = new Color() { Val = "000000" };
            Kern kern97 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize97 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript97 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties25.Append(runFonts109);
            paragraphMarkRunProperties25.Append(bold89);
            paragraphMarkRunProperties25.Append(color86);
            paragraphMarkRunProperties25.Append(kern97);
            paragraphMarkRunProperties25.Append(fontSize97);
            paragraphMarkRunProperties25.Append(fontSizeComplexScript97);

            paragraphProperties25.Append(widowControl25);
            paragraphProperties25.Append(adjustRightIndent25);
            paragraphProperties25.Append(snapToGrid25);
            paragraphProperties25.Append(indentation25);
            paragraphProperties25.Append(justification25);
            paragraphProperties25.Append(paragraphMarkRunProperties25);

            paragraph25.Append(paragraphProperties25);

            tableCell20.Append(tableCellProperties20);
            tableCell20.Append(paragraph25);

            TableCell tableCell21 = new TableCell();

            TableCellProperties tableCellProperties21 = new TableCellProperties();
            TableCellWidth tableCellWidth21 = new TableCellWidth() { Width = "3186", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders13 = new TableCellBorders();
            TopBorder topBorder14 = new TopBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder14 = new LeftBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder14 = new BottomBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder14 = new RightBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders13.Append(topBorder14);
            tableCellBorders13.Append(leftBorder14);
            tableCellBorders13.Append(bottomBorder14);
            tableCellBorders13.Append(rightBorder14);
            Shading shading15 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment18 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties21.Append(tableCellWidth21);
            tableCellProperties21.Append(tableCellBorders13);
            tableCellProperties21.Append(shading15);
            tableCellProperties21.Append(tableCellVerticalAlignment18);

            Paragraph paragraph26 = new Paragraph() { RsidParagraphMarkRevision = "00522EB6", RsidParagraphAddition = "007B3C50", RsidParagraphProperties = "006F5D65", RsidRunAdditionDefault = "007B3C50" };

            ParagraphProperties paragraphProperties26 = new ParagraphProperties();
            WidowControl widowControl26 = new WidowControl();
            AdjustRightIndent adjustRightIndent26 = new AdjustRightIndent() { Val = false };
            SnapToGrid snapToGrid26 = new SnapToGrid() { Val = false };
            Indentation indentation26 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification26 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties26 = new ParagraphMarkRunProperties();
            RunFonts runFonts110 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold90 = new Bold();
            Color color87 = new Color() { Val = "000000" };
            Kern kern98 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize98 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript98 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties26.Append(runFonts110);
            paragraphMarkRunProperties26.Append(bold90);
            paragraphMarkRunProperties26.Append(color87);
            paragraphMarkRunProperties26.Append(kern98);
            paragraphMarkRunProperties26.Append(fontSize98);
            paragraphMarkRunProperties26.Append(fontSizeComplexScript98);

            paragraphProperties26.Append(widowControl26);
            paragraphProperties26.Append(adjustRightIndent26);
            paragraphProperties26.Append(snapToGrid26);
            paragraphProperties26.Append(indentation26);
            paragraphProperties26.Append(justification26);
            paragraphProperties26.Append(paragraphMarkRunProperties26);

            Run run85 = new Run() { RsidRunProperties = "00522EB6" };

            RunProperties runProperties85 = new RunProperties();
            RunFonts runFonts111 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold91 = new Bold();
            Color color88 = new Color() { Val = "000000" };
            Kern kern99 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize99 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript99 = new FontSizeComplexScript() { Val = "24" };

            runProperties85.Append(runFonts111);
            runProperties85.Append(bold91);
            runProperties85.Append(color88);
            runProperties85.Append(kern99);
            runProperties85.Append(fontSize99);
            runProperties85.Append(fontSizeComplexScript99);
            Text text25 = new Text();
            text25.Text = "亲密";

            run85.Append(runProperties85);
            run85.Append(text25);

            Run run86 = new Run() { RsidRunProperties = "00522EB6" };

            RunProperties runProperties86 = new RunProperties();
            RunFonts runFonts112 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold92 = new Bold();
            Color color89 = new Color() { Val = "000000" };
            Kern kern100 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize100 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript100 = new FontSizeComplexScript() { Val = "24" };

            runProperties86.Append(runFonts112);
            runProperties86.Append(bold92);
            runProperties86.Append(color89);
            runProperties86.Append(kern100);
            runProperties86.Append(fontSize100);
            runProperties86.Append(fontSizeComplexScript100);
            Text text26 = new Text();
            text26.Text = "性";

            run86.Append(runProperties86);
            run86.Append(text26);

            paragraph26.Append(paragraphProperties26);
            paragraph26.Append(run85);
            paragraph26.Append(run86);

            tableCell21.Append(tableCellProperties21);
            tableCell21.Append(paragraph26);

            TableCell tableCell22 = new TableCell();

            TableCellProperties tableCellProperties22 = new TableCellProperties();
            TableCellWidth tableCellWidth22 = new TableCellWidth() { Width = "935", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders14 = new TableCellBorders();
            TopBorder topBorder15 = new TopBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder15 = new LeftBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder15 = new BottomBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder15 = new RightBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders14.Append(topBorder15);
            tableCellBorders14.Append(leftBorder15);
            tableCellBorders14.Append(bottomBorder15);
            tableCellBorders14.Append(rightBorder15);

            tableCellProperties22.Append(tableCellWidth22);
            tableCellProperties22.Append(tableCellBorders14);

            Paragraph paragraph27 = new Paragraph() { RsidParagraphMarkRevision = "00522EB6", RsidParagraphAddition = "007B3C50", RsidParagraphProperties = "001F4E3B", RsidRunAdditionDefault = "007B3C50" };

            ParagraphProperties paragraphProperties27 = new ParagraphProperties();
            WidowControl widowControl27 = new WidowControl();
            AdjustRightIndent adjustRightIndent27 = new AdjustRightIndent() { Val = false };
            SnapToGrid snapToGrid27 = new SnapToGrid() { Val = false };
            Indentation indentation27 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification27 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties27 = new ParagraphMarkRunProperties();
            RunFonts runFonts113 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold93 = new Bold();
            Color color90 = new Color() { Val = "000000" };
            Kern kern101 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize101 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript101 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties27.Append(runFonts113);
            paragraphMarkRunProperties27.Append(bold93);
            paragraphMarkRunProperties27.Append(color90);
            paragraphMarkRunProperties27.Append(kern101);
            paragraphMarkRunProperties27.Append(fontSize101);
            paragraphMarkRunProperties27.Append(fontSizeComplexScript101);

            paragraphProperties27.Append(widowControl27);
            paragraphProperties27.Append(adjustRightIndent27);
            paragraphProperties27.Append(snapToGrid27);
            paragraphProperties27.Append(indentation27);
            paragraphProperties27.Append(justification27);
            paragraphProperties27.Append(paragraphMarkRunProperties27);

            paragraph27.Append(paragraphProperties27);

            tableCell22.Append(tableCellProperties22);
            tableCell22.Append(paragraph27);

            TableCell tableCell23 = new TableCell();

            TableCellProperties tableCellProperties23 = new TableCellProperties();
            TableCellWidth tableCellWidth23 = new TableCellWidth() { Width = "935", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders15 = new TableCellBorders();
            TopBorder topBorder16 = new TopBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder16 = new LeftBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder16 = new BottomBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder16 = new RightBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders15.Append(topBorder16);
            tableCellBorders15.Append(leftBorder16);
            tableCellBorders15.Append(bottomBorder16);
            tableCellBorders15.Append(rightBorder16);
            Shading shading16 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment19 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties23.Append(tableCellWidth23);
            tableCellProperties23.Append(tableCellBorders15);
            tableCellProperties23.Append(shading16);
            tableCellProperties23.Append(tableCellVerticalAlignment19);

            Paragraph paragraph28 = new Paragraph() { RsidParagraphMarkRevision = "00522EB6", RsidParagraphAddition = "007B3C50", RsidParagraphProperties = "001F4E3B", RsidRunAdditionDefault = "007B3C50" };

            ParagraphProperties paragraphProperties28 = new ParagraphProperties();
            WidowControl widowControl28 = new WidowControl();
            AdjustRightIndent adjustRightIndent28 = new AdjustRightIndent() { Val = false };
            SnapToGrid snapToGrid28 = new SnapToGrid() { Val = false };
            Indentation indentation28 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification28 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties28 = new ParagraphMarkRunProperties();
            RunFonts runFonts114 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", EastAsia = "宋体", ComplexScript = "Times New Roman" };

            paragraphMarkRunProperties28.Append(runFonts114);

            paragraphProperties28.Append(widowControl28);
            paragraphProperties28.Append(adjustRightIndent28);
            paragraphProperties28.Append(snapToGrid28);
            paragraphProperties28.Append(indentation28);
            paragraphProperties28.Append(justification28);
            paragraphProperties28.Append(paragraphMarkRunProperties28);

            paragraph28.Append(paragraphProperties28);

            tableCell23.Append(tableCellProperties23);
            tableCell23.Append(paragraph28);

            TableCell tableCell24 = new TableCell();

            TableCellProperties tableCellProperties24 = new TableCellProperties();
            TableCellWidth tableCellWidth24 = new TableCellWidth() { Width = "1243", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders16 = new TableCellBorders();
            TopBorder topBorder17 = new TopBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder17 = new LeftBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder17 = new BottomBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder17 = new RightBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders16.Append(topBorder17);
            tableCellBorders16.Append(leftBorder17);
            tableCellBorders16.Append(bottomBorder17);
            tableCellBorders16.Append(rightBorder17);
            TableCellVerticalAlignment tableCellVerticalAlignment20 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties24.Append(tableCellWidth24);
            tableCellProperties24.Append(tableCellBorders16);
            tableCellProperties24.Append(tableCellVerticalAlignment20);

            Paragraph paragraph29 = new Paragraph() { RsidParagraphMarkRevision = "00885DA2", RsidParagraphAddition = "007B3C50", RsidParagraphProperties = "001F4E3B", RsidRunAdditionDefault = "007B3C50" };

            ParagraphProperties paragraphProperties29 = new ParagraphProperties();
            WidowControl widowControl29 = new WidowControl();
            AdjustRightIndent adjustRightIndent29 = new AdjustRightIndent() { Val = false };
            SnapToGrid snapToGrid29 = new SnapToGrid() { Val = false };
            Indentation indentation29 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification29 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties29 = new ParagraphMarkRunProperties();
            RunFonts runFonts115 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", EastAsia = "宋体", ComplexScript = "Times New Roman" };
            Color color91 = new Color() { Val = "FF0000" };
            Kern kern102 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize102 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript102 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties29.Append(runFonts115);
            paragraphMarkRunProperties29.Append(color91);
            paragraphMarkRunProperties29.Append(kern102);
            paragraphMarkRunProperties29.Append(fontSize102);
            paragraphMarkRunProperties29.Append(fontSizeComplexScript102);

            paragraphProperties29.Append(widowControl29);
            paragraphProperties29.Append(adjustRightIndent29);
            paragraphProperties29.Append(snapToGrid29);
            paragraphProperties29.Append(indentation29);
            paragraphProperties29.Append(justification29);
            paragraphProperties29.Append(paragraphMarkRunProperties29);

            Run run87 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties87 = new RunProperties();
            RunFonts runFonts116 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold94 = new Bold();
            Color color92 = new Color() { Val = "FF0000" };
            Kern kern103 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize103 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript103 = new FontSizeComplexScript() { Val = "24" };

            runProperties87.Append(runFonts116);
            runProperties87.Append(bold94);
            runProperties87.Append(color92);
            runProperties87.Append(kern103);
            runProperties87.Append(fontSize103);
            runProperties87.Append(fontSizeComplexScript103);
            FieldChar fieldChar7 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

            run87.Append(runProperties87);
            run87.Append(fieldChar7);

            Run run88 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties88 = new RunProperties();
            RunFonts runFonts117 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold95 = new Bold();
            Color color93 = new Color() { Val = "FF0000" };
            Kern kern104 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize104 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript104 = new FontSizeComplexScript() { Val = "24" };

            runProperties88.Append(runFonts117);
            runProperties88.Append(bold95);
            runProperties88.Append(color93);
            runProperties88.Append(kern104);
            runProperties88.Append(fontSize104);
            runProperties88.Append(fontSizeComplexScript104);
            FieldCode fieldCode55 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
            fieldCode55.Text = " LINK Excel.Sheet.12";

            run88.Append(runProperties88);
            run88.Append(fieldCode55);

            Run run89 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties89 = new RunProperties();
            RunFonts runFonts118 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold96 = new Bold();
            Color color94 = new Color() { Val = "FF0000" };
            Kern kern105 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize105 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript105 = new FontSizeComplexScript() { Val = "24" };

            runProperties89.Append(runFonts118);
            runProperties89.Append(bold96);
            runProperties89.Append(color94);
            runProperties89.Append(kern105);
            runProperties89.Append(fontSize105);
            runProperties89.Append(fontSizeComplexScript105);
            FieldCode fieldCode56 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
            fieldCode56.Text = " E:\\\\";

            run89.Append(runProperties89);
            run89.Append(fieldCode56);

            Run run90 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties90 = new RunProperties();
            RunFonts runFonts119 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold97 = new Bold();
            Color color95 = new Color() { Val = "FF0000" };
            Kern kern106 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize106 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript106 = new FontSizeComplexScript() { Val = "24" };

            runProperties90.Append(runFonts119);
            runProperties90.Append(bold97);
            runProperties90.Append(color95);
            runProperties90.Append(kern106);
            runProperties90.Append(fontSize106);
            runProperties90.Append(fontSizeComplexScript106);
            FieldCode fieldCode57 = new FieldCode();
            fieldCode57.Text = "【";

            run90.Append(runProperties90);
            run90.Append(fieldCode57);

            Run run91 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties91 = new RunProperties();
            RunFonts runFonts120 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold98 = new Bold();
            Color color96 = new Color() { Val = "FF0000" };
            Kern kern107 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize107 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript107 = new FontSizeComplexScript() { Val = "24" };

            runProperties91.Append(runFonts120);
            runProperties91.Append(bold98);
            runProperties91.Append(color96);
            runProperties91.Append(kern107);
            runProperties91.Append(fontSize107);
            runProperties91.Append(fontSizeComplexScript107);
            FieldCode fieldCode58 = new FieldCode();
            fieldCode58.Text = "5";

            run91.Append(runProperties91);
            run91.Append(fieldCode58);

            Run run92 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties92 = new RunProperties();
            RunFonts runFonts121 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold99 = new Bold();
            Color color97 = new Color() { Val = "FF0000" };
            Kern kern108 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize108 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript108 = new FontSizeComplexScript() { Val = "24" };

            runProperties92.Append(runFonts121);
            runProperties92.Append(bold99);
            runProperties92.Append(color97);
            runProperties92.Append(kern108);
            runProperties92.Append(fontSize108);
            runProperties92.Append(fontSizeComplexScript108);
            FieldCode fieldCode59 = new FieldCode();
            fieldCode59.Text = "】产品";

            run92.Append(runProperties92);
            run92.Append(fieldCode59);

            Run run93 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties93 = new RunProperties();
            RunFonts runFonts122 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold100 = new Bold();
            Color color98 = new Color() { Val = "FF0000" };
            Kern kern109 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize109 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript109 = new FontSizeComplexScript() { Val = "24" };

            runProperties93.Append(runFonts122);
            runProperties93.Append(bold100);
            runProperties93.Append(color98);
            runProperties93.Append(kern109);
            runProperties93.Append(fontSize109);
            runProperties93.Append(fontSizeComplexScript109);
            FieldCode fieldCode60 = new FieldCode();
            fieldCode60.Text = "\\\\";

            run93.Append(runProperties93);
            run93.Append(fieldCode60);

            Run run94 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties94 = new RunProperties();
            RunFonts runFonts123 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold101 = new Bold();
            Color color99 = new Color() { Val = "FF0000" };
            Kern kern110 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize110 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript110 = new FontSizeComplexScript() { Val = "24" };

            runProperties94.Append(runFonts123);
            runProperties94.Append(bold101);
            runProperties94.Append(color99);
            runProperties94.Append(kern110);
            runProperties94.Append(fontSize110);
            runProperties94.Append(fontSizeComplexScript110);
            FieldCode fieldCode61 = new FieldCode();
            fieldCode61.Text = "产品";

            run94.Append(runProperties94);
            run94.Append(fieldCode61);

            Run run95 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties95 = new RunProperties();
            RunFonts runFonts124 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold102 = new Bold();
            Color color100 = new Color() { Val = "FF0000" };
            Kern kern111 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize111 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript111 = new FontSizeComplexScript() { Val = "24" };

            runProperties95.Append(runFonts124);
            runProperties95.Append(bold102);
            runProperties95.Append(color100);
            runProperties95.Append(kern111);
            runProperties95.Append(fontSize111);
            runProperties95.Append(fontSizeComplexScript111);
            FieldCode fieldCode62 = new FieldCode();
            fieldCode62.Text = "-";

            run95.Append(runProperties95);
            run95.Append(fieldCode62);

            Run run96 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties96 = new RunProperties();
            RunFonts runFonts125 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold103 = new Bold();
            Color color101 = new Color() { Val = "FF0000" };
            Kern kern112 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize112 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript112 = new FontSizeComplexScript() { Val = "24" };

            runProperties96.Append(runFonts125);
            runProperties96.Append(bold103);
            runProperties96.Append(color101);
            runProperties96.Append(kern112);
            runProperties96.Append(fontSize112);
            runProperties96.Append(fontSizeComplexScript112);
            FieldCode fieldCode63 = new FieldCode();
            fieldCode63.Text = "测评系统";

            run96.Append(runProperties96);
            run96.Append(fieldCode63);

            Run run97 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties97 = new RunProperties();
            RunFonts runFonts126 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold104 = new Bold();
            Color color102 = new Color() { Val = "FF0000" };
            Kern kern113 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize113 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript113 = new FontSizeComplexScript() { Val = "24" };

            runProperties97.Append(runFonts126);
            runProperties97.Append(bold104);
            runProperties97.Append(color102);
            runProperties97.Append(kern113);
            runProperties97.Append(fontSize113);
            runProperties97.Append(fontSizeComplexScript113);
            FieldCode fieldCode64 = new FieldCode();
            fieldCode64.Text = "\\\\";

            run97.Append(runProperties97);
            run97.Append(fieldCode64);

            Run run98 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties98 = new RunProperties();
            RunFonts runFonts127 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold105 = new Bold();
            Color color103 = new Color() { Val = "FF0000" };
            Kern kern114 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize114 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript114 = new FontSizeComplexScript() { Val = "24" };

            runProperties98.Append(runFonts127);
            runProperties98.Append(bold105);
            runProperties98.Append(color103);
            runProperties98.Append(kern114);
            runProperties98.Append(fontSize114);
            runProperties98.Append(fontSizeComplexScript114);
            FieldCode fieldCode65 = new FieldCode();
            fieldCode65.Text = "学业心理素质评估系统";

            run98.Append(runProperties98);
            run98.Append(fieldCode65);

            Run run99 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties99 = new RunProperties();
            RunFonts runFonts128 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold106 = new Bold();
            Color color104 = new Color() { Val = "FF0000" };
            Kern kern115 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize115 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript115 = new FontSizeComplexScript() { Val = "24" };

            runProperties99.Append(runFonts128);
            runProperties99.Append(bold106);
            runProperties99.Append(color104);
            runProperties99.Append(kern115);
            runProperties99.Append(fontSize115);
            runProperties99.Append(fontSizeComplexScript115);
            FieldCode fieldCode66 = new FieldCode();
            fieldCode66.Text = "\\\\";

            run99.Append(runProperties99);
            run99.Append(fieldCode66);

            Run run100 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties100 = new RunProperties();
            RunFonts runFonts129 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold107 = new Bold();
            Color color105 = new Color() { Val = "FF0000" };
            Kern kern116 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize116 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript116 = new FontSizeComplexScript() { Val = "24" };

            runProperties100.Append(runFonts129);
            runProperties100.Append(bold107);
            runProperties100.Append(color105);
            runProperties100.Append(kern116);
            runProperties100.Append(fontSize116);
            runProperties100.Append(fontSizeComplexScript116);
            FieldCode fieldCode67 = new FieldCode();
            fieldCode67.Text = "学业支持系统";

            run100.Append(runProperties100);
            run100.Append(fieldCode67);

            Run run101 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties101 = new RunProperties();
            RunFonts runFonts130 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold108 = new Bold();
            Color color106 = new Color() { Val = "FF0000" };
            Kern kern117 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize117 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript117 = new FontSizeComplexScript() { Val = "24" };

            runProperties101.Append(runFonts130);
            runProperties101.Append(bold108);
            runProperties101.Append(color106);
            runProperties101.Append(kern117);
            runProperties101.Append(fontSize117);
            runProperties101.Append(fontSizeComplexScript117);
            FieldCode fieldCode68 = new FieldCode();
            fieldCode68.Text = "\\\\";

            run101.Append(runProperties101);
            run101.Append(fieldCode68);

            Run run102 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties102 = new RunProperties();
            RunFonts runFonts131 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold109 = new Bold();
            Color color107 = new Color() { Val = "FF0000" };
            Kern kern118 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize118 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript118 = new FontSizeComplexScript() { Val = "24" };

            runProperties102.Append(runFonts131);
            runProperties102.Append(bold109);
            runProperties102.Append(color107);
            runProperties102.Append(kern118);
            runProperties102.Append(fontSize118);
            runProperties102.Append(fontSizeComplexScript118);
            FieldCode fieldCode69 = new FieldCode();
            fieldCode69.Text = "测评系统分析模板";

            run102.Append(runProperties102);
            run102.Append(fieldCode69);

            Run run103 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties103 = new RunProperties();
            RunFonts runFonts132 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold110 = new Bold();
            Color color108 = new Color() { Val = "FF0000" };
            Kern kern119 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize119 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript119 = new FontSizeComplexScript() { Val = "24" };

            runProperties103.Append(runFonts132);
            runProperties103.Append(bold110);
            runProperties103.Append(color108);
            runProperties103.Append(kern119);
            runProperties103.Append(fontSize119);
            runProperties103.Append(fontSizeComplexScript119);
            FieldCode fieldCode70 = new FieldCode();
            fieldCode70.Text = "-";

            run103.Append(runProperties103);
            run103.Append(fieldCode70);

            Run run104 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties104 = new RunProperties();
            RunFonts runFonts133 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold111 = new Bold();
            Color color109 = new Color() { Val = "FF0000" };
            Kern kern120 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize120 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript120 = new FontSizeComplexScript() { Val = "24" };

            runProperties104.Append(runFonts133);
            runProperties104.Append(bold111);
            runProperties104.Append(color109);
            runProperties104.Append(kern120);
            runProperties104.Append(fontSize120);
            runProperties104.Append(fontSizeComplexScript120);
            FieldCode fieldCode71 = new FieldCode();
            fieldCode71.Text = "学业支持系统";

            run104.Append(runProperties104);
            run104.Append(fieldCode71);

            Run run105 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties105 = new RunProperties();
            RunFonts runFonts134 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold112 = new Bold();
            Color color110 = new Color() { Val = "FF0000" };
            Kern kern121 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize121 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript121 = new FontSizeComplexScript() { Val = "24" };

            runProperties105.Append(runFonts134);
            runProperties105.Append(bold112);
            runProperties105.Append(color110);
            runProperties105.Append(kern121);
            runProperties105.Append(fontSize121);
            runProperties105.Append(fontSizeComplexScript121);
            FieldCode fieldCode72 = new FieldCode();
            fieldCode72.Text = "\\\\";

            run105.Append(runProperties105);
            run105.Append(fieldCode72);

            Run run106 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties106 = new RunProperties();
            RunFonts runFonts135 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold113 = new Bold();
            Color color111 = new Color() { Val = "FF0000" };
            Kern kern122 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize122 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript122 = new FontSizeComplexScript() { Val = "24" };

            runProperties106.Append(runFonts135);
            runProperties106.Append(bold113);
            runProperties106.Append(color111);
            runProperties106.Append(kern122);
            runProperties106.Append(fontSize122);
            runProperties106.Append(fontSizeComplexScript122);
            FieldCode fieldCode73 = new FieldCode();
            fieldCode73.Text = "少量出报告模板";

            run106.Append(runProperties106);
            run106.Append(fieldCode73);

            Run run107 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties107 = new RunProperties();
            RunFonts runFonts136 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold114 = new Bold();
            Color color112 = new Color() { Val = "FF0000" };
            Kern kern123 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize123 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript123 = new FontSizeComplexScript() { Val = "24" };

            runProperties107.Append(runFonts136);
            runProperties107.Append(bold114);
            runProperties107.Append(color112);
            runProperties107.Append(kern123);
            runProperties107.Append(fontSize123);
            runProperties107.Append(fontSizeComplexScript123);
            FieldCode fieldCode74 = new FieldCode();
            fieldCode74.Text = "\\\\";

            run107.Append(runProperties107);
            run107.Append(fieldCode74);

            Run run108 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties108 = new RunProperties();
            RunFonts runFonts137 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold115 = new Bold();
            Color color113 = new Color() { Val = "FF0000" };
            Kern kern124 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize124 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript124 = new FontSizeComplexScript() { Val = "24" };

            runProperties108.Append(runFonts137);
            runProperties108.Append(bold115);
            runProperties108.Append(color113);
            runProperties108.Append(kern124);
            runProperties108.Append(fontSize124);
            runProperties108.Append(fontSizeComplexScript124);
            FieldCode fieldCode75 = new FieldCode();
            fieldCode75.Text = "学业支持测评系统";

            run108.Append(runProperties108);
            run108.Append(fieldCode75);

            Run run109 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties109 = new RunProperties();
            RunFonts runFonts138 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold116 = new Bold();
            Color color114 = new Color() { Val = "FF0000" };
            Kern kern125 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize125 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript125 = new FontSizeComplexScript() { Val = "24" };

            runProperties109.Append(runFonts138);
            runProperties109.Append(bold116);
            runProperties109.Append(color114);
            runProperties109.Append(kern125);
            runProperties109.Append(fontSize125);
            runProperties109.Append(fontSizeComplexScript125);
            FieldCode fieldCode76 = new FieldCode();
            fieldCode76.Text = "excel";

            run109.Append(runProperties109);
            run109.Append(fieldCode76);

            Run run110 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties110 = new RunProperties();
            RunFonts runFonts139 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold117 = new Bold();
            Color color115 = new Color() { Val = "FF0000" };
            Kern kern126 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize126 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript126 = new FontSizeComplexScript() { Val = "24" };

            runProperties110.Append(runFonts139);
            runProperties110.Append(bold117);
            runProperties110.Append(color115);
            runProperties110.Append(kern126);
            runProperties110.Append(fontSize126);
            runProperties110.Append(fontSizeComplexScript126);
            FieldCode fieldCode77 = new FieldCode();
            fieldCode77.Text = "模板";

            run110.Append(runProperties110);
            run110.Append(fieldCode77);

            Run run111 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties111 = new RunProperties();
            RunFonts runFonts140 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold118 = new Bold();
            Color color116 = new Color() { Val = "FF0000" };
            Kern kern127 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize127 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript127 = new FontSizeComplexScript() { Val = "24" };

            runProperties111.Append(runFonts140);
            runProperties111.Append(bold118);
            runProperties111.Append(color116);
            runProperties111.Append(kern127);
            runProperties111.Append(fontSize127);
            runProperties111.Append(fontSizeComplexScript127);
            FieldCode fieldCode78 = new FieldCode();
            fieldCode78.Text = "-20160908-";

            run111.Append(runProperties111);
            run111.Append(fieldCode78);

            Run run112 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties112 = new RunProperties();
            RunFonts runFonts141 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold119 = new Bold();
            Color color117 = new Color() { Val = "FF0000" };
            Kern kern128 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize128 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript128 = new FontSizeComplexScript() { Val = "24" };

            runProperties112.Append(runFonts141);
            runProperties112.Append(bold119);
            runProperties112.Append(color117);
            runProperties112.Append(kern128);
            runProperties112.Append(fontSize128);
            runProperties112.Append(fontSizeComplexScript128);
            FieldCode fieldCode79 = new FieldCode();
            fieldCode79.Text = "少量出报告";

            run112.Append(runProperties112);
            run112.Append(fieldCode79);

            Run run113 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties113 = new RunProperties();
            RunFonts runFonts142 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold120 = new Bold();
            Color color118 = new Color() { Val = "FF0000" };
            Kern kern129 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize129 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript129 = new FontSizeComplexScript() { Val = "24" };

            runProperties113.Append(runFonts142);
            runProperties113.Append(bold120);
            runProperties113.Append(color118);
            runProperties113.Append(kern129);
            runProperties113.Append(fontSize129);
            runProperties113.Append(fontSizeComplexScript129);
            FieldCode fieldCode80 = new FieldCode();
            fieldCode80.Text = ".xlsx";

            run113.Append(runProperties113);
            run113.Append(fieldCode80);

            Run run114 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties114 = new RunProperties();
            RunFonts runFonts143 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold121 = new Bold();
            Color color119 = new Color() { Val = "FF0000" };
            Kern kern130 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize130 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript130 = new FontSizeComplexScript() { Val = "24" };

            runProperties114.Append(runFonts143);
            runProperties114.Append(bold121);
            runProperties114.Append(color119);
            runProperties114.Append(kern130);
            runProperties114.Append(fontSize130);
            runProperties114.Append(fontSizeComplexScript130);
            FieldCode fieldCode81 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
            fieldCode81.Text = " Sheet1!R2C319 \\a \\f 5 \\h  \\* MERGEFORMAT ";

            run114.Append(runProperties114);
            run114.Append(fieldCode81);

            Run run115 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties115 = new RunProperties();
            RunFonts runFonts144 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold122 = new Bold();
            Color color120 = new Color() { Val = "FF0000" };
            Kern kern131 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize131 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript131 = new FontSizeComplexScript() { Val = "24" };

            runProperties115.Append(runFonts144);
            runProperties115.Append(bold122);
            runProperties115.Append(color120);
            runProperties115.Append(kern131);
            runProperties115.Append(fontSize131);
            runProperties115.Append(fontSizeComplexScript131);
            FieldChar fieldChar8 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

            run115.Append(runProperties115);
            run115.Append(fieldChar8);

            paragraph29.Append(paragraphProperties29);
            paragraph29.Append(run87);
            paragraph29.Append(run88);
            paragraph29.Append(run89);
            paragraph29.Append(run90);
            paragraph29.Append(run91);
            paragraph29.Append(run92);
            paragraph29.Append(run93);
            paragraph29.Append(run94);
            paragraph29.Append(run95);
            paragraph29.Append(run96);
            paragraph29.Append(run97);
            paragraph29.Append(run98);
            paragraph29.Append(run99);
            paragraph29.Append(run100);
            paragraph29.Append(run101);
            paragraph29.Append(run102);
            paragraph29.Append(run103);
            paragraph29.Append(run104);
            paragraph29.Append(run105);
            paragraph29.Append(run106);
            paragraph29.Append(run107);
            paragraph29.Append(run108);
            paragraph29.Append(run109);
            paragraph29.Append(run110);
            paragraph29.Append(run111);
            paragraph29.Append(run112);
            paragraph29.Append(run113);
            paragraph29.Append(run114);
            paragraph29.Append(run115);

            Paragraph paragraph30 = new Paragraph() { RsidParagraphMarkRevision = "00885DA2", RsidParagraphAddition = "007B3C50", RsidParagraphProperties = "001F4E3B", RsidRunAdditionDefault = "007B3C50" };

            ParagraphProperties paragraphProperties30 = new ParagraphProperties();
            WidowControl widowControl30 = new WidowControl();
            AdjustRightIndent adjustRightIndent30 = new AdjustRightIndent() { Val = false };
            SnapToGrid snapToGrid30 = new SnapToGrid() { Val = false };
            Indentation indentation30 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification30 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties30 = new ParagraphMarkRunProperties();
            RunFonts runFonts145 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Color color121 = new Color() { Val = "FF0000" };
            Kern kern132 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize132 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript132 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties30.Append(runFonts145);
            paragraphMarkRunProperties30.Append(color121);
            paragraphMarkRunProperties30.Append(kern132);
            paragraphMarkRunProperties30.Append(fontSize132);
            paragraphMarkRunProperties30.Append(fontSizeComplexScript132);

            paragraphProperties30.Append(widowControl30);
            paragraphProperties30.Append(adjustRightIndent30);
            paragraphProperties30.Append(snapToGrid30);
            paragraphProperties30.Append(indentation30);
            paragraphProperties30.Append(justification30);
            paragraphProperties30.Append(paragraphMarkRunProperties30);

            Run run116 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties116 = new RunProperties();
            RunFonts runFonts146 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Color color122 = new Color() { Val = "FF0000" };
            Kern kern133 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize133 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript133 = new FontSizeComplexScript() { Val = "24" };

            runProperties116.Append(runFonts146);
            runProperties116.Append(color122);
            runProperties116.Append(kern133);
            runProperties116.Append(fontSize133);
            runProperties116.Append(fontSizeComplexScript133);
            Text text27 = new Text();
            text27.Text = "■";

            run116.Append(runProperties116);
            run116.Append(text27);

            paragraph30.Append(paragraphProperties30);
            paragraph30.Append(run116);

            Paragraph paragraph31 = new Paragraph() { RsidParagraphMarkRevision = "00885DA2", RsidParagraphAddition = "007B3C50", RsidParagraphProperties = "001F4E3B", RsidRunAdditionDefault = "007B3C50" };

            ParagraphProperties paragraphProperties31 = new ParagraphProperties();
            WidowControl widowControl31 = new WidowControl();
            AdjustRightIndent adjustRightIndent31 = new AdjustRightIndent() { Val = false };
            SnapToGrid snapToGrid31 = new SnapToGrid() { Val = false };
            Indentation indentation31 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification31 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties31 = new ParagraphMarkRunProperties();
            RunFonts runFonts147 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold123 = new Bold();
            Color color123 = new Color() { Val = "FF0000" };
            Kern kern134 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize134 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript134 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties31.Append(runFonts147);
            paragraphMarkRunProperties31.Append(bold123);
            paragraphMarkRunProperties31.Append(color123);
            paragraphMarkRunProperties31.Append(kern134);
            paragraphMarkRunProperties31.Append(fontSize134);
            paragraphMarkRunProperties31.Append(fontSizeComplexScript134);

            paragraphProperties31.Append(widowControl31);
            paragraphProperties31.Append(adjustRightIndent31);
            paragraphProperties31.Append(snapToGrid31);
            paragraphProperties31.Append(indentation31);
            paragraphProperties31.Append(justification31);
            paragraphProperties31.Append(paragraphMarkRunProperties31);

            Run run117 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties117 = new RunProperties();
            RunFonts runFonts148 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold124 = new Bold();
            Color color124 = new Color() { Val = "FF0000" };
            Kern kern135 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize135 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript135 = new FontSizeComplexScript() { Val = "24" };

            runProperties117.Append(runFonts148);
            runProperties117.Append(bold124);
            runProperties117.Append(color124);
            runProperties117.Append(kern135);
            runProperties117.Append(fontSize135);
            runProperties117.Append(fontSizeComplexScript135);
            FieldChar fieldChar9 = new FieldChar() { FieldCharType = FieldCharValues.End };

            run117.Append(runProperties117);
            run117.Append(fieldChar9);

            paragraph31.Append(paragraphProperties31);
            paragraph31.Append(run117);

            tableCell24.Append(tableCellProperties24);
            tableCell24.Append(paragraph29);
            tableCell24.Append(paragraph30);
            tableCell24.Append(paragraph31);

            tableRow5.Append(tableRowProperties5);
            tableRow5.Append(tableCell20);
            tableRow5.Append(tableCell21);
            tableRow5.Append(tableCell22);
            tableRow5.Append(tableCell23);
            tableRow5.Append(tableCell24);

            TableRow tableRow6 = new TableRow() { RsidTableRowMarkRevision = "00885DA2", RsidTableRowAddition = "007B3C50", RsidTableRowProperties = "00522EB6" };

            TableRowProperties tableRowProperties6 = new TableRowProperties();
            TableRowHeight tableRowHeight6 = new TableRowHeight() { Val = (UInt32Value)367U };
            TableJustification tableJustification7 = new TableJustification() { Val = TableRowAlignmentValues.Center };

            tableRowProperties6.Append(tableRowHeight6);
            tableRowProperties6.Append(tableJustification7);

            TableCell tableCell25 = new TableCell();

            TableCellProperties tableCellProperties25 = new TableCellProperties();
            TableCellWidth tableCellWidth25 = new TableCellWidth() { Width = "1986", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge5 = new VerticalMerge();
            Shading shading17 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment21 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties25.Append(tableCellWidth25);
            tableCellProperties25.Append(verticalMerge5);
            tableCellProperties25.Append(shading17);
            tableCellProperties25.Append(tableCellVerticalAlignment21);

            Paragraph paragraph32 = new Paragraph() { RsidParagraphMarkRevision = "00885DA2", RsidParagraphAddition = "007B3C50", RsidParagraphProperties = "001F4E3B", RsidRunAdditionDefault = "007B3C50" };

            ParagraphProperties paragraphProperties32 = new ParagraphProperties();
            WidowControl widowControl32 = new WidowControl();
            AdjustRightIndent adjustRightIndent32 = new AdjustRightIndent() { Val = false };
            SnapToGrid snapToGrid32 = new SnapToGrid() { Val = false };
            Indentation indentation32 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification32 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties32 = new ParagraphMarkRunProperties();
            RunFonts runFonts149 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold125 = new Bold();
            Color color125 = new Color() { Val = "000000" };
            Kern kern136 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize136 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript136 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties32.Append(runFonts149);
            paragraphMarkRunProperties32.Append(bold125);
            paragraphMarkRunProperties32.Append(color125);
            paragraphMarkRunProperties32.Append(kern136);
            paragraphMarkRunProperties32.Append(fontSize136);
            paragraphMarkRunProperties32.Append(fontSizeComplexScript136);

            paragraphProperties32.Append(widowControl32);
            paragraphProperties32.Append(adjustRightIndent32);
            paragraphProperties32.Append(snapToGrid32);
            paragraphProperties32.Append(indentation32);
            paragraphProperties32.Append(justification32);
            paragraphProperties32.Append(paragraphMarkRunProperties32);

            paragraph32.Append(paragraphProperties32);

            tableCell25.Append(tableCellProperties25);
            tableCell25.Append(paragraph32);

            TableCell tableCell26 = new TableCell();

            TableCellProperties tableCellProperties26 = new TableCellProperties();
            TableCellWidth tableCellWidth26 = new TableCellWidth() { Width = "3186", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders17 = new TableCellBorders();
            TopBorder topBorder18 = new TopBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder18 = new LeftBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder18 = new BottomBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder18 = new RightBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders17.Append(topBorder18);
            tableCellBorders17.Append(leftBorder18);
            tableCellBorders17.Append(bottomBorder18);
            tableCellBorders17.Append(rightBorder18);
            Shading shading18 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment22 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties26.Append(tableCellWidth26);
            tableCellProperties26.Append(tableCellBorders17);
            tableCellProperties26.Append(shading18);
            tableCellProperties26.Append(tableCellVerticalAlignment22);

            Paragraph paragraph33 = new Paragraph() { RsidParagraphMarkRevision = "00522EB6", RsidParagraphAddition = "007B3C50", RsidParagraphProperties = "006F5D65", RsidRunAdditionDefault = "007B3C50" };

            ParagraphProperties paragraphProperties33 = new ParagraphProperties();
            WidowControl widowControl33 = new WidowControl();
            AdjustRightIndent adjustRightIndent33 = new AdjustRightIndent() { Val = false };
            SnapToGrid snapToGrid33 = new SnapToGrid() { Val = false };
            Indentation indentation33 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification33 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties33 = new ParagraphMarkRunProperties();
            RunFonts runFonts150 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold126 = new Bold();
            Color color126 = new Color() { Val = "000000" };
            Kern kern137 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize137 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript137 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties33.Append(runFonts150);
            paragraphMarkRunProperties33.Append(bold126);
            paragraphMarkRunProperties33.Append(color126);
            paragraphMarkRunProperties33.Append(kern137);
            paragraphMarkRunProperties33.Append(fontSize137);
            paragraphMarkRunProperties33.Append(fontSizeComplexScript137);

            paragraphProperties33.Append(widowControl33);
            paragraphProperties33.Append(adjustRightIndent33);
            paragraphProperties33.Append(snapToGrid33);
            paragraphProperties33.Append(indentation33);
            paragraphProperties33.Append(justification33);
            paragraphProperties33.Append(paragraphMarkRunProperties33);

            Run run118 = new Run() { RsidRunProperties = "00522EB6" };

            RunProperties runProperties118 = new RunProperties();
            RunFonts runFonts151 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold127 = new Bold();
            Color color127 = new Color() { Val = "000000" };
            Kern kern138 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize138 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript138 = new FontSizeComplexScript() { Val = "24" };

            runProperties118.Append(runFonts151);
            runProperties118.Append(bold127);
            runProperties118.Append(color127);
            runProperties118.Append(kern138);
            runProperties118.Append(fontSize138);
            runProperties118.Append(fontSizeComplexScript138);
            Text text28 = new Text();
            text28.Text = "依恋";

            run118.Append(runProperties118);
            run118.Append(text28);

            Run run119 = new Run() { RsidRunProperties = "00522EB6" };

            RunProperties runProperties119 = new RunProperties();
            RunFonts runFonts152 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold128 = new Bold();
            Color color128 = new Color() { Val = "000000" };
            Kern kern139 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize139 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript139 = new FontSizeComplexScript() { Val = "24" };

            runProperties119.Append(runFonts152);
            runProperties119.Append(bold128);
            runProperties119.Append(color128);
            runProperties119.Append(kern139);
            runProperties119.Append(fontSize139);
            runProperties119.Append(fontSizeComplexScript139);
            Text text29 = new Text();
            text29.Text = "性";

            run119.Append(runProperties119);
            run119.Append(text29);

            paragraph33.Append(paragraphProperties33);
            paragraph33.Append(run118);
            paragraph33.Append(run119);

            tableCell26.Append(tableCellProperties26);
            tableCell26.Append(paragraph33);

            TableCell tableCell27 = new TableCell();

            TableCellProperties tableCellProperties27 = new TableCellProperties();
            TableCellWidth tableCellWidth27 = new TableCellWidth() { Width = "935", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders18 = new TableCellBorders();
            TopBorder topBorder19 = new TopBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder19 = new LeftBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder19 = new BottomBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder19 = new RightBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders18.Append(topBorder19);
            tableCellBorders18.Append(leftBorder19);
            tableCellBorders18.Append(bottomBorder19);
            tableCellBorders18.Append(rightBorder19);

            tableCellProperties27.Append(tableCellWidth27);
            tableCellProperties27.Append(tableCellBorders18);

            Paragraph paragraph34 = new Paragraph() { RsidParagraphMarkRevision = "007B3C50", RsidParagraphAddition = "007B3C50", RsidParagraphProperties = "001F4E3B", RsidRunAdditionDefault = "007B3C50" };

            ParagraphProperties paragraphProperties34 = new ParagraphProperties();
            WidowControl widowControl34 = new WidowControl();
            AdjustRightIndent adjustRightIndent34 = new AdjustRightIndent() { Val = false };
            SnapToGrid snapToGrid34 = new SnapToGrid() { Val = false };
            Indentation indentation34 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification34 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties34 = new ParagraphMarkRunProperties();
            RunFonts runFonts153 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", EastAsia = "宋体", ComplexScript = "Times New Roman" };

            paragraphMarkRunProperties34.Append(runFonts153);

            paragraphProperties34.Append(widowControl34);
            paragraphProperties34.Append(adjustRightIndent34);
            paragraphProperties34.Append(snapToGrid34);
            paragraphProperties34.Append(indentation34);
            paragraphProperties34.Append(justification34);
            paragraphProperties34.Append(paragraphMarkRunProperties34);

            paragraph34.Append(paragraphProperties34);

            tableCell27.Append(tableCellProperties27);
            tableCell27.Append(paragraph34);

            TableCell tableCell28 = new TableCell();

            TableCellProperties tableCellProperties28 = new TableCellProperties();
            TableCellWidth tableCellWidth28 = new TableCellWidth() { Width = "935", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders19 = new TableCellBorders();
            TopBorder topBorder20 = new TopBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder20 = new LeftBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder20 = new BottomBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder20 = new RightBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders19.Append(topBorder20);
            tableCellBorders19.Append(leftBorder20);
            tableCellBorders19.Append(bottomBorder20);
            tableCellBorders19.Append(rightBorder20);
            Shading shading19 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment23 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties28.Append(tableCellWidth28);
            tableCellProperties28.Append(tableCellBorders19);
            tableCellProperties28.Append(shading19);
            tableCellProperties28.Append(tableCellVerticalAlignment23);

            Paragraph paragraph35 = new Paragraph() { RsidParagraphMarkRevision = "00522EB6", RsidParagraphAddition = "007B3C50", RsidParagraphProperties = "001F4E3B", RsidRunAdditionDefault = "007B3C50" };

            ParagraphProperties paragraphProperties35 = new ParagraphProperties();
            WidowControl widowControl35 = new WidowControl();
            AdjustRightIndent adjustRightIndent35 = new AdjustRightIndent() { Val = false };
            SnapToGrid snapToGrid35 = new SnapToGrid() { Val = false };
            Indentation indentation35 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification35 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties35 = new ParagraphMarkRunProperties();
            RunFonts runFonts154 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", EastAsia = "宋体", ComplexScript = "Times New Roman" };

            paragraphMarkRunProperties35.Append(runFonts154);

            paragraphProperties35.Append(widowControl35);
            paragraphProperties35.Append(adjustRightIndent35);
            paragraphProperties35.Append(snapToGrid35);
            paragraphProperties35.Append(indentation35);
            paragraphProperties35.Append(justification35);
            paragraphProperties35.Append(paragraphMarkRunProperties35);

            paragraph35.Append(paragraphProperties35);

            tableCell28.Append(tableCellProperties28);
            tableCell28.Append(paragraph35);

            TableCell tableCell29 = new TableCell();

            TableCellProperties tableCellProperties29 = new TableCellProperties();
            TableCellWidth tableCellWidth29 = new TableCellWidth() { Width = "1243", Type = TableWidthUnitValues.Dxa };

            TableCellBorders tableCellBorders20 = new TableCellBorders();
            TopBorder topBorder21 = new TopBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder21 = new LeftBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder21 = new BottomBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder21 = new RightBorder() { Val = BorderValues.Single, Color = "C00000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableCellBorders20.Append(topBorder21);
            tableCellBorders20.Append(leftBorder21);
            tableCellBorders20.Append(bottomBorder21);
            tableCellBorders20.Append(rightBorder21);
            TableCellVerticalAlignment tableCellVerticalAlignment24 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties29.Append(tableCellWidth29);
            tableCellProperties29.Append(tableCellBorders20);
            tableCellProperties29.Append(tableCellVerticalAlignment24);

            Paragraph paragraph36 = new Paragraph() { RsidParagraphMarkRevision = "00885DA2", RsidParagraphAddition = "007B3C50", RsidParagraphProperties = "001F4E3B", RsidRunAdditionDefault = "007B3C50" };

            ParagraphProperties paragraphProperties36 = new ParagraphProperties();
            WidowControl widowControl36 = new WidowControl();
            AdjustRightIndent adjustRightIndent36 = new AdjustRightIndent() { Val = false };
            SnapToGrid snapToGrid36 = new SnapToGrid() { Val = false };
            Indentation indentation36 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification36 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties36 = new ParagraphMarkRunProperties();
            RunFonts runFonts155 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", EastAsia = "宋体", ComplexScript = "Times New Roman" };
            Color color129 = new Color() { Val = "FF0000" };
            Kern kern140 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize140 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript140 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties36.Append(runFonts155);
            paragraphMarkRunProperties36.Append(color129);
            paragraphMarkRunProperties36.Append(kern140);
            paragraphMarkRunProperties36.Append(fontSize140);
            paragraphMarkRunProperties36.Append(fontSizeComplexScript140);

            paragraphProperties36.Append(widowControl36);
            paragraphProperties36.Append(adjustRightIndent36);
            paragraphProperties36.Append(snapToGrid36);
            paragraphProperties36.Append(indentation36);
            paragraphProperties36.Append(justification36);
            paragraphProperties36.Append(paragraphMarkRunProperties36);

            Run run120 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties120 = new RunProperties();
            RunFonts runFonts156 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold129 = new Bold();
            Color color130 = new Color() { Val = "FF0000" };
            Kern kern141 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize141 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript141 = new FontSizeComplexScript() { Val = "24" };

            runProperties120.Append(runFonts156);
            runProperties120.Append(bold129);
            runProperties120.Append(color130);
            runProperties120.Append(kern141);
            runProperties120.Append(fontSize141);
            runProperties120.Append(fontSizeComplexScript141);
            FieldChar fieldChar10 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

            run120.Append(runProperties120);
            run120.Append(fieldChar10);

            Run run121 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties121 = new RunProperties();
            RunFonts runFonts157 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold130 = new Bold();
            Color color131 = new Color() { Val = "FF0000" };
            Kern kern142 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize142 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript142 = new FontSizeComplexScript() { Val = "24" };

            runProperties121.Append(runFonts157);
            runProperties121.Append(bold130);
            runProperties121.Append(color131);
            runProperties121.Append(kern142);
            runProperties121.Append(fontSize142);
            runProperties121.Append(fontSizeComplexScript142);
            FieldCode fieldCode82 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
            fieldCode82.Text = " LINK Excel.Sheet.12";

            run121.Append(runProperties121);
            run121.Append(fieldCode82);

            Run run122 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties122 = new RunProperties();
            RunFonts runFonts158 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold131 = new Bold();
            Color color132 = new Color() { Val = "FF0000" };
            Kern kern143 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize143 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript143 = new FontSizeComplexScript() { Val = "24" };

            runProperties122.Append(runFonts158);
            runProperties122.Append(bold131);
            runProperties122.Append(color132);
            runProperties122.Append(kern143);
            runProperties122.Append(fontSize143);
            runProperties122.Append(fontSizeComplexScript143);
            FieldCode fieldCode83 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
            fieldCode83.Text = " E:\\\\";

            run122.Append(runProperties122);
            run122.Append(fieldCode83);

            Run run123 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties123 = new RunProperties();
            RunFonts runFonts159 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold132 = new Bold();
            Color color133 = new Color() { Val = "FF0000" };
            Kern kern144 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize144 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript144 = new FontSizeComplexScript() { Val = "24" };

            runProperties123.Append(runFonts159);
            runProperties123.Append(bold132);
            runProperties123.Append(color133);
            runProperties123.Append(kern144);
            runProperties123.Append(fontSize144);
            runProperties123.Append(fontSizeComplexScript144);
            FieldCode fieldCode84 = new FieldCode();
            fieldCode84.Text = "【";

            run123.Append(runProperties123);
            run123.Append(fieldCode84);

            Run run124 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties124 = new RunProperties();
            RunFonts runFonts160 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold133 = new Bold();
            Color color134 = new Color() { Val = "FF0000" };
            Kern kern145 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize145 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript145 = new FontSizeComplexScript() { Val = "24" };

            runProperties124.Append(runFonts160);
            runProperties124.Append(bold133);
            runProperties124.Append(color134);
            runProperties124.Append(kern145);
            runProperties124.Append(fontSize145);
            runProperties124.Append(fontSizeComplexScript145);
            FieldCode fieldCode85 = new FieldCode();
            fieldCode85.Text = "5";

            run124.Append(runProperties124);
            run124.Append(fieldCode85);

            Run run125 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties125 = new RunProperties();
            RunFonts runFonts161 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold134 = new Bold();
            Color color135 = new Color() { Val = "FF0000" };
            Kern kern146 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize146 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript146 = new FontSizeComplexScript() { Val = "24" };

            runProperties125.Append(runFonts161);
            runProperties125.Append(bold134);
            runProperties125.Append(color135);
            runProperties125.Append(kern146);
            runProperties125.Append(fontSize146);
            runProperties125.Append(fontSizeComplexScript146);
            FieldCode fieldCode86 = new FieldCode();
            fieldCode86.Text = "】产品";

            run125.Append(runProperties125);
            run125.Append(fieldCode86);

            Run run126 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties126 = new RunProperties();
            RunFonts runFonts162 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold135 = new Bold();
            Color color136 = new Color() { Val = "FF0000" };
            Kern kern147 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize147 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript147 = new FontSizeComplexScript() { Val = "24" };

            runProperties126.Append(runFonts162);
            runProperties126.Append(bold135);
            runProperties126.Append(color136);
            runProperties126.Append(kern147);
            runProperties126.Append(fontSize147);
            runProperties126.Append(fontSizeComplexScript147);
            FieldCode fieldCode87 = new FieldCode();
            fieldCode87.Text = "\\\\";

            run126.Append(runProperties126);
            run126.Append(fieldCode87);

            Run run127 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties127 = new RunProperties();
            RunFonts runFonts163 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold136 = new Bold();
            Color color137 = new Color() { Val = "FF0000" };
            Kern kern148 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize148 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript148 = new FontSizeComplexScript() { Val = "24" };

            runProperties127.Append(runFonts163);
            runProperties127.Append(bold136);
            runProperties127.Append(color137);
            runProperties127.Append(kern148);
            runProperties127.Append(fontSize148);
            runProperties127.Append(fontSizeComplexScript148);
            FieldCode fieldCode88 = new FieldCode();
            fieldCode88.Text = "产品";

            run127.Append(runProperties127);
            run127.Append(fieldCode88);

            Run run128 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties128 = new RunProperties();
            RunFonts runFonts164 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold137 = new Bold();
            Color color138 = new Color() { Val = "FF0000" };
            Kern kern149 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize149 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript149 = new FontSizeComplexScript() { Val = "24" };

            runProperties128.Append(runFonts164);
            runProperties128.Append(bold137);
            runProperties128.Append(color138);
            runProperties128.Append(kern149);
            runProperties128.Append(fontSize149);
            runProperties128.Append(fontSizeComplexScript149);
            FieldCode fieldCode89 = new FieldCode();
            fieldCode89.Text = "-";

            run128.Append(runProperties128);
            run128.Append(fieldCode89);

            Run run129 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties129 = new RunProperties();
            RunFonts runFonts165 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold138 = new Bold();
            Color color139 = new Color() { Val = "FF0000" };
            Kern kern150 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize150 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript150 = new FontSizeComplexScript() { Val = "24" };

            runProperties129.Append(runFonts165);
            runProperties129.Append(bold138);
            runProperties129.Append(color139);
            runProperties129.Append(kern150);
            runProperties129.Append(fontSize150);
            runProperties129.Append(fontSizeComplexScript150);
            FieldCode fieldCode90 = new FieldCode();
            fieldCode90.Text = "测评系统";

            run129.Append(runProperties129);
            run129.Append(fieldCode90);

            Run run130 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties130 = new RunProperties();
            RunFonts runFonts166 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold139 = new Bold();
            Color color140 = new Color() { Val = "FF0000" };
            Kern kern151 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize151 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript151 = new FontSizeComplexScript() { Val = "24" };

            runProperties130.Append(runFonts166);
            runProperties130.Append(bold139);
            runProperties130.Append(color140);
            runProperties130.Append(kern151);
            runProperties130.Append(fontSize151);
            runProperties130.Append(fontSizeComplexScript151);
            FieldCode fieldCode91 = new FieldCode();
            fieldCode91.Text = "\\\\";

            run130.Append(runProperties130);
            run130.Append(fieldCode91);

            Run run131 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties131 = new RunProperties();
            RunFonts runFonts167 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold140 = new Bold();
            Color color141 = new Color() { Val = "FF0000" };
            Kern kern152 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize152 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript152 = new FontSizeComplexScript() { Val = "24" };

            runProperties131.Append(runFonts167);
            runProperties131.Append(bold140);
            runProperties131.Append(color141);
            runProperties131.Append(kern152);
            runProperties131.Append(fontSize152);
            runProperties131.Append(fontSizeComplexScript152);
            FieldCode fieldCode92 = new FieldCode();
            fieldCode92.Text = "学业心理素质评估系统";

            run131.Append(runProperties131);
            run131.Append(fieldCode92);

            Run run132 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties132 = new RunProperties();
            RunFonts runFonts168 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold141 = new Bold();
            Color color142 = new Color() { Val = "FF0000" };
            Kern kern153 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize153 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript153 = new FontSizeComplexScript() { Val = "24" };

            runProperties132.Append(runFonts168);
            runProperties132.Append(bold141);
            runProperties132.Append(color142);
            runProperties132.Append(kern153);
            runProperties132.Append(fontSize153);
            runProperties132.Append(fontSizeComplexScript153);
            FieldCode fieldCode93 = new FieldCode();
            fieldCode93.Text = "\\\\";

            run132.Append(runProperties132);
            run132.Append(fieldCode93);

            Run run133 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties133 = new RunProperties();
            RunFonts runFonts169 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold142 = new Bold();
            Color color143 = new Color() { Val = "FF0000" };
            Kern kern154 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize154 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript154 = new FontSizeComplexScript() { Val = "24" };

            runProperties133.Append(runFonts169);
            runProperties133.Append(bold142);
            runProperties133.Append(color143);
            runProperties133.Append(kern154);
            runProperties133.Append(fontSize154);
            runProperties133.Append(fontSizeComplexScript154);
            FieldCode fieldCode94 = new FieldCode();
            fieldCode94.Text = "学业支持系统";

            run133.Append(runProperties133);
            run133.Append(fieldCode94);

            Run run134 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties134 = new RunProperties();
            RunFonts runFonts170 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold143 = new Bold();
            Color color144 = new Color() { Val = "FF0000" };
            Kern kern155 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize155 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript155 = new FontSizeComplexScript() { Val = "24" };

            runProperties134.Append(runFonts170);
            runProperties134.Append(bold143);
            runProperties134.Append(color144);
            runProperties134.Append(kern155);
            runProperties134.Append(fontSize155);
            runProperties134.Append(fontSizeComplexScript155);
            FieldCode fieldCode95 = new FieldCode();
            fieldCode95.Text = "\\\\";

            run134.Append(runProperties134);
            run134.Append(fieldCode95);

            Run run135 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties135 = new RunProperties();
            RunFonts runFonts171 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold144 = new Bold();
            Color color145 = new Color() { Val = "FF0000" };
            Kern kern156 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize156 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript156 = new FontSizeComplexScript() { Val = "24" };

            runProperties135.Append(runFonts171);
            runProperties135.Append(bold144);
            runProperties135.Append(color145);
            runProperties135.Append(kern156);
            runProperties135.Append(fontSize156);
            runProperties135.Append(fontSizeComplexScript156);
            FieldCode fieldCode96 = new FieldCode();
            fieldCode96.Text = "测评系统分析模板";

            run135.Append(runProperties135);
            run135.Append(fieldCode96);

            Run run136 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties136 = new RunProperties();
            RunFonts runFonts172 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold145 = new Bold();
            Color color146 = new Color() { Val = "FF0000" };
            Kern kern157 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize157 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript157 = new FontSizeComplexScript() { Val = "24" };

            runProperties136.Append(runFonts172);
            runProperties136.Append(bold145);
            runProperties136.Append(color146);
            runProperties136.Append(kern157);
            runProperties136.Append(fontSize157);
            runProperties136.Append(fontSizeComplexScript157);
            FieldCode fieldCode97 = new FieldCode();
            fieldCode97.Text = "-";

            run136.Append(runProperties136);
            run136.Append(fieldCode97);

            Run run137 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties137 = new RunProperties();
            RunFonts runFonts173 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold146 = new Bold();
            Color color147 = new Color() { Val = "FF0000" };
            Kern kern158 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize158 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript158 = new FontSizeComplexScript() { Val = "24" };

            runProperties137.Append(runFonts173);
            runProperties137.Append(bold146);
            runProperties137.Append(color147);
            runProperties137.Append(kern158);
            runProperties137.Append(fontSize158);
            runProperties137.Append(fontSizeComplexScript158);
            FieldCode fieldCode98 = new FieldCode();
            fieldCode98.Text = "学业支持系统";

            run137.Append(runProperties137);
            run137.Append(fieldCode98);

            Run run138 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties138 = new RunProperties();
            RunFonts runFonts174 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold147 = new Bold();
            Color color148 = new Color() { Val = "FF0000" };
            Kern kern159 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize159 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript159 = new FontSizeComplexScript() { Val = "24" };

            runProperties138.Append(runFonts174);
            runProperties138.Append(bold147);
            runProperties138.Append(color148);
            runProperties138.Append(kern159);
            runProperties138.Append(fontSize159);
            runProperties138.Append(fontSizeComplexScript159);
            FieldCode fieldCode99 = new FieldCode();
            fieldCode99.Text = "\\\\";

            run138.Append(runProperties138);
            run138.Append(fieldCode99);

            Run run139 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties139 = new RunProperties();
            RunFonts runFonts175 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold148 = new Bold();
            Color color149 = new Color() { Val = "FF0000" };
            Kern kern160 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize160 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript160 = new FontSizeComplexScript() { Val = "24" };

            runProperties139.Append(runFonts175);
            runProperties139.Append(bold148);
            runProperties139.Append(color149);
            runProperties139.Append(kern160);
            runProperties139.Append(fontSize160);
            runProperties139.Append(fontSizeComplexScript160);
            FieldCode fieldCode100 = new FieldCode();
            fieldCode100.Text = "少量出报告模板";

            run139.Append(runProperties139);
            run139.Append(fieldCode100);

            Run run140 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties140 = new RunProperties();
            RunFonts runFonts176 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold149 = new Bold();
            Color color150 = new Color() { Val = "FF0000" };
            Kern kern161 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize161 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript161 = new FontSizeComplexScript() { Val = "24" };

            runProperties140.Append(runFonts176);
            runProperties140.Append(bold149);
            runProperties140.Append(color150);
            runProperties140.Append(kern161);
            runProperties140.Append(fontSize161);
            runProperties140.Append(fontSizeComplexScript161);
            FieldCode fieldCode101 = new FieldCode();
            fieldCode101.Text = "\\\\";

            run140.Append(runProperties140);
            run140.Append(fieldCode101);

            Run run141 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties141 = new RunProperties();
            RunFonts runFonts177 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold150 = new Bold();
            Color color151 = new Color() { Val = "FF0000" };
            Kern kern162 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize162 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript162 = new FontSizeComplexScript() { Val = "24" };

            runProperties141.Append(runFonts177);
            runProperties141.Append(bold150);
            runProperties141.Append(color151);
            runProperties141.Append(kern162);
            runProperties141.Append(fontSize162);
            runProperties141.Append(fontSizeComplexScript162);
            FieldCode fieldCode102 = new FieldCode();
            fieldCode102.Text = "学业支持测评系统";

            run141.Append(runProperties141);
            run141.Append(fieldCode102);

            Run run142 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties142 = new RunProperties();
            RunFonts runFonts178 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold151 = new Bold();
            Color color152 = new Color() { Val = "FF0000" };
            Kern kern163 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize163 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript163 = new FontSizeComplexScript() { Val = "24" };

            runProperties142.Append(runFonts178);
            runProperties142.Append(bold151);
            runProperties142.Append(color152);
            runProperties142.Append(kern163);
            runProperties142.Append(fontSize163);
            runProperties142.Append(fontSizeComplexScript163);
            FieldCode fieldCode103 = new FieldCode();
            fieldCode103.Text = "excel";

            run142.Append(runProperties142);
            run142.Append(fieldCode103);

            Run run143 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties143 = new RunProperties();
            RunFonts runFonts179 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold152 = new Bold();
            Color color153 = new Color() { Val = "FF0000" };
            Kern kern164 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize164 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript164 = new FontSizeComplexScript() { Val = "24" };

            runProperties143.Append(runFonts179);
            runProperties143.Append(bold152);
            runProperties143.Append(color153);
            runProperties143.Append(kern164);
            runProperties143.Append(fontSize164);
            runProperties143.Append(fontSizeComplexScript164);
            FieldCode fieldCode104 = new FieldCode();
            fieldCode104.Text = "模板";

            run143.Append(runProperties143);
            run143.Append(fieldCode104);

            Run run144 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties144 = new RunProperties();
            RunFonts runFonts180 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold153 = new Bold();
            Color color154 = new Color() { Val = "FF0000" };
            Kern kern165 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize165 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript165 = new FontSizeComplexScript() { Val = "24" };

            runProperties144.Append(runFonts180);
            runProperties144.Append(bold153);
            runProperties144.Append(color154);
            runProperties144.Append(kern165);
            runProperties144.Append(fontSize165);
            runProperties144.Append(fontSizeComplexScript165);
            FieldCode fieldCode105 = new FieldCode();
            fieldCode105.Text = "-20160908-";

            run144.Append(runProperties144);
            run144.Append(fieldCode105);

            Run run145 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties145 = new RunProperties();
            RunFonts runFonts181 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold154 = new Bold();
            Color color155 = new Color() { Val = "FF0000" };
            Kern kern166 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize166 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript166 = new FontSizeComplexScript() { Val = "24" };

            runProperties145.Append(runFonts181);
            runProperties145.Append(bold154);
            runProperties145.Append(color155);
            runProperties145.Append(kern166);
            runProperties145.Append(fontSize166);
            runProperties145.Append(fontSizeComplexScript166);
            FieldCode fieldCode106 = new FieldCode();
            fieldCode106.Text = "少量出报告";

            run145.Append(runProperties145);
            run145.Append(fieldCode106);

            Run run146 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties146 = new RunProperties();
            RunFonts runFonts182 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold155 = new Bold();
            Color color156 = new Color() { Val = "FF0000" };
            Kern kern167 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize167 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript167 = new FontSizeComplexScript() { Val = "24" };

            runProperties146.Append(runFonts182);
            runProperties146.Append(bold155);
            runProperties146.Append(color156);
            runProperties146.Append(kern167);
            runProperties146.Append(fontSize167);
            runProperties146.Append(fontSizeComplexScript167);
            FieldCode fieldCode107 = new FieldCode();
            fieldCode107.Text = ".xlsx";

            run146.Append(runProperties146);
            run146.Append(fieldCode107);

            Run run147 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties147 = new RunProperties();
            RunFonts runFonts183 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold156 = new Bold();
            Color color157 = new Color() { Val = "FF0000" };
            Kern kern168 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize168 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript168 = new FontSizeComplexScript() { Val = "24" };

            runProperties147.Append(runFonts183);
            runProperties147.Append(bold156);
            runProperties147.Append(color157);
            runProperties147.Append(kern168);
            runProperties147.Append(fontSize168);
            runProperties147.Append(fontSizeComplexScript168);
            FieldCode fieldCode108 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
            fieldCode108.Text = " Sheet1!R2C327 \\a \\f 5 \\h  \\* MERGEFORMAT ";

            run147.Append(runProperties147);
            run147.Append(fieldCode108);

            Run run148 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties148 = new RunProperties();
            RunFonts runFonts184 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold157 = new Bold();
            Color color158 = new Color() { Val = "FF0000" };
            Kern kern169 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize169 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript169 = new FontSizeComplexScript() { Val = "24" };

            runProperties148.Append(runFonts184);
            runProperties148.Append(bold157);
            runProperties148.Append(color158);
            runProperties148.Append(kern169);
            runProperties148.Append(fontSize169);
            runProperties148.Append(fontSizeComplexScript169);
            FieldChar fieldChar11 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

            run148.Append(runProperties148);
            run148.Append(fieldChar11);

            paragraph36.Append(paragraphProperties36);
            paragraph36.Append(run120);
            paragraph36.Append(run121);
            paragraph36.Append(run122);
            paragraph36.Append(run123);
            paragraph36.Append(run124);
            paragraph36.Append(run125);
            paragraph36.Append(run126);
            paragraph36.Append(run127);
            paragraph36.Append(run128);
            paragraph36.Append(run129);
            paragraph36.Append(run130);
            paragraph36.Append(run131);
            paragraph36.Append(run132);
            paragraph36.Append(run133);
            paragraph36.Append(run134);
            paragraph36.Append(run135);
            paragraph36.Append(run136);
            paragraph36.Append(run137);
            paragraph36.Append(run138);
            paragraph36.Append(run139);
            paragraph36.Append(run140);
            paragraph36.Append(run141);
            paragraph36.Append(run142);
            paragraph36.Append(run143);
            paragraph36.Append(run144);
            paragraph36.Append(run145);
            paragraph36.Append(run146);
            paragraph36.Append(run147);
            paragraph36.Append(run148);

            Paragraph paragraph37 = new Paragraph() { RsidParagraphMarkRevision = "00885DA2", RsidParagraphAddition = "007B3C50", RsidParagraphProperties = "001F4E3B", RsidRunAdditionDefault = "007B3C50" };

            ParagraphProperties paragraphProperties37 = new ParagraphProperties();
            WidowControl widowControl37 = new WidowControl();
            AdjustRightIndent adjustRightIndent37 = new AdjustRightIndent() { Val = false };
            SnapToGrid snapToGrid37 = new SnapToGrid() { Val = false };
            Indentation indentation37 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification37 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties37 = new ParagraphMarkRunProperties();
            RunFonts runFonts185 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Color color159 = new Color() { Val = "FF0000" };
            Kern kern170 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize170 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript170 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties37.Append(runFonts185);
            paragraphMarkRunProperties37.Append(color159);
            paragraphMarkRunProperties37.Append(kern170);
            paragraphMarkRunProperties37.Append(fontSize170);
            paragraphMarkRunProperties37.Append(fontSizeComplexScript170);

            paragraphProperties37.Append(widowControl37);
            paragraphProperties37.Append(adjustRightIndent37);
            paragraphProperties37.Append(snapToGrid37);
            paragraphProperties37.Append(indentation37);
            paragraphProperties37.Append(justification37);
            paragraphProperties37.Append(paragraphMarkRunProperties37);

            Run run149 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties149 = new RunProperties();
            RunFonts runFonts186 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Color color160 = new Color() { Val = "FF0000" };
            Kern kern171 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize171 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript171 = new FontSizeComplexScript() { Val = "24" };

            runProperties149.Append(runFonts186);
            runProperties149.Append(color160);
            runProperties149.Append(kern171);
            runProperties149.Append(fontSize171);
            runProperties149.Append(fontSizeComplexScript171);
            Text text30 = new Text();
            text30.Text = "■";

            run149.Append(runProperties149);
            run149.Append(text30);

            paragraph37.Append(paragraphProperties37);
            paragraph37.Append(run149);

            Paragraph paragraph38 = new Paragraph() { RsidParagraphMarkRevision = "00885DA2", RsidParagraphAddition = "007B3C50", RsidParagraphProperties = "001F4E3B", RsidRunAdditionDefault = "007B3C50" };

            ParagraphProperties paragraphProperties38 = new ParagraphProperties();
            WidowControl widowControl38 = new WidowControl();
            AdjustRightIndent adjustRightIndent38 = new AdjustRightIndent() { Val = false };
            SnapToGrid snapToGrid38 = new SnapToGrid() { Val = false };
            Indentation indentation38 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification38 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties38 = new ParagraphMarkRunProperties();
            RunFonts runFonts187 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold158 = new Bold();
            Color color161 = new Color() { Val = "FF0000" };
            Kern kern172 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize172 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript172 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties38.Append(runFonts187);
            paragraphMarkRunProperties38.Append(bold158);
            paragraphMarkRunProperties38.Append(color161);
            paragraphMarkRunProperties38.Append(kern172);
            paragraphMarkRunProperties38.Append(fontSize172);
            paragraphMarkRunProperties38.Append(fontSizeComplexScript172);

            paragraphProperties38.Append(widowControl38);
            paragraphProperties38.Append(adjustRightIndent38);
            paragraphProperties38.Append(snapToGrid38);
            paragraphProperties38.Append(indentation38);
            paragraphProperties38.Append(justification38);
            paragraphProperties38.Append(paragraphMarkRunProperties38);

            Run run150 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties150 = new RunProperties();
            RunFonts runFonts188 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", EastAsia = "宋体", ComplexScript = "Tahoma" };
            Bold bold159 = new Bold();
            Color color162 = new Color() { Val = "FF0000" };
            Kern kern173 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize173 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript173 = new FontSizeComplexScript() { Val = "24" };

            runProperties150.Append(runFonts188);
            runProperties150.Append(bold159);
            runProperties150.Append(color162);
            runProperties150.Append(kern173);
            runProperties150.Append(fontSize173);
            runProperties150.Append(fontSizeComplexScript173);
            FieldChar fieldChar12 = new FieldChar() { FieldCharType = FieldCharValues.End };

            run150.Append(runProperties150);
            run150.Append(fieldChar12);

            paragraph38.Append(paragraphProperties38);
            paragraph38.Append(run150);

            tableCell29.Append(tableCellProperties29);
            tableCell29.Append(paragraph36);
            tableCell29.Append(paragraph37);
            tableCell29.Append(paragraph38);

            tableRow6.Append(tableRowProperties6);
            tableRow6.Append(tableCell25);
            tableRow6.Append(tableCell26);
            tableRow6.Append(tableCell27);
            tableRow6.Append(tableCell28);
            tableRow6.Append(tableCell29);*/
#endregion

           
            return table1;
        }


    }

}

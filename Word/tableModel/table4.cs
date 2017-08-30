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
        public Table GenerateTable4(Dictionary<string, Dictionary<string, ValueTuple<string, string, eInfluence>>> content)
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

            Paragraph paragraph1 = new Paragraph() { RsidParagraphMarkRevision = "007642DC", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "001D4055" };

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
            BookmarkStart bookmarkStart1 = new BookmarkStart() { Name = "_GoBack", Id = "0" };
            BookmarkEnd bookmarkEnd1 = new BookmarkEnd() { Id = "0" };

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
            text1.Text = "学业归因量表";//TODO量表名称

            run1.Append(runProperties1);
            run1.Append(text1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(bookmarkStart1);
            paragraph1.Append(bookmarkEnd1);
            paragraph1.Append(run1);

            tableCell1.Append(tableCellProperties1);
            tableCell1.Append(paragraph1);

            tableRow1.Append(tableRowProperties1);
            tableRow1.Append(tableCell1);

            TableRow tableRow2 = new TableRow() { RsidTableRowMarkRevision = "007642DC", RsidTableRowAddition = "001D4055", RsidTableRowProperties = "006A0186" };

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

            Paragraph paragraph2 = new Paragraph() { RsidParagraphMarkRevision = "007642DC", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "001D4055" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
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
            TableCellWidth tableCellWidth3 = new TableCellWidth() { Width = "682", Type = TableWidthUnitValues.Dxa };
            Shading shading3 = new Shading() { Val = ShadingPatternValues.Clear, Color = "000000", Fill = "D99795" };
            TableCellVerticalAlignment tableCellVerticalAlignment2 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties3.Append(tableCellWidth3);
            tableCellProperties3.Append(shading3);
            tableCellProperties3.Append(tableCellVerticalAlignment2);

            Paragraph paragraph3 = new Paragraph() { RsidParagraphMarkRevision = "007642DC", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "001D4055" };

            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
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
            TableCellWidth tableCellWidth4 = new TableCellWidth() { Width = "1082", Type = TableWidthUnitValues.Dxa };
            Shading shading4 = new Shading() { Val = ShadingPatternValues.Clear, Color = "000000", Fill = "D99795" };
            TableCellVerticalAlignment tableCellVerticalAlignment3 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties4.Append(tableCellWidth4);
            tableCellProperties4.Append(shading4);
            tableCellProperties4.Append(tableCellVerticalAlignment3);

            Paragraph paragraph4 = new Paragraph() { RsidParagraphMarkRevision = "007642DC", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "001D4055" };

            ParagraphProperties paragraphProperties4 = new ParagraphProperties();
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
            TableCellWidth tableCellWidth5 = new TableCellWidth() { Width = "882", Type = TableWidthUnitValues.Dxa };
            Shading shading5 = new Shading() { Val = ShadingPatternValues.Clear, Color = "000000", Fill = "D99795" };
            TableCellVerticalAlignment tableCellVerticalAlignment4 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties5.Append(tableCellWidth5);
            tableCellProperties5.Append(shading5);
            tableCellProperties5.Append(tableCellVerticalAlignment4);

            Paragraph paragraph5 = new Paragraph() { RsidParagraphMarkRevision = "007642DC", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "001D4055" };

            ParagraphProperties paragraphProperties5 = new ParagraphProperties();
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
            text8.Text = "对学习的影响";

            run8.Append(runProperties8);
            run8.Append(text8);

            paragraph5.Append(paragraphProperties5);
            paragraph5.Append(run8);

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

            Paragraph paragraph6 = new Paragraph() { RsidParagraphMarkRevision = "007642DC", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "001D4055" };

            ParagraphProperties paragraphProperties6 = new ParagraphProperties();
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

            TableRow tableRow3 = new TableRow() { RsidTableRowMarkRevision = "007642DC", RsidTableRowAddition = "001D4055", RsidTableRowProperties = "006A0186" };

            TableRowProperties tableRowProperties3 = new TableRowProperties();
            TableRowHeight tableRowHeight3 = new TableRowHeight() { Val = (UInt32Value)4349U };

            tableRowProperties3.Append(tableRowHeight3);

            TableCell tableCell7 = new TableCell();

            TableCellProperties tableCellProperties7 = new TableCellProperties();
            TableCellWidth tableCellWidth7 = new TableCellWidth() { Width = "2690", Type = TableWidthUnitValues.Dxa };
            Shading shading7 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            TableCellVerticalAlignment tableCellVerticalAlignment6 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties7.Append(tableCellWidth7);
            tableCellProperties7.Append(shading7);
            tableCellProperties7.Append(tableCellVerticalAlignment6);

            Paragraph paragraph7 = new Paragraph() { RsidParagraphMarkRevision = "004F3E28", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "001D4055" };

            ParagraphProperties paragraphProperties7 = new ParagraphProperties();
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
            text11.Text = "学业成功";//TODO量表名称

            run11.Append(runProperties11);
            run11.Append(text11);

            Run run12 = new Run();

            RunProperties runProperties12 = new RunProperties();
            RunFonts runFonts19 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Bold bold19 = new Bold();
            Color color5 = new Color() { Val = "000000" };
            Kern kern19 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize19 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript19 = new FontSizeComplexScript() { Val = "24" };

            runProperties12.Append(runFonts19);
            runProperties12.Append(bold19);
            runProperties12.Append(color5);
            runProperties12.Append(kern19);
            runProperties12.Append(fontSize19);
            runProperties12.Append(fontSizeComplexScript19);
            Text text12 = new Text();
            text12.Text = "-";

            run12.Append(runProperties12);
            run12.Append(text12);

            Run run13 = new Run();

            RunProperties runProperties13 = new RunProperties();
            RunFonts runFonts20 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Bold bold20 = new Bold();
            Color color6 = new Color() { Val = "000000" };
            Kern kern20 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize20 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript20 = new FontSizeComplexScript() { Val = "24" };

            runProperties13.Append(runFonts20);
            runProperties13.Append(bold20);
            runProperties13.Append(color6);
            runProperties13.Append(kern20);
            runProperties13.Append(fontSize20);
            runProperties13.Append(fontSizeComplexScript20);
            Text text13 = new Text();
            text13.Text = "内归因";//TODO维度名称

            run13.Append(runProperties13);
            run13.Append(text13);

            paragraph7.Append(paragraphProperties7);
            paragraph7.Append(run11);
            paragraph7.Append(run12);
            paragraph7.Append(run13);

            tableCell7.Append(tableCellProperties7);
            tableCell7.Append(paragraph7);

            TableCell tableCell8 = new TableCell();

            TableCellProperties tableCellProperties8 = new TableCellProperties();
            TableCellWidth tableCellWidth8 = new TableCellWidth() { Width = "682", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment7 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties8.Append(tableCellWidth8);
            tableCellProperties8.Append(tableCellVerticalAlignment7);

            Paragraph paragraph8 = new Paragraph() { RsidParagraphMarkRevision = "001B2170", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "001D4055" };

            ParagraphProperties paragraphProperties8 = new ParagraphProperties();
            Justification justification8 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties8 = new ParagraphMarkRunProperties();
            FontSize fontSize23 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript23 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties8.Append(fontSize23);
            paragraphMarkRunProperties8.Append(fontSizeComplexScript23);

            paragraphProperties8.Append(justification8);
            paragraphProperties8.Append(paragraphMarkRunProperties8);


            paragraph8.Append(paragraphProperties8);

            Paragraph paragraph9 = new Paragraph() { RsidParagraphMarkRevision = "001B2170", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "001D4055" };

            ParagraphProperties paragraphProperties9 = new ParagraphProperties();
            Justification justification9 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties9 = new ParagraphMarkRunProperties();
            RunFonts runFonts26 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "宋体" };
            Color color9 = new Color() { Val = "000000" };
            Kern kern23 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize29 = new FontSize() { Val = "22" };

            paragraphMarkRunProperties9.Append(runFonts26);
            paragraphMarkRunProperties9.Append(color9);
            paragraphMarkRunProperties9.Append(kern23);
            paragraphMarkRunProperties9.Append(fontSize29);

            paragraphProperties9.Append(justification9);
            paragraphProperties9.Append(paragraphMarkRunProperties9);

            Run run21 = new Run() { RsidRunProperties = "001B2170" };

            RunProperties runProperties21 = new RunProperties();
            RunFonts runFonts27 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "宋体" };
            Color color10 = new Color() { Val = "000000" };
            Kern kern24 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize30 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript24 = new FontSizeComplexScript() { Val = "24" };

            runProperties21.Append(runFonts27);
            runProperties21.Append(color10);
            runProperties21.Append(kern24);
            runProperties21.Append(fontSize30);
            runProperties21.Append(fontSizeComplexScript24);
            Text text16 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text16.Text = "61.40 ";//TODO 得分

            run21.Append(runProperties21);
            run21.Append(text16);

            paragraph9.Append(paragraphProperties9);
            paragraph9.Append(run21);

            Paragraph paragraph10 = new Paragraph() { RsidParagraphMarkRevision = "001B2170", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "001D4055" };

            ParagraphProperties paragraphProperties10 = new ParagraphProperties();
            Justification justification10 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties10 = new ParagraphMarkRunProperties();
            RunFonts runFonts28 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Bold bold23 = new Bold();
            Color color11 = new Color() { Val = "000000" };
            Kern kern25 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize31 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript25 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties10.Append(runFonts28);
            paragraphMarkRunProperties10.Append(bold23);
            paragraphMarkRunProperties10.Append(color11);
            paragraphMarkRunProperties10.Append(kern25);
            paragraphMarkRunProperties10.Append(fontSize31);
            paragraphMarkRunProperties10.Append(fontSizeComplexScript25);

            paragraphProperties10.Append(justification10);
            paragraphProperties10.Append(paragraphMarkRunProperties10);

            Run run22 = new Run() { RsidRunProperties = "001B2170" };

            RunProperties runProperties22 = new RunProperties();
            RunFonts runFonts29 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Bold bold24 = new Bold();
            Color color12 = new Color() { Val = "000000" };
            Kern kern26 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize32 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript26 = new FontSizeComplexScript() { Val = "24" };

            runProperties22.Append(runFonts29);
            runProperties22.Append(bold24);
            runProperties22.Append(color12);
            runProperties22.Append(kern26);
            runProperties22.Append(fontSize32);
            runProperties22.Append(fontSizeComplexScript26);
            FieldChar fieldChar3 = new FieldChar() { FieldCharType = FieldCharValues.End };

            run22.Append(runProperties22);
            run22.Append(fieldChar3);

            paragraph10.Append(paragraphProperties10);
            paragraph10.Append(run22);

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

            Paragraph paragraph11 = new Paragraph() { RsidParagraphMarkRevision = "001B2170", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "001D4055" };

            ParagraphProperties paragraphProperties11 = new ParagraphProperties();
            Justification justification11 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties11 = new ParagraphMarkRunProperties();
            RunFonts runFonts30 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Bold bold25 = new Bold();
            Color color13 = new Color() { Val = "000000" };
            Kern kern27 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize33 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript27 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties11.Append(runFonts30);
            paragraphMarkRunProperties11.Append(bold25);
            paragraphMarkRunProperties11.Append(color13);
            paragraphMarkRunProperties11.Append(kern27);
            paragraphMarkRunProperties11.Append(fontSize33);
            paragraphMarkRunProperties11.Append(fontSizeComplexScript27);

            paragraphProperties11.Append(justification11);
            paragraphProperties11.Append(paragraphMarkRunProperties11);

            Run run23 = new Run() { RsidRunProperties = "001B2170" };

            RunProperties runProperties23 = new RunProperties();
            RunFonts runFonts31 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Bold bold26 = new Bold();
            Color color14 = new Color() { Val = "000000" };
            Kern kern28 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize34 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript28 = new FontSizeComplexScript() { Val = "24" };

            runProperties23.Append(runFonts31);
            runProperties23.Append(bold26);
            runProperties23.Append(color14);
            runProperties23.Append(kern28);
            runProperties23.Append(fontSize34);
            runProperties23.Append(fontSizeComplexScript28);
            Text text17 = new Text();
            text17.Text = "高水平";//TODO水平

            run23.Append(runProperties23);
            run23.Append(text17);


            paragraph11.Append(paragraphProperties11);
            paragraph11.Append(run23);

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

            Paragraph paragraph12 = new Paragraph() { RsidParagraphMarkRevision = "001B2170", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "001D4055" };

            ParagraphProperties paragraphProperties12 = new ParagraphProperties();
            Justification justification12 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties12 = new ParagraphMarkRunProperties();
            RunFonts runFonts33 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Bold bold28 = new Bold();
            Color color16 = new Color() { Val = "000000" };
            Kern kern30 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize36 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript30 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties12.Append(runFonts33);
            paragraphMarkRunProperties12.Append(bold28);
            paragraphMarkRunProperties12.Append(color16);
            paragraphMarkRunProperties12.Append(kern30);
            paragraphMarkRunProperties12.Append(fontSize36);
            paragraphMarkRunProperties12.Append(fontSizeComplexScript30);

            paragraphProperties12.Append(justification12);
            paragraphProperties12.Append(paragraphMarkRunProperties12);

            Run run25 = new Run() { RsidRunProperties = "001B2170" };

            RunProperties runProperties25 = new RunProperties();
            RunFonts runFonts34 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Color color17 = new Color() { Val = "000000" };
            Kern kern31 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize37 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript31 = new FontSizeComplexScript() { Val = "24" };

            runProperties25.Append(runFonts34);
            runProperties25.Append(color17);
            runProperties25.Append(kern31);
            runProperties25.Append(fontSize37);
            runProperties25.Append(fontSizeComplexScript31);
            Text text19 = new Text();
            text19.Text = "- - -";

            run25.Append(runProperties25);
            run25.Append(text19);           

            paragraph12.Append(paragraphProperties12);
            paragraph12.Append(run25);

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

            Paragraph paragraph13 = new Paragraph() { RsidParagraphMarkRevision = "00C16BA1", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "001D4055" };

            ParagraphProperties paragraphProperties13 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties13 = new ParagraphMarkRunProperties();
            FontSize fontSize41 = new FontSize() { Val = "4" };
            FontSizeComplexScript fontSizeComplexScript35 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties13.Append(fontSize41);
            paragraphMarkRunProperties13.Append(fontSizeComplexScript35);

            paragraphProperties13.Append(paragraphMarkRunProperties13);

            paragraph13.Append(paragraphProperties13);

            Paragraph paragraph14 = new Paragraph() { RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "001D4055" };

            ParagraphProperties paragraphProperties14 = new ParagraphProperties();
            Indentation indentation1 = new Indentation() { FirstLine = "480" };

            ParagraphMarkRunProperties paragraphMarkRunProperties14 = new ParagraphMarkRunProperties();
            RunFonts runFonts41 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "宋体" };
            Color color21 = new Color() { Val = "000000" };
            Kern kern35 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize44 = new FontSize() { Val = "24" };

            paragraphMarkRunProperties14.Append(runFonts41);
            paragraphMarkRunProperties14.Append(color21);
            paragraphMarkRunProperties14.Append(kern35);
            paragraphMarkRunProperties14.Append(fontSize44);

            paragraphProperties14.Append(indentation1);
            paragraphProperties14.Append(paragraphMarkRunProperties14);

            Run run34 = new Run() { RsidRunProperties = "00C16BA1" };

            RunProperties runProperties32 = new RunProperties();
            RunFonts runFonts42 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "宋体" };
            Color color22 = new Color() { Val = "000000" };
            Kern kern36 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize45 = new FontSize() { Val = "24" };

            runProperties32.Append(runFonts42);
            runProperties32.Append(color22);
            runProperties32.Append(kern36);
            runProperties32.Append(fontSize45);
            Text text23 = new Text();
            text23.Text = "学业成功时的归因方式是内归因，是一种比较积极的归因方式。倾向于把学业成功归因于自身学习努力和学习能力这样的可控因素，是对自己能力的一种肯定，他们对于学习能够更有掌控感和成就感，从而增强了对于学习的控制和对成功的渴望，那么他们就会积极地争取成功，会感到满意、自豪等积极情绪，有助于维持和激发其随后的学习动机，他们就更加愿意去学习。在这种情况下，更容易完成我们所设的任务，而不是选择逃避。";

            run34.Append(runProperties32);
            run34.Append(text23);

            paragraph14.Append(paragraphProperties14);
            paragraph14.Append(run34);

            Paragraph paragraph15 = new Paragraph() { RsidParagraphMarkRevision = "00C16BA1", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "001D4055" };

            ParagraphProperties paragraphProperties15 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties15 = new ParagraphMarkRunProperties();
            FontSize fontSize46 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript36 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties15.Append(fontSize46);
            paragraphMarkRunProperties15.Append(fontSizeComplexScript36);

            paragraphProperties15.Append(paragraphMarkRunProperties15);
            paragraph15.Append(paragraphProperties15);

            Paragraph paragraph16 = new Paragraph() { RsidParagraphMarkRevision = "00C16BA1", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "001D4055" };

            ParagraphProperties paragraphProperties16 = new ParagraphProperties();
            Indentation indentation2 = new Indentation() { FirstLine = "480" };

            ParagraphMarkRunProperties paragraphMarkRunProperties16 = new ParagraphMarkRunProperties();
            RunFonts runFonts46 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "宋体" };
            Color color23 = new Color() { Val = "000000" };
            Kern kern37 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize49 = new FontSize() { Val = "22" };

            paragraphMarkRunProperties16.Append(runFonts46);
            paragraphMarkRunProperties16.Append(color23);
            paragraphMarkRunProperties16.Append(kern37);
            paragraphMarkRunProperties16.Append(fontSize49);

            paragraphProperties16.Append(indentation2);
            paragraphProperties16.Append(paragraphMarkRunProperties16);

            Run run40 = new Run() { RsidRunProperties = "00C16BA1" };

            RunProperties runProperties36 = new RunProperties();
            RunFonts runFonts47 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "宋体" };
            Color color24 = new Color() { Val = "000000" };
            Kern kern38 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize50 = new FontSize() { Val = "24" };

            runProperties36.Append(runFonts47);
            runProperties36.Append(color24);
            runProperties36.Append(kern38);
            runProperties36.Append(fontSize50);
            Text text24 = new Text();
            text24.Text = "学业失败时的归因方式是外归因，是一种相对积极的归因方式。学习失败时归因于运气不佳，题目难度太大等外部因素时，学生就不会产生内疚或羞愧的心理。即使考试不理想，也不会否认自己的学习能力，不会打击继续努力信心。这种归因方式对于学业成绩的提高有一定的推动作用。";

            run40.Append(runProperties36);
            run40.Append(text24);

            paragraph16.Append(paragraphProperties16);
            paragraph16.Append(run40);

            Paragraph paragraph17 = new Paragraph() { RsidParagraphMarkRevision = "00C16BA1", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "006A0186", RsidRunAdditionDefault = "001D4055" };

            ParagraphProperties paragraphProperties17 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties17 = new ParagraphMarkRunProperties();
            RunFonts runFonts48 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "宋体" };
            Color color25 = new Color() { Val = "000000" };
            Kern kern39 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize51 = new FontSize() { Val = "24" };

            paragraphMarkRunProperties17.Append(runFonts48);
            paragraphMarkRunProperties17.Append(color25);
            paragraphMarkRunProperties17.Append(kern39);
            paragraphMarkRunProperties17.Append(fontSize51);

            paragraphProperties17.Append(paragraphMarkRunProperties17);

            Run run41 = new Run();

            RunProperties runProperties37 = new RunProperties();
            RunFonts runFonts49 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Color color26 = new Color() { Val = "000000" };
            Kern kern40 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize52 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript37 = new FontSizeComplexScript() { Val = "24" };

            runProperties37.Append(runFonts49);
            runProperties37.Append(color26);
            runProperties37.Append(kern40);
            runProperties37.Append(fontSize52);
            runProperties37.Append(fontSizeComplexScript37);
            FieldChar fieldChar8 = new FieldChar() { FieldCharType = FieldCharValues.End };

            run41.Append(runProperties37);
            run41.Append(fieldChar8);

            paragraph17.Append(paragraphProperties17);
            paragraph17.Append(run41);

            Paragraph paragraph18 = new Paragraph() { RsidParagraphMarkRevision = "007D1A31", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "001D4055" };

            ParagraphProperties paragraphProperties18 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties18 = new ParagraphMarkRunProperties();
            RunFonts runFonts50 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Color color27 = new Color() { Val = "000000" };
            Kern kern41 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize53 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript38 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties18.Append(runFonts50);
            paragraphMarkRunProperties18.Append(color27);
            paragraphMarkRunProperties18.Append(kern41);
            paragraphMarkRunProperties18.Append(fontSize53);
            paragraphMarkRunProperties18.Append(fontSizeComplexScript38);

            paragraphProperties18.Append(paragraphMarkRunProperties18);
            paragraph18.Append(paragraphProperties18);

            tableCell11.Append(tableCellProperties11);
            tableCell11.Append(paragraph13);
            tableCell11.Append(paragraph14);
            tableCell11.Append(paragraph15);
            tableCell11.Append(paragraph16);
            tableCell11.Append(paragraph17);
            tableCell11.Append(paragraph18);

            tableRow3.Append(tableRowProperties3);
            tableRow3.Append(tableCell7);
            tableRow3.Append(tableCell8);
            tableRow3.Append(tableCell9);
            tableRow3.Append(tableCell10);
            tableRow3.Append(tableCell11);

            TableRow tableRow4 = new TableRow() { RsidTableRowMarkRevision = "007642DC", RsidTableRowAddition = "001D4055", RsidTableRowProperties = "006A0186" };

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

            Paragraph paragraph19 = new Paragraph() { RsidParagraphMarkRevision = "00973E52", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "001D4055" };

            ParagraphProperties paragraphProperties19 = new ParagraphProperties();
            Justification justification13 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties19 = new ParagraphMarkRunProperties();
            RunFonts runFonts52 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Kern kern43 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize55 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript40 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties19.Append(runFonts52);
            paragraphMarkRunProperties19.Append(kern43);
            paragraphMarkRunProperties19.Append(fontSize55);
            paragraphMarkRunProperties19.Append(fontSizeComplexScript40);

            paragraphProperties19.Append(justification13);
            paragraphProperties19.Append(paragraphMarkRunProperties19);

            Run run43 = new Run();

            RunProperties runProperties39 = new RunProperties();
            RunFonts runFonts53 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Kern kern44 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize56 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript41 = new FontSizeComplexScript() { Val = "24" };

            runProperties39.Append(runFonts53);
            runProperties39.Append(kern44);
            runProperties39.Append(fontSize56);
            runProperties39.Append(fontSizeComplexScript41);
            Text text25 = new Text();
            text25.Text = "学业成功";//TODO量表名

            run43.Append(runProperties39);
            run43.Append(text25);

            Run run45 = new Run();

            RunProperties runProperties41 = new RunProperties();
            RunFonts runFonts55 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Kern kern46 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize58 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript43 = new FontSizeComplexScript() { Val = "24" };

            runProperties41.Append(runFonts55);
            runProperties41.Append(kern46);
            runProperties41.Append(fontSize58);
            runProperties41.Append(fontSizeComplexScript43);
            Text text27 = new Text();
            text27.Text = "-";

            run45.Append(runProperties41);
            run45.Append(text27);

            Run run46 = new Run() { RsidRunProperties = "00973E52" };

            RunProperties runProperties42 = new RunProperties();
            RunFonts runFonts56 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Kern kern47 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize59 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript44 = new FontSizeComplexScript() { Val = "24" };

            runProperties42.Append(runFonts56);
            runProperties42.Append(kern47);
            runProperties42.Append(fontSize59);
            runProperties42.Append(fontSizeComplexScript44);
            Text text28 = new Text();
            text28.Text = "外归因";//TODO维度名

            run46.Append(runProperties42);
            run46.Append(text28);

            paragraph19.Append(paragraphProperties19);
            paragraph19.Append(run43);
            paragraph19.Append(run45);
            paragraph19.Append(run46);

            tableCell12.Append(tableCellProperties12);
            tableCell12.Append(paragraph19);

            TableCell tableCell13 = new TableCell();

            TableCellProperties tableCellProperties13 = new TableCellProperties();
            TableCellWidth tableCellWidth13 = new TableCellWidth() { Width = "682", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment12 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties13.Append(tableCellWidth13);
            tableCellProperties13.Append(tableCellVerticalAlignment12);

            Paragraph paragraph20 = new Paragraph() { RsidParagraphMarkRevision = "001B2170", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "001D4055" };

            ParagraphProperties paragraphProperties20 = new ParagraphProperties();
            Justification justification14 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties20 = new ParagraphMarkRunProperties();
            FontSize fontSize61 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript46 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties20.Append(fontSize61);
            paragraphMarkRunProperties20.Append(fontSizeComplexScript46);

            paragraphProperties20.Append(justification14);
            paragraphProperties20.Append(paragraphMarkRunProperties20);
            paragraph20.Append(paragraphProperties20);

            Paragraph paragraph21 = new Paragraph() { RsidParagraphMarkRevision = "001B2170", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "001D4055" };

            ParagraphProperties paragraphProperties21 = new ParagraphProperties();
            Justification justification15 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties21 = new ParagraphMarkRunProperties();
            RunFonts runFonts61 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "宋体" };
            Color color29 = new Color() { Val = "000000" };
            Kern kern49 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize67 = new FontSize() { Val = "22" };

            paragraphMarkRunProperties21.Append(runFonts61);
            paragraphMarkRunProperties21.Append(color29);
            paragraphMarkRunProperties21.Append(kern49);
            paragraphMarkRunProperties21.Append(fontSize67);

            paragraphProperties21.Append(justification15);
            paragraphProperties21.Append(paragraphMarkRunProperties21);

            Run run53 = new Run() { RsidRunProperties = "001B2170" };

            RunProperties runProperties49 = new RunProperties();
            RunFonts runFonts62 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "宋体" };
            Color color30 = new Color() { Val = "000000" };
            Kern kern50 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize68 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript47 = new FontSizeComplexScript() { Val = "24" };

            runProperties49.Append(runFonts62);
            runProperties49.Append(color30);
            runProperties49.Append(kern50);
            runProperties49.Append(fontSize68);
            runProperties49.Append(fontSizeComplexScript47);
            Text text30 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text30.Text = "45.00 ";//TODO得分

            run53.Append(runProperties49);
            run53.Append(text30);

            paragraph21.Append(paragraphProperties21);
            paragraph21.Append(run53);

            Paragraph paragraph22 = new Paragraph() { RsidParagraphMarkRevision = "001B2170", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "001D4055" };

            ParagraphProperties paragraphProperties22 = new ParagraphProperties();
            Justification justification16 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties22 = new ParagraphMarkRunProperties();
            RunFonts runFonts63 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Bold bold29 = new Bold();
            Color color31 = new Color() { Val = "000000" };
            Kern kern51 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize69 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript48 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties22.Append(runFonts63);
            paragraphMarkRunProperties22.Append(bold29);
            paragraphMarkRunProperties22.Append(color31);
            paragraphMarkRunProperties22.Append(kern51);
            paragraphMarkRunProperties22.Append(fontSize69);
            paragraphMarkRunProperties22.Append(fontSizeComplexScript48);

            paragraphProperties22.Append(justification16);
            paragraphProperties22.Append(paragraphMarkRunProperties22);

            Run run54 = new Run() { RsidRunProperties = "001B2170" };

            RunProperties runProperties50 = new RunProperties();
            RunFonts runFonts64 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Bold bold30 = new Bold();
            Color color32 = new Color() { Val = "000000" };
            Kern kern52 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize70 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript49 = new FontSizeComplexScript() { Val = "24" };

            runProperties50.Append(runFonts64);
            runProperties50.Append(bold30);
            runProperties50.Append(color32);
            runProperties50.Append(kern52);
            runProperties50.Append(fontSize70);
            runProperties50.Append(fontSizeComplexScript49);
            FieldChar fieldChar12 = new FieldChar() { FieldCharType = FieldCharValues.End };

            run54.Append(runProperties50);
            run54.Append(fieldChar12);

            paragraph22.Append(paragraphProperties22);
            paragraph22.Append(run54);

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

            Paragraph paragraph23 = new Paragraph() { RsidParagraphMarkRevision = "001B2170", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "001D4055" };

            ParagraphProperties paragraphProperties23 = new ParagraphProperties();
            Justification justification17 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties23 = new ParagraphMarkRunProperties();
            RunFonts runFonts65 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Bold bold31 = new Bold();
            Color color33 = new Color() { Val = "000000" };
            Kern kern53 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize71 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript50 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties23.Append(runFonts65);
            paragraphMarkRunProperties23.Append(bold31);
            paragraphMarkRunProperties23.Append(color33);
            paragraphMarkRunProperties23.Append(kern53);
            paragraphMarkRunProperties23.Append(fontSize71);
            paragraphMarkRunProperties23.Append(fontSizeComplexScript50);

            paragraphProperties23.Append(justification17);
            paragraphProperties23.Append(paragraphMarkRunProperties23);

            Run run55 = new Run() { RsidRunProperties = "001B2170" };

            RunProperties runProperties51 = new RunProperties();
            RunFonts runFonts66 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Bold bold32 = new Bold();
            Color color34 = new Color() { Val = "000000" };
            Kern kern54 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize72 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript51 = new FontSizeComplexScript() { Val = "24" };

            runProperties51.Append(runFonts66);
            runProperties51.Append(bold32);
            runProperties51.Append(color34);
            runProperties51.Append(kern54);
            runProperties51.Append(fontSize72);
            runProperties51.Append(fontSizeComplexScript51);
            Text text31 = new Text();
            text31.Text = "中等";

            run55.Append(runProperties51);
            run55.Append(text31);

            Run run56 = new Run() { RsidRunProperties = "001B2170" };

            RunProperties runProperties52 = new RunProperties();
            RunFonts runFonts67 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Bold bold33 = new Bold();
            Color color35 = new Color() { Val = "000000" };
            Kern kern55 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize73 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript52 = new FontSizeComplexScript() { Val = "24" };

            runProperties52.Append(runFonts67);
            runProperties52.Append(bold33);
            runProperties52.Append(color35);
            runProperties52.Append(kern55);
            runProperties52.Append(fontSize73);
            runProperties52.Append(fontSizeComplexScript52);
            Text text32 = new Text();
            text32.Text = "水平";

            run56.Append(runProperties52);
            run56.Append(text32);

            paragraph23.Append(paragraphProperties23);
            paragraph23.Append(run55);
            paragraph23.Append(run56);

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

            Paragraph paragraph24 = new Paragraph() { RsidParagraphMarkRevision = "001B2170", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "001D4055" };

            ParagraphProperties paragraphProperties24 = new ParagraphProperties();
            Justification justification18 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties24 = new ParagraphMarkRunProperties();
            RunFonts runFonts68 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            Bold bold34 = new Bold();
            Color color36 = new Color() { Val = "000000" };
            Kern kern56 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize74 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript53 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties24.Append(runFonts68);
            paragraphMarkRunProperties24.Append(bold34);
            paragraphMarkRunProperties24.Append(color36);
            paragraphMarkRunProperties24.Append(kern56);
            paragraphMarkRunProperties24.Append(fontSize74);
            paragraphMarkRunProperties24.Append(fontSizeComplexScript53);

            paragraphProperties24.Append(justification18);
            paragraphProperties24.Append(paragraphMarkRunProperties24);

            Run run57 = new Run() { RsidRunProperties = "001B2170" };

            RunProperties runProperties53 = new RunProperties();
            RunFonts runFonts69 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Color color37 = new Color() { Val = "000000" };
            Kern kern57 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize75 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript54 = new FontSizeComplexScript() { Val = "24" };

            runProperties53.Append(runFonts69);
            runProperties53.Append(color37);
            runProperties53.Append(kern57);
            runProperties53.Append(fontSize75);
            runProperties53.Append(fontSizeComplexScript54);
            Text text33 = new Text();
            text33.Text = "-";

            run57.Append(runProperties53);
            run57.Append(text33);

            Run run58 = new Run() { RsidRunProperties = "001B2170" };

            RunProperties runProperties54 = new RunProperties();
            RunFonts runFonts70 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Color color38 = new Color() { Val = "000000" };
            Kern kern58 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize76 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript55 = new FontSizeComplexScript() { Val = "24" };

            runProperties54.Append(runFonts70);
            runProperties54.Append(color38);
            runProperties54.Append(kern58);
            runProperties54.Append(fontSize76);
            runProperties54.Append(fontSizeComplexScript55);
            Text text34 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text34.Text = " ";

            run58.Append(runProperties54);
            run58.Append(text34);

            Run run59 = new Run() { RsidRunProperties = "001B2170" };

            RunProperties runProperties55 = new RunProperties();
            RunFonts runFonts71 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Color color39 = new Color() { Val = "000000" };
            Kern kern59 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize77 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript56 = new FontSizeComplexScript() { Val = "24" };

            runProperties55.Append(runFonts71);
            runProperties55.Append(color39);
            runProperties55.Append(kern59);
            runProperties55.Append(fontSize77);
            runProperties55.Append(fontSizeComplexScript56);
            Text text35 = new Text();
            text35.Text = "-";

            run59.Append(runProperties55);
            run59.Append(text35);

            Run run60 = new Run() { RsidRunProperties = "001B2170" };

            RunProperties runProperties56 = new RunProperties();
            RunFonts runFonts72 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Color color40 = new Color() { Val = "000000" };
            Kern kern60 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize78 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript57 = new FontSizeComplexScript() { Val = "24" };

            runProperties56.Append(runFonts72);
            runProperties56.Append(color40);
            runProperties56.Append(kern60);
            runProperties56.Append(fontSize78);
            runProperties56.Append(fontSizeComplexScript57);
            Text text36 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text36.Text = " -";

            run60.Append(runProperties56);
            run60.Append(text36);

            paragraph24.Append(paragraphProperties24);
            paragraph24.Append(run57);
            paragraph24.Append(run58);
            paragraph24.Append(run59);
            paragraph24.Append(run60);

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

            Paragraph paragraph25 = new Paragraph() { RsidParagraphMarkRevision = "002A093C", RsidParagraphAddition = "001D4055", RsidParagraphProperties = "001231E8", RsidRunAdditionDefault = "001D4055" };

            ParagraphProperties paragraphProperties25 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties25 = new ParagraphMarkRunProperties();
            RunFonts runFonts73 = new RunFonts() { Ascii = "宋体", HighAnsi = "宋体", ComplexScript = "Tahoma" };
            Color color41 = new Color() { Val = "000000" };
            Kern kern61 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize79 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript58 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties25.Append(runFonts73);
            paragraphMarkRunProperties25.Append(color41);
            paragraphMarkRunProperties25.Append(kern61);
            paragraphMarkRunProperties25.Append(fontSize79);
            paragraphMarkRunProperties25.Append(fontSizeComplexScript58);

            paragraphProperties25.Append(paragraphMarkRunProperties25);

            paragraph25.Append(paragraphProperties25);

            tableCell16.Append(tableCellProperties16);
            tableCell16.Append(paragraph25);

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

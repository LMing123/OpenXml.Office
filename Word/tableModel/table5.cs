using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Word.tableModel
{
    using DocumentFormat.OpenXml.Wordprocessing;
    using DocumentFormat.OpenXml;
    using Lsj.Util.Collections;
    using Word.Enum;
    public partial class GeneratedClass
    {
        public Table GenerateTable5(string title, string evaluate, SafeDictionary<string, (string, string, eInfluence, double)> content)
        {
            Table table1 = new Table();

            TableProperties tableProperties1 = new TableProperties();
            TablePositionProperties tablePositionProperties1 = new TablePositionProperties() { LeftFromText = 180, RightFromText = 180, VerticalAnchor = VerticalAnchorValues.Text, HorizontalAnchor = HorizontalAnchorValues.Margin, TablePositionXAlignment = HorizontalAlignmentValues.Center, TablePositionY = 391 };
            TableWidth tableWidth1 = new TableWidth() { Width = "9981", Type = TableWidthUnitValues.Dxa };

            TableBorders tableBorders1 = new TableBorders();
            TopBorder topBorder1 = new TopBorder() { Val = BorderValues.Double, Color = "76923C", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.Double, Color = "76923C", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.Double, Color = "76923C", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.Double, Color = "76923C", Size = (UInt32Value)6U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder1 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "9BBB59", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder1 = new InsideVerticalBorder() { Val = BorderValues.Single, Color = "9BBB59", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

            tableBorders1.Append(topBorder1);
            tableBorders1.Append(leftBorder1);
            tableBorders1.Append(bottomBorder1);
            tableBorders1.Append(rightBorder1);
            tableBorders1.Append(insideHorizontalBorder1);
            tableBorders1.Append(insideVerticalBorder1);
            TableLook tableLook1 = new TableLook() { Val = "00A0", FirstRow = true, LastRow = false, FirstColumn = true, LastColumn = false, NoHorizontalBand = false, NoVerticalBand = false };

            tableProperties1.Append(tablePositionProperties1);
            tableProperties1.Append(tableWidth1);
            tableProperties1.Append(tableBorders1);
            tableProperties1.Append(tableLook1);

            TableGrid tableGrid1 = new TableGrid();
            GridColumn gridColumn1 = new GridColumn() { Width = "1678" };
            GridColumn gridColumn2 = new GridColumn() { Width = "4453" };
            GridColumn gridColumn3 = new GridColumn() { Width = "756" };
            GridColumn gridColumn4 = new GridColumn() { Width = "755" };
            GridColumn gridColumn5 = new GridColumn() { Width = "756" };
            GridColumn gridColumn6 = new GridColumn() { Width = "756" };
            GridColumn gridColumn7 = new GridColumn() { Width = "827" };

            tableGrid1.Append(gridColumn1);
            tableGrid1.Append(gridColumn2);
            tableGrid1.Append(gridColumn3);
            tableGrid1.Append(gridColumn4);
            tableGrid1.Append(gridColumn5);
            tableGrid1.Append(gridColumn6);
            tableGrid1.Append(gridColumn7);

            TableRow tableRow1 = new TableRow() { RsidTableRowMarkRevision = "00C709B4", RsidTableRowAddition = "00D061CC", RsidTableRowProperties = "00D061CC" };

            TableRowProperties tableRowProperties1 = new TableRowProperties();
            TableRowHeight tableRowHeight1 = new TableRowHeight() { Val = (UInt32Value)239U };

            tableRowProperties1.Append(tableRowHeight1);

            TableCell tableCell1 = new TableCell();

            TableCellProperties tableCellProperties1 = new TableCellProperties();
            TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = "9981", Type = TableWidthUnitValues.Dxa };
            GridSpan gridSpan1 = new GridSpan() { Val = 7 };

            TableCellBorders tableCellBorders1 = new TableCellBorders();
            TopBorder topBorder2 = new TopBorder() { Val = BorderValues.Double, Color = "76923C", Size = (UInt32Value)6U, Space = (UInt32Value)0U };

            tableCellBorders1.Append(topBorder2);
            Shading shading1 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "D6E3BC" };
            TableCellVerticalAlignment tableCellVerticalAlignment1 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties1.Append(tableCellWidth1);
            tableCellProperties1.Append(gridSpan1);
            tableCellProperties1.Append(tableCellBorders1);
            tableCellProperties1.Append(shading1);
            tableCellProperties1.Append(tableCellVerticalAlignment1);

            Paragraph paragraph1 = new Paragraph() { RsidParagraphMarkRevision = "00C709B4", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "009E6323", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            Indentation indentation1 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification1 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
            Bold bold1 = new Bold();
            Kern kern1 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties1.Append(bold1);
            paragraphMarkRunProperties1.Append(kern1);
            paragraphMarkRunProperties1.Append(fontSize1);
            paragraphMarkRunProperties1.Append(fontSizeComplexScript1);

            paragraphProperties1.Append(indentation1);
            paragraphProperties1.Append(justification1);
            paragraphProperties1.Append(paragraphMarkRunProperties1);

            Run run1 = new Run();

            RunProperties runProperties1 = new RunProperties();
            RunFonts runFonts1 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold2 = new Bold();
            Kern kern2 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize2 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "20" };

            runProperties1.Append(runFonts1);
            runProperties1.Append(bold2);
            runProperties1.Append(kern2);
            runProperties1.Append(fontSize2);
            runProperties1.Append(fontSizeComplexScript2);
            Text text1 = new Text();
            text1.Text = "心理健康量表";

            run1.Append(runProperties1);
            run1.Append(text1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            tableCell1.Append(tableCellProperties1);
            tableCell1.Append(paragraph1);

            tableRow1.Append(tableRowProperties1);
            tableRow1.Append(tableCell1);

            TableRow tableRow2 = new TableRow() { RsidTableRowMarkRevision = "00C709B4", RsidTableRowAddition = "00D061CC", RsidTableRowProperties = "00D061CC" };

            TableRowProperties tableRowProperties2 = new TableRowProperties();
            TableRowHeight tableRowHeight2 = new TableRowHeight() { Val = (UInt32Value)303U };

            tableRowProperties2.Append(tableRowHeight2);

            TableCell tableCell2 = new TableCell();

            TableCellProperties tableCellProperties2 = new TableCellProperties();
            TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = "1678", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge1 = new VerticalMerge() { Val = MergedCellValues.Restart };
            Shading shading2 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "D6E3BC" };
            TableCellVerticalAlignment tableCellVerticalAlignment2 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties2.Append(tableCellWidth2);
            tableCellProperties2.Append(verticalMerge1);
            tableCellProperties2.Append(shading2);
            tableCellProperties2.Append(tableCellVerticalAlignment2);

            Paragraph paragraph2 = new Paragraph() { RsidParagraphMarkRevision = "00C709B4", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            Indentation indentation2 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification2 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
            Bold bold7 = new Bold();
            Kern kern7 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize7 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript7 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties2.Append(bold7);
            paragraphMarkRunProperties2.Append(kern7);
            paragraphMarkRunProperties2.Append(fontSize7);
            paragraphMarkRunProperties2.Append(fontSizeComplexScript7);

            paragraphProperties2.Append(indentation2);
            paragraphProperties2.Append(justification2);
            paragraphProperties2.Append(paragraphMarkRunProperties2);

            Run run6 = new Run() { RsidRunProperties = "00C709B4" };

            RunProperties runProperties6 = new RunProperties();
            RunFonts runFonts3 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold8 = new Bold();
            Kern kern8 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize8 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript8 = new FontSizeComplexScript() { Val = "20" };

            runProperties6.Append(runFonts3);
            runProperties6.Append(bold8);
            runProperties6.Append(kern8);
            runProperties6.Append(fontSize8);
            runProperties6.Append(fontSizeComplexScript8);
            Text text6 = new Text();
            text6.Text = "因子名称";

            run6.Append(runProperties6);
            run6.Append(text6);

            paragraph2.Append(paragraphProperties2);
            paragraph2.Append(run6);

            tableCell2.Append(tableCellProperties2);
            tableCell2.Append(paragraph2);

            TableCell tableCell3 = new TableCell();

            TableCellProperties tableCellProperties3 = new TableCellProperties();
            TableCellWidth tableCellWidth3 = new TableCellWidth() { Width = "4453", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge2 = new VerticalMerge() { Val = MergedCellValues.Restart };
            Shading shading3 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "D6E3BC" };
            TableCellVerticalAlignment tableCellVerticalAlignment3 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties3.Append(tableCellWidth3);
            tableCellProperties3.Append(verticalMerge2);
            tableCellProperties3.Append(shading3);
            tableCellProperties3.Append(tableCellVerticalAlignment3);

            Paragraph paragraph3 = new Paragraph() { RsidParagraphMarkRevision = "00C709B4", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
            Indentation indentation3 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification3 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
            Bold bold9 = new Bold();
            Kern kern9 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize9 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript9 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties3.Append(bold9);
            paragraphMarkRunProperties3.Append(kern9);
            paragraphMarkRunProperties3.Append(fontSize9);
            paragraphMarkRunProperties3.Append(fontSizeComplexScript9);

            paragraphProperties3.Append(indentation3);
            paragraphProperties3.Append(justification3);
            paragraphProperties3.Append(paragraphMarkRunProperties3);

            Run run7 = new Run() { RsidRunProperties = "00C709B4" };

            RunProperties runProperties7 = new RunProperties();
            RunFonts runFonts4 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold10 = new Bold();
            Kern kern10 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize10 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript10 = new FontSizeComplexScript() { Val = "20" };

            runProperties7.Append(runFonts4);
            runProperties7.Append(bold10);
            runProperties7.Append(kern10);
            runProperties7.Append(fontSize10);
            runProperties7.Append(fontSizeComplexScript10);
            Text text7 = new Text();
            text7.Text = "因子表现";

            run7.Append(runProperties7);
            run7.Append(text7);

            paragraph3.Append(paragraphProperties3);
            paragraph3.Append(run7);

            tableCell3.Append(tableCellProperties3);
            tableCell3.Append(paragraph3);

            TableCell tableCell4 = new TableCell();

            TableCellProperties tableCellProperties4 = new TableCellProperties();
            TableCellWidth tableCellWidth4 = new TableCellWidth() { Width = "3850", Type = TableWidthUnitValues.Dxa };
            GridSpan gridSpan2 = new GridSpan() { Val = 5 };
            Shading shading4 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "D6E3BC" };
            TableCellVerticalAlignment tableCellVerticalAlignment4 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties4.Append(tableCellWidth4);
            tableCellProperties4.Append(gridSpan2);
            tableCellProperties4.Append(shading4);
            tableCellProperties4.Append(tableCellVerticalAlignment4);

            Paragraph paragraph4 = new Paragraph() { RsidParagraphMarkRevision = "00C709B4", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties4 = new ParagraphProperties();
            Indentation indentation4 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification4 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties4 = new ParagraphMarkRunProperties();
            Bold bold11 = new Bold();
            Kern kern11 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize11 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript11 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties4.Append(bold11);
            paragraphMarkRunProperties4.Append(kern11);
            paragraphMarkRunProperties4.Append(fontSize11);
            paragraphMarkRunProperties4.Append(fontSizeComplexScript11);

            paragraphProperties4.Append(indentation4);
            paragraphProperties4.Append(justification4);
            paragraphProperties4.Append(paragraphMarkRunProperties4);

            Run run8 = new Run() { RsidRunProperties = "00C709B4" };

            RunProperties runProperties8 = new RunProperties();
            RunFonts runFonts5 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold12 = new Bold();
            Kern kern12 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize12 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript12 = new FontSizeComplexScript() { Val = "20" };

            runProperties8.Append(runFonts5);
            runProperties8.Append(bold12);
            runProperties8.Append(kern12);
            runProperties8.Append(fontSize12);
            runProperties8.Append(fontSizeComplexScript12);
            Text text8 = new Text();
            text8.Text = "程度自评";

            run8.Append(runProperties8);
            run8.Append(text8);

            paragraph4.Append(paragraphProperties4);
            paragraph4.Append(run8);

            tableCell4.Append(tableCellProperties4);
            tableCell4.Append(paragraph4);

            tableRow2.Append(tableRowProperties2);
            tableRow2.Append(tableCell2);
            tableRow2.Append(tableCell3);
            tableRow2.Append(tableCell4);

            TableRow tableRow3 = new TableRow() { RsidTableRowMarkRevision = "00C709B4", RsidTableRowAddition = "00D061CC", RsidTableRowProperties = "009E6323" };

            TableRowProperties tableRowProperties3 = new TableRowProperties();
            TableRowHeight tableRowHeight3 = new TableRowHeight() { Val = (UInt32Value)432U };

            tableRowProperties3.Append(tableRowHeight3);

            TableCell tableCell5 = new TableCell();

            TableCellProperties tableCellProperties5 = new TableCellProperties();
            TableCellWidth tableCellWidth5 = new TableCellWidth() { Width = "1678", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge3 = new VerticalMerge();
            Shading shading5 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "D6E3BC" };
            TableCellVerticalAlignment tableCellVerticalAlignment5 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties5.Append(tableCellWidth5);
            tableCellProperties5.Append(verticalMerge3);
            tableCellProperties5.Append(shading5);
            tableCellProperties5.Append(tableCellVerticalAlignment5);

            Paragraph paragraph5 = new Paragraph() { RsidParagraphMarkRevision = "00C709B4", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties5 = new ParagraphProperties();
            Indentation indentation5 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification5 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties5 = new ParagraphMarkRunProperties();
            Bold bold13 = new Bold();
            Kern kern13 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize13 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript13 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties5.Append(bold13);
            paragraphMarkRunProperties5.Append(kern13);
            paragraphMarkRunProperties5.Append(fontSize13);
            paragraphMarkRunProperties5.Append(fontSizeComplexScript13);

            paragraphProperties5.Append(indentation5);
            paragraphProperties5.Append(justification5);
            paragraphProperties5.Append(paragraphMarkRunProperties5);

            paragraph5.Append(paragraphProperties5);

            tableCell5.Append(tableCellProperties5);
            tableCell5.Append(paragraph5);

            TableCell tableCell6 = new TableCell();

            TableCellProperties tableCellProperties6 = new TableCellProperties();
            TableCellWidth tableCellWidth6 = new TableCellWidth() { Width = "4453", Type = TableWidthUnitValues.Dxa };
            VerticalMerge verticalMerge4 = new VerticalMerge();
            Shading shading6 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "D6E3BC" };
            TableCellVerticalAlignment tableCellVerticalAlignment6 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties6.Append(tableCellWidth6);
            tableCellProperties6.Append(verticalMerge4);
            tableCellProperties6.Append(shading6);
            tableCellProperties6.Append(tableCellVerticalAlignment6);

            Paragraph paragraph6 = new Paragraph() { RsidParagraphMarkRevision = "00C709B4", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties6 = new ParagraphProperties();
            Indentation indentation6 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification6 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties6 = new ParagraphMarkRunProperties();
            Bold bold14 = new Bold();
            Kern kern14 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize14 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript14 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties6.Append(bold14);
            paragraphMarkRunProperties6.Append(kern14);
            paragraphMarkRunProperties6.Append(fontSize14);
            paragraphMarkRunProperties6.Append(fontSizeComplexScript14);

            paragraphProperties6.Append(indentation6);
            paragraphProperties6.Append(justification6);
            paragraphProperties6.Append(paragraphMarkRunProperties6);

            paragraph6.Append(paragraphProperties6);

            tableCell6.Append(tableCellProperties6);
            tableCell6.Append(paragraph6);

            TableCell tableCell7 = new TableCell();

            TableCellProperties tableCellProperties7 = new TableCellProperties();
            TableCellWidth tableCellWidth7 = new TableCellWidth() { Width = "756", Type = TableWidthUnitValues.Dxa };
            Shading shading7 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "D6E3BC" };
            TableCellVerticalAlignment tableCellVerticalAlignment7 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties7.Append(tableCellWidth7);
            tableCellProperties7.Append(shading7);
            tableCellProperties7.Append(tableCellVerticalAlignment7);

            Paragraph paragraph7 = new Paragraph() { RsidParagraphMarkRevision = "00C709B4", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties7 = new ParagraphProperties();
            Indentation indentation7 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification7 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties7 = new ParagraphMarkRunProperties();
            Bold bold15 = new Bold();
            Kern kern15 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize15 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript15 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties7.Append(bold15);
            paragraphMarkRunProperties7.Append(kern15);
            paragraphMarkRunProperties7.Append(fontSize15);
            paragraphMarkRunProperties7.Append(fontSizeComplexScript15);

            paragraphProperties7.Append(indentation7);
            paragraphProperties7.Append(justification7);
            paragraphProperties7.Append(paragraphMarkRunProperties7);

            Run run9 = new Run() { RsidRunProperties = "00C709B4" };

            RunProperties runProperties9 = new RunProperties();
            RunFonts runFonts6 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold16 = new Bold();
            Kern kern16 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize16 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript16 = new FontSizeComplexScript() { Val = "20" };

            runProperties9.Append(runFonts6);
            runProperties9.Append(bold16);
            runProperties9.Append(kern16);
            runProperties9.Append(fontSize16);
            runProperties9.Append(fontSizeComplexScript16);
            Text text9 = new Text();
            text9.Text = "没有";

            run9.Append(runProperties9);
            run9.Append(text9);

            paragraph7.Append(paragraphProperties7);
            paragraph7.Append(run9);

            tableCell7.Append(tableCellProperties7);
            tableCell7.Append(paragraph7);

            TableCell tableCell8 = new TableCell();

            TableCellProperties tableCellProperties8 = new TableCellProperties();
            TableCellWidth tableCellWidth8 = new TableCellWidth() { Width = "755", Type = TableWidthUnitValues.Dxa };
            Shading shading8 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "D6E3BC" };
            TableCellVerticalAlignment tableCellVerticalAlignment8 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties8.Append(tableCellWidth8);
            tableCellProperties8.Append(shading8);
            tableCellProperties8.Append(tableCellVerticalAlignment8);

            Paragraph paragraph8 = new Paragraph() { RsidParagraphMarkRevision = "00C709B4", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties8 = new ParagraphProperties();
            Indentation indentation8 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification8 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties8 = new ParagraphMarkRunProperties();
            Bold bold17 = new Bold();
            Kern kern17 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize17 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript17 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties8.Append(bold17);
            paragraphMarkRunProperties8.Append(kern17);
            paragraphMarkRunProperties8.Append(fontSize17);
            paragraphMarkRunProperties8.Append(fontSizeComplexScript17);

            paragraphProperties8.Append(indentation8);
            paragraphProperties8.Append(justification8);
            paragraphProperties8.Append(paragraphMarkRunProperties8);

            Run run10 = new Run() { RsidRunProperties = "00C709B4" };

            RunProperties runProperties10 = new RunProperties();
            RunFonts runFonts7 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold18 = new Bold();
            Kern kern18 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize18 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript18 = new FontSizeComplexScript() { Val = "20" };

            runProperties10.Append(runFonts7);
            runProperties10.Append(bold18);
            runProperties10.Append(kern18);
            runProperties10.Append(fontSize18);
            runProperties10.Append(fontSizeComplexScript18);
            Text text10 = new Text();
            text10.Text = "很轻";

            run10.Append(runProperties10);
            run10.Append(text10);

            paragraph8.Append(paragraphProperties8);
            paragraph8.Append(run10);

            tableCell8.Append(tableCellProperties8);
            tableCell8.Append(paragraph8);

            TableCell tableCell9 = new TableCell();

            TableCellProperties tableCellProperties9 = new TableCellProperties();
            TableCellWidth tableCellWidth9 = new TableCellWidth() { Width = "756", Type = TableWidthUnitValues.Dxa };
            Shading shading9 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "D6E3BC" };
            TableCellVerticalAlignment tableCellVerticalAlignment9 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties9.Append(tableCellWidth9);
            tableCellProperties9.Append(shading9);
            tableCellProperties9.Append(tableCellVerticalAlignment9);

            Paragraph paragraph9 = new Paragraph() { RsidParagraphMarkRevision = "00C709B4", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties9 = new ParagraphProperties();
            Indentation indentation9 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification9 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties9 = new ParagraphMarkRunProperties();
            Bold bold19 = new Bold();
            Kern kern19 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize19 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript19 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties9.Append(bold19);
            paragraphMarkRunProperties9.Append(kern19);
            paragraphMarkRunProperties9.Append(fontSize19);
            paragraphMarkRunProperties9.Append(fontSizeComplexScript19);

            paragraphProperties9.Append(indentation9);
            paragraphProperties9.Append(justification9);
            paragraphProperties9.Append(paragraphMarkRunProperties9);

            Run run11 = new Run() { RsidRunProperties = "00C709B4" };

            RunProperties runProperties11 = new RunProperties();
            RunFonts runFonts8 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold20 = new Bold();
            Kern kern20 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize20 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript20 = new FontSizeComplexScript() { Val = "20" };

            runProperties11.Append(runFonts8);
            runProperties11.Append(bold20);
            runProperties11.Append(kern20);
            runProperties11.Append(fontSize20);
            runProperties11.Append(fontSizeComplexScript20);
            Text text11 = new Text();
            text11.Text = "中度";

            run11.Append(runProperties11);
            run11.Append(text11);

            paragraph9.Append(paragraphProperties9);
            paragraph9.Append(run11);

            tableCell9.Append(tableCellProperties9);
            tableCell9.Append(paragraph9);

            TableCell tableCell10 = new TableCell();

            TableCellProperties tableCellProperties10 = new TableCellProperties();
            TableCellWidth tableCellWidth10 = new TableCellWidth() { Width = "756", Type = TableWidthUnitValues.Dxa };
            Shading shading10 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "D6E3BC" };
            TableCellVerticalAlignment tableCellVerticalAlignment10 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties10.Append(tableCellWidth10);
            tableCellProperties10.Append(shading10);
            tableCellProperties10.Append(tableCellVerticalAlignment10);

            Paragraph paragraph10 = new Paragraph() { RsidParagraphMarkRevision = "00C709B4", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties10 = new ParagraphProperties();
            Indentation indentation10 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification10 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties10 = new ParagraphMarkRunProperties();
            Bold bold21 = new Bold();
            Kern kern21 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize21 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript21 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties10.Append(bold21);
            paragraphMarkRunProperties10.Append(kern21);
            paragraphMarkRunProperties10.Append(fontSize21);
            paragraphMarkRunProperties10.Append(fontSizeComplexScript21);

            paragraphProperties10.Append(indentation10);
            paragraphProperties10.Append(justification10);
            paragraphProperties10.Append(paragraphMarkRunProperties10);

            Run run12 = new Run() { RsidRunProperties = "00C709B4" };

            RunProperties runProperties12 = new RunProperties();
            RunFonts runFonts9 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold22 = new Bold();
            Kern kern22 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize22 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript22 = new FontSizeComplexScript() { Val = "20" };

            runProperties12.Append(runFonts9);
            runProperties12.Append(bold22);
            runProperties12.Append(kern22);
            runProperties12.Append(fontSize22);
            runProperties12.Append(fontSizeComplexScript22);
            Text text12 = new Text();
            text12.Text = "较重";

            run12.Append(runProperties12);
            run12.Append(text12);

            paragraph10.Append(paragraphProperties10);
            paragraph10.Append(run12);

            tableCell10.Append(tableCellProperties10);
            tableCell10.Append(paragraph10);

            TableCell tableCell11 = new TableCell();

            TableCellProperties tableCellProperties11 = new TableCellProperties();
            TableCellWidth tableCellWidth11 = new TableCellWidth() { Width = "827", Type = TableWidthUnitValues.Dxa };
            Shading shading11 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "D6E3BC" };
            TableCellVerticalAlignment tableCellVerticalAlignment11 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties11.Append(tableCellWidth11);
            tableCellProperties11.Append(shading11);
            tableCellProperties11.Append(tableCellVerticalAlignment11);

            Paragraph paragraph11 = new Paragraph() { RsidParagraphMarkRevision = "00C709B4", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties11 = new ParagraphProperties();
            Indentation indentation11 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification11 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties11 = new ParagraphMarkRunProperties();
            Bold bold23 = new Bold();
            Kern kern23 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize23 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript23 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties11.Append(bold23);
            paragraphMarkRunProperties11.Append(kern23);
            paragraphMarkRunProperties11.Append(fontSize23);
            paragraphMarkRunProperties11.Append(fontSizeComplexScript23);

            paragraphProperties11.Append(indentation11);
            paragraphProperties11.Append(justification11);
            paragraphProperties11.Append(paragraphMarkRunProperties11);

            Run run13 = new Run() { RsidRunProperties = "00C709B4" };

            RunProperties runProperties13 = new RunProperties();
            RunFonts runFonts10 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold24 = new Bold();
            Kern kern24 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize24 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript24 = new FontSizeComplexScript() { Val = "20" };

            runProperties13.Append(runFonts10);
            runProperties13.Append(bold24);
            runProperties13.Append(kern24);
            runProperties13.Append(fontSize24);
            runProperties13.Append(fontSizeComplexScript24);
            Text text13 = new Text();
            text13.Text = "严重";

            run13.Append(runProperties13);
            run13.Append(text13);

            paragraph11.Append(paragraphProperties11);
            paragraph11.Append(run13);

            tableCell11.Append(tableCellProperties11);
            tableCell11.Append(paragraph11);

            tableRow3.Append(tableRowProperties3);
            tableRow3.Append(tableCell5);
            tableRow3.Append(tableCell6);
            tableRow3.Append(tableCell7);
            tableRow3.Append(tableCell8);
            tableRow3.Append(tableCell9);
            tableRow3.Append(tableCell10);
            tableRow3.Append(tableCell11);

            TableRow tableRow4 = new TableRow() { RsidTableRowMarkRevision = "00C709B4", RsidTableRowAddition = "00D061CC", RsidTableRowProperties = "00D061CC" };

            TableRowProperties tableRowProperties4 = new TableRowProperties();
            TableRowHeight tableRowHeight4 = new TableRowHeight() { Val = (UInt32Value)870U };

            tableRowProperties4.Append(tableRowHeight4);

            TableCell tableCell12 = new TableCell();

            TableCellProperties tableCellProperties12 = new TableCellProperties();
            TableCellWidth tableCellWidth12 = new TableCellWidth() { Width = "1678", Type = TableWidthUnitValues.Dxa };
            Shading shading12 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "D6E3BC" };
            TableCellVerticalAlignment tableCellVerticalAlignment12 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties12.Append(tableCellWidth12);
            tableCellProperties12.Append(shading12);
            tableCellProperties12.Append(tableCellVerticalAlignment12);

            Paragraph paragraph12 = new Paragraph() { RsidParagraphMarkRevision = "00C709B4", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties12 = new ParagraphProperties();
            Indentation indentation12 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification12 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties12 = new ParagraphMarkRunProperties();
            Bold bold25 = new Bold();
            Kern kern25 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize25 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript25 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties12.Append(bold25);
            paragraphMarkRunProperties12.Append(kern25);
            paragraphMarkRunProperties12.Append(fontSize25);
            paragraphMarkRunProperties12.Append(fontSizeComplexScript25);

            paragraphProperties12.Append(indentation12);
            paragraphProperties12.Append(justification12);
            paragraphProperties12.Append(paragraphMarkRunProperties12);

            Run run14 = new Run() { RsidRunProperties = "00C709B4" };

            RunProperties runProperties14 = new RunProperties();
            RunFonts runFonts11 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold26 = new Bold();
            Kern kern26 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize26 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript26 = new FontSizeComplexScript() { Val = "20" };

            runProperties14.Append(runFonts11);
            runProperties14.Append(bold26);
            runProperties14.Append(kern26);
            runProperties14.Append(fontSize26);
            runProperties14.Append(fontSizeComplexScript26);
            Text text14 = new Text();
            text14.Text = "强迫症状";

            run14.Append(runProperties14);
            run14.Append(text14);

            paragraph12.Append(paragraphProperties12);
            paragraph12.Append(run14);

            tableCell12.Append(tableCellProperties12);
            tableCell12.Append(paragraph12);

            TableCell tableCell13 = new TableCell();

            TableCellProperties tableCellProperties13 = new TableCellProperties();
            TableCellWidth tableCellWidth13 = new TableCellWidth() { Width = "4453", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment13 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties13.Append(tableCellWidth13);
            tableCellProperties13.Append(tableCellVerticalAlignment13);

            Paragraph paragraph13 = new Paragraph() { RsidParagraphMarkRevision = "00C709B4", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties13 = new ParagraphProperties();
            Indentation indentation13 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };

            ParagraphMarkRunProperties paragraphMarkRunProperties13 = new ParagraphMarkRunProperties();
            Kern kern27 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize27 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript27 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties13.Append(kern27);
            paragraphMarkRunProperties13.Append(fontSize27);
            paragraphMarkRunProperties13.Append(fontSizeComplexScript27);

            paragraphProperties13.Append(indentation13);
            paragraphProperties13.Append(paragraphMarkRunProperties13);

            Run run15 = new Run();

            RunProperties runProperties15 = new RunProperties();
            RunFonts runFonts12 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Kern kern28 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize28 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript28 = new FontSizeComplexScript() { Val = "20" };

            runProperties15.Append(runFonts12);
            runProperties15.Append(kern28);
            runProperties15.Append(fontSize28);
            runProperties15.Append(fontSizeComplexScript28);
            Text text15 = new Text();
            text15.Text = "做作业反复检查，反复数数";

            run15.Append(runProperties15);
            run15.Append(text15);

            Run run16 = new Run() { RsidRunProperties = "00C709B4" };

            RunProperties runProperties16 = new RunProperties();
            RunFonts runFonts13 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Kern kern29 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize29 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript29 = new FontSizeComplexScript() { Val = "20" };

            runProperties16.Append(runFonts13);
            runProperties16.Append(kern29);
            runProperties16.Append(fontSize29);
            runProperties16.Append(fontSizeComplexScript29);
            Text text16 = new Text();
            text16.Text = "，总是想一些不必要的事情，害怕考试成绩不理想等";

            run16.Append(runProperties16);
            run16.Append(text16);

            paragraph13.Append(paragraphProperties13);
            paragraph13.Append(run15);
            paragraph13.Append(run16);

            tableCell13.Append(tableCellProperties13);
            tableCell13.Append(paragraph13);

            TableCell tableCell14 = new TableCell();

            TableCellProperties tableCellProperties14 = new TableCellProperties();
            TableCellWidth tableCellWidth14 = new TableCellWidth() { Width = "756", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment14 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties14.Append(tableCellWidth14);
            tableCellProperties14.Append(tableCellVerticalAlignment14);

            Paragraph paragraph14 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties14 = new ParagraphProperties();
            Indentation indentation14 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification13 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties14 = new ParagraphMarkRunProperties();
            FontSize fontSize30 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript30 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties14.Append(fontSize30);
            paragraphMarkRunProperties14.Append(fontSizeComplexScript30);

            paragraphProperties14.Append(indentation14);
            paragraphProperties14.Append(justification13);
            paragraphProperties14.Append(paragraphMarkRunProperties14);
            paragraph14.Append(paragraphProperties14);


            Paragraph paragraph15 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties15 = new ParagraphProperties();
            Indentation indentation15 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification14 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties15 = new ParagraphMarkRunProperties();
            Kern kern53 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize54 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript54 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties15.Append(kern53);
            paragraphMarkRunProperties15.Append(fontSize54);
            paragraphMarkRunProperties15.Append(fontSizeComplexScript54);

            paragraphProperties15.Append(indentation15);
            paragraphProperties15.Append(justification14);
            paragraphProperties15.Append(paragraphMarkRunProperties15);

            Run run40 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties40 = new RunProperties();
            Kern kern54 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize55 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript55 = new FontSizeComplexScript() { Val = "20" };

            runProperties40.Append(kern54);
            runProperties40.Append(fontSize55);
            runProperties40.Append(fontSizeComplexScript55);
            Text text17 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text17.Text = content["强迫症状"].Item2 == "没有" ? "★" : " ";

            run40.Append(runProperties40);
            run40.Append(text17);

            paragraph15.Append(paragraphProperties15);
            paragraph15.Append(run40);

            Paragraph paragraph16 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties16 = new ParagraphProperties();
            Indentation indentation16 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification15 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties16 = new ParagraphMarkRunProperties();
            Kern kern55 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize56 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript56 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties16.Append(kern55);
            paragraphMarkRunProperties16.Append(fontSize56);
            paragraphMarkRunProperties16.Append(fontSizeComplexScript56);

            paragraphProperties16.Append(indentation16);
            paragraphProperties16.Append(justification15);
            paragraphProperties16.Append(paragraphMarkRunProperties16);

            Run run41 = new Run();

            RunProperties runProperties41 = new RunProperties();
            Kern kern56 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize57 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript57 = new FontSizeComplexScript() { Val = "20" };

            runProperties41.Append(kern56);
            runProperties41.Append(fontSize57);
            runProperties41.Append(fontSizeComplexScript57);
            FieldChar fieldChar3 = new FieldChar() { FieldCharType = FieldCharValues.End };

            run41.Append(runProperties41);
            run41.Append(fieldChar3);

            paragraph16.Append(paragraphProperties16);
            paragraph16.Append(run41);

            tableCell14.Append(tableCellProperties14);
            tableCell14.Append(paragraph14);
            tableCell14.Append(paragraph15);
            tableCell14.Append(paragraph16);

            TableCell tableCell15 = new TableCell();

            TableCellProperties tableCellProperties15 = new TableCellProperties();
            TableCellWidth tableCellWidth15 = new TableCellWidth() { Width = "755", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment15 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties15.Append(tableCellWidth15);
            tableCellProperties15.Append(tableCellVerticalAlignment15);

            Paragraph paragraph17 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties17 = new ParagraphProperties();
            Indentation indentation17 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification16 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties17 = new ParagraphMarkRunProperties();
            FontSize fontSize58 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript58 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties17.Append(fontSize58);
            paragraphMarkRunProperties17.Append(fontSizeComplexScript58);

            paragraphProperties17.Append(indentation17);
            paragraphProperties17.Append(justification16);
            paragraphProperties17.Append(paragraphMarkRunProperties17);


            paragraph17.Append(paragraphProperties17);

            Paragraph paragraph18 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties18 = new ParagraphProperties();
            Indentation indentation18 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification17 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties18 = new ParagraphMarkRunProperties();
            Kern kern80 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize82 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript82 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties18.Append(kern80);
            paragraphMarkRunProperties18.Append(fontSize82);
            paragraphMarkRunProperties18.Append(fontSizeComplexScript82);

            paragraphProperties18.Append(indentation18);
            paragraphProperties18.Append(justification17);
            paragraphProperties18.Append(paragraphMarkRunProperties18);

            Run run65 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties65 = new RunProperties();
            RunFonts runFonts52 = new RunFonts() { Ascii = "Segoe UI Symbol", HighAnsi = "Segoe UI Symbol", ComplexScript = "Segoe UI Symbol" };
            Kern kern81 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize83 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript83 = new FontSizeComplexScript() { Val = "20" };

            runProperties65.Append(runFonts52);
            runProperties65.Append(kern81);
            runProperties65.Append(fontSize83);
            runProperties65.Append(fontSizeComplexScript83);
            Text text18 = new Text();
            text18.Text = content["强迫症状"].Item2 == "很轻" ? "★" : " ";

            run65.Append(runProperties65);
            run65.Append(text18);

            paragraph18.Append(paragraphProperties18);
            paragraph18.Append(run65);

            Paragraph paragraph19 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties19 = new ParagraphProperties();
            Indentation indentation19 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification18 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties19 = new ParagraphMarkRunProperties();
            Kern kern82 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize84 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript84 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties19.Append(kern82);
            paragraphMarkRunProperties19.Append(fontSize84);
            paragraphMarkRunProperties19.Append(fontSizeComplexScript84);

            paragraphProperties19.Append(indentation19);
            paragraphProperties19.Append(justification18);
            paragraphProperties19.Append(paragraphMarkRunProperties19);

            Run run66 = new Run();

            RunProperties runProperties66 = new RunProperties();
            Kern kern83 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize85 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript85 = new FontSizeComplexScript() { Val = "20" };

            runProperties66.Append(kern83);
            runProperties66.Append(fontSize85);
            runProperties66.Append(fontSizeComplexScript85);
            FieldChar fieldChar6 = new FieldChar() { FieldCharType = FieldCharValues.End };

            run66.Append(runProperties66);
            run66.Append(fieldChar6);

            paragraph19.Append(paragraphProperties19);
            paragraph19.Append(run66);

            tableCell15.Append(tableCellProperties15);
            tableCell15.Append(paragraph17);
            tableCell15.Append(paragraph18);
            tableCell15.Append(paragraph19);

            TableCell tableCell16 = new TableCell();

            TableCellProperties tableCellProperties16 = new TableCellProperties();
            TableCellWidth tableCellWidth16 = new TableCellWidth() { Width = "756", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment16 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties16.Append(tableCellWidth16);
            tableCellProperties16.Append(tableCellVerticalAlignment16);

            Paragraph paragraph20 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties20 = new ParagraphProperties();
            Indentation indentation20 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification19 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties20 = new ParagraphMarkRunProperties();
            FontSize fontSize86 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript86 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties20.Append(fontSize86);
            paragraphMarkRunProperties20.Append(fontSizeComplexScript86);

            paragraphProperties20.Append(indentation20);
            paragraphProperties20.Append(justification19);
            paragraphProperties20.Append(paragraphMarkRunProperties20);

            paragraph20.Append(paragraphProperties20);

            Paragraph paragraph21 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties21 = new ParagraphProperties();
            Indentation indentation21 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification20 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties21 = new ParagraphMarkRunProperties();
            Kern kern107 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize110 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript110 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties21.Append(kern107);
            paragraphMarkRunProperties21.Append(fontSize110);
            paragraphMarkRunProperties21.Append(fontSizeComplexScript110);

            paragraphProperties21.Append(indentation21);
            paragraphProperties21.Append(justification20);
            paragraphProperties21.Append(paragraphMarkRunProperties21);

            Run run90 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties90 = new RunProperties();
            RunFonts runFonts72 = new RunFonts() { Ascii = "Segoe UI Symbol", HighAnsi = "Segoe UI Symbol", ComplexScript = "Segoe UI Symbol" };
            Kern kern108 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize111 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript111 = new FontSizeComplexScript() { Val = "20" };

            runProperties90.Append(runFonts72);
            runProperties90.Append(kern108);
            runProperties90.Append(fontSize111);
            runProperties90.Append(fontSizeComplexScript111);
            Text text19 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text19.Text = content["强迫症状"].Item2 == "中度" ? "★" : " ";

            run90.Append(runProperties90);
            run90.Append(text19);

            paragraph21.Append(paragraphProperties21);
            paragraph21.Append(run90);

            Paragraph paragraph22 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties22 = new ParagraphProperties();
            Indentation indentation22 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification21 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties22 = new ParagraphMarkRunProperties();
            Kern kern109 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize112 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript112 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties22.Append(kern109);
            paragraphMarkRunProperties22.Append(fontSize112);
            paragraphMarkRunProperties22.Append(fontSizeComplexScript112);

            paragraphProperties22.Append(indentation22);
            paragraphProperties22.Append(justification21);
            paragraphProperties22.Append(paragraphMarkRunProperties22);

            Run run91 = new Run();

            RunProperties runProperties91 = new RunProperties();
            Kern kern110 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize113 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript113 = new FontSizeComplexScript() { Val = "20" };

            runProperties91.Append(kern110);
            runProperties91.Append(fontSize113);
            runProperties91.Append(fontSizeComplexScript113);
            FieldChar fieldChar9 = new FieldChar() { FieldCharType = FieldCharValues.End };

            run91.Append(runProperties91);
            run91.Append(fieldChar9);

            paragraph22.Append(paragraphProperties22);
            paragraph22.Append(run91);

            tableCell16.Append(tableCellProperties16);
            tableCell16.Append(paragraph20);
            tableCell16.Append(paragraph21);
            tableCell16.Append(paragraph22);

            TableCell tableCell17 = new TableCell();

            TableCellProperties tableCellProperties17 = new TableCellProperties();
            TableCellWidth tableCellWidth17 = new TableCellWidth() { Width = "756", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment17 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties17.Append(tableCellWidth17);
            tableCellProperties17.Append(tableCellVerticalAlignment17);

            Paragraph paragraph23 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties23 = new ParagraphProperties();
            Indentation indentation23 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification22 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties23 = new ParagraphMarkRunProperties();
            FontSize fontSize114 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript114 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties23.Append(fontSize114);
            paragraphMarkRunProperties23.Append(fontSizeComplexScript114);

            paragraphProperties23.Append(indentation23);
            paragraphProperties23.Append(justification22);
            paragraphProperties23.Append(paragraphMarkRunProperties23);
            paragraph23.Append(paragraphProperties23);

            Paragraph paragraph24 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties24 = new ParagraphProperties();
            Indentation indentation24 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification23 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties24 = new ParagraphMarkRunProperties();
            Kern kern134 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize138 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript138 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties24.Append(kern134);
            paragraphMarkRunProperties24.Append(fontSize138);
            paragraphMarkRunProperties24.Append(fontSizeComplexScript138);

            paragraphProperties24.Append(indentation24);
            paragraphProperties24.Append(justification23);
            paragraphProperties24.Append(paragraphMarkRunProperties24);

            Run run115 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties115 = new RunProperties();
            Kern kern135 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize139 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript139 = new FontSizeComplexScript() { Val = "20" };

            runProperties115.Append(kern135);
            runProperties115.Append(fontSize139);
            runProperties115.Append(fontSizeComplexScript139);
            Text text20 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text20.Text = content["强迫症状"].Item2 == "较重" ? "★" : " ";

            run115.Append(runProperties115);
            run115.Append(text20);

            paragraph24.Append(paragraphProperties24);
            paragraph24.Append(run115);

            Paragraph paragraph25 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties25 = new ParagraphProperties();
            Indentation indentation25 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification24 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties25 = new ParagraphMarkRunProperties();
            Kern kern136 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize140 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript140 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties25.Append(kern136);
            paragraphMarkRunProperties25.Append(fontSize140);
            paragraphMarkRunProperties25.Append(fontSizeComplexScript140);

            paragraphProperties25.Append(indentation25);
            paragraphProperties25.Append(justification24);
            paragraphProperties25.Append(paragraphMarkRunProperties25);

            paragraph25.Append(paragraphProperties25);

            tableCell17.Append(tableCellProperties17);
            tableCell17.Append(paragraph23);
            tableCell17.Append(paragraph24);
            tableCell17.Append(paragraph25);

            TableCell tableCell18 = new TableCell();

            TableCellProperties tableCellProperties18 = new TableCellProperties();
            TableCellWidth tableCellWidth18 = new TableCellWidth() { Width = "827", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment18 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties18.Append(tableCellWidth18);
            tableCellProperties18.Append(tableCellVerticalAlignment18);

            Paragraph paragraph26 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties26 = new ParagraphProperties();
            Indentation indentation26 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification25 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties26 = new ParagraphMarkRunProperties();
            FontSize fontSize142 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript142 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties26.Append(fontSize142);
            paragraphMarkRunProperties26.Append(fontSizeComplexScript142);

            paragraphProperties26.Append(indentation26);
            paragraphProperties26.Append(justification25);
            paragraphProperties26.Append(paragraphMarkRunProperties26);

            paragraph26.Append(paragraphProperties26);

            Paragraph paragraph27 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties27 = new ParagraphProperties();
            Indentation indentation27 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification26 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties27 = new ParagraphMarkRunProperties();
            Kern kern161 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize166 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript166 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties27.Append(kern161);
            paragraphMarkRunProperties27.Append(fontSize166);
            paragraphMarkRunProperties27.Append(fontSizeComplexScript166);

            paragraphProperties27.Append(indentation27);
            paragraphProperties27.Append(justification26);
            paragraphProperties27.Append(paragraphMarkRunProperties27);

            Run run140 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties140 = new RunProperties();
            Kern kern162 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize167 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript167 = new FontSizeComplexScript() { Val = "20" };

            runProperties140.Append(kern162);
            runProperties140.Append(fontSize167);
            runProperties140.Append(fontSizeComplexScript167);
            Text text21 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text21.Text = content["强迫症状"].Item2 == "严重" ? "★" : " ";

            run140.Append(runProperties140);
            run140.Append(text21);

            paragraph27.Append(paragraphProperties27);
            paragraph27.Append(run140);

            Paragraph paragraph28 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties28 = new ParagraphProperties();
            Indentation indentation28 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification27 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties28 = new ParagraphMarkRunProperties();
            Kern kern163 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize168 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript168 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties28.Append(kern163);
            paragraphMarkRunProperties28.Append(fontSize168);
            paragraphMarkRunProperties28.Append(fontSizeComplexScript168);

            paragraphProperties28.Append(indentation28);
            paragraphProperties28.Append(justification27);
            paragraphProperties28.Append(paragraphMarkRunProperties28);

            paragraph28.Append(paragraphProperties28);

            tableCell18.Append(tableCellProperties18);
            tableCell18.Append(paragraph26);
            tableCell18.Append(paragraph27);
            tableCell18.Append(paragraph28);

            tableRow4.Append(tableRowProperties4);
            tableRow4.Append(tableCell12);
            tableRow4.Append(tableCell13);
            tableRow4.Append(tableCell14);
            tableRow4.Append(tableCell15);
            tableRow4.Append(tableCell16);
            tableRow4.Append(tableCell17);
            tableRow4.Append(tableCell18);

            TableRow tableRow5 = new TableRow() { RsidTableRowMarkRevision = "00C709B4", RsidTableRowAddition = "00D061CC", RsidTableRowProperties = "009E6323" };

            TableRowProperties tableRowProperties5 = new TableRowProperties();
            TableRowHeight tableRowHeight5 = new TableRowHeight() { Val = (UInt32Value)765U };

            tableRowProperties5.Append(tableRowHeight5);

            TableCell tableCell19 = new TableCell();

            TableCellProperties tableCellProperties19 = new TableCellProperties();
            TableCellWidth tableCellWidth19 = new TableCellWidth() { Width = "1678", Type = TableWidthUnitValues.Dxa };
            Shading shading13 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "D6E3BC" };
            TableCellVerticalAlignment tableCellVerticalAlignment19 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties19.Append(tableCellWidth19);
            tableCellProperties19.Append(shading13);
            tableCellProperties19.Append(tableCellVerticalAlignment19);

            Paragraph paragraph29 = new Paragraph() { RsidParagraphMarkRevision = "00C709B4", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties29 = new ParagraphProperties();
            Indentation indentation29 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification28 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties29 = new ParagraphMarkRunProperties();
            Bold bold27 = new Bold();
            Kern kern165 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize170 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript170 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties29.Append(bold27);
            paragraphMarkRunProperties29.Append(kern165);
            paragraphMarkRunProperties29.Append(fontSize170);
            paragraphMarkRunProperties29.Append(fontSizeComplexScript170);

            paragraphProperties29.Append(indentation29);
            paragraphProperties29.Append(justification28);
            paragraphProperties29.Append(paragraphMarkRunProperties29);

            Run run142 = new Run() { RsidRunProperties = "00C709B4" };

            RunProperties runProperties142 = new RunProperties();
            RunFonts runFonts111 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold28 = new Bold();
            Kern kern166 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize171 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript171 = new FontSizeComplexScript() { Val = "20" };

            runProperties142.Append(runFonts111);
            runProperties142.Append(bold28);
            runProperties142.Append(kern166);
            runProperties142.Append(fontSize171);
            runProperties142.Append(fontSizeComplexScript171);
            Text text22 = new Text();
            text22.Text = "偏执";

            run142.Append(runProperties142);
            run142.Append(text22);

            paragraph29.Append(paragraphProperties29);
            paragraph29.Append(run142);

            tableCell19.Append(tableCellProperties19);
            tableCell19.Append(paragraph29);

            TableCell tableCell20 = new TableCell();

            TableCellProperties tableCellProperties20 = new TableCellProperties();
            TableCellWidth tableCellWidth20 = new TableCellWidth() { Width = "4453", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment20 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties20.Append(tableCellWidth20);
            tableCellProperties20.Append(tableCellVerticalAlignment20);

            Paragraph paragraph30 = new Paragraph() { RsidParagraphMarkRevision = "00C709B4", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties30 = new ParagraphProperties();
            Indentation indentation30 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };

            ParagraphMarkRunProperties paragraphMarkRunProperties30 = new ParagraphMarkRunProperties();
            Kern kern167 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize172 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript172 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties30.Append(kern167);
            paragraphMarkRunProperties30.Append(fontSize172);
            paragraphMarkRunProperties30.Append(fontSizeComplexScript172);

            paragraphProperties30.Append(indentation30);
            paragraphProperties30.Append(paragraphMarkRunProperties30);

            Run run143 = new Run() { RsidRunProperties = "00C709B4" };

            RunProperties runProperties143 = new RunProperties();
            RunFonts runFonts112 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Kern kern168 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize173 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript173 = new FontSizeComplexScript() { Val = "20" };

            runProperties143.Append(runFonts112);
            runProperties143.Append(kern168);
            runProperties143.Append(fontSize173);
            runProperties143.Append(fontSizeComplexScript173);
            Text text23 = new Text();
            text23.Text = "觉的别人占自己便宜，别人在背后议论自己，对多数人不信任，别人对自己的评价不适当，别人跟自己作对等";

            run143.Append(runProperties143);
            run143.Append(text23);

            paragraph30.Append(paragraphProperties30);
            paragraph30.Append(run143);

            tableCell20.Append(tableCellProperties20);
            tableCell20.Append(paragraph30);

            TableCell tableCell21 = new TableCell();

            TableCellProperties tableCellProperties21 = new TableCellProperties();
            TableCellWidth tableCellWidth21 = new TableCellWidth() { Width = "756", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment21 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties21.Append(tableCellWidth21);
            tableCellProperties21.Append(tableCellVerticalAlignment21);

            Paragraph paragraph31 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties31 = new ParagraphProperties();
            Indentation indentation31 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification29 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties31 = new ParagraphMarkRunProperties();
            FontSize fontSize174 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript174 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties31.Append(fontSize174);
            paragraphMarkRunProperties31.Append(fontSizeComplexScript174);

            paragraphProperties31.Append(indentation31);
            paragraphProperties31.Append(justification29);
            paragraphProperties31.Append(paragraphMarkRunProperties31);
            paragraph31.Append(paragraphProperties31);

            Paragraph paragraph32 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties32 = new ParagraphProperties();
            Indentation indentation32 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification30 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties32 = new ParagraphMarkRunProperties();
            Kern kern192 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize198 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript198 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties32.Append(kern192);
            paragraphMarkRunProperties32.Append(fontSize198);
            paragraphMarkRunProperties32.Append(fontSizeComplexScript198);

            paragraphProperties32.Append(indentation32);
            paragraphProperties32.Append(justification30);
            paragraphProperties32.Append(paragraphMarkRunProperties32);

            Run run167 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties167 = new RunProperties();
            Kern kern193 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize199 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript199 = new FontSizeComplexScript() { Val = "20" };

            runProperties167.Append(kern193);
            runProperties167.Append(fontSize199);
            runProperties167.Append(fontSizeComplexScript199);
            Text text24 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text24.Text = content["偏执"].Item2 == "没有" ? "★" : " ";

            run167.Append(runProperties167);
            run167.Append(text24);

            paragraph32.Append(paragraphProperties32);
            paragraph32.Append(run167);

            Paragraph paragraph33 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties33 = new ParagraphProperties();
            Indentation indentation33 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification31 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties33 = new ParagraphMarkRunProperties();
            Kern kern194 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize200 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript200 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties33.Append(kern194);
            paragraphMarkRunProperties33.Append(fontSize200);
            paragraphMarkRunProperties33.Append(fontSizeComplexScript200);

            paragraphProperties33.Append(indentation33);
            paragraphProperties33.Append(justification31);
            paragraphProperties33.Append(paragraphMarkRunProperties33);

            Run run168 = new Run();

            RunProperties runProperties168 = new RunProperties();
            Kern kern195 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize201 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript201 = new FontSizeComplexScript() { Val = "20" };

            runProperties168.Append(kern195);
            runProperties168.Append(fontSize201);
            runProperties168.Append(fontSizeComplexScript201);
            FieldChar fieldChar18 = new FieldChar() { FieldCharType = FieldCharValues.End };

            run168.Append(runProperties168);
            run168.Append(fieldChar18);

            paragraph33.Append(paragraphProperties33);
            paragraph33.Append(run168);

            tableCell21.Append(tableCellProperties21);
            tableCell21.Append(paragraph31);
            tableCell21.Append(paragraph32);
            tableCell21.Append(paragraph33);

            TableCell tableCell22 = new TableCell();

            TableCellProperties tableCellProperties22 = new TableCellProperties();
            TableCellWidth tableCellWidth22 = new TableCellWidth() { Width = "755", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment22 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties22.Append(tableCellWidth22);
            tableCellProperties22.Append(tableCellVerticalAlignment22);

            Paragraph paragraph34 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties34 = new ParagraphProperties();
            Indentation indentation34 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification32 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties34 = new ParagraphMarkRunProperties();
            FontSize fontSize202 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript202 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties34.Append(fontSize202);
            paragraphMarkRunProperties34.Append(fontSizeComplexScript202);

            paragraphProperties34.Append(indentation34);
            paragraphProperties34.Append(justification32);
            paragraphProperties34.Append(paragraphMarkRunProperties34);
            paragraph34.Append(paragraphProperties34);

            Paragraph paragraph35 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties35 = new ParagraphProperties();
            Indentation indentation35 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification33 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties35 = new ParagraphMarkRunProperties();
            Kern kern219 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize226 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript226 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties35.Append(kern219);
            paragraphMarkRunProperties35.Append(fontSize226);
            paragraphMarkRunProperties35.Append(fontSizeComplexScript226);

            paragraphProperties35.Append(indentation35);
            paragraphProperties35.Append(justification33);
            paragraphProperties35.Append(paragraphMarkRunProperties35);

            Run run192 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties192 = new RunProperties();
            RunFonts runFonts151 = new RunFonts() { Ascii = "Segoe UI Symbol", HighAnsi = "Segoe UI Symbol", ComplexScript = "Segoe UI Symbol" };
            Kern kern220 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize227 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript227 = new FontSizeComplexScript() { Val = "20" };

            runProperties192.Append(runFonts151);
            runProperties192.Append(kern220);
            runProperties192.Append(fontSize227);
            runProperties192.Append(fontSizeComplexScript227);
            Text text25 = new Text();
            text25.Text = content["偏执"].Item2 == "很轻" ? "★" : " ";

            run192.Append(runProperties192);
            run192.Append(text25);

            paragraph35.Append(paragraphProperties35);
            paragraph35.Append(run192);

            Paragraph paragraph36 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties36 = new ParagraphProperties();
            Indentation indentation36 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification34 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties36 = new ParagraphMarkRunProperties();
            Kern kern221 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize228 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript228 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties36.Append(kern221);
            paragraphMarkRunProperties36.Append(fontSize228);
            paragraphMarkRunProperties36.Append(fontSizeComplexScript228);

            paragraphProperties36.Append(indentation36);
            paragraphProperties36.Append(justification34);
            paragraphProperties36.Append(paragraphMarkRunProperties36);

            Run run193 = new Run();

            RunProperties runProperties193 = new RunProperties();
            Kern kern222 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize229 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript229 = new FontSizeComplexScript() { Val = "20" };

            runProperties193.Append(kern222);
            runProperties193.Append(fontSize229);
            runProperties193.Append(fontSizeComplexScript229);
            FieldChar fieldChar21 = new FieldChar() { FieldCharType = FieldCharValues.End };

            run193.Append(runProperties193);
            run193.Append(fieldChar21);

            paragraph36.Append(paragraphProperties36);
            paragraph36.Append(run193);

            tableCell22.Append(tableCellProperties22);
            tableCell22.Append(paragraph34);
            tableCell22.Append(paragraph35);
            tableCell22.Append(paragraph36);

            TableCell tableCell23 = new TableCell();

            TableCellProperties tableCellProperties23 = new TableCellProperties();
            TableCellWidth tableCellWidth23 = new TableCellWidth() { Width = "756", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment23 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties23.Append(tableCellWidth23);
            tableCellProperties23.Append(tableCellVerticalAlignment23);

            Paragraph paragraph37 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties37 = new ParagraphProperties();
            Indentation indentation37 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification35 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties37 = new ParagraphMarkRunProperties();
            FontSize fontSize230 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript230 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties37.Append(fontSize230);
            paragraphMarkRunProperties37.Append(fontSizeComplexScript230);

            paragraphProperties37.Append(indentation37);
            paragraphProperties37.Append(justification35);
            paragraphProperties37.Append(paragraphMarkRunProperties37);
            paragraph37.Append(paragraphProperties37);

            Paragraph paragraph38 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties38 = new ParagraphProperties();
            Indentation indentation38 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification36 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties38 = new ParagraphMarkRunProperties();
            Kern kern246 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize254 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript254 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties38.Append(kern246);
            paragraphMarkRunProperties38.Append(fontSize254);
            paragraphMarkRunProperties38.Append(fontSizeComplexScript254);

            paragraphProperties38.Append(indentation38);
            paragraphProperties38.Append(justification36);
            paragraphProperties38.Append(paragraphMarkRunProperties38);

            Run run217 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties217 = new RunProperties();
            Kern kern247 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize255 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript255 = new FontSizeComplexScript() { Val = "20" };

            runProperties217.Append(kern247);
            runProperties217.Append(fontSize255);
            runProperties217.Append(fontSizeComplexScript255);
            Text text26 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text26.Text = content["偏执"].Item2 == "中度" ? "★" : " "; ;

            run217.Append(runProperties217);
            run217.Append(text26);

            paragraph38.Append(paragraphProperties38);
            paragraph38.Append(run217);

            Paragraph paragraph39 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties39 = new ParagraphProperties();
            Indentation indentation39 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification37 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties39 = new ParagraphMarkRunProperties();
            Kern kern248 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize256 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript256 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties39.Append(kern248);
            paragraphMarkRunProperties39.Append(fontSize256);
            paragraphMarkRunProperties39.Append(fontSizeComplexScript256);

            paragraphProperties39.Append(indentation39);
            paragraphProperties39.Append(justification37);
            paragraphProperties39.Append(paragraphMarkRunProperties39);
            paragraph39.Append(paragraphProperties39);

            tableCell23.Append(tableCellProperties23);
            tableCell23.Append(paragraph37);
            tableCell23.Append(paragraph38);
            tableCell23.Append(paragraph39);

            TableCell tableCell24 = new TableCell();

            TableCellProperties tableCellProperties24 = new TableCellProperties();
            TableCellWidth tableCellWidth24 = new TableCellWidth() { Width = "756", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment24 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties24.Append(tableCellWidth24);
            tableCellProperties24.Append(tableCellVerticalAlignment24);

            Paragraph paragraph40 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties40 = new ParagraphProperties();
            Indentation indentation40 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification38 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties40 = new ParagraphMarkRunProperties();
            FontSize fontSize258 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript258 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties40.Append(fontSize258);
            paragraphMarkRunProperties40.Append(fontSizeComplexScript258);

            paragraphProperties40.Append(indentation40);
            paragraphProperties40.Append(justification38);
            paragraphProperties40.Append(paragraphMarkRunProperties40);
            paragraph40.Append(paragraphProperties40);

            Paragraph paragraph41 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties41 = new ParagraphProperties();
            Indentation indentation41 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification39 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties41 = new ParagraphMarkRunProperties();
            Kern kern273 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize282 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript282 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties41.Append(kern273);
            paragraphMarkRunProperties41.Append(fontSize282);
            paragraphMarkRunProperties41.Append(fontSizeComplexScript282);

            paragraphProperties41.Append(indentation41);
            paragraphProperties41.Append(justification39);
            paragraphProperties41.Append(paragraphMarkRunProperties41);

            Run run242 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties242 = new RunProperties();
            Kern kern274 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize283 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript283 = new FontSizeComplexScript() { Val = "20" };

            runProperties242.Append(kern274);
            runProperties242.Append(fontSize283);
            runProperties242.Append(fontSizeComplexScript283);
            Text text27 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text27.Text = content["偏执"].Item2 == "较重" ? "★" : " ";

            run242.Append(runProperties242);
            run242.Append(text27);

            paragraph41.Append(paragraphProperties41);
            paragraph41.Append(run242);

            Paragraph paragraph42 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties42 = new ParagraphProperties();
            Indentation indentation42 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification40 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties42 = new ParagraphMarkRunProperties();
            Kern kern275 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize284 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript284 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties42.Append(kern275);
            paragraphMarkRunProperties42.Append(fontSize284);
            paragraphMarkRunProperties42.Append(fontSizeComplexScript284);

            paragraphProperties42.Append(indentation42);
            paragraphProperties42.Append(justification40);
            paragraphProperties42.Append(paragraphMarkRunProperties42);

            paragraph42.Append(paragraphProperties42);

            tableCell24.Append(tableCellProperties24);
            tableCell24.Append(paragraph40);
            tableCell24.Append(paragraph41);
            tableCell24.Append(paragraph42);

            TableCell tableCell25 = new TableCell();

            TableCellProperties tableCellProperties25 = new TableCellProperties();
            TableCellWidth tableCellWidth25 = new TableCellWidth() { Width = "827", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment25 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties25.Append(tableCellWidth25);
            tableCellProperties25.Append(tableCellVerticalAlignment25);

            Paragraph paragraph43 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties43 = new ParagraphProperties();
            Indentation indentation43 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification41 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties43 = new ParagraphMarkRunProperties();
            FontSize fontSize286 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript286 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties43.Append(fontSize286);
            paragraphMarkRunProperties43.Append(fontSizeComplexScript286);

            paragraphProperties43.Append(indentation43);
            paragraphProperties43.Append(justification41);
            paragraphProperties43.Append(paragraphMarkRunProperties43);

            paragraph43.Append(paragraphProperties43);

            Paragraph paragraph44 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties44 = new ParagraphProperties();
            Indentation indentation44 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification42 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties44 = new ParagraphMarkRunProperties();
            Kern kern300 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize310 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript310 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties44.Append(kern300);
            paragraphMarkRunProperties44.Append(fontSize310);
            paragraphMarkRunProperties44.Append(fontSizeComplexScript310);

            paragraphProperties44.Append(indentation44);
            paragraphProperties44.Append(justification42);
            paragraphProperties44.Append(paragraphMarkRunProperties44);

            Run run267 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties267 = new RunProperties();
            Kern kern301 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize311 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript311 = new FontSizeComplexScript() { Val = "20" };

            runProperties267.Append(kern301);
            runProperties267.Append(fontSize311);
            runProperties267.Append(fontSizeComplexScript311);
            Text text28 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text28.Text = content["偏执"].Item2 == "严重" ? "★" : " ";

            run267.Append(runProperties267);
            run267.Append(text28);

            paragraph44.Append(paragraphProperties44);
            paragraph44.Append(run267);

            Paragraph paragraph45 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties45 = new ParagraphProperties();
            Indentation indentation45 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification43 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties45 = new ParagraphMarkRunProperties();
            Kern kern302 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize312 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript312 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties45.Append(kern302);
            paragraphMarkRunProperties45.Append(fontSize312);
            paragraphMarkRunProperties45.Append(fontSizeComplexScript312);

            paragraphProperties45.Append(indentation45);
            paragraphProperties45.Append(justification43);
            paragraphProperties45.Append(paragraphMarkRunProperties45);

            paragraph45.Append(paragraphProperties45);

            tableCell25.Append(tableCellProperties25);
            tableCell25.Append(paragraph43);
            tableCell25.Append(paragraph44);
            tableCell25.Append(paragraph45);

            tableRow5.Append(tableRowProperties5);
            tableRow5.Append(tableCell19);
            tableRow5.Append(tableCell20);
            tableRow5.Append(tableCell21);
            tableRow5.Append(tableCell22);
            tableRow5.Append(tableCell23);
            tableRow5.Append(tableCell24);
            tableRow5.Append(tableCell25);

            TableRow tableRow6 = new TableRow() { RsidTableRowMarkRevision = "00C709B4", RsidTableRowAddition = "00D061CC", RsidTableRowProperties = "009E6323" };

            TableRowProperties tableRowProperties6 = new TableRowProperties();
            TableRowHeight tableRowHeight6 = new TableRowHeight() { Val = (UInt32Value)800U };

            tableRowProperties6.Append(tableRowHeight6);

            TableCell tableCell26 = new TableCell();

            TableCellProperties tableCellProperties26 = new TableCellProperties();
            TableCellWidth tableCellWidth26 = new TableCellWidth() { Width = "1678", Type = TableWidthUnitValues.Dxa };
            Shading shading14 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "D6E3BC" };
            TableCellVerticalAlignment tableCellVerticalAlignment26 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties26.Append(tableCellWidth26);
            tableCellProperties26.Append(shading14);
            tableCellProperties26.Append(tableCellVerticalAlignment26);

            Paragraph paragraph46 = new Paragraph() { RsidParagraphMarkRevision = "00C709B4", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties46 = new ParagraphProperties();
            Indentation indentation46 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification44 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties46 = new ParagraphMarkRunProperties();
            Bold bold29 = new Bold();
            Kern kern304 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize314 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript314 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties46.Append(bold29);
            paragraphMarkRunProperties46.Append(kern304);
            paragraphMarkRunProperties46.Append(fontSize314);
            paragraphMarkRunProperties46.Append(fontSizeComplexScript314);

            paragraphProperties46.Append(indentation46);
            paragraphProperties46.Append(justification44);
            paragraphProperties46.Append(paragraphMarkRunProperties46);

            Run run269 = new Run() { RsidRunProperties = "00C709B4" };

            RunProperties runProperties269 = new RunProperties();
            RunFonts runFonts209 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold30 = new Bold();
            Kern kern305 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize315 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript315 = new FontSizeComplexScript() { Val = "20" };

            runProperties269.Append(runFonts209);
            runProperties269.Append(bold30);
            runProperties269.Append(kern305);
            runProperties269.Append(fontSize315);
            runProperties269.Append(fontSizeComplexScript315);
            Text text29 = new Text();
            text29.Text = "敌对";

            run269.Append(runProperties269);
            run269.Append(text29);

            paragraph46.Append(paragraphProperties46);
            paragraph46.Append(run269);

            tableCell26.Append(tableCellProperties26);
            tableCell26.Append(paragraph46);

            TableCell tableCell27 = new TableCell();

            TableCellProperties tableCellProperties27 = new TableCellProperties();
            TableCellWidth tableCellWidth27 = new TableCellWidth() { Width = "4453", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment27 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties27.Append(tableCellWidth27);
            tableCellProperties27.Append(tableCellVerticalAlignment27);

            Paragraph paragraph47 = new Paragraph() { RsidParagraphMarkRevision = "00C709B4", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties47 = new ParagraphProperties();
            Indentation indentation47 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };

            ParagraphMarkRunProperties paragraphMarkRunProperties47 = new ParagraphMarkRunProperties();
            Kern kern306 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize316 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript316 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties47.Append(kern306);
            paragraphMarkRunProperties47.Append(fontSize316);
            paragraphMarkRunProperties47.Append(fontSizeComplexScript316);

            paragraphProperties47.Append(indentation47);
            paragraphProperties47.Append(paragraphMarkRunProperties47);

            Run run270 = new Run() { RsidRunProperties = "00C709B4" };

            RunProperties runProperties270 = new RunProperties();
            RunFonts runFonts210 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Kern kern307 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize317 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript317 = new FontSizeComplexScript() { Val = "20" };

            runProperties270.Append(runFonts210);
            runProperties270.Append(kern307);
            runProperties270.Append(fontSize317);
            runProperties270.Append(fontSizeComplexScript317);
            Text text30 = new Text();
            text30.Text = "控制不住自己的脾气，经常与别人争论，容易激动，有摔东西的冲动";

            run270.Append(runProperties270);
            run270.Append(text30);

            paragraph47.Append(paragraphProperties47);
            paragraph47.Append(run270);

            tableCell27.Append(tableCellProperties27);
            tableCell27.Append(paragraph47);

            TableCell tableCell28 = new TableCell();

            TableCellProperties tableCellProperties28 = new TableCellProperties();
            TableCellWidth tableCellWidth28 = new TableCellWidth() { Width = "756", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment28 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties28.Append(tableCellWidth28);
            tableCellProperties28.Append(tableCellVerticalAlignment28);

            Paragraph paragraph48 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties48 = new ParagraphProperties();
            Indentation indentation48 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification45 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties48 = new ParagraphMarkRunProperties();
            FontSize fontSize318 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript318 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties48.Append(fontSize318);
            paragraphMarkRunProperties48.Append(fontSizeComplexScript318);

            paragraphProperties48.Append(indentation48);
            paragraphProperties48.Append(justification45);
            paragraphProperties48.Append(paragraphMarkRunProperties48);

            paragraph48.Append(paragraphProperties48);

            Paragraph paragraph49 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties49 = new ParagraphProperties();
            Indentation indentation49 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification46 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties49 = new ParagraphMarkRunProperties();
            Kern kern331 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize342 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript342 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties49.Append(kern331);
            paragraphMarkRunProperties49.Append(fontSize342);
            paragraphMarkRunProperties49.Append(fontSizeComplexScript342);

            paragraphProperties49.Append(indentation49);
            paragraphProperties49.Append(justification46);
            paragraphProperties49.Append(paragraphMarkRunProperties49);

            Run run294 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties294 = new RunProperties();
            Kern kern332 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize343 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript343 = new FontSizeComplexScript() { Val = "20" };

            runProperties294.Append(kern332);
            runProperties294.Append(fontSize343);
            runProperties294.Append(fontSizeComplexScript343);
            Text text31 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text31.Text = content["敌对"].Item2 == "没有" ? "★" : " ";

            run294.Append(runProperties294);
            run294.Append(text31);

            paragraph49.Append(paragraphProperties49);
            paragraph49.Append(run294);

            Paragraph paragraph50 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties50 = new ParagraphProperties();
            Indentation indentation50 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification47 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties50 = new ParagraphMarkRunProperties();
            Kern kern333 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize344 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript344 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties50.Append(kern333);
            paragraphMarkRunProperties50.Append(fontSize344);
            paragraphMarkRunProperties50.Append(fontSizeComplexScript344);

            paragraphProperties50.Append(indentation50);
            paragraphProperties50.Append(justification47);
            paragraphProperties50.Append(paragraphMarkRunProperties50);

            paragraph50.Append(paragraphProperties50);

            tableCell28.Append(tableCellProperties28);
            tableCell28.Append(paragraph48);
            tableCell28.Append(paragraph49);
            tableCell28.Append(paragraph50);

            TableCell tableCell29 = new TableCell();

            TableCellProperties tableCellProperties29 = new TableCellProperties();
            TableCellWidth tableCellWidth29 = new TableCellWidth() { Width = "755", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment29 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties29.Append(tableCellWidth29);
            tableCellProperties29.Append(tableCellVerticalAlignment29);

            Paragraph paragraph51 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties51 = new ParagraphProperties();
            Indentation indentation51 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification48 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties51 = new ParagraphMarkRunProperties();
            FontSize fontSize346 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript346 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties51.Append(fontSize346);
            paragraphMarkRunProperties51.Append(fontSizeComplexScript346);

            paragraphProperties51.Append(indentation51);
            paragraphProperties51.Append(justification48);
            paragraphProperties51.Append(paragraphMarkRunProperties51);
            paragraph51.Append(paragraphProperties51);


            Paragraph paragraph52 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties52 = new ParagraphProperties();
            Indentation indentation52 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification49 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties52 = new ParagraphMarkRunProperties();
            Kern kern358 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize370 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript370 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties52.Append(kern358);
            paragraphMarkRunProperties52.Append(fontSize370);
            paragraphMarkRunProperties52.Append(fontSizeComplexScript370);

            paragraphProperties52.Append(indentation52);
            paragraphProperties52.Append(justification49);
            paragraphProperties52.Append(paragraphMarkRunProperties52);

            Run run319 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties319 = new RunProperties();
            RunFonts runFonts249 = new RunFonts() { Ascii = "Segoe UI Symbol", HighAnsi = "Segoe UI Symbol", ComplexScript = "Segoe UI Symbol" };
            Kern kern359 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize371 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript371 = new FontSizeComplexScript() { Val = "20" };

            runProperties319.Append(runFonts249);
            runProperties319.Append(kern359);
            runProperties319.Append(fontSize371);
            runProperties319.Append(fontSizeComplexScript371);
            Text text32 = new Text();
            text32.Text = content["敌对"].Item2 == "轻度" ? "★" : " ";

            run319.Append(runProperties319);
            run319.Append(text32);

            paragraph52.Append(paragraphProperties52);
            paragraph52.Append(run319);

            Paragraph paragraph53 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties53 = new ParagraphProperties();
            Indentation indentation53 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification50 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties53 = new ParagraphMarkRunProperties();
            Kern kern360 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize372 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript372 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties53.Append(kern360);
            paragraphMarkRunProperties53.Append(fontSize372);
            paragraphMarkRunProperties53.Append(fontSizeComplexScript372);

            paragraphProperties53.Append(indentation53);
            paragraphProperties53.Append(justification50);
            paragraphProperties53.Append(paragraphMarkRunProperties53);

            paragraph53.Append(paragraphProperties53);

            tableCell29.Append(tableCellProperties29);
            tableCell29.Append(paragraph51);
            tableCell29.Append(paragraph52);
            tableCell29.Append(paragraph53);

            TableCell tableCell30 = new TableCell();

            TableCellProperties tableCellProperties30 = new TableCellProperties();
            TableCellWidth tableCellWidth30 = new TableCellWidth() { Width = "756", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment30 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties30.Append(tableCellWidth30);
            tableCellProperties30.Append(tableCellVerticalAlignment30);

            Paragraph paragraph54 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties54 = new ParagraphProperties();
            Indentation indentation54 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification51 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties54 = new ParagraphMarkRunProperties();
            FontSize fontSize374 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript374 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties54.Append(fontSize374);
            paragraphMarkRunProperties54.Append(fontSizeComplexScript374);

            paragraphProperties54.Append(indentation54);
            paragraphProperties54.Append(justification51);
            paragraphProperties54.Append(paragraphMarkRunProperties54);

            paragraph54.Append(paragraphProperties54);

            Paragraph paragraph55 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties55 = new ParagraphProperties();
            Indentation indentation55 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification52 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties55 = new ParagraphMarkRunProperties();
            Kern kern385 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize398 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript398 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties55.Append(kern385);
            paragraphMarkRunProperties55.Append(fontSize398);
            paragraphMarkRunProperties55.Append(fontSizeComplexScript398);

            paragraphProperties55.Append(indentation55);
            paragraphProperties55.Append(justification52);
            paragraphProperties55.Append(paragraphMarkRunProperties55);

            Run run344 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties344 = new RunProperties();
            Kern kern386 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize399 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript399 = new FontSizeComplexScript() { Val = "20" };

            runProperties344.Append(kern386);
            runProperties344.Append(fontSize399);
            runProperties344.Append(fontSizeComplexScript399);
            Text text33 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text33.Text = content["敌对"].Item2 == "中度" ? "★" : " ";

            run344.Append(runProperties344);
            run344.Append(text33);

            paragraph55.Append(paragraphProperties55);
            paragraph55.Append(run344);

            Paragraph paragraph56 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties56 = new ParagraphProperties();
            Indentation indentation56 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification53 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties56 = new ParagraphMarkRunProperties();
            Kern kern387 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize400 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript400 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties56.Append(kern387);
            paragraphMarkRunProperties56.Append(fontSize400);
            paragraphMarkRunProperties56.Append(fontSizeComplexScript400);

            paragraphProperties56.Append(indentation56);
            paragraphProperties56.Append(justification53);
            paragraphProperties56.Append(paragraphMarkRunProperties56);
            paragraph56.Append(paragraphProperties56);

            tableCell30.Append(tableCellProperties30);
            tableCell30.Append(paragraph54);
            tableCell30.Append(paragraph55);
            tableCell30.Append(paragraph56);

            TableCell tableCell31 = new TableCell();

            TableCellProperties tableCellProperties31 = new TableCellProperties();
            TableCellWidth tableCellWidth31 = new TableCellWidth() { Width = "756", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment31 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties31.Append(tableCellWidth31);
            tableCellProperties31.Append(tableCellVerticalAlignment31);

            Paragraph paragraph57 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties57 = new ParagraphProperties();
            Indentation indentation57 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification54 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties57 = new ParagraphMarkRunProperties();
            FontSize fontSize402 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript402 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties57.Append(fontSize402);
            paragraphMarkRunProperties57.Append(fontSizeComplexScript402);

            paragraphProperties57.Append(indentation57);
            paragraphProperties57.Append(justification54);
            paragraphProperties57.Append(paragraphMarkRunProperties57);

            paragraph57.Append(paragraphProperties57);

            Paragraph paragraph58 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties58 = new ParagraphProperties();
            Indentation indentation58 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification55 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties58 = new ParagraphMarkRunProperties();
            Kern kern412 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize426 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript426 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties58.Append(kern412);
            paragraphMarkRunProperties58.Append(fontSize426);
            paragraphMarkRunProperties58.Append(fontSizeComplexScript426);

            paragraphProperties58.Append(indentation58);
            paragraphProperties58.Append(justification55);
            paragraphProperties58.Append(paragraphMarkRunProperties58);

            Run run369 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties369 = new RunProperties();
            Kern kern413 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize427 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript427 = new FontSizeComplexScript() { Val = "20" };

            runProperties369.Append(kern413);
            runProperties369.Append(fontSize427);
            runProperties369.Append(fontSizeComplexScript427);
            Text text34 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text34.Text = content["敌对"].Item2 == "较重" ? "★" : " ";

            run369.Append(runProperties369);
            run369.Append(text34);

            paragraph58.Append(paragraphProperties58);
            paragraph58.Append(run369);

            Paragraph paragraph59 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties59 = new ParagraphProperties();
            Indentation indentation59 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification56 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties59 = new ParagraphMarkRunProperties();
            Kern kern414 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize428 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript428 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties59.Append(kern414);
            paragraphMarkRunProperties59.Append(fontSize428);
            paragraphMarkRunProperties59.Append(fontSizeComplexScript428);

            paragraphProperties59.Append(indentation59);
            paragraphProperties59.Append(justification56);
            paragraphProperties59.Append(paragraphMarkRunProperties59);

            paragraph59.Append(paragraphProperties59);

            tableCell31.Append(tableCellProperties31);
            tableCell31.Append(paragraph57);
            tableCell31.Append(paragraph58);
            tableCell31.Append(paragraph59);

            TableCell tableCell32 = new TableCell();

            TableCellProperties tableCellProperties32 = new TableCellProperties();
            TableCellWidth tableCellWidth32 = new TableCellWidth() { Width = "827", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment32 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties32.Append(tableCellWidth32);
            tableCellProperties32.Append(tableCellVerticalAlignment32);

            Paragraph paragraph60 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties60 = new ParagraphProperties();
            Indentation indentation60 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification57 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties60 = new ParagraphMarkRunProperties();
            FontSize fontSize430 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript430 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties60.Append(fontSize430);
            paragraphMarkRunProperties60.Append(fontSizeComplexScript430);

            paragraphProperties60.Append(indentation60);
            paragraphProperties60.Append(justification57);
            paragraphProperties60.Append(paragraphMarkRunProperties60);

            paragraph60.Append(paragraphProperties60);

            Paragraph paragraph61 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties61 = new ParagraphProperties();
            Indentation indentation61 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification58 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties61 = new ParagraphMarkRunProperties();
            Kern kern439 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize454 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript454 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties61.Append(kern439);
            paragraphMarkRunProperties61.Append(fontSize454);
            paragraphMarkRunProperties61.Append(fontSizeComplexScript454);

            paragraphProperties61.Append(indentation61);
            paragraphProperties61.Append(justification58);
            paragraphProperties61.Append(paragraphMarkRunProperties61);

            Run run394 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties394 = new RunProperties();
            Kern kern440 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize455 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript455 = new FontSizeComplexScript() { Val = "20" };

            runProperties394.Append(kern440);
            runProperties394.Append(fontSize455);
            runProperties394.Append(fontSizeComplexScript455);
            Text text35 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text35.Text = content["敌对"].Item2 == "严重" ? "★" : " ";

            run394.Append(runProperties394);
            run394.Append(text35);

            paragraph61.Append(paragraphProperties61);
            paragraph61.Append(run394);

            Paragraph paragraph62 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties62 = new ParagraphProperties();
            Indentation indentation62 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification59 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties62 = new ParagraphMarkRunProperties();
            Kern kern441 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize456 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript456 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties62.Append(kern441);
            paragraphMarkRunProperties62.Append(fontSize456);
            paragraphMarkRunProperties62.Append(fontSizeComplexScript456);

            paragraphProperties62.Append(indentation62);
            paragraphProperties62.Append(justification59);
            paragraphProperties62.Append(paragraphMarkRunProperties62);

            paragraph62.Append(paragraphProperties62);

            tableCell32.Append(tableCellProperties32);
            tableCell32.Append(paragraph60);
            tableCell32.Append(paragraph61);
            tableCell32.Append(paragraph62);

            tableRow6.Append(tableRowProperties6);
            tableRow6.Append(tableCell26);
            tableRow6.Append(tableCell27);
            tableRow6.Append(tableCell28);
            tableRow6.Append(tableCell29);
            tableRow6.Append(tableCell30);
            tableRow6.Append(tableCell31);
            tableRow6.Append(tableCell32);

            TableRow tableRow7 = new TableRow() { RsidTableRowMarkRevision = "00C709B4", RsidTableRowAddition = "00D061CC", RsidTableRowProperties = "009E6323" };

            TableRowProperties tableRowProperties7 = new TableRowProperties();
            TableRowHeight tableRowHeight7 = new TableRowHeight() { Val = (UInt32Value)800U };

            tableRowProperties7.Append(tableRowHeight7);

            TableCell tableCell33 = new TableCell();

            TableCellProperties tableCellProperties33 = new TableCellProperties();
            TableCellWidth tableCellWidth33 = new TableCellWidth() { Width = "1678", Type = TableWidthUnitValues.Dxa };
            Shading shading15 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "D6E3BC" };
            TableCellVerticalAlignment tableCellVerticalAlignment33 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties33.Append(tableCellWidth33);
            tableCellProperties33.Append(shading15);
            tableCellProperties33.Append(tableCellVerticalAlignment33);

            Paragraph paragraph63 = new Paragraph() { RsidParagraphMarkRevision = "00C709B4", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties63 = new ParagraphProperties();
            Indentation indentation63 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification60 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties63 = new ParagraphMarkRunProperties();
            Bold bold31 = new Bold();
            Kern kern443 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize458 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript458 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties63.Append(bold31);
            paragraphMarkRunProperties63.Append(kern443);
            paragraphMarkRunProperties63.Append(fontSize458);
            paragraphMarkRunProperties63.Append(fontSizeComplexScript458);

            paragraphProperties63.Append(indentation63);
            paragraphProperties63.Append(justification60);
            paragraphProperties63.Append(paragraphMarkRunProperties63);

            Run run396 = new Run() { RsidRunProperties = "00C709B4" };

            RunProperties runProperties396 = new RunProperties();
            RunFonts runFonts307 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold32 = new Bold();
            Kern kern444 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize459 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript459 = new FontSizeComplexScript() { Val = "20" };

            runProperties396.Append(runFonts307);
            runProperties396.Append(bold32);
            runProperties396.Append(kern444);
            runProperties396.Append(fontSize459);
            runProperties396.Append(fontSizeComplexScript459);
            Text text36 = new Text();
            text36.Text = "人际关系紧张与敏感";

            run396.Append(runProperties396);
            run396.Append(text36);

            paragraph63.Append(paragraphProperties63);
            paragraph63.Append(run396);
            tableCell33.Append(tableCellProperties33);
            tableCell33.Append(paragraph63);

            TableCell tableCell34 = new TableCell();

            TableCellProperties tableCellProperties34 = new TableCellProperties();
            TableCellWidth tableCellWidth34 = new TableCellWidth() { Width = "4453", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment34 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties34.Append(tableCellWidth34);
            tableCellProperties34.Append(tableCellVerticalAlignment34);

            Paragraph paragraph65 = new Paragraph() { RsidParagraphMarkRevision = "00C709B4", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties65 = new ParagraphProperties();
            Indentation indentation65 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };

            ParagraphMarkRunProperties paragraphMarkRunProperties65 = new ParagraphMarkRunProperties();
            Kern kern447 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize462 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript462 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties65.Append(kern447);
            paragraphMarkRunProperties65.Append(fontSize462);
            paragraphMarkRunProperties65.Append(fontSizeComplexScript462);

            paragraphProperties65.Append(indentation65);
            paragraphProperties65.Append(paragraphMarkRunProperties65);

            Run run398 = new Run() { RsidRunProperties = "00C709B4" };

            RunProperties runProperties398 = new RunProperties();
            RunFonts runFonts309 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Kern kern448 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize463 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript463 = new FontSizeComplexScript() { Val = "20" };

            runProperties398.Append(runFonts309);
            runProperties398.Append(kern448);
            runProperties398.Append(fontSize463);
            runProperties398.Append(fontSizeComplexScript463);
            Text text38 = new Text();
            text38.Text = "总是认为别人不理解自己，别人对自己不友好，感情容易受到别人伤害，对别人求全责备，同异性在一起感到不自在";

            run398.Append(runProperties398);
            run398.Append(text38);

            paragraph65.Append(paragraphProperties65);
            paragraph65.Append(run398);

            tableCell34.Append(tableCellProperties34);
            tableCell34.Append(paragraph65);

            TableCell tableCell35 = new TableCell();

            TableCellProperties tableCellProperties35 = new TableCellProperties();
            TableCellWidth tableCellWidth35 = new TableCellWidth() { Width = "756", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment35 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties35.Append(tableCellWidth35);
            tableCellProperties35.Append(tableCellVerticalAlignment35);

            Paragraph paragraph66 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties66 = new ParagraphProperties();
            Indentation indentation66 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification62 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties66 = new ParagraphMarkRunProperties();
            FontSize fontSize464 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript464 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties66.Append(fontSize464);
            paragraphMarkRunProperties66.Append(fontSizeComplexScript464);

            paragraphProperties66.Append(indentation66);
            paragraphProperties66.Append(justification62);
            paragraphProperties66.Append(paragraphMarkRunProperties66);
            paragraph66.Append(paragraphProperties66);

            Paragraph paragraph67 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties67 = new ParagraphProperties();
            Indentation indentation67 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification63 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties67 = new ParagraphMarkRunProperties();
            Kern kern472 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize488 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript488 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties67.Append(kern472);
            paragraphMarkRunProperties67.Append(fontSize488);
            paragraphMarkRunProperties67.Append(fontSizeComplexScript488);

            paragraphProperties67.Append(indentation67);
            paragraphProperties67.Append(justification63);
            paragraphProperties67.Append(paragraphMarkRunProperties67);

            Run run422 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties422 = new RunProperties();
            Kern kern473 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize489 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript489 = new FontSizeComplexScript() { Val = "20" };

            runProperties422.Append(kern473);
            runProperties422.Append(fontSize489);
            runProperties422.Append(fontSizeComplexScript489);
            Text text39 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text39.Text = content["人际关系紧张与敏感"].Item2 == "没有" ? "★" : " ";

            run422.Append(runProperties422);
            run422.Append(text39);

            paragraph67.Append(paragraphProperties67);
            paragraph67.Append(run422);

            Paragraph paragraph68 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties68 = new ParagraphProperties();
            Indentation indentation68 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification64 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties68 = new ParagraphMarkRunProperties();
            Kern kern474 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize490 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript490 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties68.Append(kern474);
            paragraphMarkRunProperties68.Append(fontSize490);
            paragraphMarkRunProperties68.Append(fontSizeComplexScript490);

            paragraphProperties68.Append(indentation68);
            paragraphProperties68.Append(justification64);
            paragraphProperties68.Append(paragraphMarkRunProperties68);
            paragraph68.Append(paragraphProperties68);
            tableCell35.Append(tableCellProperties35);

            tableCell35.Append(paragraph66);
            tableCell35.Append(paragraph67);
            tableCell35.Append(paragraph68);

            TableCell tableCell36 = new TableCell();

            TableCellProperties tableCellProperties36 = new TableCellProperties();
            TableCellWidth tableCellWidth36 = new TableCellWidth() { Width = "755", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment36 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties36.Append(tableCellWidth36);
            tableCellProperties36.Append(tableCellVerticalAlignment36);

            Paragraph paragraph69 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties69 = new ParagraphProperties();
            Indentation indentation69 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification65 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties69 = new ParagraphMarkRunProperties();
            FontSize fontSize492 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript492 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties69.Append(fontSize492);
            paragraphMarkRunProperties69.Append(fontSizeComplexScript492);

            paragraphProperties69.Append(indentation69);
            paragraphProperties69.Append(justification65);
            paragraphProperties69.Append(paragraphMarkRunProperties69);

            paragraph69.Append(paragraphProperties69);

            Paragraph paragraph70 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties70 = new ParagraphProperties();
            Indentation indentation70 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification66 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties70 = new ParagraphMarkRunProperties();
            Kern kern499 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize516 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript516 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties70.Append(kern499);
            paragraphMarkRunProperties70.Append(fontSize516);
            paragraphMarkRunProperties70.Append(fontSizeComplexScript516);

            paragraphProperties70.Append(indentation70);
            paragraphProperties70.Append(justification66);
            paragraphProperties70.Append(paragraphMarkRunProperties70);

            Run run447 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties447 = new RunProperties();
            Kern kern500 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize517 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript517 = new FontSizeComplexScript() { Val = "20" };

            runProperties447.Append(kern500);
            runProperties447.Append(fontSize517);
            runProperties447.Append(fontSizeComplexScript517);
            Text text40 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text40.Text = content["人际关系紧张与敏感"].Item2 == "很轻" ? "★" : " ";

            run447.Append(runProperties447);
            run447.Append(text40);

            paragraph70.Append(paragraphProperties70);
            paragraph70.Append(run447);

            Paragraph paragraph71 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties71 = new ParagraphProperties();
            Indentation indentation71 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification67 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties71 = new ParagraphMarkRunProperties();
            Kern kern501 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize518 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript518 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties71.Append(kern501);
            paragraphMarkRunProperties71.Append(fontSize518);
            paragraphMarkRunProperties71.Append(fontSizeComplexScript518);

            paragraphProperties71.Append(indentation71);
            paragraphProperties71.Append(justification67);
            paragraphProperties71.Append(paragraphMarkRunProperties71);

            paragraph71.Append(paragraphProperties71);

            tableCell36.Append(tableCellProperties36);
            tableCell36.Append(paragraph69);
            tableCell36.Append(paragraph70);
            tableCell36.Append(paragraph71);

            TableCell tableCell37 = new TableCell();

            TableCellProperties tableCellProperties37 = new TableCellProperties();
            TableCellWidth tableCellWidth37 = new TableCellWidth() { Width = "756", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment37 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties37.Append(tableCellWidth37);
            tableCellProperties37.Append(tableCellVerticalAlignment37);

            Paragraph paragraph72 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties72 = new ParagraphProperties();
            Indentation indentation72 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification68 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties72 = new ParagraphMarkRunProperties();
            FontSize fontSize520 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript520 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties72.Append(fontSize520);
            paragraphMarkRunProperties72.Append(fontSizeComplexScript520);

            paragraphProperties72.Append(indentation72);
            paragraphProperties72.Append(justification68);
            paragraphProperties72.Append(paragraphMarkRunProperties72);

            paragraph72.Append(paragraphProperties72);

            Paragraph paragraph73 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties73 = new ParagraphProperties();
            Indentation indentation73 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification69 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties73 = new ParagraphMarkRunProperties();
            Kern kern526 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize544 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript544 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties73.Append(kern526);
            paragraphMarkRunProperties73.Append(fontSize544);
            paragraphMarkRunProperties73.Append(fontSizeComplexScript544);

            paragraphProperties73.Append(indentation73);
            paragraphProperties73.Append(justification69);
            paragraphProperties73.Append(paragraphMarkRunProperties73);

            Run run472 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties472 = new RunProperties();
            RunFonts runFonts367 = new RunFonts() { Ascii = "Segoe UI Symbol", HighAnsi = "Segoe UI Symbol", ComplexScript = "Segoe UI Symbol" };
            Kern kern527 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize545 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript545 = new FontSizeComplexScript() { Val = "20" };

            runProperties472.Append(runFonts367);
            runProperties472.Append(kern527);
            runProperties472.Append(fontSize545);
            runProperties472.Append(fontSizeComplexScript545);
            Text text41 = new Text();
            text41.Text = content["人际关系紧张与敏感"].Item2 == "中度" ? "★" : " ";

            run472.Append(runProperties472);
            run472.Append(text41);

            paragraph73.Append(paragraphProperties73);
            paragraph73.Append(run472);

            Paragraph paragraph74 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties74 = new ParagraphProperties();
            Indentation indentation74 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification70 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties74 = new ParagraphMarkRunProperties();
            Kern kern528 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize546 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript546 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties74.Append(kern528);
            paragraphMarkRunProperties74.Append(fontSize546);
            paragraphMarkRunProperties74.Append(fontSizeComplexScript546);

            paragraphProperties74.Append(indentation74);
            paragraphProperties74.Append(justification70);
            paragraphProperties74.Append(paragraphMarkRunProperties74);

            paragraph74.Append(paragraphProperties74);

            tableCell37.Append(tableCellProperties37);
            tableCell37.Append(paragraph72);
            tableCell37.Append(paragraph73);
            tableCell37.Append(paragraph74);

            TableCell tableCell38 = new TableCell();

            TableCellProperties tableCellProperties38 = new TableCellProperties();
            TableCellWidth tableCellWidth38 = new TableCellWidth() { Width = "756", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment38 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties38.Append(tableCellWidth38);
            tableCellProperties38.Append(tableCellVerticalAlignment38);

            Paragraph paragraph75 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties75 = new ParagraphProperties();
            Indentation indentation75 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification71 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties75 = new ParagraphMarkRunProperties();
            FontSize fontSize548 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript548 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties75.Append(fontSize548);
            paragraphMarkRunProperties75.Append(fontSizeComplexScript548);

            paragraphProperties75.Append(indentation75);
            paragraphProperties75.Append(justification71);
            paragraphProperties75.Append(paragraphMarkRunProperties75);

            paragraph75.Append(paragraphProperties75);

            Paragraph paragraph76 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties76 = new ParagraphProperties();
            Indentation indentation76 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification72 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties76 = new ParagraphMarkRunProperties();
            Kern kern553 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize572 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript572 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties76.Append(kern553);
            paragraphMarkRunProperties76.Append(fontSize572);
            paragraphMarkRunProperties76.Append(fontSizeComplexScript572);

            paragraphProperties76.Append(indentation76);
            paragraphProperties76.Append(justification72);
            paragraphProperties76.Append(paragraphMarkRunProperties76);

            Run run497 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties497 = new RunProperties();
            Kern kern554 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize573 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript573 = new FontSizeComplexScript() { Val = "20" };

            runProperties497.Append(kern554);
            runProperties497.Append(fontSize573);
            runProperties497.Append(fontSizeComplexScript573);
            Text text42 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text42.Text = content["人际关系紧张与敏感"].Item2 == "较重" ? "★" : " ";

            run497.Append(runProperties497);
            run497.Append(text42);

            paragraph76.Append(paragraphProperties76);
            paragraph76.Append(run497);

            Paragraph paragraph77 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties77 = new ParagraphProperties();
            Indentation indentation77 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification73 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties77 = new ParagraphMarkRunProperties();
            Kern kern555 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize574 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript574 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties77.Append(kern555);
            paragraphMarkRunProperties77.Append(fontSize574);
            paragraphMarkRunProperties77.Append(fontSizeComplexScript574);

            paragraphProperties77.Append(indentation77);
            paragraphProperties77.Append(justification73);
            paragraphProperties77.Append(paragraphMarkRunProperties77);

            paragraph77.Append(paragraphProperties77);

            tableCell38.Append(tableCellProperties38);
            tableCell38.Append(paragraph75);
            tableCell38.Append(paragraph76);
            tableCell38.Append(paragraph77);

            TableCell tableCell39 = new TableCell();

            TableCellProperties tableCellProperties39 = new TableCellProperties();
            TableCellWidth tableCellWidth39 = new TableCellWidth() { Width = "827", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment39 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties39.Append(tableCellWidth39);
            tableCellProperties39.Append(tableCellVerticalAlignment39);

            Paragraph paragraph78 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties78 = new ParagraphProperties();
            Indentation indentation78 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification74 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties78 = new ParagraphMarkRunProperties();
            FontSize fontSize576 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript576 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties78.Append(fontSize576);
            paragraphMarkRunProperties78.Append(fontSizeComplexScript576);

            paragraphProperties78.Append(indentation78);
            paragraphProperties78.Append(justification74);
            paragraphProperties78.Append(paragraphMarkRunProperties78);


            paragraph78.Append(paragraphProperties78);

            Paragraph paragraph79 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties79 = new ParagraphProperties();
            Indentation indentation79 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification75 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties79 = new ParagraphMarkRunProperties();
            Kern kern580 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize600 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript600 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties79.Append(kern580);
            paragraphMarkRunProperties79.Append(fontSize600);
            paragraphMarkRunProperties79.Append(fontSizeComplexScript600);

            paragraphProperties79.Append(indentation79);
            paragraphProperties79.Append(justification75);
            paragraphProperties79.Append(paragraphMarkRunProperties79);

            Run run522 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties522 = new RunProperties();
            Kern kern581 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize601 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript601 = new FontSizeComplexScript() { Val = "20" };

            runProperties522.Append(kern581);
            runProperties522.Append(fontSize601);
            runProperties522.Append(fontSizeComplexScript601);
            Text text43 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text43.Text = content["人际关系紧张与敏感"].Item2 == "严重" ? "★" : " ";

            run522.Append(runProperties522);
            run522.Append(text43);

            paragraph79.Append(paragraphProperties79);
            paragraph79.Append(run522);

            Paragraph paragraph80 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties80 = new ParagraphProperties();
            Indentation indentation80 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification76 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties80 = new ParagraphMarkRunProperties();
            Kern kern582 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize602 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript602 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties80.Append(kern582);
            paragraphMarkRunProperties80.Append(fontSize602);
            paragraphMarkRunProperties80.Append(fontSizeComplexScript602);

            paragraphProperties80.Append(indentation80);
            paragraphProperties80.Append(justification76);
            paragraphProperties80.Append(paragraphMarkRunProperties80);
            paragraph80.Append(paragraphProperties80);

            tableCell39.Append(tableCellProperties39);
            tableCell39.Append(paragraph78);
            tableCell39.Append(paragraph79);
            tableCell39.Append(paragraph80);

            tableRow7.Append(tableRowProperties7);
            tableRow7.Append(tableCell33);
            tableRow7.Append(tableCell34);
            tableRow7.Append(tableCell35);
            tableRow7.Append(tableCell36);
            tableRow7.Append(tableCell37);
            tableRow7.Append(tableCell38);
            tableRow7.Append(tableCell39);

            TableRow tableRow8 = new TableRow() { RsidTableRowMarkRevision = "00C709B4", RsidTableRowAddition = "00D061CC", RsidTableRowProperties = "009E6323" };

            TableRowProperties tableRowProperties8 = new TableRowProperties();
            TableRowHeight tableRowHeight8 = new TableRowHeight() { Val = (UInt32Value)765U };

            tableRowProperties8.Append(tableRowHeight8);

            TableCell tableCell40 = new TableCell();

            TableCellProperties tableCellProperties40 = new TableCellProperties();
            TableCellWidth tableCellWidth40 = new TableCellWidth() { Width = "1678", Type = TableWidthUnitValues.Dxa };
            Shading shading16 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "D6E3BC" };
            TableCellVerticalAlignment tableCellVerticalAlignment40 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties40.Append(tableCellWidth40);
            tableCellProperties40.Append(shading16);
            tableCellProperties40.Append(tableCellVerticalAlignment40);

            Paragraph paragraph81 = new Paragraph() { RsidParagraphMarkRevision = "00C709B4", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties81 = new ParagraphProperties();
            Indentation indentation81 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification77 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties81 = new ParagraphMarkRunProperties();
            Bold bold35 = new Bold();
            Kern kern584 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize604 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript604 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties81.Append(bold35);
            paragraphMarkRunProperties81.Append(kern584);
            paragraphMarkRunProperties81.Append(fontSize604);
            paragraphMarkRunProperties81.Append(fontSizeComplexScript604);

            paragraphProperties81.Append(indentation81);
            paragraphProperties81.Append(justification77);
            paragraphProperties81.Append(paragraphMarkRunProperties81);

            Run run524 = new Run() { RsidRunProperties = "00C709B4" };

            RunProperties runProperties524 = new RunProperties();
            RunFonts runFonts406 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold36 = new Bold();
            Kern kern585 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize605 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript605 = new FontSizeComplexScript() { Val = "20" };

            runProperties524.Append(runFonts406);
            runProperties524.Append(bold36);
            runProperties524.Append(kern585);
            runProperties524.Append(fontSize605);
            runProperties524.Append(fontSizeComplexScript605);
            Text text44 = new Text();
            text44.Text = "抑郁";

            run524.Append(runProperties524);
            run524.Append(text44);

            paragraph81.Append(paragraphProperties81);
            paragraph81.Append(run524);

            tableCell40.Append(tableCellProperties40);
            tableCell40.Append(paragraph81);

            TableCell tableCell41 = new TableCell();

            TableCellProperties tableCellProperties41 = new TableCellProperties();
            TableCellWidth tableCellWidth41 = new TableCellWidth() { Width = "4453", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment41 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties41.Append(tableCellWidth41);
            tableCellProperties41.Append(tableCellVerticalAlignment41);

            Paragraph paragraph82 = new Paragraph() { RsidParagraphMarkRevision = "00C709B4", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties82 = new ParagraphProperties();
            Indentation indentation82 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };

            ParagraphMarkRunProperties paragraphMarkRunProperties82 = new ParagraphMarkRunProperties();
            Kern kern586 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize606 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript606 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties82.Append(kern586);
            paragraphMarkRunProperties82.Append(fontSize606);
            paragraphMarkRunProperties82.Append(fontSizeComplexScript606);

            paragraphProperties82.Append(indentation82);
            paragraphProperties82.Append(paragraphMarkRunProperties82);

            Run run525 = new Run() { RsidRunProperties = "00C709B4" };

            RunProperties runProperties525 = new RunProperties();
            RunFonts runFonts407 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Kern kern587 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize607 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript607 = new FontSizeComplexScript() { Val = "20" };

            runProperties525.Append(runFonts407);
            runProperties525.Append(kern587);
            runProperties525.Append(fontSize607);
            runProperties525.Append(fontSizeComplexScript607);
            Text text45 = new Text();
            text45.Text = "感到生活单调，感到自己没有前途，容易哭泣，责备自己，无精打采";

            run525.Append(runProperties525);
            run525.Append(text45);

            paragraph82.Append(paragraphProperties82);
            paragraph82.Append(run525);

            tableCell41.Append(tableCellProperties41);
            tableCell41.Append(paragraph82);

            TableCell tableCell42 = new TableCell();

            TableCellProperties tableCellProperties42 = new TableCellProperties();
            TableCellWidth tableCellWidth42 = new TableCellWidth() { Width = "756", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment42 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties42.Append(tableCellWidth42);
            tableCellProperties42.Append(tableCellVerticalAlignment42);

            Paragraph paragraph83 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties83 = new ParagraphProperties();
            Indentation indentation83 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification78 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties83 = new ParagraphMarkRunProperties();
            FontSize fontSize608 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript608 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties83.Append(fontSize608);
            paragraphMarkRunProperties83.Append(fontSizeComplexScript608);

            paragraphProperties83.Append(indentation83);
            paragraphProperties83.Append(justification78);
            paragraphProperties83.Append(paragraphMarkRunProperties83);

            paragraph83.Append(paragraphProperties83);

            Paragraph paragraph84 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties84 = new ParagraphProperties();
            Indentation indentation84 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification79 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties84 = new ParagraphMarkRunProperties();
            Kern kern611 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize632 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript632 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties84.Append(kern611);
            paragraphMarkRunProperties84.Append(fontSize632);
            paragraphMarkRunProperties84.Append(fontSizeComplexScript632);

            paragraphProperties84.Append(indentation84);
            paragraphProperties84.Append(justification79);
            paragraphProperties84.Append(paragraphMarkRunProperties84);

            Run run549 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties549 = new RunProperties();
            Kern kern612 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize633 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript633 = new FontSizeComplexScript() { Val = "20" };

            runProperties549.Append(kern612);
            runProperties549.Append(fontSize633);
            runProperties549.Append(fontSizeComplexScript633);
            Text text46 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text46.Text = content["抑郁"].Item2 == "没有" ? "★" : " ";

            run549.Append(runProperties549);
            run549.Append(text46);

            paragraph84.Append(paragraphProperties84);
            paragraph84.Append(run549);

            Paragraph paragraph85 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties85 = new ParagraphProperties();
            Indentation indentation85 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification80 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties85 = new ParagraphMarkRunProperties();
            Kern kern613 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize634 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript634 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties85.Append(kern613);
            paragraphMarkRunProperties85.Append(fontSize634);
            paragraphMarkRunProperties85.Append(fontSizeComplexScript634);

            paragraphProperties85.Append(indentation85);
            paragraphProperties85.Append(justification80);
            paragraphProperties85.Append(paragraphMarkRunProperties85);

            paragraph85.Append(paragraphProperties85);

            tableCell42.Append(tableCellProperties42);
            tableCell42.Append(paragraph83);
            tableCell42.Append(paragraph84);
            tableCell42.Append(paragraph85);

            TableCell tableCell43 = new TableCell();

            TableCellProperties tableCellProperties43 = new TableCellProperties();
            TableCellWidth tableCellWidth43 = new TableCellWidth() { Width = "755", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment43 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties43.Append(tableCellWidth43);
            tableCellProperties43.Append(tableCellVerticalAlignment43);

            Paragraph paragraph86 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties86 = new ParagraphProperties();
            Indentation indentation86 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification81 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties86 = new ParagraphMarkRunProperties();
            FontSize fontSize636 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript636 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties86.Append(fontSize636);
            paragraphMarkRunProperties86.Append(fontSizeComplexScript636);

            paragraphProperties86.Append(indentation86);
            paragraphProperties86.Append(justification81);
            paragraphProperties86.Append(paragraphMarkRunProperties86);

            paragraph86.Append(paragraphProperties86);

            Paragraph paragraph87 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties87 = new ParagraphProperties();
            Indentation indentation87 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification82 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties87 = new ParagraphMarkRunProperties();
            Kern kern638 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize660 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript660 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties87.Append(kern638);
            paragraphMarkRunProperties87.Append(fontSize660);
            paragraphMarkRunProperties87.Append(fontSizeComplexScript660);

            paragraphProperties87.Append(indentation87);
            paragraphProperties87.Append(justification82);
            paragraphProperties87.Append(paragraphMarkRunProperties87);

            Run run574 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties574 = new RunProperties();
            RunFonts runFonts446 = new RunFonts() { Ascii = "Segoe UI Symbol", HighAnsi = "Segoe UI Symbol", ComplexScript = "Segoe UI Symbol" };
            Kern kern639 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize661 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript661 = new FontSizeComplexScript() { Val = "20" };

            runProperties574.Append(runFonts446);
            runProperties574.Append(kern639);
            runProperties574.Append(fontSize661);
            runProperties574.Append(fontSizeComplexScript661);
            Text text47 = new Text();
            text47.Text = content["抑郁"].Item2 == "轻度" ? "★" : " ";

            run574.Append(runProperties574);
            run574.Append(text47);

            paragraph87.Append(paragraphProperties87);
            paragraph87.Append(run574);

            Paragraph paragraph88 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties88 = new ParagraphProperties();
            Indentation indentation88 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification83 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties88 = new ParagraphMarkRunProperties();
            Kern kern640 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize662 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript662 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties88.Append(kern640);
            paragraphMarkRunProperties88.Append(fontSize662);
            paragraphMarkRunProperties88.Append(fontSizeComplexScript662);

            paragraphProperties88.Append(indentation88);
            paragraphProperties88.Append(justification83);
            paragraphProperties88.Append(paragraphMarkRunProperties88);

            Run run575 = new Run();

            RunProperties runProperties575 = new RunProperties();
            Kern kern641 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize663 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript663 = new FontSizeComplexScript() { Val = "20" };

            runProperties575.Append(kern641);
            runProperties575.Append(fontSize663);
            runProperties575.Append(fontSizeComplexScript663);
            FieldChar fieldChar66 = new FieldChar() { FieldCharType = FieldCharValues.End };

            run575.Append(runProperties575);
            run575.Append(fieldChar66);

            paragraph88.Append(paragraphProperties88);
            paragraph88.Append(run575);

            tableCell43.Append(tableCellProperties43);
            tableCell43.Append(paragraph86);
            tableCell43.Append(paragraph87);
            tableCell43.Append(paragraph88);

            TableCell tableCell44 = new TableCell();

            TableCellProperties tableCellProperties44 = new TableCellProperties();
            TableCellWidth tableCellWidth44 = new TableCellWidth() { Width = "756", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment44 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties44.Append(tableCellWidth44);
            tableCellProperties44.Append(tableCellVerticalAlignment44);

            Paragraph paragraph89 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties89 = new ParagraphProperties();
            Indentation indentation89 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification84 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties89 = new ParagraphMarkRunProperties();
            FontSize fontSize664 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript664 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties89.Append(fontSize664);
            paragraphMarkRunProperties89.Append(fontSizeComplexScript664);

            paragraphProperties89.Append(indentation89);
            paragraphProperties89.Append(justification84);
            paragraphProperties89.Append(paragraphMarkRunProperties89);

            paragraph89.Append(paragraphProperties89);

            Paragraph paragraph90 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties90 = new ParagraphProperties();
            Indentation indentation90 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification85 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties90 = new ParagraphMarkRunProperties();
            Kern kern665 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize688 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript688 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties90.Append(kern665);
            paragraphMarkRunProperties90.Append(fontSize688);
            paragraphMarkRunProperties90.Append(fontSizeComplexScript688);

            paragraphProperties90.Append(indentation90);
            paragraphProperties90.Append(justification85);
            paragraphProperties90.Append(paragraphMarkRunProperties90);

            Run run599 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties599 = new RunProperties();
            Kern kern666 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize689 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript689 = new FontSizeComplexScript() { Val = "20" };

            runProperties599.Append(kern666);
            runProperties599.Append(fontSize689);
            runProperties599.Append(fontSizeComplexScript689);
            Text text48 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text48.Text = content["抑郁"].Item2 == "中度" ? "★" : " ";

            run599.Append(runProperties599);
            run599.Append(text48);

            paragraph90.Append(paragraphProperties90);
            paragraph90.Append(run599);

            Paragraph paragraph91 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties91 = new ParagraphProperties();
            Indentation indentation91 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification86 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties91 = new ParagraphMarkRunProperties();
            Kern kern667 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize690 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript690 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties91.Append(kern667);
            paragraphMarkRunProperties91.Append(fontSize690);
            paragraphMarkRunProperties91.Append(fontSizeComplexScript690);

            paragraphProperties91.Append(indentation91);
            paragraphProperties91.Append(justification86);
            paragraphProperties91.Append(paragraphMarkRunProperties91);

            paragraph91.Append(paragraphProperties91);

            tableCell44.Append(tableCellProperties44);
            tableCell44.Append(paragraph89);
            tableCell44.Append(paragraph90);
            tableCell44.Append(paragraph91);

            TableCell tableCell45 = new TableCell();

            TableCellProperties tableCellProperties45 = new TableCellProperties();
            TableCellWidth tableCellWidth45 = new TableCellWidth() { Width = "756", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment45 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties45.Append(tableCellWidth45);
            tableCellProperties45.Append(tableCellVerticalAlignment45);

            Paragraph paragraph92 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties92 = new ParagraphProperties();
            Indentation indentation92 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification87 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties92 = new ParagraphMarkRunProperties();
            FontSize fontSize692 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript692 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties92.Append(fontSize692);
            paragraphMarkRunProperties92.Append(fontSizeComplexScript692);

            paragraphProperties92.Append(indentation92);
            paragraphProperties92.Append(justification87);
            paragraphProperties92.Append(paragraphMarkRunProperties92);

            paragraph92.Append(paragraphProperties92);

            Paragraph paragraph93 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties93 = new ParagraphProperties();
            Indentation indentation93 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification88 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties93 = new ParagraphMarkRunProperties();
            Kern kern692 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize716 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript716 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties93.Append(kern692);
            paragraphMarkRunProperties93.Append(fontSize716);
            paragraphMarkRunProperties93.Append(fontSizeComplexScript716);

            paragraphProperties93.Append(indentation93);
            paragraphProperties93.Append(justification88);
            paragraphProperties93.Append(paragraphMarkRunProperties93);

            Run run624 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties624 = new RunProperties();
            Kern kern693 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize717 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript717 = new FontSizeComplexScript() { Val = "20" };

            runProperties624.Append(kern693);
            runProperties624.Append(fontSize717);
            runProperties624.Append(fontSizeComplexScript717);
            Text text49 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text49.Text = content["抑郁"].Item2 == "较重" ? "★" : " ";

            run624.Append(runProperties624);
            run624.Append(text49);

            paragraph93.Append(paragraphProperties93);
            paragraph93.Append(run624);

            Paragraph paragraph94 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties94 = new ParagraphProperties();
            Indentation indentation94 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification89 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties94 = new ParagraphMarkRunProperties();
            Kern kern694 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize718 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript718 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties94.Append(kern694);
            paragraphMarkRunProperties94.Append(fontSize718);
            paragraphMarkRunProperties94.Append(fontSizeComplexScript718);

            paragraphProperties94.Append(indentation94);
            paragraphProperties94.Append(justification89);
            paragraphProperties94.Append(paragraphMarkRunProperties94);

            paragraph94.Append(paragraphProperties94);

            tableCell45.Append(tableCellProperties45);
            tableCell45.Append(paragraph92);
            tableCell45.Append(paragraph93);
            tableCell45.Append(paragraph94);

            TableCell tableCell46 = new TableCell();

            TableCellProperties tableCellProperties46 = new TableCellProperties();
            TableCellWidth tableCellWidth46 = new TableCellWidth() { Width = "827", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment46 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties46.Append(tableCellWidth46);
            tableCellProperties46.Append(tableCellVerticalAlignment46);

            Paragraph paragraph95 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties95 = new ParagraphProperties();
            Indentation indentation95 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification90 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties95 = new ParagraphMarkRunProperties();
            FontSize fontSize720 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript720 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties95.Append(fontSize720);
            paragraphMarkRunProperties95.Append(fontSizeComplexScript720);

            paragraphProperties95.Append(indentation95);
            paragraphProperties95.Append(justification90);
            paragraphProperties95.Append(paragraphMarkRunProperties95);

            paragraph95.Append(paragraphProperties95);

            Paragraph paragraph96 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties96 = new ParagraphProperties();
            Indentation indentation96 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification91 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties96 = new ParagraphMarkRunProperties();
            Kern kern719 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize744 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript744 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties96.Append(kern719);
            paragraphMarkRunProperties96.Append(fontSize744);
            paragraphMarkRunProperties96.Append(fontSizeComplexScript744);

            paragraphProperties96.Append(indentation96);
            paragraphProperties96.Append(justification91);
            paragraphProperties96.Append(paragraphMarkRunProperties96);

            Run run649 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties649 = new RunProperties();
            Kern kern720 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize745 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript745 = new FontSizeComplexScript() { Val = "20" };

            runProperties649.Append(kern720);
            runProperties649.Append(fontSize745);
            runProperties649.Append(fontSizeComplexScript745);
            Text text50 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text50.Text = content["抑郁"].Item2 == "严重" ? "★" : " ";

            run649.Append(runProperties649);
            run649.Append(text50);

            paragraph96.Append(paragraphProperties96);
            paragraph96.Append(run649);

            Paragraph paragraph97 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties97 = new ParagraphProperties();
            Indentation indentation97 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification92 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties97 = new ParagraphMarkRunProperties();
            Kern kern721 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize746 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript746 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties97.Append(kern721);
            paragraphMarkRunProperties97.Append(fontSize746);
            paragraphMarkRunProperties97.Append(fontSizeComplexScript746);

            paragraphProperties97.Append(indentation97);
            paragraphProperties97.Append(justification92);
            paragraphProperties97.Append(paragraphMarkRunProperties97);

            paragraph97.Append(paragraphProperties97);

            tableCell46.Append(tableCellProperties46);
            tableCell46.Append(paragraph95);
            tableCell46.Append(paragraph96);
            tableCell46.Append(paragraph97);

            tableRow8.Append(tableRowProperties8);
            tableRow8.Append(tableCell40);
            tableRow8.Append(tableCell41);
            tableRow8.Append(tableCell42);
            tableRow8.Append(tableCell43);
            tableRow8.Append(tableCell44);
            tableRow8.Append(tableCell45);
            tableRow8.Append(tableCell46);

            TableRow tableRow9 = new TableRow() { RsidTableRowMarkRevision = "00C709B4", RsidTableRowAddition = "00D061CC", RsidTableRowProperties = "009E6323" };

            TableRowProperties tableRowProperties9 = new TableRowProperties();
            TableRowHeight tableRowHeight9 = new TableRowHeight() { Val = (UInt32Value)800U };

            tableRowProperties9.Append(tableRowHeight9);

            TableCell tableCell47 = new TableCell();

            TableCellProperties tableCellProperties47 = new TableCellProperties();
            TableCellWidth tableCellWidth47 = new TableCellWidth() { Width = "1678", Type = TableWidthUnitValues.Dxa };
            Shading shading17 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "D6E3BC" };
            TableCellVerticalAlignment tableCellVerticalAlignment47 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties47.Append(tableCellWidth47);
            tableCellProperties47.Append(shading17);
            tableCellProperties47.Append(tableCellVerticalAlignment47);

            Paragraph paragraph98 = new Paragraph() { RsidParagraphMarkRevision = "00C709B4", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties98 = new ParagraphProperties();
            Indentation indentation98 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification93 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties98 = new ParagraphMarkRunProperties();
            Bold bold37 = new Bold();
            Kern kern723 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize748 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript748 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties98.Append(bold37);
            paragraphMarkRunProperties98.Append(kern723);
            paragraphMarkRunProperties98.Append(fontSize748);
            paragraphMarkRunProperties98.Append(fontSizeComplexScript748);

            paragraphProperties98.Append(indentation98);
            paragraphProperties98.Append(justification93);
            paragraphProperties98.Append(paragraphMarkRunProperties98);

            Run run651 = new Run() { RsidRunProperties = "00C709B4" };

            RunProperties runProperties651 = new RunProperties();
            RunFonts runFonts504 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold38 = new Bold();
            Kern kern724 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize749 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript749 = new FontSizeComplexScript() { Val = "20" };

            runProperties651.Append(runFonts504);
            runProperties651.Append(bold38);
            runProperties651.Append(kern724);
            runProperties651.Append(fontSize749);
            runProperties651.Append(fontSizeComplexScript749);
            Text text51 = new Text();
            text51.Text = "焦虑";

            run651.Append(runProperties651);
            run651.Append(text51);

            paragraph98.Append(paragraphProperties98);
            paragraph98.Append(run651);

            tableCell47.Append(tableCellProperties47);
            tableCell47.Append(paragraph98);

            TableCell tableCell48 = new TableCell();

            TableCellProperties tableCellProperties48 = new TableCellProperties();
            TableCellWidth tableCellWidth48 = new TableCellWidth() { Width = "4453", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment48 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties48.Append(tableCellWidth48);
            tableCellProperties48.Append(tableCellVerticalAlignment48);

            Paragraph paragraph99 = new Paragraph() { RsidParagraphMarkRevision = "00C709B4", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties99 = new ParagraphProperties();
            Indentation indentation99 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };

            ParagraphMarkRunProperties paragraphMarkRunProperties99 = new ParagraphMarkRunProperties();
            Kern kern725 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize750 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript750 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties99.Append(kern725);
            paragraphMarkRunProperties99.Append(fontSize750);
            paragraphMarkRunProperties99.Append(fontSizeComplexScript750);

            paragraphProperties99.Append(indentation99);
            paragraphProperties99.Append(paragraphMarkRunProperties99);

            Run run652 = new Run() { RsidRunProperties = "00C709B4" };

            RunProperties runProperties652 = new RunProperties();
            RunFonts runFonts505 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Kern kern726 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize751 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript751 = new FontSizeComplexScript() { Val = "20" };

            runProperties652.Append(runFonts505);
            runProperties652.Append(kern726);
            runProperties652.Append(fontSize751);
            runProperties652.Append(fontSizeComplexScript751);
            Text text52 = new Text();
            text52.Text = "感到紧张，心神不定，无缘无故的害怕，心里烦躁，心里不踏实";

            run652.Append(runProperties652);
            run652.Append(text52);

            paragraph99.Append(paragraphProperties99);
            paragraph99.Append(run652);

            tableCell48.Append(tableCellProperties48);
            tableCell48.Append(paragraph99);

            TableCell tableCell49 = new TableCell();

            TableCellProperties tableCellProperties49 = new TableCellProperties();
            TableCellWidth tableCellWidth49 = new TableCellWidth() { Width = "756", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment49 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties49.Append(tableCellWidth49);
            tableCellProperties49.Append(tableCellVerticalAlignment49);

            Paragraph paragraph100 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties100 = new ParagraphProperties();
            Indentation indentation100 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification94 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties100 = new ParagraphMarkRunProperties();
            FontSize fontSize752 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript752 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties100.Append(fontSize752);
            paragraphMarkRunProperties100.Append(fontSizeComplexScript752);

            paragraphProperties100.Append(indentation100);
            paragraphProperties100.Append(justification94);
            paragraphProperties100.Append(paragraphMarkRunProperties100);

            paragraph100.Append(paragraphProperties100);

            Paragraph paragraph101 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties101 = new ParagraphProperties();
            Indentation indentation101 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification95 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties101 = new ParagraphMarkRunProperties();
            Kern kern750 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize776 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript776 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties101.Append(kern750);
            paragraphMarkRunProperties101.Append(fontSize776);
            paragraphMarkRunProperties101.Append(fontSizeComplexScript776);

            paragraphProperties101.Append(indentation101);
            paragraphProperties101.Append(justification95);
            paragraphProperties101.Append(paragraphMarkRunProperties101);

            Run run676 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties676 = new RunProperties();
            Kern kern751 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize777 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript777 = new FontSizeComplexScript() { Val = "20" };

            runProperties676.Append(kern751);
            runProperties676.Append(fontSize777);
            runProperties676.Append(fontSizeComplexScript777);
            Text text53 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text53.Text = content["焦虑"].Item2 == "没有" ? "★" : " ";

            run676.Append(runProperties676);
            run676.Append(text53);

            paragraph101.Append(paragraphProperties101);
            paragraph101.Append(run676);

            Paragraph paragraph102 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties102 = new ParagraphProperties();
            Indentation indentation102 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification96 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties102 = new ParagraphMarkRunProperties();
            Kern kern752 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize778 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript778 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties102.Append(kern752);
            paragraphMarkRunProperties102.Append(fontSize778);
            paragraphMarkRunProperties102.Append(fontSizeComplexScript778);

            paragraphProperties102.Append(indentation102);
            paragraphProperties102.Append(justification96);
            paragraphProperties102.Append(paragraphMarkRunProperties102);

            paragraph102.Append(paragraphProperties102);

            tableCell49.Append(tableCellProperties49);
            tableCell49.Append(paragraph100);
            tableCell49.Append(paragraph101);
            tableCell49.Append(paragraph102);

            TableCell tableCell50 = new TableCell();

            TableCellProperties tableCellProperties50 = new TableCellProperties();
            TableCellWidth tableCellWidth50 = new TableCellWidth() { Width = "755", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment50 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties50.Append(tableCellWidth50);
            tableCellProperties50.Append(tableCellVerticalAlignment50);

            Paragraph paragraph103 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties103 = new ParagraphProperties();
            Indentation indentation103 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification97 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties103 = new ParagraphMarkRunProperties();
            FontSize fontSize780 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript780 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties103.Append(fontSize780);
            paragraphMarkRunProperties103.Append(fontSizeComplexScript780);

            paragraphProperties103.Append(indentation103);
            paragraphProperties103.Append(justification97);
            paragraphProperties103.Append(paragraphMarkRunProperties103);

            paragraph103.Append(paragraphProperties103);

            Paragraph paragraph104 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties104 = new ParagraphProperties();
            Indentation indentation104 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification98 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties104 = new ParagraphMarkRunProperties();
            Kern kern777 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize804 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript804 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties104.Append(kern777);
            paragraphMarkRunProperties104.Append(fontSize804);
            paragraphMarkRunProperties104.Append(fontSizeComplexScript804);

            paragraphProperties104.Append(indentation104);
            paragraphProperties104.Append(justification98);
            paragraphProperties104.Append(paragraphMarkRunProperties104);

            Run run701 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties701 = new RunProperties();
            RunFonts runFonts544 = new RunFonts() { Ascii = "Segoe UI Symbol", HighAnsi = "Segoe UI Symbol", ComplexScript = "Segoe UI Symbol" };
            Kern kern778 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize805 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript805 = new FontSizeComplexScript() { Val = "20" };

            runProperties701.Append(runFonts544);
            runProperties701.Append(kern778);
            runProperties701.Append(fontSize805);
            runProperties701.Append(fontSizeComplexScript805);
            Text text54 = new Text();
            text54.Text = content["焦虑"].Item2 == "轻度" ? "★" : " ";

            run701.Append(runProperties701);
            run701.Append(text54);

            paragraph104.Append(paragraphProperties104);
            paragraph104.Append(run701);

            Paragraph paragraph105 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties105 = new ParagraphProperties();
            Indentation indentation105 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification99 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties105 = new ParagraphMarkRunProperties();
            Kern kern779 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize806 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript806 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties105.Append(kern779);
            paragraphMarkRunProperties105.Append(fontSize806);
            paragraphMarkRunProperties105.Append(fontSizeComplexScript806);

            paragraphProperties105.Append(indentation105);
            paragraphProperties105.Append(justification99);
            paragraphProperties105.Append(paragraphMarkRunProperties105);

            paragraph105.Append(paragraphProperties105);

            tableCell50.Append(tableCellProperties50);
            tableCell50.Append(paragraph103);
            tableCell50.Append(paragraph104);
            tableCell50.Append(paragraph105);

            TableCell tableCell51 = new TableCell();

            TableCellProperties tableCellProperties51 = new TableCellProperties();
            TableCellWidth tableCellWidth51 = new TableCellWidth() { Width = "756", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment51 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties51.Append(tableCellWidth51);
            tableCellProperties51.Append(tableCellVerticalAlignment51);

            Paragraph paragraph106 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties106 = new ParagraphProperties();
            Indentation indentation106 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification100 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties106 = new ParagraphMarkRunProperties();
            FontSize fontSize808 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript808 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties106.Append(fontSize808);
            paragraphMarkRunProperties106.Append(fontSizeComplexScript808);

            paragraphProperties106.Append(indentation106);
            paragraphProperties106.Append(justification100);
            paragraphProperties106.Append(paragraphMarkRunProperties106);

            paragraph106.Append(paragraphProperties106);

            Paragraph paragraph107 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties107 = new ParagraphProperties();
            Indentation indentation107 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification101 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties107 = new ParagraphMarkRunProperties();
            Kern kern804 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize832 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript832 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties107.Append(kern804);
            paragraphMarkRunProperties107.Append(fontSize832);
            paragraphMarkRunProperties107.Append(fontSizeComplexScript832);

            paragraphProperties107.Append(indentation107);
            paragraphProperties107.Append(justification101);
            paragraphProperties107.Append(paragraphMarkRunProperties107);

            Run run726 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties726 = new RunProperties();
            Kern kern805 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize833 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript833 = new FontSizeComplexScript() { Val = "20" };

            runProperties726.Append(kern805);
            runProperties726.Append(fontSize833);
            runProperties726.Append(fontSizeComplexScript833);
            Text text55 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text55.Text = content["焦虑"].Item2 == "中度" ? "★" : " ";

            run726.Append(runProperties726);
            run726.Append(text55);

            paragraph107.Append(paragraphProperties107);
            paragraph107.Append(run726);

            Paragraph paragraph108 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties108 = new ParagraphProperties();
            Indentation indentation108 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification102 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties108 = new ParagraphMarkRunProperties();
            Kern kern806 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize834 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript834 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties108.Append(kern806);
            paragraphMarkRunProperties108.Append(fontSize834);
            paragraphMarkRunProperties108.Append(fontSizeComplexScript834);

            paragraphProperties108.Append(indentation108);
            paragraphProperties108.Append(justification102);
            paragraphProperties108.Append(paragraphMarkRunProperties108);

            paragraph108.Append(paragraphProperties108);

            tableCell51.Append(tableCellProperties51);
            tableCell51.Append(paragraph106);
            tableCell51.Append(paragraph107);
            tableCell51.Append(paragraph108);

            TableCell tableCell52 = new TableCell();

            TableCellProperties tableCellProperties52 = new TableCellProperties();
            TableCellWidth tableCellWidth52 = new TableCellWidth() { Width = "756", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment52 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties52.Append(tableCellWidth52);
            tableCellProperties52.Append(tableCellVerticalAlignment52);

            Paragraph paragraph109 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties109 = new ParagraphProperties();
            Indentation indentation109 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification103 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties109 = new ParagraphMarkRunProperties();
            FontSize fontSize836 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript836 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties109.Append(fontSize836);
            paragraphMarkRunProperties109.Append(fontSizeComplexScript836);

            paragraphProperties109.Append(indentation109);
            paragraphProperties109.Append(justification103);
            paragraphProperties109.Append(paragraphMarkRunProperties109);

            paragraph109.Append(paragraphProperties109);

            Paragraph paragraph110 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties110 = new ParagraphProperties();
            Indentation indentation110 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification104 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties110 = new ParagraphMarkRunProperties();
            Kern kern831 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize860 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript860 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties110.Append(kern831);
            paragraphMarkRunProperties110.Append(fontSize860);
            paragraphMarkRunProperties110.Append(fontSizeComplexScript860);

            paragraphProperties110.Append(indentation110);
            paragraphProperties110.Append(justification104);
            paragraphProperties110.Append(paragraphMarkRunProperties110);

            Run run751 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties751 = new RunProperties();
            Kern kern832 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize861 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript861 = new FontSizeComplexScript() { Val = "20" };

            runProperties751.Append(kern832);
            runProperties751.Append(fontSize861);
            runProperties751.Append(fontSizeComplexScript861);
            Text text56 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text56.Text = content["焦虑"].Item2 == "较重" ? "★" : " ";

            run751.Append(runProperties751);
            run751.Append(text56);

            paragraph110.Append(paragraphProperties110);
            paragraph110.Append(run751);

            Paragraph paragraph111 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties111 = new ParagraphProperties();
            Indentation indentation111 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification105 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties111 = new ParagraphMarkRunProperties();
            Kern kern833 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize862 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript862 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties111.Append(kern833);
            paragraphMarkRunProperties111.Append(fontSize862);
            paragraphMarkRunProperties111.Append(fontSizeComplexScript862);

            paragraphProperties111.Append(indentation111);
            paragraphProperties111.Append(justification105);
            paragraphProperties111.Append(paragraphMarkRunProperties111);

            paragraph111.Append(paragraphProperties111);

            tableCell52.Append(tableCellProperties52);
            tableCell52.Append(paragraph109);
            tableCell52.Append(paragraph110);
            tableCell52.Append(paragraph111);

            TableCell tableCell53 = new TableCell();

            TableCellProperties tableCellProperties53 = new TableCellProperties();
            TableCellWidth tableCellWidth53 = new TableCellWidth() { Width = "827", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment53 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties53.Append(tableCellWidth53);
            tableCellProperties53.Append(tableCellVerticalAlignment53);

            Paragraph paragraph112 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties112 = new ParagraphProperties();
            Indentation indentation112 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification106 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties112 = new ParagraphMarkRunProperties();
            FontSize fontSize864 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript864 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties112.Append(fontSize864);
            paragraphMarkRunProperties112.Append(fontSizeComplexScript864);

            paragraphProperties112.Append(indentation112);
            paragraphProperties112.Append(justification106);
            paragraphProperties112.Append(paragraphMarkRunProperties112);

            paragraph112.Append(paragraphProperties112);

            Paragraph paragraph113 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties113 = new ParagraphProperties();
            Indentation indentation113 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification107 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties113 = new ParagraphMarkRunProperties();
            Kern kern858 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize888 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript888 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties113.Append(kern858);
            paragraphMarkRunProperties113.Append(fontSize888);
            paragraphMarkRunProperties113.Append(fontSizeComplexScript888);

            paragraphProperties113.Append(indentation113);
            paragraphProperties113.Append(justification107);
            paragraphProperties113.Append(paragraphMarkRunProperties113);

            Run run776 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties776 = new RunProperties();
            Kern kern859 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize889 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript889 = new FontSizeComplexScript() { Val = "20" };

            runProperties776.Append(kern859);
            runProperties776.Append(fontSize889);
            runProperties776.Append(fontSizeComplexScript889);
            Text text57 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text57.Text = content["焦虑"].Item2 == "严重" ? "★" : " ";

            run776.Append(runProperties776);
            run776.Append(text57);

            paragraph113.Append(paragraphProperties113);
            paragraph113.Append(run776);

            Paragraph paragraph114 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties114 = new ParagraphProperties();
            Indentation indentation114 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification108 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties114 = new ParagraphMarkRunProperties();
            Kern kern860 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize890 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript890 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties114.Append(kern860);
            paragraphMarkRunProperties114.Append(fontSize890);
            paragraphMarkRunProperties114.Append(fontSizeComplexScript890);

            paragraphProperties114.Append(indentation114);
            paragraphProperties114.Append(justification108);
            paragraphProperties114.Append(paragraphMarkRunProperties114);

            paragraph114.Append(paragraphProperties114);

            tableCell53.Append(tableCellProperties53);
            tableCell53.Append(paragraph112);
            tableCell53.Append(paragraph113);
            tableCell53.Append(paragraph114);

            tableRow9.Append(tableRowProperties9);
            tableRow9.Append(tableCell47);
            tableRow9.Append(tableCell48);
            tableRow9.Append(tableCell49);
            tableRow9.Append(tableCell50);
            tableRow9.Append(tableCell51);
            tableRow9.Append(tableCell52);
            tableRow9.Append(tableCell53);

            TableRow tableRow10 = new TableRow() { RsidTableRowMarkRevision = "00C709B4", RsidTableRowAddition = "00D061CC", RsidTableRowProperties = "009E6323" };

            TableRowProperties tableRowProperties10 = new TableRowProperties();
            TableRowHeight tableRowHeight10 = new TableRowHeight() { Val = (UInt32Value)800U };

            tableRowProperties10.Append(tableRowHeight10);

            TableCell tableCell54 = new TableCell();

            TableCellProperties tableCellProperties54 = new TableCellProperties();
            TableCellWidth tableCellWidth54 = new TableCellWidth() { Width = "1678", Type = TableWidthUnitValues.Dxa };
            Shading shading18 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "D6E3BC" };
            TableCellVerticalAlignment tableCellVerticalAlignment54 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties54.Append(tableCellWidth54);
            tableCellProperties54.Append(shading18);
            tableCellProperties54.Append(tableCellVerticalAlignment54);

            Paragraph paragraph115 = new Paragraph() { RsidParagraphMarkRevision = "00C709B4", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties115 = new ParagraphProperties();
            Indentation indentation115 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification109 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties115 = new ParagraphMarkRunProperties();
            Bold bold39 = new Bold();
            Kern kern862 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize892 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript892 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties115.Append(bold39);
            paragraphMarkRunProperties115.Append(kern862);
            paragraphMarkRunProperties115.Append(fontSize892);
            paragraphMarkRunProperties115.Append(fontSizeComplexScript892);

            paragraphProperties115.Append(indentation115);
            paragraphProperties115.Append(justification109);
            paragraphProperties115.Append(paragraphMarkRunProperties115);
            BookmarkStart bookmarkStart1 = new BookmarkStart() { Name = "_GoBack", ColumnFirst = 5, ColumnLast = 5, Id = "0" };

            Run run778 = new Run() { RsidRunProperties = "00C709B4" };

            RunProperties runProperties778 = new RunProperties();
            RunFonts runFonts602 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold40 = new Bold();
            Kern kern863 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize893 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript893 = new FontSizeComplexScript() { Val = "20" };

            runProperties778.Append(runFonts602);
            runProperties778.Append(bold40);
            runProperties778.Append(kern863);
            runProperties778.Append(fontSize893);
            runProperties778.Append(fontSizeComplexScript893);
            Text text58 = new Text();
            text58.Text = "学习压力";

            run778.Append(runProperties778);
            run778.Append(text58);

            paragraph115.Append(paragraphProperties115);
            paragraph115.Append(bookmarkStart1);
            paragraph115.Append(run778);

            tableCell54.Append(tableCellProperties54);
            tableCell54.Append(paragraph115);

            TableCell tableCell55 = new TableCell();

            TableCellProperties tableCellProperties55 = new TableCellProperties();
            TableCellWidth tableCellWidth55 = new TableCellWidth() { Width = "4453", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment55 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties55.Append(tableCellWidth55);
            tableCellProperties55.Append(tableCellVerticalAlignment55);

            Paragraph paragraph116 = new Paragraph() { RsidParagraphMarkRevision = "00C709B4", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties116 = new ParagraphProperties();
            Indentation indentation116 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };

            ParagraphMarkRunProperties paragraphMarkRunProperties116 = new ParagraphMarkRunProperties();
            Kern kern864 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize894 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript894 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties116.Append(kern864);
            paragraphMarkRunProperties116.Append(fontSize894);
            paragraphMarkRunProperties116.Append(fontSizeComplexScript894);

            paragraphProperties116.Append(indentation116);
            paragraphProperties116.Append(paragraphMarkRunProperties116);

            Run run779 = new Run() { RsidRunProperties = "00C709B4" };

            RunProperties runProperties779 = new RunProperties();
            RunFonts runFonts603 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Kern kern865 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize895 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript895 = new FontSizeComplexScript() { Val = "20" };

            runProperties779.Append(runFonts603);
            runProperties779.Append(kern865);
            runProperties779.Append(fontSize895);
            runProperties779.Append(fontSizeComplexScript895);
            Text text59 = new Text();
            text59.Text = "感到学习负担重，怕老师提问，讨厌做作业，讨厌上学，害怕和讨厌考试";

            run779.Append(runProperties779);
            run779.Append(text59);

            paragraph116.Append(paragraphProperties116);
            paragraph116.Append(run779);

            tableCell55.Append(tableCellProperties55);
            tableCell55.Append(paragraph116);

            TableCell tableCell56 = new TableCell();

            TableCellProperties tableCellProperties56 = new TableCellProperties();
            TableCellWidth tableCellWidth56 = new TableCellWidth() { Width = "756", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment56 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties56.Append(tableCellWidth56);
            tableCellProperties56.Append(tableCellVerticalAlignment56);

            Paragraph paragraph117 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties117 = new ParagraphProperties();
            Indentation indentation117 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification110 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties117 = new ParagraphMarkRunProperties();
            FontSize fontSize896 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript896 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties117.Append(fontSize896);
            paragraphMarkRunProperties117.Append(fontSizeComplexScript896);

            paragraphProperties117.Append(indentation117);
            paragraphProperties117.Append(justification110);
            paragraphProperties117.Append(paragraphMarkRunProperties117);

            paragraph117.Append(paragraphProperties117);

            Paragraph paragraph118 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties118 = new ParagraphProperties();
            Indentation indentation118 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification111 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties118 = new ParagraphMarkRunProperties();
            Kern kern889 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize920 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript920 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties118.Append(kern889);
            paragraphMarkRunProperties118.Append(fontSize920);
            paragraphMarkRunProperties118.Append(fontSizeComplexScript920);

            paragraphProperties118.Append(indentation118);
            paragraphProperties118.Append(justification111);
            paragraphProperties118.Append(paragraphMarkRunProperties118);

            Run run803 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties803 = new RunProperties();
            Kern kern890 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize921 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript921 = new FontSizeComplexScript() { Val = "20" };

            runProperties803.Append(kern890);
            runProperties803.Append(fontSize921);
            runProperties803.Append(fontSizeComplexScript921);
            Text text60 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text60.Text = content["学习压力"].Item2 == "没有" ? "★" : " ";

            run803.Append(runProperties803);
            run803.Append(text60);

            paragraph118.Append(paragraphProperties118);
            paragraph118.Append(run803);

            Paragraph paragraph119 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties119 = new ParagraphProperties();
            Indentation indentation119 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification112 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties119 = new ParagraphMarkRunProperties();
            Kern kern891 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize922 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript922 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties119.Append(kern891);
            paragraphMarkRunProperties119.Append(fontSize922);
            paragraphMarkRunProperties119.Append(fontSizeComplexScript922);

            paragraphProperties119.Append(indentation119);
            paragraphProperties119.Append(justification112);
            paragraphProperties119.Append(paragraphMarkRunProperties119);

            paragraph119.Append(paragraphProperties119);

            tableCell56.Append(tableCellProperties56);
            tableCell56.Append(paragraph117);
            tableCell56.Append(paragraph118);
            tableCell56.Append(paragraph119);

            TableCell tableCell57 = new TableCell();

            TableCellProperties tableCellProperties57 = new TableCellProperties();
            TableCellWidth tableCellWidth57 = new TableCellWidth() { Width = "755", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment57 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties57.Append(tableCellWidth57);
            tableCellProperties57.Append(tableCellVerticalAlignment57);

            Paragraph paragraph120 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties120 = new ParagraphProperties();
            Indentation indentation120 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification113 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties120 = new ParagraphMarkRunProperties();
            FontSize fontSize924 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript924 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties120.Append(fontSize924);
            paragraphMarkRunProperties120.Append(fontSizeComplexScript924);

            paragraphProperties120.Append(indentation120);
            paragraphProperties120.Append(justification113);
            paragraphProperties120.Append(paragraphMarkRunProperties120);

            paragraph120.Append(paragraphProperties120);

            Paragraph paragraph121 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties121 = new ParagraphProperties();
            Indentation indentation121 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification114 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties121 = new ParagraphMarkRunProperties();
            Kern kern916 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize948 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript948 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties121.Append(kern916);
            paragraphMarkRunProperties121.Append(fontSize948);
            paragraphMarkRunProperties121.Append(fontSizeComplexScript948);

            paragraphProperties121.Append(indentation121);
            paragraphProperties121.Append(justification114);
            paragraphProperties121.Append(paragraphMarkRunProperties121);

            Run run828 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties828 = new RunProperties();
            RunFonts runFonts642 = new RunFonts() { Ascii = "Segoe UI Symbol", HighAnsi = "Segoe UI Symbol", ComplexScript = "Segoe UI Symbol" };
            Kern kern917 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize949 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript949 = new FontSizeComplexScript() { Val = "20" };

            runProperties828.Append(runFonts642);
            runProperties828.Append(kern917);
            runProperties828.Append(fontSize949);
            runProperties828.Append(fontSizeComplexScript949);
            Text text61 = new Text();
            text61.Text = content["学习压力"].Item2 == "轻度" ? "★" : " ";

            run828.Append(runProperties828);
            run828.Append(text61);

            paragraph121.Append(paragraphProperties121);
            paragraph121.Append(run828);

            Paragraph paragraph122 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties122 = new ParagraphProperties();
            Indentation indentation122 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification115 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties122 = new ParagraphMarkRunProperties();
            Kern kern918 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize950 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript950 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties122.Append(kern918);
            paragraphMarkRunProperties122.Append(fontSize950);
            paragraphMarkRunProperties122.Append(fontSizeComplexScript950);

            paragraphProperties122.Append(indentation122);
            paragraphProperties122.Append(justification115);
            paragraphProperties122.Append(paragraphMarkRunProperties122);

            paragraph122.Append(paragraphProperties122);

            tableCell57.Append(tableCellProperties57);
            tableCell57.Append(paragraph120);
            tableCell57.Append(paragraph121);
            tableCell57.Append(paragraph122);

            TableCell tableCell58 = new TableCell();

            TableCellProperties tableCellProperties58 = new TableCellProperties();
            TableCellWidth tableCellWidth58 = new TableCellWidth() { Width = "756", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment58 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties58.Append(tableCellWidth58);
            tableCellProperties58.Append(tableCellVerticalAlignment58);

            Paragraph paragraph123 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties123 = new ParagraphProperties();
            Indentation indentation123 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification116 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties123 = new ParagraphMarkRunProperties();
            FontSize fontSize952 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript952 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties123.Append(fontSize952);
            paragraphMarkRunProperties123.Append(fontSizeComplexScript952);

            paragraphProperties123.Append(indentation123);
            paragraphProperties123.Append(justification116);
            paragraphProperties123.Append(paragraphMarkRunProperties123);
            paragraph123.Append(paragraphProperties123);

            Paragraph paragraph124 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties124 = new ParagraphProperties();
            Indentation indentation124 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification117 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties124 = new ParagraphMarkRunProperties();
            Kern kern943 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize976 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript976 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties124.Append(kern943);
            paragraphMarkRunProperties124.Append(fontSize976);
            paragraphMarkRunProperties124.Append(fontSizeComplexScript976);

            paragraphProperties124.Append(indentation124);
            paragraphProperties124.Append(justification117);
            paragraphProperties124.Append(paragraphMarkRunProperties124);

            Run run853 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties853 = new RunProperties();
            Kern kern944 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize977 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript977 = new FontSizeComplexScript() { Val = "20" };

            runProperties853.Append(kern944);
            runProperties853.Append(fontSize977);
            runProperties853.Append(fontSizeComplexScript977);
            Text text62 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text62.Text = content["学习压力"].Item2 == "中度" ? "★" : " ";

            run853.Append(runProperties853);
            run853.Append(text62);

            paragraph124.Append(paragraphProperties124);
            paragraph124.Append(run853);

            Paragraph paragraph125 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties125 = new ParagraphProperties();
            Indentation indentation125 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification118 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties125 = new ParagraphMarkRunProperties();
            Kern kern945 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize978 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript978 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties125.Append(kern945);
            paragraphMarkRunProperties125.Append(fontSize978);
            paragraphMarkRunProperties125.Append(fontSizeComplexScript978);

            paragraphProperties125.Append(indentation125);
            paragraphProperties125.Append(justification118);
            paragraphProperties125.Append(paragraphMarkRunProperties125);

            paragraph125.Append(paragraphProperties125);

            tableCell58.Append(tableCellProperties58);
            tableCell58.Append(paragraph123);
            tableCell58.Append(paragraph124);
            tableCell58.Append(paragraph125);

            TableCell tableCell59 = new TableCell();

            TableCellProperties tableCellProperties59 = new TableCellProperties();
            TableCellWidth tableCellWidth59 = new TableCellWidth() { Width = "756", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment59 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties59.Append(tableCellWidth59);
            tableCellProperties59.Append(tableCellVerticalAlignment59);

            Paragraph paragraph126 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties126 = new ParagraphProperties();
            Indentation indentation126 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification119 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties126 = new ParagraphMarkRunProperties();
            FontSize fontSize980 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript980 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties126.Append(fontSize980);
            paragraphMarkRunProperties126.Append(fontSizeComplexScript980);

            paragraphProperties126.Append(indentation126);
            paragraphProperties126.Append(justification119);
            paragraphProperties126.Append(paragraphMarkRunProperties126);

            paragraph126.Append(paragraphProperties126);

            Paragraph paragraph127 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties127 = new ParagraphProperties();
            Indentation indentation127 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification120 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties127 = new ParagraphMarkRunProperties();
            Kern kern970 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1004 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1004 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties127.Append(kern970);
            paragraphMarkRunProperties127.Append(fontSize1004);
            paragraphMarkRunProperties127.Append(fontSizeComplexScript1004);

            paragraphProperties127.Append(indentation127);
            paragraphProperties127.Append(justification120);
            paragraphProperties127.Append(paragraphMarkRunProperties127);

            Run run878 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties878 = new RunProperties();
            Kern kern971 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1005 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1005 = new FontSizeComplexScript() { Val = "20" };

            runProperties878.Append(kern971);
            runProperties878.Append(fontSize1005);
            runProperties878.Append(fontSizeComplexScript1005);
            Text text63 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text63.Text = content["学习压力"].Item2 == "严重" ? "★" : " ";

            run878.Append(runProperties878);
            run878.Append(text63);

            paragraph127.Append(paragraphProperties127);
            paragraph127.Append(run878);

            Paragraph paragraph128 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties128 = new ParagraphProperties();
            Indentation indentation128 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification121 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties128 = new ParagraphMarkRunProperties();
            Kern kern972 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1006 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript1006 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties128.Append(kern972);
            paragraphMarkRunProperties128.Append(fontSize1006);
            paragraphMarkRunProperties128.Append(fontSizeComplexScript1006);

            paragraphProperties128.Append(indentation128);
            paragraphProperties128.Append(justification121);
            paragraphProperties128.Append(paragraphMarkRunProperties128);

            paragraph128.Append(paragraphProperties128);

            tableCell59.Append(tableCellProperties59);
            tableCell59.Append(paragraph126);
            tableCell59.Append(paragraph127);
            tableCell59.Append(paragraph128);

            TableCell tableCell60 = new TableCell();

            TableCellProperties tableCellProperties60 = new TableCellProperties();
            TableCellWidth tableCellWidth60 = new TableCellWidth() { Width = "827", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment60 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties60.Append(tableCellWidth60);
            tableCellProperties60.Append(tableCellVerticalAlignment60);

            Paragraph paragraph129 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties129 = new ParagraphProperties();
            Indentation indentation129 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification122 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties129 = new ParagraphMarkRunProperties();
            FontSize fontSize1008 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript1008 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties129.Append(fontSize1008);
            paragraphMarkRunProperties129.Append(fontSizeComplexScript1008);

            paragraphProperties129.Append(indentation129);
            paragraphProperties129.Append(justification122);
            paragraphProperties129.Append(paragraphMarkRunProperties129);

            paragraph129.Append(paragraphProperties129);

            Paragraph paragraph130 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties130 = new ParagraphProperties();
            Indentation indentation130 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification123 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties130 = new ParagraphMarkRunProperties();
            Kern kern997 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1032 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1032 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties130.Append(kern997);
            paragraphMarkRunProperties130.Append(fontSize1032);
            paragraphMarkRunProperties130.Append(fontSizeComplexScript1032);

            paragraphProperties130.Append(indentation130);
            paragraphProperties130.Append(justification123);
            paragraphProperties130.Append(paragraphMarkRunProperties130);

            Run run903 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties903 = new RunProperties();
            Kern kern998 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1033 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1033 = new FontSizeComplexScript() { Val = "20" };

            runProperties903.Append(kern998);
            runProperties903.Append(fontSize1033);
            runProperties903.Append(fontSizeComplexScript1033);
            Text text64 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text64.Text = content["学习压力"].Item2 == "严重" ? "★" : " ";

            run903.Append(runProperties903);
            run903.Append(text64);

            paragraph130.Append(paragraphProperties130);
            paragraph130.Append(run903);

            Paragraph paragraph131 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties131 = new ParagraphProperties();
            Indentation indentation131 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification124 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties131 = new ParagraphMarkRunProperties();
            Kern kern999 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1034 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript1034 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties131.Append(kern999);
            paragraphMarkRunProperties131.Append(fontSize1034);
            paragraphMarkRunProperties131.Append(fontSizeComplexScript1034);

            paragraphProperties131.Append(indentation131);
            paragraphProperties131.Append(justification124);
            paragraphProperties131.Append(paragraphMarkRunProperties131);
            paragraph131.Append(paragraphProperties131);

            tableCell60.Append(tableCellProperties60);
            tableCell60.Append(paragraph129);
            tableCell60.Append(paragraph130);
            tableCell60.Append(paragraph131);

            tableRow10.Append(tableRowProperties10);
            tableRow10.Append(tableCell54);
            tableRow10.Append(tableCell55);
            tableRow10.Append(tableCell56);
            tableRow10.Append(tableCell57);
            tableRow10.Append(tableCell58);
            tableRow10.Append(tableCell59);
            tableRow10.Append(tableCell60);
            BookmarkEnd bookmarkEnd1 = new BookmarkEnd() { Id = "0" };

            TableRow tableRow11 = new TableRow() { RsidTableRowMarkRevision = "00C709B4", RsidTableRowAddition = "00D061CC", RsidTableRowProperties = "009E6323" };

            TableRowProperties tableRowProperties11 = new TableRowProperties();
            TableRowHeight tableRowHeight11 = new TableRowHeight() { Val = (UInt32Value)800U };

            tableRowProperties11.Append(tableRowHeight11);

            TableCell tableCell61 = new TableCell();

            TableCellProperties tableCellProperties61 = new TableCellProperties();
            TableCellWidth tableCellWidth61 = new TableCellWidth() { Width = "1678", Type = TableWidthUnitValues.Dxa };
            Shading shading19 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "D6E3BC" };
            TableCellVerticalAlignment tableCellVerticalAlignment61 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties61.Append(tableCellWidth61);
            tableCellProperties61.Append(shading19);
            tableCellProperties61.Append(tableCellVerticalAlignment61);

            Paragraph paragraph132 = new Paragraph() { RsidParagraphMarkRevision = "00C709B4", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties132 = new ParagraphProperties();
            Indentation indentation132 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification125 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties132 = new ParagraphMarkRunProperties();
            Bold bold41 = new Bold();
            Kern kern1001 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1036 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript1036 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties132.Append(bold41);
            paragraphMarkRunProperties132.Append(kern1001);
            paragraphMarkRunProperties132.Append(fontSize1036);
            paragraphMarkRunProperties132.Append(fontSizeComplexScript1036);

            paragraphProperties132.Append(indentation132);
            paragraphProperties132.Append(justification125);
            paragraphProperties132.Append(paragraphMarkRunProperties132);

            Run run905 = new Run() { RsidRunProperties = "00C709B4" };

            RunProperties runProperties905 = new RunProperties();
            RunFonts runFonts700 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold42 = new Bold();
            Kern kern1002 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1037 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript1037 = new FontSizeComplexScript() { Val = "20" };

            runProperties905.Append(runFonts700);
            runProperties905.Append(bold42);
            runProperties905.Append(kern1002);
            runProperties905.Append(fontSize1037);
            runProperties905.Append(fontSizeComplexScript1037);
            LastRenderedPageBreak lastRenderedPageBreak1 = new LastRenderedPageBreak();
            Text text65 = new Text();
            text65.Text = "适应不良";

            run905.Append(runProperties905);
            run905.Append(lastRenderedPageBreak1);
            run905.Append(text65);

            paragraph132.Append(paragraphProperties132);
            paragraph132.Append(run905);

            tableCell61.Append(tableCellProperties61);
            tableCell61.Append(paragraph132);

            TableCell tableCell62 = new TableCell();

            TableCellProperties tableCellProperties62 = new TableCellProperties();
            TableCellWidth tableCellWidth62 = new TableCellWidth() { Width = "4453", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment62 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties62.Append(tableCellWidth62);
            tableCellProperties62.Append(tableCellVerticalAlignment62);

            Paragraph paragraph133 = new Paragraph() { RsidParagraphMarkRevision = "00C709B4", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties133 = new ParagraphProperties();
            Indentation indentation133 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };

            ParagraphMarkRunProperties paragraphMarkRunProperties133 = new ParagraphMarkRunProperties();
            Kern kern1003 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1038 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1038 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties133.Append(kern1003);
            paragraphMarkRunProperties133.Append(fontSize1038);
            paragraphMarkRunProperties133.Append(fontSizeComplexScript1038);

            paragraphProperties133.Append(indentation133);
            paragraphProperties133.Append(paragraphMarkRunProperties133);

            Run run906 = new Run() { RsidRunProperties = "00C709B4" };

            RunProperties runProperties906 = new RunProperties();
            RunFonts runFonts701 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Kern kern1004 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1039 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1039 = new FontSizeComplexScript() { Val = "20" };

            runProperties906.Append(runFonts701);
            runProperties906.Append(kern1004);
            runProperties906.Append(fontSize1039);
            runProperties906.Append(fontSizeComplexScript1039);
            Text text66 = new Text();
            text66.Text = "对学校适应不良，不愿意参加课外活动，不适应老师教学方法，不适应家里学习环境";

            run906.Append(runProperties906);
            run906.Append(text66);

            paragraph133.Append(paragraphProperties133);
            paragraph133.Append(run906);

            tableCell62.Append(tableCellProperties62);
            tableCell62.Append(paragraph133);

            TableCell tableCell63 = new TableCell();

            TableCellProperties tableCellProperties63 = new TableCellProperties();
            TableCellWidth tableCellWidth63 = new TableCellWidth() { Width = "756", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment63 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties63.Append(tableCellWidth63);
            tableCellProperties63.Append(tableCellVerticalAlignment63);

            Paragraph paragraph134 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties134 = new ParagraphProperties();
            Indentation indentation134 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification126 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties134 = new ParagraphMarkRunProperties();
            FontSize fontSize1040 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript1040 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties134.Append(fontSize1040);
            paragraphMarkRunProperties134.Append(fontSizeComplexScript1040);

            paragraphProperties134.Append(indentation134);
            paragraphProperties134.Append(justification126);
            paragraphProperties134.Append(paragraphMarkRunProperties134);

            paragraph134.Append(paragraphProperties134);

            Paragraph paragraph135 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties135 = new ParagraphProperties();
            Indentation indentation135 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification127 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties135 = new ParagraphMarkRunProperties();
            Kern kern1028 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1064 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1064 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties135.Append(kern1028);
            paragraphMarkRunProperties135.Append(fontSize1064);
            paragraphMarkRunProperties135.Append(fontSizeComplexScript1064);

            paragraphProperties135.Append(indentation135);
            paragraphProperties135.Append(justification127);
            paragraphProperties135.Append(paragraphMarkRunProperties135);

            Run run930 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties930 = new RunProperties();
            Kern kern1029 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1065 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1065 = new FontSizeComplexScript() { Val = "20" };

            runProperties930.Append(kern1029);
            runProperties930.Append(fontSize1065);
            runProperties930.Append(fontSizeComplexScript1065);
            Text text67 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text67.Text = content["适应不良"].Item2 == "没有" ? "★" : " ";

            run930.Append(runProperties930);
            run930.Append(text67);

            paragraph135.Append(paragraphProperties135);
            paragraph135.Append(run930);

            Paragraph paragraph136 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties136 = new ParagraphProperties();
            Indentation indentation136 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification128 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties136 = new ParagraphMarkRunProperties();
            Kern kern1030 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1066 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript1066 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties136.Append(kern1030);
            paragraphMarkRunProperties136.Append(fontSize1066);
            paragraphMarkRunProperties136.Append(fontSizeComplexScript1066);

            paragraphProperties136.Append(indentation136);
            paragraphProperties136.Append(justification128);
            paragraphProperties136.Append(paragraphMarkRunProperties136);

            paragraph136.Append(paragraphProperties136);

            tableCell63.Append(tableCellProperties63);
            tableCell63.Append(paragraph134);
            tableCell63.Append(paragraph135);
            tableCell63.Append(paragraph136);

            TableCell tableCell64 = new TableCell();

            TableCellProperties tableCellProperties64 = new TableCellProperties();
            TableCellWidth tableCellWidth64 = new TableCellWidth() { Width = "755", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment64 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties64.Append(tableCellWidth64);
            tableCellProperties64.Append(tableCellVerticalAlignment64);

            Paragraph paragraph137 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties137 = new ParagraphProperties();
            Indentation indentation137 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification129 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties137 = new ParagraphMarkRunProperties();
            FontSize fontSize1068 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript1068 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties137.Append(fontSize1068);
            paragraphMarkRunProperties137.Append(fontSizeComplexScript1068);

            paragraphProperties137.Append(indentation137);
            paragraphProperties137.Append(justification129);
            paragraphProperties137.Append(paragraphMarkRunProperties137);
            paragraph137.Append(paragraphProperties137);

            Paragraph paragraph138 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties138 = new ParagraphProperties();
            Indentation indentation138 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification130 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties138 = new ParagraphMarkRunProperties();
            Kern kern1055 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1092 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1092 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties138.Append(kern1055);
            paragraphMarkRunProperties138.Append(fontSize1092);
            paragraphMarkRunProperties138.Append(fontSizeComplexScript1092);

            paragraphProperties138.Append(indentation138);
            paragraphProperties138.Append(justification130);
            paragraphProperties138.Append(paragraphMarkRunProperties138);

            Run run955 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties955 = new RunProperties();
            RunFonts runFonts740 = new RunFonts() { Ascii = "Segoe UI Symbol", HighAnsi = "Segoe UI Symbol", ComplexScript = "Segoe UI Symbol" };
            Kern kern1056 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1093 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1093 = new FontSizeComplexScript() { Val = "20" };

            runProperties955.Append(runFonts740);
            runProperties955.Append(kern1056);
            runProperties955.Append(fontSize1093);
            runProperties955.Append(fontSizeComplexScript1093);
            Text text68 = new Text();
            text68.Text = content["轻度"].Item2 == "没有" ? "★" : " ";

            run955.Append(runProperties955);
            run955.Append(text68);

            paragraph138.Append(paragraphProperties138);
            paragraph138.Append(run955);

            Paragraph paragraph139 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties139 = new ParagraphProperties();
            Indentation indentation139 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification131 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties139 = new ParagraphMarkRunProperties();
            Kern kern1057 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1094 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript1094 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties139.Append(kern1057);
            paragraphMarkRunProperties139.Append(fontSize1094);
            paragraphMarkRunProperties139.Append(fontSizeComplexScript1094);

            paragraphProperties139.Append(indentation139);
            paragraphProperties139.Append(justification131);
            paragraphProperties139.Append(paragraphMarkRunProperties139);

            paragraph139.Append(paragraphProperties139);

            tableCell64.Append(tableCellProperties64);
            tableCell64.Append(paragraph137);
            tableCell64.Append(paragraph138);
            tableCell64.Append(paragraph139);

            TableCell tableCell65 = new TableCell();

            TableCellProperties tableCellProperties65 = new TableCellProperties();
            TableCellWidth tableCellWidth65 = new TableCellWidth() { Width = "756", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment65 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties65.Append(tableCellWidth65);
            tableCellProperties65.Append(tableCellVerticalAlignment65);

            Paragraph paragraph140 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties140 = new ParagraphProperties();
            Indentation indentation140 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification132 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties140 = new ParagraphMarkRunProperties();
            FontSize fontSize1096 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript1096 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties140.Append(fontSize1096);
            paragraphMarkRunProperties140.Append(fontSizeComplexScript1096);

            paragraphProperties140.Append(indentation140);
            paragraphProperties140.Append(justification132);
            paragraphProperties140.Append(paragraphMarkRunProperties140);
            paragraph140.Append(paragraphProperties140);

            Paragraph paragraph141 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties141 = new ParagraphProperties();
            Indentation indentation141 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification133 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties141 = new ParagraphMarkRunProperties();
            Kern kern1082 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1120 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1120 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties141.Append(kern1082);
            paragraphMarkRunProperties141.Append(fontSize1120);
            paragraphMarkRunProperties141.Append(fontSizeComplexScript1120);

            paragraphProperties141.Append(indentation141);
            paragraphProperties141.Append(justification133);
            paragraphProperties141.Append(paragraphMarkRunProperties141);

            Run run980 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties980 = new RunProperties();
            Kern kern1083 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1121 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1121 = new FontSizeComplexScript() { Val = "20" };

            runProperties980.Append(kern1083);
            runProperties980.Append(fontSize1121);
            runProperties980.Append(fontSizeComplexScript1121);
            Text text69 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text69.Text = content["适应不良"].Item2 == "中度" ? "★" : " ";

            run980.Append(runProperties980);
            run980.Append(text69);

            paragraph141.Append(paragraphProperties141);
            paragraph141.Append(run980);

            Paragraph paragraph142 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties142 = new ParagraphProperties();
            Indentation indentation142 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification134 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties142 = new ParagraphMarkRunProperties();
            Kern kern1084 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1122 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript1122 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties142.Append(kern1084);
            paragraphMarkRunProperties142.Append(fontSize1122);
            paragraphMarkRunProperties142.Append(fontSizeComplexScript1122);

            paragraphProperties142.Append(indentation142);
            paragraphProperties142.Append(justification134);
            paragraphProperties142.Append(paragraphMarkRunProperties142);

            paragraph142.Append(paragraphProperties142);

            tableCell65.Append(tableCellProperties65);
            tableCell65.Append(paragraph140);
            tableCell65.Append(paragraph141);
            tableCell65.Append(paragraph142);

            TableCell tableCell66 = new TableCell();

            TableCellProperties tableCellProperties66 = new TableCellProperties();
            TableCellWidth tableCellWidth66 = new TableCellWidth() { Width = "756", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment66 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties66.Append(tableCellWidth66);
            tableCellProperties66.Append(tableCellVerticalAlignment66);

            Paragraph paragraph143 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties143 = new ParagraphProperties();
            Indentation indentation143 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification135 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties143 = new ParagraphMarkRunProperties();
            FontSize fontSize1124 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript1124 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties143.Append(fontSize1124);
            paragraphMarkRunProperties143.Append(fontSizeComplexScript1124);

            paragraphProperties143.Append(indentation143);
            paragraphProperties143.Append(justification135);
            paragraphProperties143.Append(paragraphMarkRunProperties143);

            paragraph143.Append(paragraphProperties143);

            Paragraph paragraph144 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties144 = new ParagraphProperties();
            Indentation indentation144 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification136 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties144 = new ParagraphMarkRunProperties();
            Kern kern1109 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1148 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1148 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties144.Append(kern1109);
            paragraphMarkRunProperties144.Append(fontSize1148);
            paragraphMarkRunProperties144.Append(fontSizeComplexScript1148);

            paragraphProperties144.Append(indentation144);
            paragraphProperties144.Append(justification136);
            paragraphProperties144.Append(paragraphMarkRunProperties144);

            Run run1005 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties1005 = new RunProperties();
            Kern kern1110 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1149 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1149 = new FontSizeComplexScript() { Val = "20" };

            runProperties1005.Append(kern1110);
            runProperties1005.Append(fontSize1149);
            runProperties1005.Append(fontSizeComplexScript1149);
            Text text70 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text70.Text = content["适应不良"].Item2 == "较重" ? "★" : " ";

            run1005.Append(runProperties1005);
            run1005.Append(text70);

            paragraph144.Append(paragraphProperties144);
            paragraph144.Append(run1005);

            Paragraph paragraph145 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties145 = new ParagraphProperties();
            Indentation indentation145 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification137 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties145 = new ParagraphMarkRunProperties();
            Kern kern1111 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1150 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript1150 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties145.Append(kern1111);
            paragraphMarkRunProperties145.Append(fontSize1150);
            paragraphMarkRunProperties145.Append(fontSizeComplexScript1150);

            paragraphProperties145.Append(indentation145);
            paragraphProperties145.Append(justification137);
            paragraphProperties145.Append(paragraphMarkRunProperties145);

            paragraph145.Append(paragraphProperties145);

            tableCell66.Append(tableCellProperties66);
            tableCell66.Append(paragraph143);
            tableCell66.Append(paragraph144);
            tableCell66.Append(paragraph145);

            TableCell tableCell67 = new TableCell();

            TableCellProperties tableCellProperties67 = new TableCellProperties();
            TableCellWidth tableCellWidth67 = new TableCellWidth() { Width = "827", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment67 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties67.Append(tableCellWidth67);
            tableCellProperties67.Append(tableCellVerticalAlignment67);

            Paragraph paragraph146 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties146 = new ParagraphProperties();
            Indentation indentation146 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification138 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties146 = new ParagraphMarkRunProperties();
            FontSize fontSize1152 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript1152 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties146.Append(fontSize1152);
            paragraphMarkRunProperties146.Append(fontSizeComplexScript1152);

            paragraphProperties146.Append(indentation146);
            paragraphProperties146.Append(justification138);
            paragraphProperties146.Append(paragraphMarkRunProperties146);
            paragraph146.Append(paragraphProperties146);

            Paragraph paragraph147 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties147 = new ParagraphProperties();
            Indentation indentation147 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification139 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties147 = new ParagraphMarkRunProperties();
            Kern kern1136 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1176 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1176 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties147.Append(kern1136);
            paragraphMarkRunProperties147.Append(fontSize1176);
            paragraphMarkRunProperties147.Append(fontSizeComplexScript1176);

            paragraphProperties147.Append(indentation147);
            paragraphProperties147.Append(justification139);
            paragraphProperties147.Append(paragraphMarkRunProperties147);

            Run run1030 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties1030 = new RunProperties();
            Kern kern1137 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1177 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1177 = new FontSizeComplexScript() { Val = "20" };

            runProperties1030.Append(kern1137);
            runProperties1030.Append(fontSize1177);
            runProperties1030.Append(fontSizeComplexScript1177);
            Text text71 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text71.Text = content["适应不良"].Item2 == "严重" ? "★" : " ";

            run1030.Append(runProperties1030);
            run1030.Append(text71);

            paragraph147.Append(paragraphProperties147);
            paragraph147.Append(run1030);

            Paragraph paragraph148 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties148 = new ParagraphProperties();
            Indentation indentation148 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification140 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties148 = new ParagraphMarkRunProperties();
            Kern kern1138 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1178 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript1178 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties148.Append(kern1138);
            paragraphMarkRunProperties148.Append(fontSize1178);
            paragraphMarkRunProperties148.Append(fontSizeComplexScript1178);

            paragraphProperties148.Append(indentation148);
            paragraphProperties148.Append(justification140);
            paragraphProperties148.Append(paragraphMarkRunProperties148);

            paragraph148.Append(paragraphProperties148);

            tableCell67.Append(tableCellProperties67);
            tableCell67.Append(paragraph146);
            tableCell67.Append(paragraph147);
            tableCell67.Append(paragraph148);

            tableRow11.Append(tableRowProperties11);
            tableRow11.Append(tableCell61);
            tableRow11.Append(tableCell62);
            tableRow11.Append(tableCell63);
            tableRow11.Append(tableCell64);
            tableRow11.Append(tableCell65);
            tableRow11.Append(tableCell66);
            tableRow11.Append(tableCell67);

            TableRow tableRow12 = new TableRow() { RsidTableRowMarkRevision = "00C709B4", RsidTableRowAddition = "00D061CC", RsidTableRowProperties = "009E6323" };

            TableRowProperties tableRowProperties12 = new TableRowProperties();
            TableRowHeight tableRowHeight12 = new TableRowHeight() { Val = (UInt32Value)800U };

            tableRowProperties12.Append(tableRowHeight12);

            TableCell tableCell68 = new TableCell();

            TableCellProperties tableCellProperties68 = new TableCellProperties();
            TableCellWidth tableCellWidth68 = new TableCellWidth() { Width = "1678", Type = TableWidthUnitValues.Dxa };
            Shading shading20 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "D6E3BC" };
            TableCellVerticalAlignment tableCellVerticalAlignment68 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties68.Append(tableCellWidth68);
            tableCellProperties68.Append(shading20);
            tableCellProperties68.Append(tableCellVerticalAlignment68);

            Paragraph paragraph149 = new Paragraph() { RsidParagraphMarkRevision = "00C709B4", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties149 = new ParagraphProperties();
            Indentation indentation149 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification141 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties149 = new ParagraphMarkRunProperties();
            Bold bold43 = new Bold();
            Kern kern1140 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1180 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript1180 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties149.Append(bold43);
            paragraphMarkRunProperties149.Append(kern1140);
            paragraphMarkRunProperties149.Append(fontSize1180);
            paragraphMarkRunProperties149.Append(fontSizeComplexScript1180);

            paragraphProperties149.Append(indentation149);
            paragraphProperties149.Append(justification141);
            paragraphProperties149.Append(paragraphMarkRunProperties149);

            Run run1032 = new Run() { RsidRunProperties = "00C709B4" };

            RunProperties runProperties1032 = new RunProperties();
            RunFonts runFonts798 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold44 = new Bold();
            Kern kern1141 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1181 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript1181 = new FontSizeComplexScript() { Val = "20" };

            runProperties1032.Append(runFonts798);
            runProperties1032.Append(bold44);
            runProperties1032.Append(kern1141);
            runProperties1032.Append(fontSize1181);
            runProperties1032.Append(fontSizeComplexScript1181);
            Text text72 = new Text();
            text72.Text = "情绪不平衡";

            run1032.Append(runProperties1032);
            run1032.Append(text72);

            paragraph149.Append(paragraphProperties149);
            paragraph149.Append(run1032);

            tableCell68.Append(tableCellProperties68);
            tableCell68.Append(paragraph149);

            TableCell tableCell69 = new TableCell();

            TableCellProperties tableCellProperties69 = new TableCellProperties();
            TableCellWidth tableCellWidth69 = new TableCellWidth() { Width = "4453", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment69 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties69.Append(tableCellWidth69);
            tableCellProperties69.Append(tableCellVerticalAlignment69);

            Paragraph paragraph151 = new Paragraph() { RsidParagraphMarkRevision = "00C709B4", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties151 = new ParagraphProperties();
            Indentation indentation151 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };

            ParagraphMarkRunProperties paragraphMarkRunProperties151 = new ParagraphMarkRunProperties();
            Kern kern1144 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1184 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1184 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties151.Append(kern1144);
            paragraphMarkRunProperties151.Append(fontSize1184);
            paragraphMarkRunProperties151.Append(fontSizeComplexScript1184);

            paragraphProperties151.Append(indentation151);
            paragraphProperties151.Append(paragraphMarkRunProperties151);

            Run run1034 = new Run() { RsidRunProperties = "00C709B4" };

            RunProperties runProperties1034 = new RunProperties();
            RunFonts runFonts800 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Kern kern1145 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1185 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1185 = new FontSizeComplexScript() { Val = "20" };

            runProperties1034.Append(runFonts800);
            runProperties1034.Append(kern1145);
            runProperties1034.Append(fontSize1185);
            runProperties1034.Append(fontSizeComplexScript1185);
            Text text74 = new Text();
            text74.Text = "情绪不稳定，对老师和同学以及父母态度多变，学习成绩忽高忽低";

            run1034.Append(runProperties1034);
            run1034.Append(text74);

            paragraph151.Append(paragraphProperties151);
            paragraph151.Append(run1034);

            tableCell69.Append(tableCellProperties69);
            tableCell69.Append(paragraph151);

            TableCell tableCell70 = new TableCell();

            TableCellProperties tableCellProperties70 = new TableCellProperties();
            TableCellWidth tableCellWidth70 = new TableCellWidth() { Width = "756", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment70 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties70.Append(tableCellWidth70);
            tableCellProperties70.Append(tableCellVerticalAlignment70);

            Paragraph paragraph152 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties152 = new ParagraphProperties();
            Indentation indentation152 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification143 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties152 = new ParagraphMarkRunProperties();
            FontSize fontSize1186 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript1186 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties152.Append(fontSize1186);
            paragraphMarkRunProperties152.Append(fontSizeComplexScript1186);

            paragraphProperties152.Append(indentation152);
            paragraphProperties152.Append(justification143);
            paragraphProperties152.Append(paragraphMarkRunProperties152);
            paragraph152.Append(paragraphProperties152);

            Paragraph paragraph153 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties153 = new ParagraphProperties();
            Indentation indentation153 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification144 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties153 = new ParagraphMarkRunProperties();
            Kern kern1169 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1210 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1210 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties153.Append(kern1169);
            paragraphMarkRunProperties153.Append(fontSize1210);
            paragraphMarkRunProperties153.Append(fontSizeComplexScript1210);

            paragraphProperties153.Append(indentation153);
            paragraphProperties153.Append(justification144);
            paragraphProperties153.Append(paragraphMarkRunProperties153);

            Run run1058 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties1058 = new RunProperties();
            Kern kern1170 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1211 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1211 = new FontSizeComplexScript() { Val = "20" };

            runProperties1058.Append(kern1170);
            runProperties1058.Append(fontSize1211);
            runProperties1058.Append(fontSizeComplexScript1211);
            Text text75 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text75.Text = content["情绪不平衡"].Item2 == "没有" ? "★" : " ";

            run1058.Append(runProperties1058);
            run1058.Append(text75);

            paragraph153.Append(paragraphProperties153);
            paragraph153.Append(run1058);

            Paragraph paragraph154 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties154 = new ParagraphProperties();
            Indentation indentation154 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification145 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties154 = new ParagraphMarkRunProperties();
            Kern kern1171 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1212 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript1212 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties154.Append(kern1171);
            paragraphMarkRunProperties154.Append(fontSize1212);
            paragraphMarkRunProperties154.Append(fontSizeComplexScript1212);

            paragraphProperties154.Append(indentation154);
            paragraphProperties154.Append(justification145);
            paragraphProperties154.Append(paragraphMarkRunProperties154);

            paragraph154.Append(paragraphProperties154);

            tableCell70.Append(tableCellProperties70);
            tableCell70.Append(paragraph152);
            tableCell70.Append(paragraph153);
            tableCell70.Append(paragraph154);

            TableCell tableCell71 = new TableCell();

            TableCellProperties tableCellProperties71 = new TableCellProperties();
            TableCellWidth tableCellWidth71 = new TableCellWidth() { Width = "755", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment71 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties71.Append(tableCellWidth71);
            tableCellProperties71.Append(tableCellVerticalAlignment71);

            Paragraph paragraph155 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties155 = new ParagraphProperties();
            Indentation indentation155 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification146 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties155 = new ParagraphMarkRunProperties();
            FontSize fontSize1214 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript1214 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties155.Append(fontSize1214);
            paragraphMarkRunProperties155.Append(fontSizeComplexScript1214);

            paragraphProperties155.Append(indentation155);
            paragraphProperties155.Append(justification146);
            paragraphProperties155.Append(paragraphMarkRunProperties155);
            paragraph155.Append(paragraphProperties155);

            Paragraph paragraph156 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties156 = new ParagraphProperties();
            Indentation indentation156 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification147 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties156 = new ParagraphMarkRunProperties();
            Kern kern1196 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1238 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1238 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties156.Append(kern1196);
            paragraphMarkRunProperties156.Append(fontSize1238);
            paragraphMarkRunProperties156.Append(fontSizeComplexScript1238);

            paragraphProperties156.Append(indentation156);
            paragraphProperties156.Append(justification147);
            paragraphProperties156.Append(paragraphMarkRunProperties156);

            Run run1083 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties1083 = new RunProperties();
            Kern kern1197 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1239 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1239 = new FontSizeComplexScript() { Val = "20" };

            runProperties1083.Append(kern1197);
            runProperties1083.Append(fontSize1239);
            runProperties1083.Append(fontSizeComplexScript1239);
            Text text76 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text76.Text = content["情绪不平衡"].Item2 == "轻度" ? "★" : " ";

            run1083.Append(runProperties1083);
            run1083.Append(text76);

            paragraph156.Append(paragraphProperties156);
            paragraph156.Append(run1083);

            Paragraph paragraph157 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties157 = new ParagraphProperties();
            Indentation indentation157 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification148 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties157 = new ParagraphMarkRunProperties();
            Kern kern1198 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1240 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript1240 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties157.Append(kern1198);
            paragraphMarkRunProperties157.Append(fontSize1240);
            paragraphMarkRunProperties157.Append(fontSizeComplexScript1240);

            paragraphProperties157.Append(indentation157);
            paragraphProperties157.Append(justification148);
            paragraphProperties157.Append(paragraphMarkRunProperties157);
            paragraph157.Append(paragraphProperties157);


            tableCell71.Append(tableCellProperties71);
            tableCell71.Append(paragraph155);
            tableCell71.Append(paragraph156);
            tableCell71.Append(paragraph157);

            TableCell tableCell72 = new TableCell();

            TableCellProperties tableCellProperties72 = new TableCellProperties();
            TableCellWidth tableCellWidth72 = new TableCellWidth() { Width = "756", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment72 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties72.Append(tableCellWidth72);
            tableCellProperties72.Append(tableCellVerticalAlignment72);

            Paragraph paragraph158 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties158 = new ParagraphProperties();
            Indentation indentation158 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification149 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties158 = new ParagraphMarkRunProperties();
            FontSize fontSize1242 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript1242 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties158.Append(fontSize1242);
            paragraphMarkRunProperties158.Append(fontSizeComplexScript1242);

            paragraphProperties158.Append(indentation158);
            paragraphProperties158.Append(justification149);
            paragraphProperties158.Append(paragraphMarkRunProperties158);
            paragraph158.Append(paragraphProperties158);

            Paragraph paragraph159 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties159 = new ParagraphProperties();
            Indentation indentation159 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification150 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties159 = new ParagraphMarkRunProperties();
            Kern kern1223 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1266 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1266 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties159.Append(kern1223);
            paragraphMarkRunProperties159.Append(fontSize1266);
            paragraphMarkRunProperties159.Append(fontSizeComplexScript1266);

            paragraphProperties159.Append(indentation159);
            paragraphProperties159.Append(justification150);
            paragraphProperties159.Append(paragraphMarkRunProperties159);

            Run run1108 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties1108 = new RunProperties();
            RunFonts runFonts858 = new RunFonts() { Ascii = "Segoe UI Symbol", HighAnsi = "Segoe UI Symbol", ComplexScript = "Segoe UI Symbol" };
            Kern kern1224 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1267 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1267 = new FontSizeComplexScript() { Val = "20" };

            runProperties1108.Append(runFonts858);
            runProperties1108.Append(kern1224);
            runProperties1108.Append(fontSize1267);
            runProperties1108.Append(fontSizeComplexScript1267);
            Text text77 = new Text();
            text77.Text = content["情绪不平衡"].Item2 == "中度" ? "★" : " ";

            run1108.Append(runProperties1108);
            run1108.Append(text77);

            paragraph159.Append(paragraphProperties159);
            paragraph159.Append(run1108);

            Paragraph paragraph160 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties160 = new ParagraphProperties();
            Indentation indentation160 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification151 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties160 = new ParagraphMarkRunProperties();
            Kern kern1225 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1268 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript1268 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties160.Append(kern1225);
            paragraphMarkRunProperties160.Append(fontSize1268);
            paragraphMarkRunProperties160.Append(fontSizeComplexScript1268);

            paragraphProperties160.Append(indentation160);
            paragraphProperties160.Append(justification151);
            paragraphProperties160.Append(paragraphMarkRunProperties160);

            paragraph160.Append(paragraphProperties160);

            tableCell72.Append(tableCellProperties72);
            tableCell72.Append(paragraph158);
            tableCell72.Append(paragraph159);
            tableCell72.Append(paragraph160);

            TableCell tableCell73 = new TableCell();

            TableCellProperties tableCellProperties73 = new TableCellProperties();
            TableCellWidth tableCellWidth73 = new TableCellWidth() { Width = "756", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment73 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties73.Append(tableCellWidth73);
            tableCellProperties73.Append(tableCellVerticalAlignment73);

            Paragraph paragraph161 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties161 = new ParagraphProperties();
            Indentation indentation161 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification152 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties161 = new ParagraphMarkRunProperties();
            FontSize fontSize1270 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript1270 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties161.Append(fontSize1270);
            paragraphMarkRunProperties161.Append(fontSizeComplexScript1270);

            paragraphProperties161.Append(indentation161);
            paragraphProperties161.Append(justification152);
            paragraphProperties161.Append(paragraphMarkRunProperties161);

            paragraph161.Append(paragraphProperties161);

            Paragraph paragraph162 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties162 = new ParagraphProperties();
            Indentation indentation162 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification153 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties162 = new ParagraphMarkRunProperties();
            Kern kern1250 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1294 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1294 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties162.Append(kern1250);
            paragraphMarkRunProperties162.Append(fontSize1294);
            paragraphMarkRunProperties162.Append(fontSizeComplexScript1294);

            paragraphProperties162.Append(indentation162);
            paragraphProperties162.Append(justification153);
            paragraphProperties162.Append(paragraphMarkRunProperties162);

            Run run1133 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties1133 = new RunProperties();
            Kern kern1251 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1295 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1295 = new FontSizeComplexScript() { Val = "20" };

            runProperties1133.Append(kern1251);
            runProperties1133.Append(fontSize1295);
            runProperties1133.Append(fontSizeComplexScript1295);
            Text text78 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text78.Text = content["情绪不平衡"].Item2 == "较重" ? "★" : " ";

            run1133.Append(runProperties1133);
            run1133.Append(text78);

            paragraph162.Append(paragraphProperties162);
            paragraph162.Append(run1133);

            Paragraph paragraph163 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties163 = new ParagraphProperties();
            Indentation indentation163 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification154 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties163 = new ParagraphMarkRunProperties();
            Kern kern1252 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1296 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript1296 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties163.Append(kern1252);
            paragraphMarkRunProperties163.Append(fontSize1296);
            paragraphMarkRunProperties163.Append(fontSizeComplexScript1296);

            paragraphProperties163.Append(indentation163);
            paragraphProperties163.Append(justification154);
            paragraphProperties163.Append(paragraphMarkRunProperties163);
            paragraph163.Append(paragraphProperties163);

            tableCell73.Append(tableCellProperties73);
            tableCell73.Append(paragraph161);
            tableCell73.Append(paragraph162);
            tableCell73.Append(paragraph163);

            TableCell tableCell74 = new TableCell();

            TableCellProperties tableCellProperties74 = new TableCellProperties();
            TableCellWidth tableCellWidth74 = new TableCellWidth() { Width = "827", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment74 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties74.Append(tableCellWidth74);
            tableCellProperties74.Append(tableCellVerticalAlignment74);

            Paragraph paragraph164 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties164 = new ParagraphProperties();
            Indentation indentation164 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification155 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties164 = new ParagraphMarkRunProperties();
            FontSize fontSize1298 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript1298 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties164.Append(fontSize1298);
            paragraphMarkRunProperties164.Append(fontSizeComplexScript1298);

            paragraphProperties164.Append(indentation164);
            paragraphProperties164.Append(justification155);
            paragraphProperties164.Append(paragraphMarkRunProperties164);

            paragraph164.Append(paragraphProperties164);

            Paragraph paragraph165 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties165 = new ParagraphProperties();
            Indentation indentation165 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification156 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties165 = new ParagraphMarkRunProperties();
            Kern kern1277 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1322 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1322 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties165.Append(kern1277);
            paragraphMarkRunProperties165.Append(fontSize1322);
            paragraphMarkRunProperties165.Append(fontSizeComplexScript1322);

            paragraphProperties165.Append(indentation165);
            paragraphProperties165.Append(justification156);
            paragraphProperties165.Append(paragraphMarkRunProperties165);

            Run run1158 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties1158 = new RunProperties();
            Kern kern1278 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1323 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1323 = new FontSizeComplexScript() { Val = "20" };

            runProperties1158.Append(kern1278);
            runProperties1158.Append(fontSize1323);
            runProperties1158.Append(fontSizeComplexScript1323);
            Text text79 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text79.Text = content["情绪不平衡"].Item2 == "严重" ? "★" : " ";

            run1158.Append(runProperties1158);
            run1158.Append(text79);

            paragraph165.Append(paragraphProperties165);
            paragraph165.Append(run1158);

            Paragraph paragraph166 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties166 = new ParagraphProperties();
            Indentation indentation166 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification157 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties166 = new ParagraphMarkRunProperties();
            Kern kern1279 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1324 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript1324 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties166.Append(kern1279);
            paragraphMarkRunProperties166.Append(fontSize1324);
            paragraphMarkRunProperties166.Append(fontSizeComplexScript1324);

            paragraphProperties166.Append(indentation166);
            paragraphProperties166.Append(justification157);
            paragraphProperties166.Append(paragraphMarkRunProperties166);
            paragraph166.Append(paragraphProperties166);

            tableCell74.Append(tableCellProperties74);
            tableCell74.Append(paragraph164);
            tableCell74.Append(paragraph165);
            tableCell74.Append(paragraph166);

            tableRow12.Append(tableRowProperties12);
            tableRow12.Append(tableCell68);
            tableRow12.Append(tableCell69);
            tableRow12.Append(tableCell70);
            tableRow12.Append(tableCell71);
            tableRow12.Append(tableCell72);
            tableRow12.Append(tableCell73);
            tableRow12.Append(tableCell74);

            TableRow tableRow13 = new TableRow() { RsidTableRowMarkRevision = "00C709B4", RsidTableRowAddition = "00D061CC", RsidTableRowProperties = "009E6323" };

            TableRowProperties tableRowProperties13 = new TableRowProperties();
            TableRowHeight tableRowHeight13 = new TableRowHeight() { Val = (UInt32Value)800U };

            tableRowProperties13.Append(tableRowHeight13);

            TableCell tableCell75 = new TableCell();

            TableCellProperties tableCellProperties75 = new TableCellProperties();
            TableCellWidth tableCellWidth75 = new TableCellWidth() { Width = "1678", Type = TableWidthUnitValues.Dxa };
            Shading shading21 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "D6E3BC" };
            TableCellVerticalAlignment tableCellVerticalAlignment75 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties75.Append(tableCellWidth75);
            tableCellProperties75.Append(shading21);
            tableCellProperties75.Append(tableCellVerticalAlignment75);

            Paragraph paragraph167 = new Paragraph() { RsidParagraphMarkRevision = "00C709B4", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties167 = new ParagraphProperties();
            Indentation indentation167 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification158 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties167 = new ParagraphMarkRunProperties();
            Bold bold47 = new Bold();
            Kern kern1281 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1326 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript1326 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties167.Append(bold47);
            paragraphMarkRunProperties167.Append(kern1281);
            paragraphMarkRunProperties167.Append(fontSize1326);
            paragraphMarkRunProperties167.Append(fontSizeComplexScript1326);

            paragraphProperties167.Append(indentation167);
            paragraphProperties167.Append(justification158);
            paragraphProperties167.Append(paragraphMarkRunProperties167);

            Run run1160 = new Run() { RsidRunProperties = "00C709B4" };

            RunProperties runProperties1160 = new RunProperties();
            RunFonts runFonts897 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold48 = new Bold();
            Kern kern1282 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1327 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript1327 = new FontSizeComplexScript() { Val = "20" };

            runProperties1160.Append(runFonts897);
            runProperties1160.Append(bold48);
            runProperties1160.Append(kern1282);
            runProperties1160.Append(fontSize1327);
            runProperties1160.Append(fontSizeComplexScript1327);
            Text text80 = new Text();
            text80.Text = "心理不平衡";

            run1160.Append(runProperties1160);
            run1160.Append(text80);

            paragraph167.Append(paragraphProperties167);
            paragraph167.Append(run1160);

            tableCell75.Append(tableCellProperties75);
            tableCell75.Append(paragraph167);

            TableCell tableCell76 = new TableCell();

            TableCellProperties tableCellProperties76 = new TableCellProperties();
            TableCellWidth tableCellWidth76 = new TableCellWidth() { Width = "4453", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment76 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties76.Append(tableCellWidth76);
            tableCellProperties76.Append(tableCellVerticalAlignment76);

            Paragraph paragraph169 = new Paragraph() { RsidParagraphMarkRevision = "00C709B4", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "0010765C", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties169 = new ParagraphProperties();
            Indentation indentation169 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification160 = new Justification() { Val = JustificationValues.Left };

            ParagraphMarkRunProperties paragraphMarkRunProperties169 = new ParagraphMarkRunProperties();
            Kern kern1285 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1330 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1330 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties169.Append(kern1285);
            paragraphMarkRunProperties169.Append(fontSize1330);
            paragraphMarkRunProperties169.Append(fontSizeComplexScript1330);

            paragraphProperties169.Append(indentation169);
            paragraphProperties169.Append(justification160);
            paragraphProperties169.Append(paragraphMarkRunProperties169);

            Run run1162 = new Run() { RsidRunProperties = "00C709B4" };

            RunProperties runProperties1162 = new RunProperties();
            RunFonts runFonts899 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Kern kern1286 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1331 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1331 = new FontSizeComplexScript() { Val = "20" };

            runProperties1162.Append(runFonts899);
            runProperties1162.Append(kern1286);
            runProperties1162.Append(fontSize1331);
            runProperties1162.Append(fontSizeComplexScript1331);
            Text text82 = new Text();
            text82.Text = "感到老师和父母对自己不公平，对同学比自己成绩好难过和不服气";

            run1162.Append(runProperties1162);
            run1162.Append(text82);

            paragraph169.Append(paragraphProperties169);
            paragraph169.Append(run1162);

            tableCell76.Append(tableCellProperties76);
            tableCell76.Append(paragraph169);

            TableCell tableCell77 = new TableCell();

            TableCellProperties tableCellProperties77 = new TableCellProperties();
            TableCellWidth tableCellWidth77 = new TableCellWidth() { Width = "756", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment77 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties77.Append(tableCellWidth77);
            tableCellProperties77.Append(tableCellVerticalAlignment77);

            Paragraph paragraph170 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties170 = new ParagraphProperties();
            Indentation indentation170 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification161 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties170 = new ParagraphMarkRunProperties();
            FontSize fontSize1332 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript1332 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties170.Append(fontSize1332);
            paragraphMarkRunProperties170.Append(fontSizeComplexScript1332);

            paragraphProperties170.Append(indentation170);
            paragraphProperties170.Append(justification161);
            paragraphProperties170.Append(paragraphMarkRunProperties170);

            paragraph170.Append(paragraphProperties170);

            Paragraph paragraph171 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties171 = new ParagraphProperties();
            Indentation indentation171 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification162 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties171 = new ParagraphMarkRunProperties();
            Kern kern1310 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1356 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1356 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties171.Append(kern1310);
            paragraphMarkRunProperties171.Append(fontSize1356);
            paragraphMarkRunProperties171.Append(fontSizeComplexScript1356);

            paragraphProperties171.Append(indentation171);
            paragraphProperties171.Append(justification162);
            paragraphProperties171.Append(paragraphMarkRunProperties171);

            Run run1186 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties1186 = new RunProperties();
            Kern kern1311 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1357 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1357 = new FontSizeComplexScript() { Val = "20" };

            runProperties1186.Append(kern1311);
            runProperties1186.Append(fontSize1357);
            runProperties1186.Append(fontSizeComplexScript1357);
            Text text83 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text83.Text = content["心理不平衡"].Item2 == "没有" ? "★" : " ";

            run1186.Append(runProperties1186);
            run1186.Append(text83);

            paragraph171.Append(paragraphProperties171);
            paragraph171.Append(run1186);

            Paragraph paragraph172 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties172 = new ParagraphProperties();
            Indentation indentation172 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification163 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties172 = new ParagraphMarkRunProperties();
            Kern kern1312 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1358 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript1358 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties172.Append(kern1312);
            paragraphMarkRunProperties172.Append(fontSize1358);
            paragraphMarkRunProperties172.Append(fontSizeComplexScript1358);

            paragraphProperties172.Append(indentation172);
            paragraphProperties172.Append(justification163);
            paragraphProperties172.Append(paragraphMarkRunProperties172);

            paragraph172.Append(paragraphProperties172);

            tableCell77.Append(tableCellProperties77);
            tableCell77.Append(paragraph170);
            tableCell77.Append(paragraph171);
            tableCell77.Append(paragraph172);

            TableCell tableCell78 = new TableCell();

            TableCellProperties tableCellProperties78 = new TableCellProperties();
            TableCellWidth tableCellWidth78 = new TableCellWidth() { Width = "755", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment78 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties78.Append(tableCellWidth78);
            tableCellProperties78.Append(tableCellVerticalAlignment78);

            Paragraph paragraph174 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties174 = new ParagraphProperties();
            Indentation indentation174 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };

            ParagraphMarkRunProperties paragraphMarkRunProperties174 = new ParagraphMarkRunProperties();
            Kern kern1337 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1384 = new FontSize() { Val = "22" };
            Justification justification666 = new Justification() { Val = JustificationValues.Center };
            FontSizeComplexScript fontSizeComplexScript1384 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties174.Append(kern1337);
            paragraphMarkRunProperties174.Append(fontSize1384);
            paragraphMarkRunProperties174.Append(fontSizeComplexScript1384);

            paragraphProperties174.Append(indentation174);
            paragraphProperties174.Append(justification666);
            paragraphProperties174.Append(paragraphMarkRunProperties174);

            Run run1211 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties1211 = new RunProperties();
            RunFonts runFonts938 = new RunFonts() { Ascii = "Segoe UI Symbol", HighAnsi = "Segoe UI Symbol", ComplexScript = "Segoe UI Symbol" };
            Kern kern1338 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1385 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1385 = new FontSizeComplexScript() { Val = "20" };

            runProperties1211.Append(runFonts938);
            runProperties1211.Append(kern1338);
            runProperties1211.Append(fontSize1385);
            runProperties1211.Append(fontSizeComplexScript1385);
            Text text84 = new Text();
            text84.Text = content["心理不平衡"].Item2 == "轻度" ? "★" : " ";

            run1211.Append(runProperties1211);
            run1211.Append(text84);

            paragraph174.Append(paragraphProperties174);
            paragraph174.Append(run1211);

            Paragraph paragraph175 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties175 = new ParagraphProperties();
            Indentation indentation175 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };

            ParagraphMarkRunProperties paragraphMarkRunProperties175 = new ParagraphMarkRunProperties();
            Kern kern1339 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1386 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript1386 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties175.Append(kern1339);
            paragraphMarkRunProperties175.Append(fontSize1386);
            paragraphMarkRunProperties175.Append(fontSizeComplexScript1386);

            paragraphProperties175.Append(indentation175);
            paragraphProperties175.Append(paragraphMarkRunProperties175);
            paragraph175.Append(paragraphProperties175);

            tableCell78.Append(tableCellProperties78);
            tableCell78.Append(paragraph174);
            tableCell78.Append(paragraph175);

            TableCell tableCell79 = new TableCell();

            TableCellProperties tableCellProperties79 = new TableCellProperties();
            TableCellWidth tableCellWidth79 = new TableCellWidth() { Width = "756", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment79 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties79.Append(tableCellWidth79);
            tableCellProperties79.Append(tableCellVerticalAlignment79);

            Paragraph paragraph176 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties176 = new ParagraphProperties();
            Indentation indentation176 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification164 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties176 = new ParagraphMarkRunProperties();
            FontSize fontSize1388 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript1388 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties176.Append(fontSize1388);
            paragraphMarkRunProperties176.Append(fontSizeComplexScript1388);

            paragraphProperties176.Append(indentation176);
            paragraphProperties176.Append(justification164);
            paragraphProperties176.Append(paragraphMarkRunProperties176);
            paragraph176.Append(paragraphProperties176);

            Paragraph paragraph177 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties177 = new ParagraphProperties();
            Indentation indentation177 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification165 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties177 = new ParagraphMarkRunProperties();
            Kern kern1364 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1412 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1412 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties177.Append(kern1364);
            paragraphMarkRunProperties177.Append(fontSize1412);
            paragraphMarkRunProperties177.Append(fontSizeComplexScript1412);

            paragraphProperties177.Append(indentation177);
            paragraphProperties177.Append(justification165);
            paragraphProperties177.Append(paragraphMarkRunProperties177);

            Run run1236 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties1236 = new RunProperties();
            Kern kern1365 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1413 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1413 = new FontSizeComplexScript() { Val = "20" };

            runProperties1236.Append(kern1365);
            runProperties1236.Append(fontSize1413);
            runProperties1236.Append(fontSizeComplexScript1413);
            Text text85 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text85.Text = content["心理不平衡"].Item2 == "中度" ? "★" : " ";

            run1236.Append(runProperties1236);
            run1236.Append(text85);

            paragraph177.Append(paragraphProperties177);
            paragraph177.Append(run1236);

            Paragraph paragraph178 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties178 = new ParagraphProperties();
            Indentation indentation178 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification166 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties178 = new ParagraphMarkRunProperties();
            Kern kern1366 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1414 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript1414 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties178.Append(kern1366);
            paragraphMarkRunProperties178.Append(fontSize1414);
            paragraphMarkRunProperties178.Append(fontSizeComplexScript1414);

            paragraphProperties178.Append(indentation178);
            paragraphProperties178.Append(justification166);
            paragraphProperties178.Append(paragraphMarkRunProperties178);
            paragraph178.Append(paragraphProperties178);

            tableCell79.Append(tableCellProperties79);
            tableCell79.Append(paragraph176);
            tableCell79.Append(paragraph177);
            tableCell79.Append(paragraph178);

            TableCell tableCell80 = new TableCell();

            TableCellProperties tableCellProperties80 = new TableCellProperties();
            TableCellWidth tableCellWidth80 = new TableCellWidth() { Width = "756", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment80 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties80.Append(tableCellWidth80);
            tableCellProperties80.Append(tableCellVerticalAlignment80);

            Paragraph paragraph179 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties179 = new ParagraphProperties();
            Indentation indentation179 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification167 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties179 = new ParagraphMarkRunProperties();
            FontSize fontSize1416 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript1416 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties179.Append(fontSize1416);
            paragraphMarkRunProperties179.Append(fontSizeComplexScript1416);

            paragraphProperties179.Append(indentation179);
            paragraphProperties179.Append(justification167);
            paragraphProperties179.Append(paragraphMarkRunProperties179);
            paragraph179.Append(paragraphProperties179);

            Paragraph paragraph180 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties180 = new ParagraphProperties();
            Indentation indentation180 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification168 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties180 = new ParagraphMarkRunProperties();
            Kern kern1391 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1440 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1440 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties180.Append(kern1391);
            paragraphMarkRunProperties180.Append(fontSize1440);
            paragraphMarkRunProperties180.Append(fontSizeComplexScript1440);

            paragraphProperties180.Append(indentation180);
            paragraphProperties180.Append(justification168);
            paragraphProperties180.Append(paragraphMarkRunProperties180);

            Run run1261 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties1261 = new RunProperties();
            Kern kern1392 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1441 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1441 = new FontSizeComplexScript() { Val = "20" };

            runProperties1261.Append(kern1392);
            runProperties1261.Append(fontSize1441);
            runProperties1261.Append(fontSizeComplexScript1441);
            Text text86 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text86.Text = content["心理不平衡"].Item2 == "较重" ? "★" : " ";

            run1261.Append(runProperties1261);
            run1261.Append(text86);

            paragraph180.Append(paragraphProperties180);
            paragraph180.Append(run1261);

            Paragraph paragraph181 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties181 = new ParagraphProperties();
            Indentation indentation181 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification169 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties181 = new ParagraphMarkRunProperties();
            Kern kern1393 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1442 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript1442 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties181.Append(kern1393);
            paragraphMarkRunProperties181.Append(fontSize1442);
            paragraphMarkRunProperties181.Append(fontSizeComplexScript1442);

            paragraphProperties181.Append(indentation181);
            paragraphProperties181.Append(justification169);
            paragraphProperties181.Append(paragraphMarkRunProperties181);
            paragraph181.Append(paragraphProperties181);

            tableCell80.Append(tableCellProperties80);
            tableCell80.Append(paragraph179);
            tableCell80.Append(paragraph180);
            tableCell80.Append(paragraph181);

            TableCell tableCell81 = new TableCell();

            TableCellProperties tableCellProperties81 = new TableCellProperties();
            TableCellWidth tableCellWidth81 = new TableCellWidth() { Width = "827", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment81 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties81.Append(tableCellWidth81);
            tableCellProperties81.Append(tableCellVerticalAlignment81);

            Paragraph paragraph182 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties182 = new ParagraphProperties();
            Indentation indentation182 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification170 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties182 = new ParagraphMarkRunProperties();
            FontSize fontSize1444 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript1444 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties182.Append(fontSize1444);
            paragraphMarkRunProperties182.Append(fontSizeComplexScript1444);

            paragraphProperties182.Append(indentation182);
            paragraphProperties182.Append(justification170);
            paragraphProperties182.Append(paragraphMarkRunProperties182);
            paragraph182.Append(paragraphProperties182);

            Paragraph paragraph183 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties183 = new ParagraphProperties();
            Indentation indentation183 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification171 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties183 = new ParagraphMarkRunProperties();
            Kern kern1418 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1468 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1468 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties183.Append(kern1418);
            paragraphMarkRunProperties183.Append(fontSize1468);
            paragraphMarkRunProperties183.Append(fontSizeComplexScript1468);

            paragraphProperties183.Append(indentation183);
            paragraphProperties183.Append(justification171);
            paragraphProperties183.Append(paragraphMarkRunProperties183);

            Run run1286 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties1286 = new RunProperties();
            Kern kern1419 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1469 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1469 = new FontSizeComplexScript() { Val = "20" };

            runProperties1286.Append(kern1419);
            runProperties1286.Append(fontSize1469);
            runProperties1286.Append(fontSizeComplexScript1469);
            Text text87 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text87.Text = content["心理不平衡"].Item2 == "严重" ? "★" : " ";

            run1286.Append(runProperties1286);
            run1286.Append(text87);

            paragraph183.Append(paragraphProperties183);
            paragraph183.Append(run1286);

            Paragraph paragraph184 = new Paragraph() { RsidParagraphMarkRevision = "00FC1624", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties184 = new ParagraphProperties();
            Indentation indentation184 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification172 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties184 = new ParagraphMarkRunProperties();
            Kern kern1420 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1470 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript1470 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties184.Append(kern1420);
            paragraphMarkRunProperties184.Append(fontSize1470);
            paragraphMarkRunProperties184.Append(fontSizeComplexScript1470);

            paragraphProperties184.Append(indentation184);
            paragraphProperties184.Append(justification172);
            paragraphProperties184.Append(paragraphMarkRunProperties184);
            paragraph184.Append(paragraphProperties184);

            tableCell81.Append(tableCellProperties81);
            tableCell81.Append(paragraph182);
            tableCell81.Append(paragraph183);
            tableCell81.Append(paragraph184);

            tableRow13.Append(tableRowProperties13);
            tableRow13.Append(tableCell75);
            tableRow13.Append(tableCell76);
            tableRow13.Append(tableCell77);
            tableRow13.Append(tableCell78);
            tableRow13.Append(tableCell79);
            tableRow13.Append(tableCell80);
            tableRow13.Append(tableCell81);

            TableRow tableRow14 = new TableRow() { RsidTableRowMarkRevision = "00C709B4", RsidTableRowAddition = "00D061CC", RsidTableRowProperties = "009E6323" };

            TableRowProperties tableRowProperties14 = new TableRowProperties();
            TableRowHeight tableRowHeight14 = new TableRowHeight() { Val = (UInt32Value)800U };

            tableRowProperties14.Append(tableRowHeight14);

            TableCell tableCell82 = new TableCell();

            TableCellProperties tableCellProperties82 = new TableCellProperties();
            TableCellWidth tableCellWidth82 = new TableCellWidth() { Width = "1678", Type = TableWidthUnitValues.Dxa };
            Shading shading22 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "D6E3BC" };
            TableCellVerticalAlignment tableCellVerticalAlignment82 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties82.Append(tableCellWidth82);
            tableCellProperties82.Append(shading22);
            tableCellProperties82.Append(tableCellVerticalAlignment82);

            Paragraph paragraph185 = new Paragraph() { RsidParagraphMarkRevision = "00C709B4", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties185 = new ParagraphProperties();
            Indentation indentation185 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification173 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties185 = new ParagraphMarkRunProperties();
            Bold bold51 = new Bold();
            Kern kern1422 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1472 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript1472 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties185.Append(bold51);
            paragraphMarkRunProperties185.Append(kern1422);
            paragraphMarkRunProperties185.Append(fontSize1472);
            paragraphMarkRunProperties185.Append(fontSizeComplexScript1472);

            paragraphProperties185.Append(indentation185);
            paragraphProperties185.Append(justification173);
            paragraphProperties185.Append(paragraphMarkRunProperties185);

            Run run1288 = new Run() { RsidRunProperties = "00C709B4" };

            RunProperties runProperties1288 = new RunProperties();
            RunFonts runFonts996 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold52 = new Bold();
            Kern kern1423 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1473 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript1473 = new FontSizeComplexScript() { Val = "20" };

            runProperties1288.Append(runFonts996);
            runProperties1288.Append(bold52);
            runProperties1288.Append(kern1423);
            runProperties1288.Append(fontSize1473);
            runProperties1288.Append(fontSizeComplexScript1473);
            Text text88 = new Text();
            text88.Text = "总体健康水平";

            run1288.Append(runProperties1288);
            run1288.Append(text88);

            paragraph185.Append(paragraphProperties185);
            paragraph185.Append(run1288);

            tableCell82.Append(tableCellProperties82);
            tableCell82.Append(paragraph185);

            TableCell tableCell83 = new TableCell();

            TableCellProperties tableCellProperties83 = new TableCellProperties();
            TableCellWidth tableCellWidth83 = new TableCellWidth() { Width = "8303", Type = TableWidthUnitValues.Dxa };
            GridSpan gridSpan3 = new GridSpan() { Val = 6 };
            TableCellVerticalAlignment tableCellVerticalAlignment83 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties83.Append(tableCellWidth83);
            tableCellProperties83.Append(gridSpan3);
            tableCellProperties83.Append(tableCellVerticalAlignment83);

            Paragraph paragraph186 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00F82954" };

            ParagraphProperties paragraphProperties186 = new ParagraphProperties();
            Indentation indentation186 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification174 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties186 = new ParagraphMarkRunProperties();
            FontSize fontSize1474 = new FontSize() { Val = "2" };
            FontSizeComplexScript fontSizeComplexScript1474 = new FontSizeComplexScript() { Val = "2" };

            paragraphMarkRunProperties186.Append(fontSize1474);
            paragraphMarkRunProperties186.Append(fontSizeComplexScript1474);

            paragraphProperties186.Append(indentation186);
            paragraphProperties186.Append(justification174);
            paragraphProperties186.Append(paragraphMarkRunProperties186);
            paragraph186.Append(paragraphProperties186);

            Paragraph paragraph187 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "009E6323", RsidParagraphProperties = "009E6323", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties187 = new ParagraphProperties();
            Indentation indentation187 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification175 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties187 = new ParagraphMarkRunProperties();
            Kern kern1447 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1498 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript1498 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties187.Append(kern1447);
            paragraphMarkRunProperties187.Append(fontSize1498);
            paragraphMarkRunProperties187.Append(fontSizeComplexScript1498);

            paragraphProperties187.Append(indentation187);
            paragraphProperties187.Append(justification175);
            paragraphProperties187.Append(paragraphMarkRunProperties187);

            Run run1312 = new Run() { RsidRunProperties = "00147D4F" };

            RunProperties runProperties1312 = new RunProperties();
            Kern kern1448 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1499 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript1499 = new FontSizeComplexScript() { Val = "20" };

            runProperties1312.Append(kern1448);
            runProperties1312.Append(fontSize1499);
            runProperties1312.Append(fontSizeComplexScript1499);
            Text text89 = new Text();
            text89.Text = evaluate;
            run1312.Append(runProperties1312);
            run1312.Append(text89);

            paragraph187.Append(paragraphProperties187);
            paragraph187.Append(run1312);

            Paragraph paragraph188 = new Paragraph() { RsidParagraphMarkRevision = "00147D4F", RsidParagraphAddition = "00D061CC", RsidParagraphProperties = "00D061CC", RsidRunAdditionDefault = "00D061CC" };

            ParagraphProperties paragraphProperties188 = new ParagraphProperties();
            Indentation indentation188 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification176 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties188 = new ParagraphMarkRunProperties();
            Kern kern1454 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1505 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript1505 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties188.Append(kern1454);
            paragraphMarkRunProperties188.Append(fontSize1505);
            paragraphMarkRunProperties188.Append(fontSizeComplexScript1505);

            paragraphProperties188.Append(indentation188);
            paragraphProperties188.Append(justification176);
            paragraphProperties188.Append(paragraphMarkRunProperties188);

            paragraph188.Append(paragraphProperties188);

            tableCell83.Append(tableCellProperties83);
            tableCell83.Append(paragraph186);
            tableCell83.Append(paragraph187);
            tableCell83.Append(paragraph188);

            tableRow14.Append(tableRowProperties14);
            tableRow14.Append(tableCell82);
            tableRow14.Append(tableCell83);

            table1.Append(tableProperties1);
            table1.Append(tableGrid1);
            table1.Append(tableRow1);
            table1.Append(tableRow2);
            table1.Append(tableRow3);
            table1.Append(tableRow4);
            table1.Append(tableRow5);
            table1.Append(tableRow6);
            table1.Append(tableRow7);
            table1.Append(tableRow8);
            table1.Append(tableRow9);
            table1.Append(tableRow10);
            table1.Append(bookmarkEnd1);
            table1.Append(tableRow11);
            table1.Append(tableRow12);
            table1.Append(tableRow13);
            table1.Append(tableRow14);
            return table1;
        }
    }
}


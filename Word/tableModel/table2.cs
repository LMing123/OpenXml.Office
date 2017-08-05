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
        // Creates an Table instance and adds its children.
        public Table GenerateTable2()
        {
            Table table1 = new Table();

            TableProperties tableProperties1 = new TableProperties();
            TablePositionProperties tablePositionProperties1 = new TablePositionProperties() { LeftFromText = 181, RightFromText = 181, VerticalAnchor = VerticalAnchorValues.Text, HorizontalAnchor = HorizontalAnchorValues.Margin, TablePositionXAlignment = HorizontalAlignmentValues.Center, TablePositionY = 143 };
            TableOverlap tableOverlap1 = new TableOverlap() { Val = TableOverlapValues.Never };
            TableWidth tableWidth1 = new TableWidth() { Width = "9341", Type = TableWidthUnitValues.Dxa };

            TableBorders tableBorders1 = new TableBorders();
            TopBorder topBorder1 = new TopBorder() { Val = BorderValues.Single, Color = "008000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.Single, Color = "008000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.Single, Color = "008000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.Single, Color = "008000", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder1 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder1 = new InsideVerticalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableBorders1.Append(topBorder1);
            tableBorders1.Append(leftBorder1);
            tableBorders1.Append(bottomBorder1);
            tableBorders1.Append(rightBorder1);
            tableBorders1.Append(insideHorizontalBorder1);
            tableBorders1.Append(insideVerticalBorder1);
            TableLook tableLook1 = new TableLook() { Val = "04A0", FirstRow = true, LastRow = false, FirstColumn = true, LastColumn = false, NoHorizontalBand = false, NoVerticalBand = true };

            tableProperties1.Append(tablePositionProperties1);
            tableProperties1.Append(tableOverlap1);
            tableProperties1.Append(tableWidth1);
            tableProperties1.Append(tableBorders1);
            tableProperties1.Append(tableLook1);

            TableGrid tableGrid1 = new TableGrid();
            GridColumn gridColumn1 = new GridColumn() { Width = "1214" };
            GridColumn gridColumn2 = new GridColumn() { Width = "2367" };
            GridColumn gridColumn3 = new GridColumn() { Width = "2633" };
            GridColumn gridColumn4 = new GridColumn() { Width = "3127" };

            tableGrid1.Append(gridColumn1);
            tableGrid1.Append(gridColumn2);
            tableGrid1.Append(gridColumn3);
            tableGrid1.Append(gridColumn4);

            TableRow tableRow1 = new TableRow() { RsidTableRowMarkRevision = "0038217E", RsidTableRowAddition = "00B97AF7", RsidTableRowProperties = "00B97AF7" };

            TableRowProperties tableRowProperties1 = new TableRowProperties();
            TableRowHeight tableRowHeight1 = new TableRowHeight() { Val = (UInt32Value)329U };

            tableRowProperties1.Append(tableRowHeight1);

            TableCell tableCell1 = new TableCell();

            TableCellProperties tableCellProperties1 = new TableCellProperties();
            TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = "9341", Type = TableWidthUnitValues.Dxa };
            GridSpan gridSpan1 = new GridSpan() { Val = 4 };

            TableCellBorders tableCellBorders1 = new TableCellBorders();
            TopBorder topBorder2 = new TopBorder() { Val = BorderValues.Single, Color = "009999", Size = (UInt32Value)18U, Space = (UInt32Value)0U };
            LeftBorder leftBorder2 = new LeftBorder() { Val = BorderValues.Single, Color = "009999", Size = (UInt32Value)18U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder2 = new BottomBorder() { Val = BorderValues.Single, Color = "009999", Size = (UInt32Value)12U, Space = (UInt32Value)0U };
            RightBorder rightBorder2 = new RightBorder() { Val = BorderValues.Single, Color = "2C95A0", Size = (UInt32Value)18U, Space = (UInt32Value)0U };

            tableCellBorders1.Append(topBorder2);
            tableCellBorders1.Append(leftBorder2);
            tableCellBorders1.Append(bottomBorder2);
            tableCellBorders1.Append(rightBorder2);
            Shading shading1 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "A6A6A6" };
            TableCellVerticalAlignment tableCellVerticalAlignment1 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Bottom };

            tableCellProperties1.Append(tableCellWidth1);
            tableCellProperties1.Append(gridSpan1);
            tableCellProperties1.Append(tableCellBorders1);
            tableCellProperties1.Append(shading1);
            tableCellProperties1.Append(tableCellVerticalAlignment1);

            Paragraph paragraph1 = new Paragraph() { RsidParagraphMarkRevision = "0038217E", RsidParagraphAddition = "00B97AF7", RsidParagraphProperties = "00B97AF7", RsidRunAdditionDefault = "00B97AF7" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            Indentation indentation1 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification1 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
            Bold bold1 = new Bold();
            BoldComplexScript boldComplexScript1 = new BoldComplexScript();
            FontSize fontSize1 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "28" };

            paragraphMarkRunProperties1.Append(bold1);
            paragraphMarkRunProperties1.Append(boldComplexScript1);
            paragraphMarkRunProperties1.Append(fontSize1);
            paragraphMarkRunProperties1.Append(fontSizeComplexScript1);

            paragraphProperties1.Append(indentation1);
            paragraphProperties1.Append(justification1);
            paragraphProperties1.Append(paragraphMarkRunProperties1);

            Run run1 = new Run();

            RunProperties runProperties1 = new RunProperties();
            RunFonts runFonts1 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold2 = new Bold();
            BoldComplexScript boldComplexScript2 = new BoldComplexScript();
            FontSize fontSize2 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "28" };

            runProperties1.Append(runFonts1);
            runProperties1.Append(bold2);
            runProperties1.Append(boldComplexScript2);
            runProperties1.Append(fontSize2);
            runProperties1.Append(fontSizeComplexScript2);
            Text text1 = new Text();
            text1.Text = "师生关系";//表名

            run1.Append(runProperties1);
            run1.Append(text1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            tableCell1.Append(tableCellProperties1);
            tableCell1.Append(paragraph1);

            tableRow1.Append(tableRowProperties1);
            tableRow1.Append(tableCell1);

            TableRow tableRow2 = new TableRow() { RsidTableRowMarkRevision = "00324142", RsidTableRowAddition = "00E16A7F", RsidTableRowProperties = "00E16A7F" };

            TableRowProperties tableRowProperties2 = new TableRowProperties();
            TableRowHeight tableRowHeight2 = new TableRowHeight() { Val = (UInt32Value)677U };

            tableRowProperties2.Append(tableRowHeight2);

            TableCell tableCell2 = new TableCell();

            TableCellProperties tableCellProperties2 = new TableCellProperties();
            TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = "9341", Type = TableWidthUnitValues.Dxa };
            GridSpan gridSpan2 = new GridSpan() { Val = 4 };

            TableCellBorders tableCellBorders2 = new TableCellBorders();
            TopBorder topBorder3 = new TopBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder3 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
            RightBorder rightBorder3 = new RightBorder() { Val = BorderValues.Single, Color = "2C95A0", Size = (UInt32Value)18U, Space = (UInt32Value)0U };

            tableCellBorders2.Append(topBorder3);
            tableCellBorders2.Append(bottomBorder3);
            tableCellBorders2.Append(rightBorder3);
            Shading shading2 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFD9" };
            TableCellVerticalAlignment tableCellVerticalAlignment2 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties2.Append(tableCellWidth2);
            tableCellProperties2.Append(gridSpan2);
            tableCellProperties2.Append(tableCellBorders2);
            tableCellProperties2.Append(shading2);
            tableCellProperties2.Append(tableCellVerticalAlignment2);

            Paragraph paragraph2 = new Paragraph() { RsidParagraphMarkRevision = "00885DA2", RsidParagraphAddition = "00E16A7F", RsidParagraphProperties = "00B97AF7", RsidRunAdditionDefault = "00E16A7F" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation2 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
            Justification justification2 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
            Bold bold4 = new Bold();
            Kern kern1 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize4 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties2.Append(bold4);
            paragraphMarkRunProperties2.Append(kern1);
            paragraphMarkRunProperties2.Append(fontSize4);
            paragraphMarkRunProperties2.Append(fontSizeComplexScript4);

            paragraphProperties2.Append(spacingBetweenLines1);
            paragraphProperties2.Append(indentation2);
            paragraphProperties2.Append(justification2);
            paragraphProperties2.Append(paragraphMarkRunProperties2);


            Run run4 = new Run() { RsidRunProperties = "00885DA2" };

            RunProperties runProperties4 = new RunProperties();
            Bold bold6 = new Bold();
            NoProof noProof2 = new NoProof();
            Color color2 = new Color() { Val = "FF0000" };

            runProperties4.Append(bold6);
            runProperties4.Append(noProof2);
            runProperties4.Append(color2);
            Text text4 = new Text();
            text4.Text = "【量表总的评价】";

            run4.Append(runProperties4);
            run4.Append(text4);

            paragraph2.Append(paragraphProperties2);
            paragraph2.Append(run4);

            Paragraph paragraph3 = new Paragraph() { RsidParagraphMarkRevision = "008E797B", RsidParagraphAddition = "00E16A7F", RsidParagraphProperties = "00B97AF7", RsidRunAdditionDefault = "00E16A7F" };

            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation3 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };

            ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
            RunFonts runFonts28 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Kern kern2 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize5 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties3.Append(runFonts28);
            paragraphMarkRunProperties3.Append(kern2);
            paragraphMarkRunProperties3.Append(fontSize5);
            paragraphMarkRunProperties3.Append(fontSizeComplexScript5);

            paragraphProperties3.Append(spacingBetweenLines2);
            paragraphProperties3.Append(indentation3);
            paragraphProperties3.Append(paragraphMarkRunProperties3);

            Run run35 = new Run() { RsidRunProperties = "008E797B" };

            RunProperties runProperties35 = new RunProperties();
            RunFonts runFonts29 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
            Kern kern3 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize6 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "20" };

            runProperties35.Append(runFonts29);
            runProperties35.Append(kern3);
            runProperties35.Append(fontSize6);
            runProperties35.Append(fontSizeComplexScript6);
            Text text5 = new Text();//总评语
            text5.Text = "该生的师生关系类型是矛盾冲突型，属于消极型师生关系。该生与教师交往中会有冲突和回避，在思想和行为上与教师有较明显的冲突，与教师之间的依赖和亲密感较低；不信任教师，怀疑教师的决定，遇到困难时，也不寻求教师的帮助，师生关系紧张、消极。师生关系是该生学业成绩的消极影响因素。";

            run35.Append(runProperties35);
            run35.Append(text5);


            paragraph3.Append(paragraphProperties3);
            paragraph3.Append(run35);

            tableCell2.Append(tableCellProperties2);
            tableCell2.Append(paragraph2);
            tableCell2.Append(paragraph3);
            tableCell2.Append(new Paragraph(new Run(new Text() { Text = "\n" })));
            tableRow2.Append(tableRowProperties2);
            tableRow2.Append(tableCell2);
            table1.Append(tableProperties1);
            table1.Append(tableGrid1);
            table1.Append(tableRow1);
            table1.Append(tableRow2);

            for (int i = 0; i < 2; i++)
            {

                TableRow tableRow3 = new TableRow() { RsidTableRowMarkRevision = "00324142", RsidTableRowAddition = "00B97AF7", RsidTableRowProperties = "00B97AF7" };

                TableRowProperties tableRowProperties3 = new TableRowProperties();
                TableRowHeight tableRowHeight3 = new TableRowHeight() { Val = (UInt32Value)387U };

                tableRowProperties3.Append(tableRowHeight3);

                TableCell tableCell3 = new TableCell();

                TableCellProperties tableCellProperties3 = new TableCellProperties();
                TableCellWidth tableCellWidth3 = new TableCellWidth() { Width = "1214", Type = TableWidthUnitValues.Dxa };
                VerticalMerge verticalMerge1 = new VerticalMerge() { Val = MergedCellValues.Restart };

                TableCellBorders tableCellBorders3 = new TableCellBorders();
                LeftBorder leftBorder3 = new LeftBorder() { Val = BorderValues.Single, Color = "009999", Size = (UInt32Value)18U, Space = (UInt32Value)0U };

                tableCellBorders3.Append(leftBorder3);
                Shading shading3 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "BFBFBF" };
                TableCellVerticalAlignment tableCellVerticalAlignment3 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

                tableCellProperties3.Append(tableCellWidth3);
                tableCellProperties3.Append(verticalMerge1);
                tableCellProperties3.Append(tableCellBorders3);
                tableCellProperties3.Append(shading3);
                tableCellProperties3.Append(tableCellVerticalAlignment3);

                Paragraph paragraph7 = new Paragraph() { RsidParagraphAddition = "00B97AF7", RsidParagraphProperties = "00B97AF7", RsidRunAdditionDefault = "00B97AF7" };

                ParagraphProperties paragraphProperties7 = new ParagraphProperties();
                Indentation indentation7 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
                Justification justification3 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties7 = new ParagraphMarkRunProperties();
                RunFonts runFonts42 = new RunFonts() { EastAsia = "黑体" };
                Bold bold37 = new Bold();
                BoldComplexScript boldComplexScript31 = new BoldComplexScript();
                FontSize fontSize34 = new FontSize() { Val = "22" };

                paragraphMarkRunProperties7.Append(runFonts42);
                paragraphMarkRunProperties7.Append(bold37);
                paragraphMarkRunProperties7.Append(boldComplexScript31);
                paragraphMarkRunProperties7.Append(fontSize34);

                paragraphProperties7.Append(indentation7);
                paragraphProperties7.Append(justification3);
                paragraphProperties7.Append(paragraphMarkRunProperties7);

                Run run60 = new Run();

                RunProperties runProperties60 = new RunProperties();
                RunFonts runFonts43 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, EastAsia = "黑体" };
                Bold bold38 = new Bold();
                BoldComplexScript boldComplexScript32 = new BoldComplexScript();
                FontSize fontSize35 = new FontSize() { Val = "22" };

                runProperties60.Append(runFonts43);
                runProperties60.Append(bold38);
                runProperties60.Append(boldComplexScript32);
                runProperties60.Append(fontSize35);
                Text text29 = new Text();
                text29.Text = "冲突性";

                run60.Append(runProperties60);
                run60.Append(text29);

                paragraph7.Append(paragraphProperties7);
                paragraph7.Append(run60);

                tableCell3.Append(tableCellProperties3);
                tableCell3.Append(paragraph7);

                TableCell tableCell4 = new TableCell();

                TableCellProperties tableCellProperties4 = new TableCellProperties();
                TableCellWidth tableCellWidth4 = new TableCellWidth() { Width = "2367", Type = TableWidthUnitValues.Dxa };

                TableCellBorders tableCellBorders4 = new TableCellBorders();
                TopBorder topBorder4 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
                BottomBorder bottomBorder4 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
                RightBorder rightBorder4 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };

                tableCellBorders4.Append(topBorder4);
                tableCellBorders4.Append(bottomBorder4);
                tableCellBorders4.Append(rightBorder4);
                Shading shading4 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "B4C6E7", ThemeFill = ThemeColorValues.Accent5, ThemeFillTint = "66" };
                TableCellVerticalAlignment tableCellVerticalAlignment4 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

                tableCellProperties4.Append(tableCellWidth4);
                tableCellProperties4.Append(tableCellBorders4);
                tableCellProperties4.Append(shading4);
                tableCellProperties4.Append(tableCellVerticalAlignment4);

                Paragraph paragraph8 = new Paragraph() { RsidParagraphMarkRevision = "00324142", RsidParagraphAddition = "00B97AF7", RsidParagraphProperties = "00B97AF7", RsidRunAdditionDefault = "00B97AF7" };

                ParagraphProperties paragraphProperties8 = new ParagraphProperties();
                Indentation indentation8 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
                Justification justification4 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties8 = new ParagraphMarkRunProperties();
                BoldComplexScript boldComplexScript34 = new BoldComplexScript();
                FontSize fontSize37 = new FontSize() { Val = "22" };
                FontSizeComplexScript fontSizeComplexScript34 = new FontSizeComplexScript() { Val = "28" };

                paragraphMarkRunProperties8.Append(boldComplexScript34);
                paragraphMarkRunProperties8.Append(fontSize37);
                paragraphMarkRunProperties8.Append(fontSizeComplexScript34);

                paragraphProperties8.Append(indentation8);
                paragraphProperties8.Append(justification4);
                paragraphProperties8.Append(paragraphMarkRunProperties8);

                Run run62 = new Run();

                RunProperties runProperties62 = new RunProperties();
                RunFonts runFonts45 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
                BoldComplexScript boldComplexScript35 = new BoldComplexScript();
                FontSize fontSize38 = new FontSize() { Val = "22" };
                FontSizeComplexScript fontSizeComplexScript35 = new FontSizeComplexScript() { Val = "28" };

                runProperties62.Append(runFonts45);
                runProperties62.Append(boldComplexScript35);
                runProperties62.Append(fontSize38);
                runProperties62.Append(fontSizeComplexScript35);
                Text text31 = new Text();
                text31.Text = "得分";

                run62.Append(runProperties62);
                run62.Append(text31);

                paragraph8.Append(paragraphProperties8);
                paragraph8.Append(run62);

                tableCell4.Append(tableCellProperties4);
                tableCell4.Append(paragraph8);

                TableCell tableCell5 = new TableCell();

                TableCellProperties tableCellProperties5 = new TableCellProperties();
                TableCellWidth tableCellWidth5 = new TableCellWidth() { Width = "2633", Type = TableWidthUnitValues.Dxa };

                TableCellBorders tableCellBorders5 = new TableCellBorders();
                TopBorder topBorder5 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
                LeftBorder leftBorder4 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
                BottomBorder bottomBorder5 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
                RightBorder rightBorder5 = new RightBorder() { Val = BorderValues.Single, Color = "2C95A0", Size = (UInt32Value)18U, Space = (UInt32Value)0U };

                tableCellBorders5.Append(topBorder5);
                tableCellBorders5.Append(leftBorder4);
                tableCellBorders5.Append(bottomBorder5);
                tableCellBorders5.Append(rightBorder5);
                Shading shading5 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "B4C6E7", ThemeFill = ThemeColorValues.Accent5, ThemeFillTint = "66" };
                TableCellVerticalAlignment tableCellVerticalAlignment5 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

                tableCellProperties5.Append(tableCellWidth5);
                tableCellProperties5.Append(tableCellBorders5);
                tableCellProperties5.Append(shading5);
                tableCellProperties5.Append(tableCellVerticalAlignment5);

                Paragraph paragraph9 = new Paragraph() { RsidParagraphMarkRevision = "00324142", RsidParagraphAddition = "00B97AF7", RsidParagraphProperties = "00B97AF7", RsidRunAdditionDefault = "00B97AF7" };

                ParagraphProperties paragraphProperties9 = new ParagraphProperties();
                Indentation indentation9 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
                Justification justification5 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties9 = new ParagraphMarkRunProperties();
                BoldComplexScript boldComplexScript36 = new BoldComplexScript();
                FontSize fontSize39 = new FontSize() { Val = "22" };
                FontSizeComplexScript fontSizeComplexScript36 = new FontSizeComplexScript() { Val = "28" };

                paragraphMarkRunProperties9.Append(boldComplexScript36);
                paragraphMarkRunProperties9.Append(fontSize39);
                paragraphMarkRunProperties9.Append(fontSizeComplexScript36);

                paragraphProperties9.Append(indentation9);
                paragraphProperties9.Append(justification5);
                paragraphProperties9.Append(paragraphMarkRunProperties9);

                Run run63 = new Run();

                RunProperties runProperties63 = new RunProperties();
                RunFonts runFonts46 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
                BoldComplexScript boldComplexScript37 = new BoldComplexScript();
                FontSize fontSize40 = new FontSize() { Val = "22" };
                FontSizeComplexScript fontSizeComplexScript37 = new FontSizeComplexScript() { Val = "28" };

                runProperties63.Append(runFonts46);
                runProperties63.Append(boldComplexScript37);
                runProperties63.Append(fontSize40);
                runProperties63.Append(fontSizeComplexScript37);
                Text text32 = new Text();
                text32.Text = "水平";

                run63.Append(runProperties63);
                run63.Append(text32);

                paragraph9.Append(paragraphProperties9);
                paragraph9.Append(run63);

                tableCell5.Append(tableCellProperties5);
                tableCell5.Append(paragraph9);

                TableCell tableCell6 = new TableCell();

                TableCellProperties tableCellProperties6 = new TableCellProperties();
                TableCellWidth tableCellWidth6 = new TableCellWidth() { Width = "3127", Type = TableWidthUnitValues.Dxa };

                TableCellBorders tableCellBorders6 = new TableCellBorders();
                TopBorder topBorder6 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
                LeftBorder leftBorder5 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
                BottomBorder bottomBorder6 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
                RightBorder rightBorder6 = new RightBorder() { Val = BorderValues.Single, Color = "2C95A0", Size = (UInt32Value)18U, Space = (UInt32Value)0U };

                tableCellBorders6.Append(topBorder6);
                tableCellBorders6.Append(leftBorder5);
                tableCellBorders6.Append(bottomBorder6);
                tableCellBorders6.Append(rightBorder6);
                Shading shading6 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "B4C6E7", ThemeFill = ThemeColorValues.Accent5, ThemeFillTint = "66" };
                TableCellVerticalAlignment tableCellVerticalAlignment6 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

                tableCellProperties6.Append(tableCellWidth6);
                tableCellProperties6.Append(tableCellBorders6);
                tableCellProperties6.Append(shading6);
                tableCellProperties6.Append(tableCellVerticalAlignment6);

                Paragraph paragraph10 = new Paragraph() { RsidParagraphMarkRevision = "00324142", RsidParagraphAddition = "00B97AF7", RsidParagraphProperties = "00B97AF7", RsidRunAdditionDefault = "00B97AF7" };

                ParagraphProperties paragraphProperties10 = new ParagraphProperties();
                Indentation indentation10 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
                Justification justification6 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties10 = new ParagraphMarkRunProperties();
                BoldComplexScript boldComplexScript38 = new BoldComplexScript();
                FontSize fontSize41 = new FontSize() { Val = "22" };
                FontSizeComplexScript fontSizeComplexScript38 = new FontSizeComplexScript() { Val = "28" };

                paragraphMarkRunProperties10.Append(boldComplexScript38);
                paragraphMarkRunProperties10.Append(fontSize41);
                paragraphMarkRunProperties10.Append(fontSizeComplexScript38);

                paragraphProperties10.Append(indentation10);
                paragraphProperties10.Append(justification6);
                paragraphProperties10.Append(paragraphMarkRunProperties10);


                Run run65 = new Run();

                RunProperties runProperties65 = new RunProperties();
                BoldComplexScript boldComplexScript40 = new BoldComplexScript();
                FontSize fontSize43 = new FontSize() { Val = "22" };
                FontSizeComplexScript fontSizeComplexScript40 = new FontSizeComplexScript() { Val = "28" };

                runProperties65.Append(boldComplexScript40);
                runProperties65.Append(fontSize43);
                runProperties65.Append(fontSizeComplexScript40);
                Text text34 = new Text();
                text34.Text = "对学习的影响";

                run65.Append(runProperties65);
                run65.Append(text34);

                paragraph10.Append(paragraphProperties10);
                paragraph10.Append(run65);

                tableCell6.Append(tableCellProperties6);
                tableCell6.Append(paragraph10);

                tableRow3.Append(tableRowProperties3);
                tableRow3.Append(tableCell3);
                tableRow3.Append(tableCell4);
                tableRow3.Append(tableCell5);
                tableRow3.Append(tableCell6);

                TableRow tableRow4 = new TableRow() { RsidTableRowAddition = "00B97AF7", RsidTableRowProperties = "00B97AF7" };

                TableRowProperties tableRowProperties4 = new TableRowProperties();
                TableRowHeight tableRowHeight4 = new TableRowHeight() { Val = (UInt32Value)387U };

                tableRowProperties4.Append(tableRowHeight4);

                TableCell tableCell7 = new TableCell();

                TableCellProperties tableCellProperties7 = new TableCellProperties();
                TableCellWidth tableCellWidth7 = new TableCellWidth() { Width = "1214", Type = TableWidthUnitValues.Dxa };
                VerticalMerge verticalMerge2 = new VerticalMerge();

                TableCellBorders tableCellBorders7 = new TableCellBorders();
                LeftBorder leftBorder6 = new LeftBorder() { Val = BorderValues.Single, Color = "009999", Size = (UInt32Value)18U, Space = (UInt32Value)0U };

                tableCellBorders7.Append(leftBorder6);
                Shading shading7 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "BFBFBF" };
                TableCellVerticalAlignment tableCellVerticalAlignment7 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

                tableCellProperties7.Append(tableCellWidth7);
                tableCellProperties7.Append(verticalMerge2);
                tableCellProperties7.Append(tableCellBorders7);
                tableCellProperties7.Append(shading7);
                tableCellProperties7.Append(tableCellVerticalAlignment7);

                Paragraph paragraph11 = new Paragraph() { RsidParagraphAddition = "00B97AF7", RsidParagraphProperties = "00B97AF7", RsidRunAdditionDefault = "00B97AF7" };

                ParagraphProperties paragraphProperties11 = new ParagraphProperties();
                Indentation indentation11 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
                Justification justification7 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties11 = new ParagraphMarkRunProperties();
                RunFonts runFonts48 = new RunFonts() { EastAsia = "黑体" };
                Bold bold40 = new Bold();
                BoldComplexScript boldComplexScript41 = new BoldComplexScript();
                FontSize fontSize44 = new FontSize() { Val = "22" };

                paragraphMarkRunProperties11.Append(runFonts48);
                paragraphMarkRunProperties11.Append(bold40);
                paragraphMarkRunProperties11.Append(boldComplexScript41);
                paragraphMarkRunProperties11.Append(fontSize44);

                paragraphProperties11.Append(indentation11);
                paragraphProperties11.Append(justification7);
                paragraphProperties11.Append(paragraphMarkRunProperties11);

                paragraph11.Append(paragraphProperties11);

                tableCell7.Append(tableCellProperties7);
                tableCell7.Append(paragraph11);

                TableCell tableCell8 = new TableCell();

                TableCellProperties tableCellProperties8 = new TableCellProperties();
                TableCellWidth tableCellWidth8 = new TableCellWidth() { Width = "2367", Type = TableWidthUnitValues.Dxa };

                TableCellBorders tableCellBorders8 = new TableCellBorders();
                TopBorder topBorder7 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
                BottomBorder bottomBorder7 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
                RightBorder rightBorder7 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };

                tableCellBorders8.Append(topBorder7);
                tableCellBorders8.Append(bottomBorder7);
                tableCellBorders8.Append(rightBorder7);
                Shading shading8 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "B4C6E7", ThemeFill = ThemeColorValues.Accent5, ThemeFillTint = "66" };
                TableCellVerticalAlignment tableCellVerticalAlignment8 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

                tableCellProperties8.Append(tableCellWidth8);
                tableCellProperties8.Append(tableCellBorders8);
                tableCellProperties8.Append(shading8);
                tableCellProperties8.Append(tableCellVerticalAlignment8);

                Paragraph paragraph12 = new Paragraph() { RsidParagraphAddition = "00B97AF7", RsidParagraphProperties = "00B97AF7", RsidRunAdditionDefault = "00B97AF7" };

                ParagraphProperties paragraphProperties12 = new ParagraphProperties();
                Indentation indentation12 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
                Justification justification8 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties12 = new ParagraphMarkRunProperties();
                BoldComplexScript boldComplexScript42 = new BoldComplexScript();
                FontSize fontSize45 = new FontSize() { Val = "22" };
                FontSizeComplexScript fontSizeComplexScript41 = new FontSizeComplexScript() { Val = "28" };

                paragraphMarkRunProperties12.Append(boldComplexScript42);
                paragraphMarkRunProperties12.Append(fontSize45);
                paragraphMarkRunProperties12.Append(fontSizeComplexScript41);

                paragraphProperties12.Append(indentation12);
                paragraphProperties12.Append(justification8);
                paragraphProperties12.Append(paragraphMarkRunProperties12);

                Run run66 = new Run();

                RunProperties runProperties66 = new RunProperties();
                RunFonts runFonts49 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
                BoldComplexScript boldComplexScript43 = new BoldComplexScript();
                FontSize fontSize46 = new FontSize() { Val = "22" };
                FontSizeComplexScript fontSizeComplexScript42 = new FontSizeComplexScript() { Val = "28" };

                runProperties66.Append(runFonts49);
                runProperties66.Append(boldComplexScript43);
                runProperties66.Append(fontSize46);
                runProperties66.Append(fontSizeComplexScript42);
                Text text35 = new Text();
                text35.Text = "82.89";//得分

                run66.Append(runProperties66);
                run66.Append(text35);

                paragraph12.Append(paragraphProperties12);
                paragraph12.Append(run66);

                tableCell8.Append(tableCellProperties8);
                tableCell8.Append(paragraph12);

                TableCell tableCell9 = new TableCell();

                TableCellProperties tableCellProperties9 = new TableCellProperties();
                TableCellWidth tableCellWidth9 = new TableCellWidth() { Width = "2633", Type = TableWidthUnitValues.Dxa };

                TableCellBorders tableCellBorders9 = new TableCellBorders();
                TopBorder topBorder8 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
                LeftBorder leftBorder7 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
                BottomBorder bottomBorder8 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
                RightBorder rightBorder8 = new RightBorder() { Val = BorderValues.Single, Color = "2C95A0", Size = (UInt32Value)18U, Space = (UInt32Value)0U };

                tableCellBorders9.Append(topBorder8);
                tableCellBorders9.Append(leftBorder7);
                tableCellBorders9.Append(bottomBorder8);
                tableCellBorders9.Append(rightBorder8);
                Shading shading9 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "B4C6E7", ThemeFill = ThemeColorValues.Accent5, ThemeFillTint = "66" };
                TableCellVerticalAlignment tableCellVerticalAlignment9 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

                tableCellProperties9.Append(tableCellWidth9);
                tableCellProperties9.Append(tableCellBorders9);
                tableCellProperties9.Append(shading9);
                tableCellProperties9.Append(tableCellVerticalAlignment9);

                Paragraph paragraph13 = new Paragraph() { RsidParagraphAddition = "00B97AF7", RsidParagraphProperties = "00B97AF7", RsidRunAdditionDefault = "00B97AF7" };

                ParagraphProperties paragraphProperties13 = new ParagraphProperties();
                Indentation indentation13 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
                Justification justification9 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties13 = new ParagraphMarkRunProperties();
                BoldComplexScript boldComplexScript44 = new BoldComplexScript();
                FontSize fontSize47 = new FontSize() { Val = "22" };
                FontSizeComplexScript fontSizeComplexScript43 = new FontSizeComplexScript() { Val = "28" };

                paragraphMarkRunProperties13.Append(boldComplexScript44);
                paragraphMarkRunProperties13.Append(fontSize47);
                paragraphMarkRunProperties13.Append(fontSizeComplexScript43);

                paragraphProperties13.Append(indentation13);
                paragraphProperties13.Append(justification9);
                paragraphProperties13.Append(paragraphMarkRunProperties13);

                Run run67 = new Run();

                RunProperties runProperties67 = new RunProperties();
                RunFonts runFonts50 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
                BoldComplexScript boldComplexScript45 = new BoldComplexScript();
                FontSize fontSize48 = new FontSize() { Val = "22" };
                FontSizeComplexScript fontSizeComplexScript44 = new FontSizeComplexScript() { Val = "28" };

                runProperties67.Append(runFonts50);
                runProperties67.Append(boldComplexScript45);
                runProperties67.Append(fontSize48);
                runProperties67.Append(fontSizeComplexScript44);
                Text text36 = new Text();
                text36.Text = "高水平";

                run67.Append(runProperties67);
                run67.Append(text36);

                paragraph13.Append(paragraphProperties13);
                paragraph13.Append(run67);

                tableCell9.Append(tableCellProperties9);
                tableCell9.Append(paragraph13);

                TableCell tableCell10 = new TableCell();

                TableCellProperties tableCellProperties10 = new TableCellProperties();
                TableCellWidth tableCellWidth10 = new TableCellWidth() { Width = "3127", Type = TableWidthUnitValues.Dxa };

                TableCellBorders tableCellBorders10 = new TableCellBorders();
                TopBorder topBorder9 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
                LeftBorder leftBorder8 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
                BottomBorder bottomBorder9 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
                RightBorder rightBorder9 = new RightBorder() { Val = BorderValues.Single, Color = "2C95A0", Size = (UInt32Value)18U, Space = (UInt32Value)0U };

                tableCellBorders10.Append(topBorder9);
                tableCellBorders10.Append(leftBorder8);
                tableCellBorders10.Append(bottomBorder9);
                tableCellBorders10.Append(rightBorder9);
                Shading shading10 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "B4C6E7", ThemeFill = ThemeColorValues.Accent5, ThemeFillTint = "66" };
                TableCellVerticalAlignment tableCellVerticalAlignment10 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

                tableCellProperties10.Append(tableCellWidth10);
                tableCellProperties10.Append(tableCellBorders10);
                tableCellProperties10.Append(shading10);
                tableCellProperties10.Append(tableCellVerticalAlignment10);

                Paragraph paragraph14 = new Paragraph() { RsidParagraphAddition = "00B97AF7", RsidParagraphProperties = "00E9225D", RsidRunAdditionDefault = "00B97AF7" };

                ParagraphProperties paragraphProperties14 = new ParagraphProperties();
                SpacingBetweenLines spacingBetweenLines6 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Auto };
                Indentation indentation14 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
                Justification justification10 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties14 = new ParagraphMarkRunProperties();
                BoldComplexScript boldComplexScript46 = new BoldComplexScript();
                FontSize fontSize49 = new FontSize() { Val = "22" };
                FontSizeComplexScript fontSizeComplexScript45 = new FontSizeComplexScript() { Val = "28" };

                paragraphMarkRunProperties14.Append(boldComplexScript46);
                paragraphMarkRunProperties14.Append(fontSize49);
                paragraphMarkRunProperties14.Append(fontSizeComplexScript45);

                paragraphProperties14.Append(spacingBetweenLines6);
                paragraphProperties14.Append(indentation14);
                paragraphProperties14.Append(justification10);
                paragraphProperties14.Append(paragraphMarkRunProperties14);

                paragraph14.Append(paragraphProperties14);

                tableCell10.Append(tableCellProperties10);
                tableCell10.Append(paragraph14);

                tableRow4.Append(tableRowProperties4);
                tableRow4.Append(tableCell7);
                tableRow4.Append(tableCell8);
                tableRow4.Append(tableCell9);
                tableRow4.Append(tableCell10);

                TableRow tableRow5 = new TableRow() { RsidTableRowMarkRevision = "0094019C", RsidTableRowAddition = "00B97AF7", RsidTableRowProperties = "00B97AF7" };

                TableRowProperties tableRowProperties5 = new TableRowProperties();
                TableRowHeight tableRowHeight5 = new TableRowHeight() { Val = (UInt32Value)759U };

                tableRowProperties5.Append(tableRowHeight5);

                TableCell tableCell11 = new TableCell();

                TableCellProperties tableCellProperties11 = new TableCellProperties();
                TableCellWidth tableCellWidth11 = new TableCellWidth() { Width = "1214", Type = TableWidthUnitValues.Dxa };
                VerticalMerge verticalMerge3 = new VerticalMerge();

                TableCellBorders tableCellBorders11 = new TableCellBorders();
                LeftBorder leftBorder9 = new LeftBorder() { Val = BorderValues.Single, Color = "009999", Size = (UInt32Value)18U, Space = (UInt32Value)0U };

                tableCellBorders11.Append(leftBorder9);
                Shading shading11 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "BFBFBF" };
                TableCellVerticalAlignment tableCellVerticalAlignment11 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

                tableCellProperties11.Append(tableCellWidth11);
                tableCellProperties11.Append(verticalMerge3);
                tableCellProperties11.Append(tableCellBorders11);
                tableCellProperties11.Append(shading11);
                tableCellProperties11.Append(tableCellVerticalAlignment11);

                Paragraph paragraph15 = new Paragraph() { RsidParagraphMarkRevision = "00324142", RsidParagraphAddition = "00B97AF7", RsidParagraphProperties = "00B97AF7", RsidRunAdditionDefault = "00B97AF7" };

                ParagraphProperties paragraphProperties15 = new ParagraphProperties();
                SpacingBetweenLines spacingBetweenLines7 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
                Indentation indentation15 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
                Justification justification11 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties15 = new ParagraphMarkRunProperties();
                RunFonts runFonts51 = new RunFonts() { EastAsia = "黑体" };
                Bold bold41 = new Bold();
                BoldComplexScript boldComplexScript47 = new BoldComplexScript();
                FontSize fontSize50 = new FontSize() { Val = "22" };

                paragraphMarkRunProperties15.Append(runFonts51);
                paragraphMarkRunProperties15.Append(bold41);
                paragraphMarkRunProperties15.Append(boldComplexScript47);
                paragraphMarkRunProperties15.Append(fontSize50);

                paragraphProperties15.Append(spacingBetweenLines7);
                paragraphProperties15.Append(indentation15);
                paragraphProperties15.Append(justification11);
                paragraphProperties15.Append(paragraphMarkRunProperties15);

                paragraph15.Append(paragraphProperties15);

                tableCell11.Append(tableCellProperties11);
                tableCell11.Append(paragraph15);

                TableCell tableCell12 = new TableCell();

                TableCellProperties tableCellProperties12 = new TableCellProperties();
                TableCellWidth tableCellWidth12 = new TableCellWidth() { Width = "8127", Type = TableWidthUnitValues.Dxa };
                GridSpan gridSpan3 = new GridSpan() { Val = 3 };

                TableCellBorders tableCellBorders12 = new TableCellBorders();
                TopBorder topBorder10 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
                BottomBorder bottomBorder10 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
                RightBorder rightBorder10 = new RightBorder() { Val = BorderValues.Single, Color = "2C95A0", Size = (UInt32Value)18U, Space = (UInt32Value)0U };

                tableCellBorders12.Append(topBorder10);
                tableCellBorders12.Append(bottomBorder10);
                tableCellBorders12.Append(rightBorder10);
                Shading shading12 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFD9" };
                TableCellVerticalAlignment tableCellVerticalAlignment12 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

                tableCellProperties12.Append(tableCellWidth12);
                tableCellProperties12.Append(gridSpan3);
                tableCellProperties12.Append(tableCellBorders12);
                tableCellProperties12.Append(shading12);
                tableCellProperties12.Append(tableCellVerticalAlignment12);

                Paragraph paragraph16 = new Paragraph() { RsidParagraphMarkRevision = "00961443", RsidParagraphAddition = "00B97AF7", RsidParagraphProperties = "00B97AF7", RsidRunAdditionDefault = "00B97AF7" };

                ParagraphProperties paragraphProperties16 = new ParagraphProperties();
                SpacingBetweenLines spacingBetweenLines8 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
                Indentation indentation16 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };
                Justification justification12 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties16 = new ParagraphMarkRunProperties();
                Bold bold42 = new Bold();
                Color color3 = new Color() { Val = "FF0000" };
                Kern kern4 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize51 = new FontSize() { Val = "20" };
                FontSizeComplexScript fontSizeComplexScript46 = new FontSizeComplexScript() { Val = "20" };

                paragraphMarkRunProperties16.Append(bold42);
                paragraphMarkRunProperties16.Append(color3);
                paragraphMarkRunProperties16.Append(kern4);
                paragraphMarkRunProperties16.Append(fontSize51);
                paragraphMarkRunProperties16.Append(fontSizeComplexScript46);

                paragraphProperties16.Append(spacingBetweenLines8);
                paragraphProperties16.Append(indentation16);
                paragraphProperties16.Append(justification12);
                paragraphProperties16.Append(paragraphMarkRunProperties16);

                Run run68 = new Run() { RsidRunProperties = "00961443" };

                RunProperties runProperties68 = new RunProperties();
                RunFonts runFonts52 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
                Bold bold43 = new Bold();
                NoProof noProof60 = new NoProof();
                Color color4 = new Color() { Val = "FF0000" };

                runProperties68.Append(runFonts52);
                runProperties68.Append(bold43);
                runProperties68.Append(noProof60);
                runProperties68.Append(color4);
                Text text37 = new Text();
                text37.Text = "【维度评价】";

                run68.Append(runProperties68);
                run68.Append(text37);

                paragraph16.Append(paragraphProperties16);
                paragraph16.Append(run68);

                Paragraph paragraph17 = new Paragraph() { RsidParagraphMarkRevision = "008E797B", RsidParagraphAddition = "00B97AF7", RsidParagraphProperties = "00B97AF7", RsidRunAdditionDefault = "00B97AF7" };

                ParagraphProperties paragraphProperties17 = new ParagraphProperties();
                SpacingBetweenLines spacingBetweenLines9 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Auto };
                Indentation indentation17 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };

                ParagraphMarkRunProperties paragraphMarkRunProperties17 = new ParagraphMarkRunProperties();
                RunFonts runFonts78 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                Kern kern5 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize52 = new FontSize() { Val = "20" };
                FontSizeComplexScript fontSizeComplexScript47 = new FontSizeComplexScript() { Val = "20" };

                paragraphMarkRunProperties17.Append(runFonts78);
                paragraphMarkRunProperties17.Append(kern5);
                paragraphMarkRunProperties17.Append(fontSize52);
                paragraphMarkRunProperties17.Append(fontSizeComplexScript47);

                paragraphProperties17.Append(spacingBetweenLines9);
                paragraphProperties17.Append(indentation17);
                paragraphProperties17.Append(paragraphMarkRunProperties17);

                Run run99 = new Run() { RsidRunProperties = "008E797B" };

                RunProperties runProperties99 = new RunProperties();
                RunFonts runFonts79 = new RunFonts() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
                Kern kern6 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize53 = new FontSize() { Val = "20" };
                FontSizeComplexScript fontSizeComplexScript48 = new FontSizeComplexScript() { Val = "20" };

                runProperties99.Append(runFonts79);
                runProperties99.Append(kern6);
                runProperties99.Append(fontSize53);
                runProperties99.Append(fontSizeComplexScript48);
                Text text39 = new Text();
                text39.Text = "该生在冲突性上得分为高水平，说明该生与教师之间缺少情感协调，容易产生摩擦；学生由于被迫屈服于教师权威或受到教师的误解而引起了心理上的不满，容易触发愤怒等消极情绪，导致学生退缩，出现如孤独、消极的学习态度。在遇到问题时，教师也不愿意找学生家长解决问题。";

                run99.Append(runProperties99);
                run99.Append(text39);

                paragraph17.Append(paragraphProperties17);
                paragraph17.Append(run99);

                Paragraph paragraph18 = new Paragraph() { RsidParagraphMarkRevision = "0094019C", RsidParagraphAddition = "00B97AF7", RsidParagraphProperties = "00B97AF7", RsidRunAdditionDefault = "00F82954" };

                ParagraphProperties paragraphProperties18 = new ParagraphProperties();
                SpacingBetweenLines spacingBetweenLines10 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
                Indentation indentation18 = new Indentation() { FirstLine = "0", FirstLineChars = 0 };

                ParagraphMarkRunProperties paragraphMarkRunProperties18 = new ParagraphMarkRunProperties();
                BoldComplexScript boldComplexScript48 = new BoldComplexScript();
                NoProof noProof91 = new NoProof();
                FontSize fontSize54 = new FontSize() { Val = "21" };
                FontSizeComplexScript fontSizeComplexScript49 = new FontSizeComplexScript() { Val = "21" };

                paragraphMarkRunProperties18.Append(boldComplexScript48);
                paragraphMarkRunProperties18.Append(noProof91);
                paragraphMarkRunProperties18.Append(fontSize54);
                paragraphMarkRunProperties18.Append(fontSizeComplexScript49);

                paragraphProperties18.Append(spacingBetweenLines10);
                paragraphProperties18.Append(indentation18);
                paragraphProperties18.Append(paragraphMarkRunProperties18);

                Run run100 = new Run();

                RunProperties runProperties100 = new RunProperties();
                BoldComplexScript boldComplexScript49 = new BoldComplexScript();
                NoProof noProof92 = new NoProof();
                FontSize fontSize55 = new FontSize() { Val = "21" };
                FontSizeComplexScript fontSizeComplexScript50 = new FontSizeComplexScript() { Val = "21" };

                runProperties100.Append(boldComplexScript49);
                runProperties100.Append(noProof92);
                runProperties100.Append(fontSize55);
                runProperties100.Append(fontSizeComplexScript50);
                FieldChar fieldChar6 = new FieldChar() { FieldCharType = FieldCharValues.End };

                run100.Append(runProperties100);
                run100.Append(fieldChar6);

                paragraph18.Append(paragraphProperties18);
                paragraph18.Append(run100);

                tableCell12.Append(tableCellProperties12);
                tableCell12.Append(paragraph16);
                tableCell12.Append(paragraph17);
                tableCell12.Append(paragraph18);

                tableRow5.Append(tableRowProperties5);
                tableRow5.Append(tableCell11);
                tableRow5.Append(tableCell12);



                table1.Append(tableRow3);
                table1.Append(tableRow4);
                table1.Append(tableRow5);
            }


            return table1;
        }


    }


}

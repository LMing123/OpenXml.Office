using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Word
{
    public class Table
    {
        private DocumentFormat.OpenXml.Wordprocessing.Table table;
        private TableProperties tableProperties;
        private int row;
        private int column;
        public Table(DocumentFormat.OpenXml.Wordprocessing.Table table, int row, int column)
        {
            this.table = table;
            this.row = row;
            this.column = column;
            tableProperties = new TableProperties();
            table.AppendChild(tableProperties);
            CreateTable(row, column);
            AddBorder();


        }
        /// <summary>
        /// 设置单元格格式
        /// 目前只底色设置
        /// </summary>
        /// <param name="row">设置的单元格所在行</param>
        /// <param name="column">设置的单元格所在列</param>
        /// <param name="groundColor">设置的单元格底色</param>
        /// <param name="hight">not finished</param>
        /// <param name="size">not finished</param>
        /// <param name="fontname">not finished</param>
        /// <param name="fontcolor">not finished</param>
        /// <param name="bold">not finished</param>
        /// <param name="italic">not finished</param>
        /// <param name="underlineValues">not finished</param>
        public void SetCellStyle(int row,int column, System.Drawing.Color? groundColor= null, float? hight = null, int? size = null, string fontname = null,/* eParagraphAlignment? alignment = null,*/ System.Drawing.Color? fontcolor = null, bool? bold = null, bool? italic = null, UnderlineValues? underlineValues = null)
        {
            int tem_row = 1;
            int tem_column = 1;
            //  var cell= table.Where(i => i.LocalName == "tr").Where(i=>i.LocalName=="tc");

            foreach (var tableRow in table)
            {
                if (tableRow.LocalName == "tr")
                {
                    if (row == tem_row)
                    {
                        tem_column = 1;
                        foreach (var cell in tableRow)
                        {
                            if (cell.LocalName == "tc")
                            {
                                if (column == tem_column)
                                {
                                    foreach (var cellProperties in cell)
                                    {
                                        if (cellProperties.LocalName == "tcPr")
                                        {
                                            var cPr = (TableCellProperties)cellProperties;
                                            if (groundColor != null)
                                            {
                                                cPr.Shading = new Shading() { Fill = String.Format("{0:X6}", groundColor.Value.R << 16 | groundColor.Value.G << 8 | groundColor.Value.B) };
                                            }
                                            return;
                                        }
                                    }
                                }
                                tem_column++;

                            }
                        }
                    }
                    tem_row++;
                }
               
            }
        }

        /// <summary>
        /// 设置边框
        /// </summary>
        /// <param name="hasBorder"></param>
        /// <param name="topBorder"></param>
        /// <param name="bottomBorder"></param>
        /// <param name="leftBorder"></param>
        /// <param name="rightBorder"></param>
        /// <param name="insideHorizontalBorder"></param>
        /// <param name="insideVerticalBorder"></param>
        public void AddBorder(bool hasBorder = true, int? topBorder = null, int? bottomBorder = null, int? leftBorder = null, int? rightBorder = null, int? insideHorizontalBorder = null, int? insideVerticalBorder = null)
        {
            if (hasBorder)
            {
                if (tableProperties.TableBorders == null)
                {
                    tableProperties.TableBorders = new TableBorders();
                }
                var borders = tableProperties.TableBorders;
                borders.TopBorder = new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 5 };
                borders.BottomBorder = new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 5 };
                borders.LeftBorder = new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 5 };
                borders.RightBorder = new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 5 };
                borders.InsideHorizontalBorder = new InsideHorizontalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 5 };
                borders.InsideVerticalBorder = new InsideVerticalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 5 };
                if (topBorder != null)
                {
                    borders.TopBorder.Size = (UInt32)topBorder;
                }
                if (bottomBorder != null)
                {
                    borders.BottomBorder.Size = (UInt32)bottomBorder;
                }
                if (leftBorder != null)
                {
                    borders.LeftBorder.Size = (UInt32)leftBorder;
                }
                if (rightBorder != null)
                {
                    borders.RightBorder.Size = (UInt32)rightBorder;
                }
                if (insideHorizontalBorder != null)
                {
                    borders.InsideHorizontalBorder.Size = (UInt32)insideHorizontalBorder;
                }
                if (insideVerticalBorder != null)
                {
                    borders.InsideVerticalBorder.Size = (UInt32)insideVerticalBorder;
                }
            }

        }
        /// <summary>
        /// 单元格添加文字
        /// </summary>
        /// <param name="row">行</param>
        /// <param name="column">列</param>
        /// <param name="text">要添加的文字</param>
        /// <param name="justificationValues">对其方式</param>
        public void CellText(int row, int column, string text, JustificationValues? justificationValues = null)
        {
            int tem_row = 1;
            int tem_column;
            foreach (var tablerow in table)
            {
                if (tablerow.LocalName != "tr" )
                {                    
                    continue;
                }
                if(tem_row!=row)
                {
                    tem_row++;
                    continue;
                }
                tem_column = 1;
                foreach (var cell in tablerow)
                {
                    if (cell.LocalName != "tc")
                    {                      
                        continue;
                    }
                    if (tem_column == column)
                    {
                        foreach (var cellPara in cell)
                        {
                            if (cellPara.LocalName == "p")
                            {
                                cell.RemoveChild(cellPara);
                            }
                        }
                        cell.Append(new Paragraph(new Run(new Text() { Text = text })) { ParagraphProperties = new ParagraphProperties() { Justification = new Justification() { Val = (DocumentFormat.OpenXml.Wordprocessing.JustificationValues)justificationValues } } });

                    }
                    tem_column++;

                }
                tem_row++;
            }
        }
        /// <summary>
        /// 合并单元格
        /// </summary>
        /// <param name="row">开始行</param>
        /// <param name="column">开始列</param>
        /// <param name="mergerow">结束行</param>
        /// <param name="mergecolumn">结束列</param>
        public void MergeCell(int row, int column, int mergerow, int mergecolumn)
        {
            int tem_row = 1;
            int tem_column = 1;
            bool begin = false;
            if (row == mergerow)
            {
                foreach (var tableRow in table)
                {
                    if (tableRow.LocalName == "tr")
                    {
                        if (row == tem_row)
                        {
                            tem_column = 1;
                            foreach (var cell in tableRow)
                            {
                                if (cell.LocalName == "tc")
                                {
                                    if (tem_column > mergecolumn)
                                    {
                                        return;
                                    }
                                    if (column == tem_column && begin == false)
                                    {
                                        foreach (var cellProperties in cell)
                                        {
                                            if (cellProperties.LocalName == "tcPr")
                                            {
                                                var cPr = (TableCellProperties)cellProperties;
                                                cPr.HorizontalMerge = new HorizontalMerge() { Val = MergedCellValues.Restart };
                                                begin = true;
                                            }

                                        }

                                    }
                                    if (column < tem_column)
                                    {
                                        foreach (var cellProperties in cell)
                                        {
                                            if (cellProperties.LocalName == "tcPr")
                                            {
                                                var cPr = (TableCellProperties)cellProperties;
                                                cPr.VerticalMerge = new VerticalMerge() { Val = MergedCellValues.Continue }; cPr.HorizontalMerge = new HorizontalMerge() { Val = MergedCellValues.Continue };
                                                //  begin = true;
                                            }
                                        }
                                    }
                                    tem_column++;
                                }
                            }

                        }
                        tem_row++;
                    }
                }
            }

            if (row != mergerow)
            {
                foreach (var tableRow in table)
                {
                    if (tableRow.LocalName == "tr")
                    {
                        if (tem_row > mergerow)
                        {
                            return;
                        }
                        if (row == tem_row)
                        {
                            tem_column = 1;
                            foreach (var cell in tableRow)
                            {
                                if (cell.LocalName == "tc")
                                {
                                    if (tem_column > mergecolumn)
                                    {
                                        break;
                                    }
                                    if (column == tem_column && begin == false)
                                    {
                                        foreach (var cellProperties in cell)
                                        {
                                            if (cellProperties.LocalName == "tcPr")
                                            {
                                                var cPr = (TableCellProperties)cellProperties;
                                                cPr.HorizontalMerge = new HorizontalMerge() { Val = MergedCellValues.Restart };
                                                cPr.VerticalMerge = new VerticalMerge() { Val = MergedCellValues.Restart };
                                                begin = true;
                                            }

                                        }

                                    }
                                    if (column < tem_column)
                                    {
                                        foreach (var cellProperties in cell)
                                        {
                                            if (cellProperties.LocalName == "tcPr")
                                            {

                                                var cPr = (TableCellProperties)cellProperties;
                                                cPr.HorizontalMerge = new HorizontalMerge() { Val = MergedCellValues.Continue };
                                                cPr.VerticalMerge = new VerticalMerge() { Val = MergedCellValues.Restart };
                                                //  begin = true;
                                            }
                                        }
                                    }
                                    tem_column++;
                                }
                            }

                        }
                        else
                        {
                            tem_column = 1;
                            begin = false;
                            foreach (var cell in tableRow)
                            {
                                if (cell.LocalName == "tc")
                                {
                                    if (tem_column > mergecolumn)
                                    {
                                        break;
                                    }
                                    if (column == tem_column && begin == false)
                                    {
                                        foreach (var cellProperties in cell)
                                        {
                                            if (cellProperties.LocalName == "tcPr")
                                            {
                                                var cPr = (TableCellProperties)cellProperties;
                                                cPr.HorizontalMerge = new HorizontalMerge() { Val = MergedCellValues.Restart };
                                                cPr.VerticalMerge = new VerticalMerge() { Val = MergedCellValues.Continue };
                                                begin = true;
                                            }

                                        }

                                    }
                                    if (column < tem_column)
                                    {
                                        foreach (var cellProperties in cell)
                                        {
                                            if (cellProperties.LocalName == "tcPr")
                                            {

                                                var cPr = (TableCellProperties)cellProperties;
                                                cPr.HorizontalMerge = new HorizontalMerge() { Val = MergedCellValues.Continue };
                                                cPr.VerticalMerge = new VerticalMerge() { Val = MergedCellValues.Continue };
                                                //  begin = true;
                                            }
                                        }
                                    }
                                    tem_column++;
                                }
                            }
                        }
                        tem_row++;
                    }
                }
            }

        }

        private void CreateTable(int row, int column)
        {
            TableRow tableRow = new TableRow();
            TableCell tableCell = new TableCell();
            tableCell.Append(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2400" }));
            tableCell.Append(new Paragraph(new Run(new Text())));
            for (int i = 0; i < column; i++)
            {
                tableCell = (TableCell)tableCell.Clone();
                tableRow.Append(tableCell);
            }
            for (int i = 0; i < row; i++)
            {
                tableRow = (TableRow)tableRow.Clone();
                table.Append(tableRow);
            }
        }
    }
}

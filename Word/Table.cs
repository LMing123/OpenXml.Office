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
                if(topBorder!=null)
                {
                    borders.TopBorder.Size =(UInt32)topBorder;
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
                if(insideVerticalBorder!=null)
                {
                    borders.InsideVerticalBorder.Size = (UInt32)insideVerticalBorder;
                }
            }

        }

        //TODO not finished

        public void MergeCell(int row, int column, int mergerow, int mergecolumn)
        {
            var i = table.Elements<TableRow>().First();
            var j = i.First();
                var m = (TableCellProperties)j.First();
                m.HorizontalMerge = new HorizontalMerge() { Val = MergedCellValues.Restart };
                var j1 = j.ElementsAfter();
            foreach(var k in j1)
            {
               var  k1 =(TableCellProperties) k.First();

               k1 .HorizontalMerge = new HorizontalMerge() { Val = MergedCellValues.Continue };
            }

            //var j = i.Elements<TableCell>().First();
            //var k = j.Elements<TableCellProperties>().First();
            //k.HorizontalMerge= new HorizontalMerge() { Val = MergedCellValues.Restart };
            //var j1 = i.Elements<TableCell>().ElementAt(2);
            //var k1 = j.Elements<TableCellProperties>().First();

            //k1.HorizontalMerge = new HorizontalMerge() { Val = MergedCellValues.Continue };


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

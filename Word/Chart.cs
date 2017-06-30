using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using d = DocumentFormat.OpenXml.Drawing;
using dc = DocumentFormat.OpenXml.Drawing.Charts;
using dw = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using C14 = DocumentFormat.OpenXml.Office2010.Drawing.Charts;
using A = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml;

namespace Word
{
    public class Chart
    {
        private ChartPart chartPart;
        public Chart(ChartPart chartpart)
        {
            this.chartPart = chartpart;
        }

        public void AddNewBarAndLineChart(List<ChartSubArea> barChartList,int top,int bottom)
        {

            ChartSpace chartSpace1 = new ChartSpace();
#region 图表开头
            chartSpace1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            chartSpace1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            chartSpace1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            chartSpace1.AddNamespaceDeclaration("c16r2", "http://schemas.microsoft.com/office/drawing/2015/06/chart");
            C.Date1904 date19041 = new C.Date1904() { Val = false };
            C.EditingLanguage editingLanguage1 = new C.EditingLanguage() { Val = "zh-CN" };
            C.RoundedCorners roundedCorners1 = new C.RoundedCorners() { Val = false };

            AlternateContent alternateContent1 = new AlternateContent();
            alternateContent1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

            AlternateContentChoice alternateContentChoice1 = new AlternateContentChoice() { Requires = "c14" };
            alternateContentChoice1.AddNamespaceDeclaration("c14", "http://schemas.microsoft.com/office/drawing/2007/8/2/chart");
            C14.Style style5 = new C14.Style() { Val = 102 };

            alternateContentChoice1.Append(style5);

            AlternateContentFallback alternateContentFallback1 = new AlternateContentFallback();
            C.Style style6 = new C.Style() { Val = 2 };

            alternateContentFallback1.Append(style6);

            alternateContent1.Append(alternateContentChoice1);
            alternateContent1.Append(alternateContentFallback1);

            C.Chart chart1 = new C.Chart();

            C.Title title1 = new C.Title();
            C.Overlay overlay1 = new C.Overlay() { Val = false };

            C.ChartShapeProperties chartShapeProperties1 = new C.ChartShapeProperties();
            A.NoFill noFill1 = new A.NoFill();

            A.Outline outline4 = new A.Outline();
            A.NoFill noFill2 = new A.NoFill();

            outline4.Append(noFill2);
            A.EffectList effectList4 = new A.EffectList();

            chartShapeProperties1.Append(noFill1);
            chartShapeProperties1.Append(outline4);
            chartShapeProperties1.Append(effectList4);

            C.TextProperties textProperties1 = new C.TextProperties();
            A.BodyProperties bodyProperties1 = new A.BodyProperties() { Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
            A.ListStyle listStyle1 = new A.ListStyle();

            A.Paragraph paragraph1 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties1 = new A.ParagraphProperties();

            A.DefaultRunProperties defaultRunProperties1 = new A.DefaultRunProperties() { FontSize = 1400, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Spacing = 0, Baseline = 0 };

            A.SolidFill solidFill7 = new A.SolidFill();

            A.SchemeColor schemeColor16 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation9 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset1 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor16.Append(luminanceModulation9);
            schemeColor16.Append(luminanceOffset1);

            solidFill7.Append(schemeColor16);
            A.LatinFont latinFont3 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont3 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont3 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties1.Append(solidFill7);
            defaultRunProperties1.Append(latinFont3);
            defaultRunProperties1.Append(eastAsianFont3);
            defaultRunProperties1.Append(complexScriptFont3);

            paragraphProperties1.Append(defaultRunProperties1);
            A.EndParagraphRunProperties endParagraphRunProperties1 = new A.EndParagraphRunProperties() { Language = "zh-CN" };

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(endParagraphRunProperties1);

            textProperties1.Append(bodyProperties1);
            textProperties1.Append(listStyle1);
            textProperties1.Append(paragraph1);

            title1.Append(overlay1);
            title1.Append(chartShapeProperties1);
            title1.Append(textProperties1);
            C.AutoTitleDeleted autoTitleDeleted1 = new C.AutoTitleDeleted() { Val = false };

            C.PlotArea plotArea1 = new C.PlotArea();
            C.Layout layout1 = new C.Layout();
#endregion

            #region 柱状图
            C.BarChart barChart1 = new C.BarChart();
            C.BarDirection barDirection1 = new C.BarDirection() { Val = C.BarDirectionValues.Column };
            C.BarGrouping barGrouping1 = new C.BarGrouping() { Val = C.BarGroupingValues.Clustered };
            C.VaryColors varyColors1 = new C.VaryColors() { Val = false };

            C.BarChartSeries barChartSeries1 = new C.BarChartSeries();
            C.Index index1 = new C.Index() { Val = (UInt32Value)0U };
            C.Order order1 = new C.Order() { Val = (UInt32Value)0U };

            C.SeriesText seriesText1 = new C.SeriesText();

            C.StringReference stringReference1 = new C.StringReference();
            C.Formula formula1 = new C.Formula();
            formula1.Text = "Sheet1!$B$1";

            C.StringCache stringCache1 = new C.StringCache();
            C.PointCount pointCount1 = new C.PointCount() { Val = (UInt32Value)1U };

            C.StringPoint stringPoint1 = new C.StringPoint() { Index = (UInt32Value)0U };
            C.NumericValue numericValue1 = new C.NumericValue();
            numericValue1.Text = "系列 1";

            stringPoint1.Append(numericValue1);

            stringCache1.Append(pointCount1);
            stringCache1.Append(stringPoint1);

            stringReference1.Append(formula1);
            stringReference1.Append(stringCache1);

            seriesText1.Append(stringReference1);

            C.ChartShapeProperties chartShapeProperties2 = new C.ChartShapeProperties();

            A.SolidFill solidFill8 = new A.SolidFill();
            A.SchemeColor schemeColor17 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            solidFill8.Append(schemeColor17);

            A.Outline outline5 = new A.Outline();
            A.NoFill noFill3 = new A.NoFill();

            outline5.Append(noFill3);
            A.EffectList effectList5 = new A.EffectList();

            chartShapeProperties2.Append(solidFill8);
            chartShapeProperties2.Append(outline5);
            chartShapeProperties2.Append(effectList5);
            C.InvertIfNegative invertIfNegative1 = new C.InvertIfNegative() { Val = false };

            C.CategoryAxisData categoryAxisData1 = new C.CategoryAxisData();


            uint barcount =(uint) barChartList.Count;

            //标签
            C.StringReference stringReference2 = new C.StringReference();
            C.Formula formula2 = new C.Formula();
            formula2.Text = "Sheet1!$A$2:$A$5";
            C.StringCache stringCache2 = new C.StringCache();
            C.PointCount pointCount2 = new C.PointCount() { Val = (uint)barcount };
            stringCache2.Append(pointCount2);
            stringReference2.Append(formula2);
            //数据
            C.Values values1 = new C.Values();
            C.NumberReference numberReference1 = new C.NumberReference();
            C.Formula formula3 = new C.Formula();
            formula3.Text = "Sheet1!$B$2:$B$5";
            C.NumberingCache numberingCache1 = new C.NumberingCache();
            C.FormatCode formatCode1 = new C.FormatCode();
            formatCode1.Text = "General";
            C.PointCount pointCount3 = new C.PointCount() { Val = (UInt32Value)barcount };
            numberingCache1.Append(formatCode1);
            numberingCache1.Append(pointCount3);
            numberReference1.Append(formula3);

            uint rowcount = 0;
            foreach (var item in barChartList)
            {
                C.StringPoint stringPoint2 = new C.StringPoint() { Index = rowcount };
                C.NumericValue numericValue2 = new C.NumericValue() { Text=item.Label};
                stringPoint2.Append(numericValue2);
                stringCache2.Append(stringPoint2);

                C.NumericPoint numericPoint1 = new C.NumericPoint() { Index = rowcount };
                C.NumericValue numericValue6 = new C.NumericValue() { Text=item.Value};
                numericPoint1.Append(numericValue6);
                numberingCache1.Append(numericPoint1);

                rowcount++;
            }

            stringReference2.Append(stringCache2);
            categoryAxisData1.Append(stringReference2);

            numberReference1.Append(numberingCache1);
            values1.Append(numberReference1);

            C.BarSerExtensionList barSerExtensionList1 = new C.BarSerExtensionList();

            C.BarSerExtension barSerExtension1 = new C.BarSerExtension() { Uri = "{C3380CC4-5D6E-409C-BE32-E72D297353CC}" };
            barSerExtension1.AddNamespaceDeclaration("c16", "http://schemas.microsoft.com/office/drawing/2014/chart");

            OpenXmlUnknownElement openXmlUnknownElement1 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<c16:uniqueId val=\"{00000000-2FE6-46C9-9BDA-484F977967C3}\" xmlns:c16=\"http://schemas.microsoft.com/office/drawing/2014/chart\" />");

            barSerExtension1.Append(openXmlUnknownElement1);

            barSerExtensionList1.Append(barSerExtension1);

            barChartSeries1.Append(index1);
            barChartSeries1.Append(order1);
            barChartSeries1.Append(seriesText1);
            barChartSeries1.Append(chartShapeProperties2);
            barChartSeries1.Append(invertIfNegative1);
            barChartSeries1.Append(categoryAxisData1);
            barChartSeries1.Append(values1);
            barChartSeries1.Append(barSerExtensionList1);

            C.DataLabels dataLabels1 = new C.DataLabels();
            C.ShowLegendKey showLegendKey1 = new C.ShowLegendKey() { Val = false };
            C.ShowValue showValue1 = new C.ShowValue() { Val = false };
            C.ShowCategoryName showCategoryName1 = new C.ShowCategoryName() { Val = false };
            C.ShowSeriesName showSeriesName1 = new C.ShowSeriesName() { Val = false };
            C.ShowPercent showPercent1 = new C.ShowPercent() { Val = false };
            C.ShowBubbleSize showBubbleSize1 = new C.ShowBubbleSize() { Val = false };

            dataLabels1.Append(showLegendKey1);
            dataLabels1.Append(showValue1);
            dataLabels1.Append(showCategoryName1);
            dataLabels1.Append(showSeriesName1);
            dataLabels1.Append(showPercent1);
            dataLabels1.Append(showBubbleSize1);
            C.GapWidth gapWidth1 = new C.GapWidth() { Val = (UInt16Value)219U };
            C.Overlap overlap1 = new C.Overlap() { Val = -27 };
            C.AxisId axisId1 = new C.AxisId() { Val = (UInt32Value)298007167U };
            C.AxisId axisId2 = new C.AxisId() { Val = (UInt32Value)298002591U };

            barChart1.Append(barDirection1);
            barChart1.Append(barGrouping1);
            barChart1.Append(varyColors1);
            barChart1.Append(barChartSeries1);
            barChart1.Append(dataLabels1);
            barChart1.Append(gapWidth1);
            barChart1.Append(overlap1);
            barChart1.Append(axisId1);
            barChart1.Append(axisId2);
            #endregion

#region 条形图
            C.LineChart lineChart1 = new C.LineChart();
            C.Grouping grouping1 = new C.Grouping() { Val = C.GroupingValues.Standard };
            C.VaryColors varyColors2 = new C.VaryColors() { Val = false };

            C.LineChartSeries lineChartSeries1 = new C.LineChartSeries();
            C.Index index2 = new C.Index() { Val = (UInt32Value)1U };
            C.Order order2 = new C.Order() { Val = (UInt32Value)1U };

            C.SeriesText seriesText2 = new C.SeriesText();

            C.StringReference stringReference3 = new C.StringReference();
            C.Formula formula4 = new C.Formula();
            formula4.Text = "Sheet1!$C$1";
            C.StringCache stringCache3 = new C.StringCache();
            C.PointCount pointCount4 = new C.PointCount() { Val = (UInt32Value)1U };
            C.StringPoint stringPoint6 = new C.StringPoint() { Index = (UInt32Value)0U };
            C.NumericValue numericValue10 = new C.NumericValue();
            numericValue10.Text = "系列 2";
            stringPoint6.Append(numericValue10);
            stringCache3.Append(pointCount4);
            stringCache3.Append(stringPoint6);

            stringReference3.Append(formula4);
            stringReference3.Append(stringCache3);
            seriesText2.Append(stringReference3);

            C.ChartShapeProperties chartShapeProperties3 = new C.ChartShapeProperties();

            A.Outline outline6 = new A.Outline() { Width = 28575, CapType = A.LineCapValues.Round };

            A.SolidFill solidFill9 = new A.SolidFill();
            A.SchemeColor schemeColor18 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent2 };

            solidFill9.Append(schemeColor18);
            A.Round round1 = new A.Round();

            outline6.Append(solidFill9);
            outline6.Append(round1);
            A.EffectList effectList6 = new A.EffectList();

            chartShapeProperties3.Append(outline6);
            chartShapeProperties3.Append(effectList6);

            C.Marker marker1 = new C.Marker();
            C.Symbol symbol1 = new C.Symbol() { Val = C.MarkerStyleValues.None };

            marker1.Append(symbol1);

            C.CategoryAxisData categoryAxisData2 = new C.CategoryAxisData();

            C.StringReference stringReference4 = new C.StringReference();
            C.Formula formula5 = new C.Formula();
            formula5.Text = "Sheet1!$A$2:$A$5";
            C.StringCache stringCache4 = new C.StringCache();
            C.PointCount pointCount5 = new C.PointCount() { Val = (UInt32Value)barcount };
            stringCache4.Append(pointCount5);
            stringReference4.Append(formula5);
            stringReference4.Append(stringCache4);

            C.Values values2 = new C.Values();
            C.NumberReference numberReference2 = new C.NumberReference();
            C.Formula formula6 = new C.Formula();
            formula6.Text = "Sheet1!$C$2:$C$5";
            C.NumberingCache numberingCache2 = new C.NumberingCache();
            C.FormatCode formatCode2 = new C.FormatCode();
            formatCode2.Text = "General";
            C.PointCount pointCount6 = new C.PointCount() { Val = (UInt32Value)barcount };
            numberingCache2.Append(formatCode2);
            numberingCache2.Append(pointCount6);
            numberReference2.Append(formula6);

            for(uint i=0;i<barcount;i++)
            {
                C.StringPoint stringPoint7 = new C.StringPoint() { Index = (UInt32Value)i };
                C.NumericValue numericValue11 = new C.NumericValue();
                numericValue11.Text = "类别 1";
                stringPoint7.Append(numericValue11);
                stringCache4.Append(stringPoint7);

                C.NumericPoint numericPoint5 = new C.NumericPoint() { Index = (UInt32Value)i };
                C.NumericValue numericValue15 = new C.NumericValue() { Text=top.ToString()};
                numericPoint5.Append(numericValue15);
                numberingCache2.Append(numericPoint5);


            }

            categoryAxisData2.Append(stringReference4);             
            numberReference2.Append(numberingCache2);

            values2.Append(numberReference2);
            C.Smooth smooth1 = new C.Smooth() { Val = false };

            C.LineSerExtensionList lineSerExtensionList1 = new C.LineSerExtensionList();

            C.LineSerExtension lineSerExtension1 = new C.LineSerExtension() { Uri = "{C3380CC4-5D6E-409C-BE32-E72D297353CC}" };
            lineSerExtension1.AddNamespaceDeclaration("c16", "http://schemas.microsoft.com/office/drawing/2014/chart");

            OpenXmlUnknownElement openXmlUnknownElement2 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<c16:uniqueId val=\"{00000001-2FE6-46C9-9BDA-484F977967C3}\" xmlns:c16=\"http://schemas.microsoft.com/office/drawing/2014/chart\" />");

            lineSerExtension1.Append(openXmlUnknownElement2);

            lineSerExtensionList1.Append(lineSerExtension1);

            lineChartSeries1.Append(index2);
            lineChartSeries1.Append(order2);
            lineChartSeries1.Append(seriesText2);
            lineChartSeries1.Append(chartShapeProperties3);
            lineChartSeries1.Append(marker1);
            lineChartSeries1.Append(categoryAxisData2);
            lineChartSeries1.Append(values2);
            lineChartSeries1.Append(smooth1);
            lineChartSeries1.Append(lineSerExtensionList1);
            #endregion
#region 第二线图
            C.LineChartSeries lineChartSeries2 = new C.LineChartSeries();
            C.Index index3 = new C.Index() { Val = (UInt32Value)2U };
            C.Order order3 = new C.Order() { Val = (UInt32Value)2U };

            C.SeriesText seriesText3 = new C.SeriesText();

            C.StringReference stringReference5 = new C.StringReference();
            C.Formula formula7 = new C.Formula();
            formula7.Text = "Sheet1!$D$1";

            C.StringCache stringCache5 = new C.StringCache();
            C.PointCount pointCount7 = new C.PointCount() { Val = (UInt32Value)1U };

            C.StringPoint stringPoint11 = new C.StringPoint() { Index = (UInt32Value)0U };
            C.NumericValue numericValue19 = new C.NumericValue();
            numericValue19.Text = "系列 3";

            stringPoint11.Append(numericValue19);

            stringCache5.Append(pointCount7);
            stringCache5.Append(stringPoint11);

            stringReference5.Append(formula7);
            stringReference5.Append(stringCache5);

            seriesText3.Append(stringReference5);

            C.ChartShapeProperties chartShapeProperties4 = new C.ChartShapeProperties();

            A.Outline outline7 = new A.Outline() { Width = 28575, CapType = A.LineCapValues.Round };

            A.SolidFill solidFill10 = new A.SolidFill();
            A.SchemeColor schemeColor19 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent3 };

            solidFill10.Append(schemeColor19);
            A.Round round2 = new A.Round();

            outline7.Append(solidFill10);
            outline7.Append(round2);
            A.EffectList effectList7 = new A.EffectList();

            chartShapeProperties4.Append(outline7);
            chartShapeProperties4.Append(effectList7);

            C.Marker marker2 = new C.Marker();
            C.Symbol symbol2 = new C.Symbol() { Val = C.MarkerStyleValues.None };

            marker2.Append(symbol2);

            C.CategoryAxisData categoryAxisData3 = new C.CategoryAxisData();
            //标签
            C.StringReference stringReference6 = new C.StringReference();
            C.Formula formula8 = new C.Formula();
            formula8.Text = "Sheet1!$A$2:$A$5";
            C.StringCache stringCache6 = new C.StringCache();
            C.PointCount pointCount8 = new C.PointCount() { Val = (UInt32Value)barcount };
            stringCache6.Append(pointCount8);
            stringReference6.Append(formula8);

            //数据
            C.Values values3 = new C.Values();
            C.NumberReference numberReference3 = new C.NumberReference();
            C.Formula formula9 = new C.Formula();
            formula9.Text = "Sheet1!$D$2:$D$5";
            C.NumberingCache numberingCache3 = new C.NumberingCache();
            C.FormatCode formatCode3 = new C.FormatCode();
            formatCode3.Text = "General";
            C.PointCount pointCount9 = new C.PointCount() { Val = (UInt32Value)barcount };
            numberingCache3.Append(formatCode3);
            numberingCache3.Append(pointCount9);
            numberReference3.Append(formula9);


            for (uint i=0;i<barcount;i++)
            {
                C.StringPoint stringPoint12 = new C.StringPoint() { Index = (UInt32Value)i };
                C.NumericValue numericValue20 = new C.NumericValue() { Text= "类别 2" };
                stringPoint12.Append(numericValue20);
                stringCache6.Append(stringPoint12);

                C.NumericPoint numericPoint9 = new C.NumericPoint() { Index = (UInt32Value)i };
                C.NumericValue numericValue24 = new C.NumericValue() { Text=bottom.ToString()};
                numericPoint9.Append(numericValue24);
                numberingCache3.Append(numericPoint9);

            }

            stringReference6.Append(stringCache6);
            numberReference3.Append(numberingCache3);
            values3.Append(numberReference3);
            categoryAxisData3.Append(stringReference6);

            C.Smooth smooth2 = new C.Smooth() { Val = false };

            C.LineSerExtensionList lineSerExtensionList2 = new C.LineSerExtensionList();

            C.LineSerExtension lineSerExtension2 = new C.LineSerExtension() { Uri = "{C3380CC4-5D6E-409C-BE32-E72D297353CC}" };
            lineSerExtension2.AddNamespaceDeclaration("c16", "http://schemas.microsoft.com/office/drawing/2014/chart");

            OpenXmlUnknownElement openXmlUnknownElement3 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<c16:uniqueId val=\"{00000002-2FE6-46C9-9BDA-484F977967C3}\" xmlns:c16=\"http://schemas.microsoft.com/office/drawing/2014/chart\" />");

            lineSerExtension2.Append(openXmlUnknownElement3);

            lineSerExtensionList2.Append(lineSerExtension2);

            lineChartSeries2.Append(index3);
            lineChartSeries2.Append(order3);
            lineChartSeries2.Append(seriesText3);
            lineChartSeries2.Append(chartShapeProperties4);
            lineChartSeries2.Append(marker2);
            lineChartSeries2.Append(categoryAxisData3);
            lineChartSeries2.Append(values3);
            lineChartSeries2.Append(smooth2);
            lineChartSeries2.Append(lineSerExtensionList2);

            C.DataLabels dataLabels2 = new C.DataLabels();
            C.ShowLegendKey showLegendKey2 = new C.ShowLegendKey() { Val = false };
            C.ShowValue showValue2 = new C.ShowValue() { Val = false };
            C.ShowCategoryName showCategoryName2 = new C.ShowCategoryName() { Val = false };
            C.ShowSeriesName showSeriesName2 = new C.ShowSeriesName() { Val = false };
            C.ShowPercent showPercent2 = new C.ShowPercent() { Val = false };
            C.ShowBubbleSize showBubbleSize2 = new C.ShowBubbleSize() { Val = false };

            dataLabels2.Append(showLegendKey2);
            dataLabels2.Append(showValue2);
            dataLabels2.Append(showCategoryName2);
            dataLabels2.Append(showSeriesName2);
            dataLabels2.Append(showPercent2);
            dataLabels2.Append(showBubbleSize2);
            C.ShowMarker showMarker1 = new C.ShowMarker() { Val = true };
            C.Smooth smooth3 = new C.Smooth() { Val = false };
            C.AxisId axisId3 = new C.AxisId() { Val = (UInt32Value)298007167U };
            C.AxisId axisId4 = new C.AxisId() { Val = (UInt32Value)298002591U };

            lineChart1.Append(grouping1);
            lineChart1.Append(varyColors2);
            lineChart1.Append(lineChartSeries1);
            lineChart1.Append(lineChartSeries2);
            lineChart1.Append(dataLabels2);
            lineChart1.Append(showMarker1);
            lineChart1.Append(smooth3);
            lineChart1.Append(axisId3);
            lineChart1.Append(axisId4);
#endregion
            C.CategoryAxis categoryAxis1 = new C.CategoryAxis();
            C.AxisId axisId5 = new C.AxisId() { Val = (UInt32Value)298007167U };

            C.Scaling scaling1 = new C.Scaling();
            C.Orientation orientation1 = new C.Orientation() { Val = C.OrientationValues.MinMax };

            scaling1.Append(orientation1);
            C.Delete delete1 = new C.Delete() { Val = false };
            C.AxisPosition axisPosition1 = new C.AxisPosition() { Val = C.AxisPositionValues.Bottom };
            C.NumberingFormat numberingFormat1 = new C.NumberingFormat() { FormatCode = "General", SourceLinked = true };
            C.MajorTickMark majorTickMark1 = new C.MajorTickMark() { Val = C.TickMarkValues.None };
            C.MinorTickMark minorTickMark1 = new C.MinorTickMark() { Val = C.TickMarkValues.None };
            C.TickLabelPosition tickLabelPosition1 = new C.TickLabelPosition() { Val = C.TickLabelPositionValues.NextTo };

            C.ChartShapeProperties chartShapeProperties5 = new C.ChartShapeProperties();
            A.NoFill noFill4 = new A.NoFill();

            A.Outline outline8 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill11 = new A.SolidFill();

            A.SchemeColor schemeColor20 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation10 = new A.LuminanceModulation() { Val = 15000 };
            A.LuminanceOffset luminanceOffset2 = new A.LuminanceOffset() { Val = 85000 };

            schemeColor20.Append(luminanceModulation10);
            schemeColor20.Append(luminanceOffset2);

            solidFill11.Append(schemeColor20);
            A.Round round3 = new A.Round();

            outline8.Append(solidFill11);
            outline8.Append(round3);
            A.EffectList effectList8 = new A.EffectList();

            chartShapeProperties5.Append(noFill4);
            chartShapeProperties5.Append(outline8);
            chartShapeProperties5.Append(effectList8);

            C.TextProperties textProperties2 = new C.TextProperties();
            A.BodyProperties bodyProperties2 = new A.BodyProperties() { Rotation = -60000000, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
            A.ListStyle listStyle2 = new A.ListStyle();

            A.Paragraph paragraph2 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties2 = new A.ParagraphProperties();

            A.DefaultRunProperties defaultRunProperties2 = new A.DefaultRunProperties() { FontSize = 900, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

            A.SolidFill solidFill12 = new A.SolidFill();

            A.SchemeColor schemeColor21 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation11 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset3 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor21.Append(luminanceModulation11);
            schemeColor21.Append(luminanceOffset3);

            solidFill12.Append(schemeColor21);
            A.LatinFont latinFont4 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont4 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont4 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties2.Append(solidFill12);
            defaultRunProperties2.Append(latinFont4);
            defaultRunProperties2.Append(eastAsianFont4);
            defaultRunProperties2.Append(complexScriptFont4);

            paragraphProperties2.Append(defaultRunProperties2);
            A.EndParagraphRunProperties endParagraphRunProperties2 = new A.EndParagraphRunProperties() { Language = "zh-CN" };

            paragraph2.Append(paragraphProperties2);
            paragraph2.Append(endParagraphRunProperties2);

            textProperties2.Append(bodyProperties2);
            textProperties2.Append(listStyle2);
            textProperties2.Append(paragraph2);
            C.CrossingAxis crossingAxis1 = new C.CrossingAxis() { Val = (UInt32Value)298002591U };
            C.Crosses crosses1 = new C.Crosses() { Val = C.CrossesValues.AutoZero };
            C.AutoLabeled autoLabeled1 = new C.AutoLabeled() { Val = true };
            C.LabelAlignment labelAlignment1 = new C.LabelAlignment() { Val = C.LabelAlignmentValues.Center };
            C.LabelOffset labelOffset1 = new C.LabelOffset() { Val = (UInt16Value)100U };
            C.NoMultiLevelLabels noMultiLevelLabels1 = new C.NoMultiLevelLabels() { Val = false };

            categoryAxis1.Append(axisId5);
            categoryAxis1.Append(scaling1);
            categoryAxis1.Append(delete1);
            categoryAxis1.Append(axisPosition1);
            categoryAxis1.Append(numberingFormat1);
            categoryAxis1.Append(majorTickMark1);
            categoryAxis1.Append(minorTickMark1);
            categoryAxis1.Append(tickLabelPosition1);
            categoryAxis1.Append(chartShapeProperties5);
            categoryAxis1.Append(textProperties2);
            categoryAxis1.Append(crossingAxis1);
            categoryAxis1.Append(crosses1);
            categoryAxis1.Append(autoLabeled1);
            categoryAxis1.Append(labelAlignment1);
            categoryAxis1.Append(labelOffset1);
            categoryAxis1.Append(noMultiLevelLabels1);

            C.ValueAxis valueAxis1 = new C.ValueAxis();
            C.AxisId axisId6 = new C.AxisId() { Val = (UInt32Value)298002591U };

            C.Scaling scaling2 = new C.Scaling();
            C.Orientation orientation2 = new C.Orientation() { Val = C.OrientationValues.MinMax };

            scaling2.Append(orientation2);
            C.Delete delete2 = new C.Delete() { Val = false };
            C.AxisPosition axisPosition2 = new C.AxisPosition() { Val = C.AxisPositionValues.Left };

            C.MajorGridlines majorGridlines1 = new C.MajorGridlines();

            C.ChartShapeProperties chartShapeProperties6 = new C.ChartShapeProperties();

            A.Outline outline9 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill13 = new A.SolidFill();

            A.SchemeColor schemeColor22 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation12 = new A.LuminanceModulation() { Val = 15000 };
            A.LuminanceOffset luminanceOffset4 = new A.LuminanceOffset() { Val = 85000 };

            schemeColor22.Append(luminanceModulation12);
            schemeColor22.Append(luminanceOffset4);

            solidFill13.Append(schemeColor22);
            A.Round round4 = new A.Round();

            outline9.Append(solidFill13);
            outline9.Append(round4);
            A.EffectList effectList9 = new A.EffectList();

            chartShapeProperties6.Append(outline9);
            chartShapeProperties6.Append(effectList9);

            majorGridlines1.Append(chartShapeProperties6);
            C.NumberingFormat numberingFormat2 = new C.NumberingFormat() { FormatCode = "General", SourceLinked = true };
            C.MajorTickMark majorTickMark2 = new C.MajorTickMark() { Val = C.TickMarkValues.None };
            C.MinorTickMark minorTickMark2 = new C.MinorTickMark() { Val = C.TickMarkValues.None };
            C.TickLabelPosition tickLabelPosition2 = new C.TickLabelPosition() { Val = C.TickLabelPositionValues.NextTo };

            C.ChartShapeProperties chartShapeProperties7 = new C.ChartShapeProperties();
            A.NoFill noFill5 = new A.NoFill();

            A.Outline outline10 = new A.Outline();
            A.NoFill noFill6 = new A.NoFill();

            outline10.Append(noFill6);
            A.EffectList effectList10 = new A.EffectList();

            chartShapeProperties7.Append(noFill5);
            chartShapeProperties7.Append(outline10);
            chartShapeProperties7.Append(effectList10);

            C.TextProperties textProperties3 = new C.TextProperties();
            A.BodyProperties bodyProperties3 = new A.BodyProperties() { Rotation = -60000000, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
            A.ListStyle listStyle3 = new A.ListStyle();

            A.Paragraph paragraph3 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties3 = new A.ParagraphProperties();

            A.DefaultRunProperties defaultRunProperties3 = new A.DefaultRunProperties() { FontSize = 900, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

            A.SolidFill solidFill14 = new A.SolidFill();

            A.SchemeColor schemeColor23 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation13 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset5 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor23.Append(luminanceModulation13);
            schemeColor23.Append(luminanceOffset5);

            solidFill14.Append(schemeColor23);
            A.LatinFont latinFont5 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont5 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont5 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties3.Append(solidFill14);
            defaultRunProperties3.Append(latinFont5);
            defaultRunProperties3.Append(eastAsianFont5);
            defaultRunProperties3.Append(complexScriptFont5);

            paragraphProperties3.Append(defaultRunProperties3);
            A.EndParagraphRunProperties endParagraphRunProperties3 = new A.EndParagraphRunProperties() { Language = "zh-CN" };

            paragraph3.Append(paragraphProperties3);
            paragraph3.Append(endParagraphRunProperties3);

            textProperties3.Append(bodyProperties3);
            textProperties3.Append(listStyle3);
            textProperties3.Append(paragraph3);
            C.CrossingAxis crossingAxis2 = new C.CrossingAxis() { Val = (UInt32Value)298007167U };
            C.Crosses crosses2 = new C.Crosses() { Val = C.CrossesValues.AutoZero };
            C.CrossBetween crossBetween1 = new C.CrossBetween() { Val = C.CrossBetweenValues.Between };

            valueAxis1.Append(axisId6);
            valueAxis1.Append(scaling2);
            valueAxis1.Append(delete2);
            valueAxis1.Append(axisPosition2);
            valueAxis1.Append(majorGridlines1);
            valueAxis1.Append(numberingFormat2);
            valueAxis1.Append(majorTickMark2);
            valueAxis1.Append(minorTickMark2);
            valueAxis1.Append(tickLabelPosition2);
            valueAxis1.Append(chartShapeProperties7);
            valueAxis1.Append(textProperties3);
            valueAxis1.Append(crossingAxis2);
            valueAxis1.Append(crosses2);
            valueAxis1.Append(crossBetween1);

            C.ShapeProperties shapeProperties1 = new C.ShapeProperties();
            A.NoFill noFill7 = new A.NoFill();

            A.Outline outline11 = new A.Outline();
            A.NoFill noFill8 = new A.NoFill();

            outline11.Append(noFill8);
            A.EffectList effectList11 = new A.EffectList();

            shapeProperties1.Append(noFill7);
            shapeProperties1.Append(outline11);
            shapeProperties1.Append(effectList11);

            plotArea1.Append(layout1);
            plotArea1.Append(barChart1);
            plotArea1.Append(lineChart1);
            plotArea1.Append(categoryAxis1);
            plotArea1.Append(valueAxis1);
            plotArea1.Append(shapeProperties1);

            C.Legend legend1 = new C.Legend();
            C.LegendPosition legendPosition1 = new C.LegendPosition() { Val = C.LegendPositionValues.Bottom };
            C.Overlay overlay2 = new C.Overlay() { Val = false };

            C.ChartShapeProperties chartShapeProperties8 = new C.ChartShapeProperties();
            A.NoFill noFill9 = new A.NoFill();

            A.Outline outline12 = new A.Outline();
            A.NoFill noFill10 = new A.NoFill();

            outline12.Append(noFill10);
            A.EffectList effectList12 = new A.EffectList();

            chartShapeProperties8.Append(noFill9);
            chartShapeProperties8.Append(outline12);
            chartShapeProperties8.Append(effectList12);

            C.TextProperties textProperties4 = new C.TextProperties();
            A.BodyProperties bodyProperties4 = new A.BodyProperties() { Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
            A.ListStyle listStyle4 = new A.ListStyle();

            A.Paragraph paragraph4 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties4 = new A.ParagraphProperties();

            A.DefaultRunProperties defaultRunProperties4 = new A.DefaultRunProperties() { FontSize = 900, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

            A.SolidFill solidFill15 = new A.SolidFill();

            A.SchemeColor schemeColor24 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation14 = new A.LuminanceModulation() { Val = 65000 };
            A.LuminanceOffset luminanceOffset6 = new A.LuminanceOffset() { Val = 35000 };

            schemeColor24.Append(luminanceModulation14);
            schemeColor24.Append(luminanceOffset6);

            solidFill15.Append(schemeColor24);
            A.LatinFont latinFont6 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont6 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont6 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties4.Append(solidFill15);
            defaultRunProperties4.Append(latinFont6);
            defaultRunProperties4.Append(eastAsianFont6);
            defaultRunProperties4.Append(complexScriptFont6);

            paragraphProperties4.Append(defaultRunProperties4);
            A.EndParagraphRunProperties endParagraphRunProperties4 = new A.EndParagraphRunProperties() { Language = "zh-CN" };

            paragraph4.Append(paragraphProperties4);
            paragraph4.Append(endParagraphRunProperties4);

            textProperties4.Append(bodyProperties4);
            textProperties4.Append(listStyle4);
            textProperties4.Append(paragraph4);

            legend1.Append(legendPosition1);
            legend1.Append(overlay2);
            legend1.Append(chartShapeProperties8);
            legend1.Append(textProperties4);
            C.PlotVisibleOnly plotVisibleOnly1 = new C.PlotVisibleOnly() { Val = true };
            C.DisplayBlanksAs displayBlanksAs1 = new C.DisplayBlanksAs() { Val = C.DisplayBlanksAsValues.Gap };
            C.ShowDataLabelsOverMaximum showDataLabelsOverMaximum1 = new C.ShowDataLabelsOverMaximum() { Val = false };

            chart1.Append(title1);
            chart1.Append(autoTitleDeleted1);
            chart1.Append(plotArea1);
            chart1.Append(legend1);
            chart1.Append(plotVisibleOnly1);
            chart1.Append(displayBlanksAs1);
            chart1.Append(showDataLabelsOverMaximum1);

            C.ShapeProperties shapeProperties2 = new C.ShapeProperties();

            A.SolidFill solidFill16 = new A.SolidFill();
            A.SchemeColor schemeColor25 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill16.Append(schemeColor25);

            A.Outline outline13 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill17 = new A.SolidFill();

            A.SchemeColor schemeColor26 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation15 = new A.LuminanceModulation() { Val = 15000 };
            A.LuminanceOffset luminanceOffset7 = new A.LuminanceOffset() { Val = 85000 };

            schemeColor26.Append(luminanceModulation15);
            schemeColor26.Append(luminanceOffset7);

            solidFill17.Append(schemeColor26);
            A.Round round5 = new A.Round();

            outline13.Append(solidFill17);
            outline13.Append(round5);
            A.EffectList effectList13 = new A.EffectList();

            shapeProperties2.Append(solidFill16);
            shapeProperties2.Append(outline13);
            shapeProperties2.Append(effectList13);

            C.TextProperties textProperties5 = new C.TextProperties();
            A.BodyProperties bodyProperties5 = new A.BodyProperties();
            A.ListStyle listStyle5 = new A.ListStyle();

            A.Paragraph paragraph5 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties5 = new A.ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties5 = new A.DefaultRunProperties();

            paragraphProperties5.Append(defaultRunProperties5);
            A.EndParagraphRunProperties endParagraphRunProperties5 = new A.EndParagraphRunProperties() { Language = "zh-CN" };

            paragraph5.Append(paragraphProperties5);
            paragraph5.Append(endParagraphRunProperties5);

            textProperties5.Append(bodyProperties5);
            textProperties5.Append(listStyle5);
            textProperties5.Append(paragraph5);

            //C.ExternalData externalData1 = new C.ExternalData() { Id = "rId3" };
            C.AutoUpdate autoUpdate1 = new C.AutoUpdate() { Val = false };

          //  externalData1.Append(autoUpdate1);

            chartSpace1.Append(date19041);
            chartSpace1.Append(editingLanguage1);
            chartSpace1.Append(roundedCorners1);
            chartSpace1.Append(alternateContent1);
            chartSpace1.Append(chart1);
            chartSpace1.Append(shapeProperties2);
            chartSpace1.Append(textProperties5);
           // chartSpace1.Append(externalData1);


            chartPart.ChartSpace = chartSpace1;
        }
    }
}

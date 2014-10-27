using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using D = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

using System.IO;
using System.IO.Packaging;
using System.Xml;
using System.Xml.Linq;

using System.Data.Common;

namespace artOPENXML
{
    // for PowerPoint functions
    static class PPTX
    {

         // this should be eventually defined in a config file
        public const string SLIDE_COLUMN_NAME = "Report_TableNo";
        public const string SHAPE_COLUMN_NAME = "Table Indicator";
        public const string FIRM_COLUMN_NAME = "Firm Indicator";
        public const string STUB_COLUMN_NAME = "Header";
        public const string BANNER_COLUMN_NAME = "Banner Display";
        public const string MEASURE_COLUMN_NAME = "GA - Measure";
        public const int HEADER_ROW = 1;
        public static readonly string[] GA_CHART_COLOR = { "508E7E", "F5E8B4", "005941", "8ABCAD", "DEB407", "FFFFFF", "000000", "808080" };
        public const string SLIDE_ID_TAGNAME = "SlideID";

        // get SlidePart by Name
        public static SlidePart getSlidePart(PresentationPart presentationPart, string slideName)
        {
            return presentationPart.SlideParts.Where(s => s.Slide.CommonSlideData.Name == slideName).SingleOrDefault();
        }

        // get SlidePart by Index
        public static SlidePart getSlidePart(PresentationPart presentationPart, int slideIndex)
        {
            // note that we can only trust SlideIdList's index order (SlideParts will give us different order)
            SlideId slideID = (SlideId)presentationPart.Presentation.SlideIdList.ElementAt(slideIndex).SingleOrDefault();
            if (slideID != null)
                return (SlidePart)presentationPart.GetPartById(slideID.RelationshipId.ToString());
            else
                return null;
        }

        // get SlidePart by matching tag (name & value pair); 
        //  It could have multiple slides that match the tag, only the first one will be returned.
        public static SlidePart getSlidePart(PresentationPart presentationPart, string tagName, string tagValue)
        {
            return (from sPart in presentationPart.SlideParts
                    from Tag tag in getTagPart(sPart).TagList
                    where tag.Name == tagName && tag.Val == tagValue
                    select sPart).FirstOrDefault();
        }

        // get SlideId
        internal static SlideId getSlideId(PresentationPart presentationPart, SlidePart slidePart)
        {
            int slideIndex = getSlideIndex(presentationPart, slidePart);
            SlideId slideId = (SlideId)presentationPart.Presentation.SlideIdList.ElementAt(slideIndex).SingleOrDefault();
            return slideId;
        }

        // get Slide Index; if not found, return -1
        public static int getSlideIndex(PresentationPart presentationPart, SlidePart slidePart)
        {
            int rtnIndex = -1, index = -1;
            foreach (var slideId in presentationPart.Presentation.SlideIdList.Elements<SlideId>())
            {
                index++;
                if (presentationPart.GetPartById(slideId.RelationshipId.ToString()).Equals(slidePart))
                {
                    rtnIndex = index;
                    break;
                }
            }
            return rtnIndex;
        }

        // get Slide Index from SlideID tag (tagName=SLIDE_ID_TAGNAME, a constant)
        public static int getSlideIndex(PresentationPart presentationPart, string slideIdTagValue)
        {
            return getSlideIndex(presentationPart,getSlidePart(presentationPart,SLIDE_ID_TAGNAME, slideIdTagValue));
        }

        // get Slide Name - it's not necessary all slide has SlideName (return null without name)
        public static string getSlideName(SlidePart slidePart)
        {
            return slidePart.Slide.CommonSlideData.Name;
        }

       // set Slide Name
        public static void setSlideName(SlidePart slidePart, string slideName)
        {
            slidePart.Slide.CommonSlideData.Name = slideName;
        }

        // get shape (excluding chart/table) tagPart 
        //  use dynamic to accommodate 4 different types (persenationPart/slidePart/graphicFrame/Shape) of "taggable" PPT objects
        internal static UserDefinedTagsPart getTagPart(dynamic part)
        {
            UserDefinedTagsPart tagPart = null;

            // get relationship from the "part"
            string relID = getTagRelationshipId(part);

            // SlidePart for graphicframe/shape/slidePart; PresentationPart for presentationPart
            dynamic pkgPart = getPackagePart(part);

            // get UserDefinedTagsPart from relationshipId
            if (relID != null)
                tagPart = (UserDefinedTagsPart)pkgPart.GetPartById(relID);

            return tagPart;
        }

        // get tagValue by tagName
        //  use dynamic to accommodate 4 different types (persenationPart/slidePart/graphicFrame/Shape) of "taggable" PPT objects
        public static string getTagValue(dynamic part, string tagName)
        {
            string tagValue = null;

            // get UserDefinedTagsPart object (generic for all tags)
            UserDefinedTagsPart tagPart = getTagPart(part);

            if (tagPart != null)
            {
                Tag tag = tagPart.TagList.Elements<Tag>().Where(t => String.Compare(t.Name, tagName, true) == 0).SingleOrDefault();
                if (tag != null)
                    tagValue = tag.Val.ToString();
            }

            return tagValue;
        }

        // get tagValue by shapeName (including graphicFrame/shape) and tagName
        public static string getTagValue(SlidePart slidePart, string shapeName, string tagName)
        {
            string tagValue = null;

            // either GraphicFrame or Shape
            GraphicFrame gFrame = getGraphicFrame(slidePart, shapeName);
            if (gFrame != null)
                tagValue = getTagValue(gFrame, tagName);
            else
            {
                Shape shape = getShape(slidePart, shapeName);
                if (shape != null)
                    tagValue = getTagValue(shape, tagName);
            }
            
            return tagValue;
        }

        // get "package" part (either slidepart or presentationpart)
        internal static dynamic getPackagePart(dynamic part)
        {
            // SlidePart for graphicframe/shape/slidePart; PresentationPart for presentationPart
            dynamic pkgPart = part;

            // use try block to simplify the code to check missing tags (null object reference error)
            try
            {
                // GraphicFrame (part)
                if (typeof(GraphicFrame) == part.GetType())
                    pkgPart = ((GraphicFrame)part).Ancestors<Slide>().Single().SlidePart;

                // Shape (part)
                if (typeof(Shape) == part.GetType())
                    pkgPart = ((Shape)part).Ancestors<Slide>().Single().SlidePart;

                //relID = cDataList.Elements<CustomerDataTags>().SingleOrDefault().Id.ToString();
            }
            catch { }

            return pkgPart;
        }

        // get Tag relationshipId (different from getIDofPart), this routine does not require TagsPart argument
        internal static string getTagRelationshipId(dynamic part)
        {
            string relID = null;

            // use try block to simplify the code to check missing tags (null object reference error)
            try
            {
                CustomerDataList cDataList = null;

                // Presentation Part
                if (typeof(PresentationPart) == part.GetType())
                    cDataList = (CustomerDataList)part.Presentation.CustomerDataList;

                // Slide Part
                if (typeof(SlidePart) == part.GetType())
                     cDataList = (CustomerDataList)part.Slide.CommonSlideData.CustomerDataList;

                // GraphicFrame (part)
                if (typeof(GraphicFrame) == part.GetType())
                    cDataList = (CustomerDataList)((GraphicFrame)part).NonVisualGraphicFrameProperties.ApplicationNonVisualDrawingProperties.Elements<CustomerDataList>().FirstOrDefault();

                // Shape (part)
                if (typeof(Shape) == part.GetType())
                    cDataList = (CustomerDataList)((Shape)part).NonVisualShapeProperties.ApplicationNonVisualDrawingProperties.Elements<CustomerDataList>();

                relID = cDataList.Elements<CustomerDataTags>().SingleOrNew().Id.ToString();

            }
            catch { }

            return relID;
        }

        // set Tag relationshipId
        internal static void setTagRelationshipId(dynamic part, string rId)
        {
            // ======================
            // presentation
            // ======================
            if (typeof(PresentationPart) == part.GetType())
            {
                CustomerDataList customerDataList1 = ((PresentationPart)part).Presentation.CustomerDataList;

                if (customerDataList1 == null)
                {
                    customerDataList1 = new CustomerDataList();
                    ((PresentationPart)part).Presentation.Append(customerDataList1);
                }

                if (customerDataList1.Elements<CustomerDataTags>().Where(t => t.Id == rId).Count() == 0)
                {
                    CustomerDataTags customerDataTags1 = new CustomerDataTags() { Id = rId };
                    customerDataList1.Append(customerDataTags1);
                }
            }

            // ======================
            // Slide
            // ======================
            if (typeof(SlidePart) == part.GetType())
            {
                CommonSlideData commonSlideData1 = ((SlidePart)part).Slide.CommonSlideData;

                if (commonSlideData1 == null)
                {
                    commonSlideData1 = new CommonSlideData();
                    ((SlidePart)part).Slide.Append(commonSlideData1);
                }

                CustomerDataList customerDataList1 = commonSlideData1.CustomerDataList;
                if (customerDataList1 == null)
                {
                    customerDataList1 = new CustomerDataList();
                    commonSlideData1.Append(customerDataList1);
                }

                if (customerDataList1.Elements<CustomerDataTags>().Where(t => t.Id == rId).Count() == 0)
                {
                    CustomerDataTags customerDataTags1 = new CustomerDataTags() { Id = rId };
                    customerDataList1.Append(customerDataTags1);
                }
            }

            // ======================
            // GraphicFrame (part)
            // ======================
            if (typeof(GraphicFrame) == part.GetType())
            {
                NonVisualGraphicFrameProperties nonVisualGraphicFrameProperties1 = ((GraphicFrame)part).NonVisualGraphicFrameProperties;
                if (nonVisualGraphicFrameProperties1 == null)
                {
                    nonVisualGraphicFrameProperties1 = new NonVisualGraphicFrameProperties();
                    ((GraphicFrame)part).Append(nonVisualGraphicFrameProperties1);
                }

                ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties1 = nonVisualGraphicFrameProperties1.ApplicationNonVisualDrawingProperties;
                if (applicationNonVisualDrawingProperties1 == null)
                {
                    applicationNonVisualDrawingProperties1 = new ApplicationNonVisualDrawingProperties();
                    nonVisualGraphicFrameProperties1.Append(nonVisualGraphicFrameProperties1);
                }

                CustomerDataList customerDataList1 = applicationNonVisualDrawingProperties1.Elements<CustomerDataList>().FirstOrDefault();
                if (customerDataList1 == null)
                {
                    customerDataList1 = new CustomerDataList();
                    applicationNonVisualDrawingProperties1.Append(customerDataList1);
                }

                if (customerDataList1.Elements<CustomerDataTags>().Where(t => t.Id == rId).Count() == 0)
                {
                    CustomerDataTags customerDataTags1 = new CustomerDataTags() { Id = rId };
                    customerDataList1.Append(customerDataTags1);
                }
            }

            // ======================
            // Shape (part)
            // ======================
            if (typeof(Shape) == part.GetType())
            {
                NonVisualShapeProperties nonVisualGraphicFrameProperties1 = ((Shape)part).NonVisualShapeProperties;
                if (nonVisualGraphicFrameProperties1 == null)
                {
                    nonVisualGraphicFrameProperties1 = new NonVisualShapeProperties();
                    ((GraphicFrame)part).Append(nonVisualGraphicFrameProperties1);
                }

                ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties1 = nonVisualGraphicFrameProperties1.ApplicationNonVisualDrawingProperties;
                if (applicationNonVisualDrawingProperties1 == null)
                {
                    applicationNonVisualDrawingProperties1 = new ApplicationNonVisualDrawingProperties();
                    nonVisualGraphicFrameProperties1.Append(nonVisualGraphicFrameProperties1);
                }

                CustomerDataList customerDataList1 = applicationNonVisualDrawingProperties1.Elements<CustomerDataList>().FirstOrDefault();
                if (customerDataList1 == null)
                {
                    customerDataList1 = new CustomerDataList();
                    applicationNonVisualDrawingProperties1.Append(customerDataList1);
                }

                if (customerDataList1.Elements<CustomerDataTags>().Where(t => t.Id == rId).Count() == 0)
                {
                    CustomerDataTags customerDataTags1 = new CustomerDataTags() { Id = rId };
                    customerDataList1.Append(customerDataTags1);
                }
            }

        }

        // get graphicFrame object (chart or table) by shapeName
        internal static GraphicFrame getGraphicFrame(SlidePart slidePart, string shapeName)
        {
            GraphicFrame gFrame = slidePart.Slide.CommonSlideData.ShapeTree.Elements<GraphicFrame>().Where(gf => gf.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Name == shapeName).SingleOrDefault();

            return gFrame;
        }

        // get Shape - Chart object
        // get ChartPart by Chart Name
        internal static ChartPart getChartPart(SlidePart slidePart, string chartName)
        {
            GraphicFrame gFrame = slidePart.Slide.CommonSlideData.ShapeTree.Descendants<GraphicFrame>().Where(g => g.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Name == chartName).FirstOrDefault();
            return (ChartPart)slidePart.GetPartById(gFrame.Graphic.GraphicData.GetFirstChild<C.ChartReference>().Id);
        }

        // get shape object by shapeName
        internal static Shape getShape(SlidePart slidePart, string shapeName)
        {
            Shape shape = slidePart.Slide.CommonSlideData.ShapeTree.Elements<Shape>().Where(gf => gf.NonVisualShapeProperties.NonVisualDrawingProperties.Name == shapeName).SingleOrDefault();
            
            return shape;
        }

        // get Table object by shapeName
        internal static D.Table getTable(SlidePart slidePart, string shapeName)
        {
            D.Table table = null;
            GraphicFrame gFrame = getGraphicFrame(slidePart, shapeName);
            if (gFrame != null)
                table = gFrame.Graphic.GraphicData.Elements<D.Table>().SingleOrDefault();

            return table;
        }

        // set tagName & tagValue pair
        //  use dynamic to accommodate 4 different types (persenationPart/slidePart/graphicFrame/Shape) of "taggable" PPT objects
        public static void setTagValue(dynamic part, string tagName, string tagValue)
        {
            // get UserDefinedTagsPart object (generic for all tags)
            UserDefinedTagsPart tagPart = getTagPart(part);

            // SlidePart for graphicframe/shape/slidePart; PresentationPart for presentationPart
            dynamic pkgPart = getPackagePart(part);

            // relationshipId
            string rId = null;

            // if tagPart does not exist, add a new one
            if (tagPart == null)
            {
                tagPart = pkgPart.AddNewPart<UserDefinedTagsPart>();
                //rId = pkgPart.CreateRelationshipToPart(tagPart);
                rId = pkgPart.GetIdOfPart(tagPart);
                pkgPart.AddPart(tagPart, rId);
            }
            else
                rId = getTagRelationshipId(part);

            // check if TagList exists
            if (tagPart.TagList == null)
            {
                // no TagList, create a new one
                TagList tagList1 = new TagList();
                tagList1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
                tagList1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                tagList1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");
                tagPart.TagList = tagList1;
            }

            Tag tag = tagPart.TagList.Elements<Tag>().Where(t => String.Compare(t.Name, tagName, true) == 0).SingleOrDefault();
            // if tag exists, overwrite it
            if (tag != null)
                tag.Val = tagValue;
            else
            {
                // no tag, create a new one
                Tag tag1 = new Tag() { Name = tagName, Val = tagValue };
                tagPart.TagList.Append(tag1);
            }

            // create Tag Relationship link to one of the presentation/slide/graphicframe/shape
            setTagRelationshipId(part, rId);

            // save presentation or slide
            if (typeof(PresentationPart) == part.GetType())
                ((PresentationPart)pkgPart).Presentation.Save();
            else
                ((SlidePart)pkgPart).Slide.Save();
        }

        // get tagValue by shapeName (including graphicFrame/shape) and tagName
        public static void setTagValue(SlidePart slidePart, string shapeName, string tagName, string tagValue)
        {
            // either GraphicFrame or Shape
            GraphicFrame gFrame = getGraphicFrame(slidePart, shapeName);
            if (gFrame != null)
                setTagValue(gFrame, tagName, tagValue);
            else
            {
                Shape shape = getShape(slidePart, shapeName);
                if (shape != null)
                    setTagValue(shape, tagName, tagValue);
            }
        }

        // remove tag - does not remove parent tags if there are empty (let the Office to do the clean up at the next Save document)
        public static void deleteTag(dynamic part, string tagName)
        {
            // get UserDefinedTagsPart object (generic for all tags)
            UserDefinedTagsPart tagPart = getTagPart(part);

            if (tagPart != null)
            {
                Tag tag = tagPart.TagList.Elements<Tag>().Where(t => String.Compare(t.Name, tagName, true) == 0).SingleOrDefault();
                if (tag != null)
                    tag.Remove();
            }
        }

        // remove tag -- override function for removing shape's tag
        public static void deleteTag(SlidePart slidePart, string shapeName, string tagName)
        {
            // either GraphicFrame or Shape
            GraphicFrame gFrame = getGraphicFrame(slidePart, shapeName);
            if (gFrame != null)
                deleteTag(gFrame, tagName);
            else
            {
                Shape shape = getShape(slidePart, shapeName);
                if (shape != null)
                    deleteTag(shape, tagName);
            }
        }

        // get PPT shape data from List<myCell>
        public static List<myCell> getShapeData(List<myCell> allCells, myCellColumnIndicator indicator, string slideID, string shapeID)
        {
            // query for a slide and shape data
            List<myCell> shapeData = (from c in allCells
                                      join c1 in allCells on c.Row equals c1.Row
                                      join c2 in allCells on c.Row equals c2.Row
                                      where c1.Column == indicator.slide && c1.Value == slideID && c2.Column == indicator.shape && c2.Value == shapeID
                                      orderby c.Column, c.Row
                                      select c).ToList<myCell>();

            return shapeData;
        }

        // set barchart data (remove existing and insert the new ones)
        //  The first column is header(stub) column (for each CategoryAxisData repeated);
        //   and the rest (2nd and thereafter) are series data (for each Values)
        internal static void setBarChartData(ChartPart chartPart, List<myCell> chartData, myCellColumnIndicator indicator, int padding = 0)
        {
            // get BarChart object
            C.ChartSpace chartSpace = chartPart.ChartSpace;
            C.BarChart barChart = chartSpace.Descendants<C.BarChart>().FirstOrDefault();
            int NumberOfSeries = barChart.Descendants<C.BarChartSeries>().Count();
            //C.BarChartSeries barChartSeries = chartSpace.Descendants<C.BarChartSeries>().Where(ser => ser.Order.Val==0).SingleOrDefault();

            // for each ShapeID, describe banners, stubs, and chart data (SerialText, CategoryAxisData, and Values)
            // use number of columns (number of distinct banners) in chart data (Values) to decide the number of Series

            // line up distinct banners in the order of Row numbers
            IEnumerable<string> banners =  chartData.Where(c => c.Column == indicator.banner)
                                                    .OrderBy(c => c.Row)
                                                    .Select(cd => cd.Value)
                                                    .Distinct();

            for (int I = 0; I < banners.Count(); I++)
            {
                // get banner for SeriesText
                string banner = banners.ElementAt(I);

                // get stubs for CategoryAxisData
                var stubs = from c in chartData
                            join c1 in chartData on c.Row equals c1.Row
                            where c1.Column == indicator.banner && c1.Value == banner && c.Column == indicator.stub
                            orderby c.Row
                            select c;

                // get measures for Values
                var measures = from c in chartData
                            join c1 in chartData on c.Row equals c1.Row
                            where c1.Column == indicator.banner && c1.Value == banner && c.Column == indicator.measure
                            orderby c.Row
                            select c;

                // check if BarChartSeries exists
                if (I < NumberOfSeries)
                {
                    // acquire the series
                    C.BarChartSeries series = barChart.Elements<C.BarChartSeries>().ElementAt(I);
                    string formatCodeText = getFormatCodeText(series);

                    // remove SeriesText, BarChartSeries CategoryAxisData and Values
                    series.RemoveAllChildren<C.SeriesText>();
                    series.RemoveAllChildren<C.CategoryAxisData>();
                    series.RemoveAllChildren<C.Values>();

                    // add SeriesText, BarChartSeries CategoryAxisData and Values
                    series.Append(addSeriesText(I, banner));
                    series.Append(addCategoryAxisData(I, (List<myCell>)stubs, padding));
                    series.Append(addValues(I, (List<myCell>)measures, formatCodeText, padding));
                }
                else    // extra Series will be added
                {
                    C.BarChartSeries prevSeries = null;
                    if (I > 0)
                        prevSeries = barChart.Elements<C.BarChartSeries>().ElementAt(I - 1);
                    barChart.Append(addSeries(I, banner, (List<myCell>)stubs, (List<myCell>)measures, prevSeries, padding));
                }
                
            }
               
            // remove excessive Series, if there are any left
            if (NumberOfSeries > banners.Count())
                for (int J = banners.Count(); J < NumberOfSeries; J++)
                    barChart.Elements<C.BarChartSeries>().ElementAt(J).Remove();

        }

        // set barchart data - override function
        public static void setBarChartData(SlidePart slidePart, string chartName, List<myCell> chartData, myCellColumnIndicator indicator, int padding = 0)
        {
            setBarChartData(getChartPart(slidePart, chartName), chartData, indicator, padding);
        }

        // create a new series
        // assume the index and order are the same
        private static C.BarChartSeries addSeries(int order, string banner, List<myCell> categoryAxisData, List<myCell> values, C.BarChartSeries prevSeries = null, int padding = 0)
        {
            C.BarChartSeries barChartSeries1 = new C.BarChartSeries();

            // add index and order
            barChartSeries1.Append(new C.Index() { Val = (UInt32Value)(UInt32)order });
            barChartSeries1.Append(new C.Order() { Val = (UInt32Value)(UInt32)order });

            // add SeriesText
            barChartSeries1.Append(addSeriesText(order, banner));

            // add ChartShapeProperties
            //  follow GA's chart standard colors with SolidFill; solid black for Outline
            barChartSeries1.Append(addChartShapeProperties(order));

            // add DataLabels
            //  check if previous series shows dataLabel; if yes, copy the properties; if not, no DataLabels
            if (prevSeries != null && prevSeries.Elements<C.DataLabels>().SingleOrDefault() != null)
                barChartSeries1.Append(prevSeries.Elements<C.DataLabels>().SingleOrDefault());

            // add CategoryAxisData
            barChartSeries1.Append(addCategoryAxisData(order, categoryAxisData, padding));

            // add Values
            //  need to copy the formatCode from the previous series, if there is one
            string formatCodeText = getFormatCodeText(prevSeries);
            barChartSeries1.Append(addValues(order, values, formatCodeText, padding));

            // return Series object
            return barChartSeries1;
        }

        // get formatCode text from a Series - default to "General"
        private static string getFormatCodeText(C.BarChartSeries series)
        {
            string formatCodeText = "General";
            try
            {
                // does not matter busting out if no formatCode for prevSeries
                if (series != null)
                    formatCodeText = series.Elements<C.Values>().SingleOrDefault().NumberReference.NumberingCache.FormatCode.Text;
            }
            finally { }

            return formatCodeText;
        }

        // add ChartShapeProperties for Chart Series- fill and color of a Series
        private static C.ChartShapeProperties addChartShapeProperties(int order)
        {
            C.ChartShapeProperties chartShapeProperties1 = new C.ChartShapeProperties();

            D.SolidFill solidFill1 = new D.SolidFill();
            D.RgbColorModelHex rgbColorModelHex1 = new D.RgbColorModelHex() { Val = GA_CHART_COLOR[order % GA_CHART_COLOR.Count()] };

            solidFill1.Append(rgbColorModelHex1);

            D.Outline outline1 = new D.Outline() { Width = 12692 };

            D.SolidFill solidFill2 = new D.SolidFill();
            D.RgbColorModelHex rgbColorModelHex2 = new D.RgbColorModelHex() { Val = "000000" };

            solidFill2.Append(rgbColorModelHex2);
            D.PresetDash presetDash1 = new D.PresetDash() { Val = D.PresetLineDashValues.Solid };

            outline1.Append(solidFill2);
            outline1.Append(presetDash1);

            chartShapeProperties1.Append(solidFill1);
            chartShapeProperties1.Append(outline1);

            return chartShapeProperties1;
        }

        // add SeriesText (banner/column header) to a BarChartSeries
        private static C.SeriesText addSeriesText(int seriesOrder, string seriesTextValue)
        {
            // seriesOrder starts 0, which should be B (2nd column in the spreadsheet)
            string formulaValue = "Sheet1!$" + XLS.getColumnName(seriesOrder + 2) + "$1";
            return addSeriesText(formulaValue, seriesTextValue);
        }

        // add SeriesText (banner/column header) to a BarChartSeries
        //  for {Blank?}, it will be treated as no text
        private static C.SeriesText addSeriesText(string formulaValue, string seriesTextValue)
        {
            C.SeriesText seriesText = new C.SeriesText();

            C.StringReference stringReference = new C.StringReference();
            Formula formula = new Formula();
            formula.Text = formulaValue;

            C.StringCache stringCache = new C.StringCache();
            C.PointCount pointCount = new C.PointCount() { Val = (UInt32Value)1U };

            C.StringPoint stringPoint = new C.StringPoint() { Index = (UInt32Value)0U };
            C.NumericValue numericValue = new C.NumericValue();
            numericValue.Text = (seriesTextValue.ToUpper().StartsWith("{BLANK") && seriesTextValue.EndsWith("}")) ? null : seriesTextValue;

            stringPoint.Append(numericValue);

            stringCache.Append(pointCount);
            stringCache.Append(stringPoint);

            stringReference.Append(formula);
            stringReference.Append(stringCache);

            seriesText.Append(stringReference);

            return seriesText;
        }

        // replace SerialText (banner/column header) Text value (inherit the same formula)
        private static void replaceSerialText(C.BarChartSeries series, string seriesTextValue)
        {
            // preserved formula first
            string formulaValue = series.SeriesText.StringReference.Formula.Text;

            // remove the old SeriesText
            series.RemoveAllChildren<C.SeriesText>();

            // add the new SeriesText
            series.Append(addSeriesText(formulaValue, seriesTextValue));        
        }

        // add CategoryAxisData (stub/row header) to a BarChartSeries
        //  padding will be based on the number of firms repeated (the chart will have a tag padding=Y)
        private static C.CategoryAxisData addCategoryAxisData(int seriesOrder, List<myCell> blockData, int padding = 0)
        {
            // calcuate the formulaHeader
            string formulaHeader = XLS.getColumnName(seriesOrder + 2);

            // the row starts from 2
            int pointCount = (padding > 0) ? (blockData.Count() + blockData.Count() / padding + 1) : (blockData.Count());
            string endOfRow = (pointCount + 2).ToString();
            UInt32 numberOfRow = (UInt32)pointCount;

            C.CategoryAxisData categoryAxisData1 = new C.CategoryAxisData();

            C.StringReference stringReference1 = new C.StringReference();

            Formula formula1 = new Formula();
            // formula text likes this "Sheet1!$B$2:$B$14"
            formula1.Text = "Sheet1!$" + formulaHeader + "$2:$" + formulaHeader + "$" + endOfRow;
            stringReference1.Append(formula1);

            C.StringCache stringCache1 = new C.StringCache();
            C.PointCount pointCount1 = new C.PointCount() { Val = (UInt32Value)numberOfRow };
            stringCache1.Append(pointCount1);

            UInt32Value index = 0U;
            foreach(myCell cell in blockData)
            {
                // the places (rows) that need to add a blank row (skip the index)
                if (index % (padding + 1) == padding)
                    index++;

                C.StringPoint stringPoint1 = new C.StringPoint() { Index = index };
                C.NumericValue numericValue1 = new C.NumericValue();
                numericValue1.Text = cell.Value;
                stringPoint1.Append(numericValue1);
                stringCache1.Append(stringPoint1);

                index++;
            }

            stringReference1.Append(stringCache1);
            categoryAxisData1.Append(stringReference1);

            // append to BarChartSeries
            return categoryAxisData1;
        }

        // add Values (chart data) to a BarChartSeries
        //  padding will be based on the number of firms repeated (the chart will have a tag padding=Y)
        private static Values addValues(int seriesOrder, List<myCell> blockData, string formatCodeText = "General", int padding = 0)
        {
            // calcuate the formulaHeader
            string formulaHeader = XLS.getColumnName(seriesOrder + 2);

            // the row starts from 2
            int pointCount = (padding > 0) ? (blockData.Count() + blockData.Count() / padding + 1) : (blockData.Count());
            string endOfRow = (pointCount + 2).ToString();
            UInt32 numberOfRow = (UInt32)pointCount;

            Values values1 = new Values();

            C.NumberReference numberReference1 = new C.NumberReference();

            Formula formula1 = new Formula();
            formula1.Text = "Sheet1!$" + formulaHeader + "$2:$" + formulaHeader + "$" + endOfRow; ;
            numberReference1.Append(formula1);

            C.NumberingCache numberingCache1 = new C.NumberingCache();
            C.FormatCode formatCode1 = new C.FormatCode();
            formatCode1.Text = formatCodeText;
            numberingCache1.Append(formatCode1);

            C.PointCount pointCount1 = new C.PointCount() { Val = (UInt32Value)numberOfRow };
            numberingCache1.Append(pointCount1);

            UInt32Value index = 0U;
            foreach (myCell cell in blockData)
            {
                // the places (rows) that need to add a blank row (skip the index)
                if (index % (padding + 1) == padding)
                    index++;

                C.NumericPoint numericPoint1 = new C.NumericPoint() { Index = index };
                C.NumericValue numericValue1 = new C.NumericValue();
                numericValue1.Text = cell.Value;
                numericPoint1.Append(numericValue1);
                numberingCache1.Append(numericPoint1);

                index++;
            }

            numberReference1.Append(numberingCache1);
            values1.Append(numberReference1);

            // append to BarChartSeries
            return values1;
        }

        // set data to Table 
        //  should preserve the total size of the table (row height and column width) - need to recalculate the average height and width if the numbers of row/column are changed
        //  should inherit the cell format (font/color/line/fill etc.)
        //  if extra cells, use the last cell's format
        //  if no pre-existing table (graphicFrame), ignore this shape
        public static void setTableData(D.Table table, List<myCell> tableData, myCellColumnIndicator indicator, int padding = 0)
        {
            // no Table; no data
            if (table == null) return;

            // preserve TableProperties if exists, or create a new one
            D.TableProperties tableProperties1 = table.TableProperties;
            if (tableProperties1 == null)
                table.Append(new D.TableProperties() { FirstRow = false, BandRow = true });

            // check if TableGrid exists; if not, add it
            D.TableGrid tableGrid = table.TableGrid;
            if (tableGrid == null)
                table.Append(new D.TableGrid());

            // preserve/clone the first cell as a template
            D.TableCell templateTableCell = (D.TableCell)table.Elements<D.TableRow>().FirstOrNew().Elements<D.TableCell>().FirstOrNew().CloneNode(true);

            // remove the text value (assign an empty value) from the templateTableCell (could have multiple paragraphs and multiple runs)
            foreach (D.Paragraph paragraph in (D.Paragraph)templateTableCell.TextBody.Descendants<D.Paragraph>())
                foreach (D.Run run in paragraph.Elements<D.Run>())
                    run.Text = new D.Text("");

            // calculate the total table width and height
            long totalColumnWidth = tableGrid.Elements<D.GridColumn>().Sum(gc => gc.Width);
            long totalRowHeight = table.Elements<D.TableRow>().Sum(tr => tr.Height);

            // get number of rows and columns from current GridColumn and dataColumn
            int numberOfGridColumns = tableGrid.Elements<D.GridColumn>().Count();
            int numberOfDataColumns = tableData.Select(t => t.Column).Distinct().Count();
            int numberOfTableRows = table.Elements<D.TableRow>().Count();
            int numberOfDataRows = tableData.Select(t => t.Row).Distinct().Count();
            // considering padding
            numberOfDataRows = (padding > 0) ? (numberOfDataRows + numberOfDataRows / padding + 1) : (numberOfDataRows);

            // get the average of ColumnWidth and RowHeight
            long averageColumnWidth = (numberOfDataColumns==0) ? 0 : totalColumnWidth / numberOfDataColumns;
            long averageRowHeight = (numberOfDataRows == 0) ? 0 : totalRowHeight / numberOfDataRows;

            // if the number of GridColumn/TableRow is equal to the DataColumn/DataRow, preserve the current heights and widths
            //  this allows users to customize (and keep) the heights and widths on each cell from the subsequent data refresh
            // different number of columns, the width will be rearranged
            if (numberOfDataColumns != numberOfGridColumns)
            {
                // remove all current GridColumns
                tableGrid.RemoveAllChildren<D.GridColumn>();

                // re-add GridColumn with new averageColumnWidth
                for (int I = 0; I < numberOfDataColumns; I++)
                        tableGrid.Append(new D.GridColumn() { Width = averageColumnWidth });
            }

            // different number of rows, the height will be rearranged and cell format(layout) will be copied from the first cell
            if (numberOfDataRows != numberOfTableRows)
            {
                // remove all current TableRows
                table.RemoveAllChildren<D.TableRow>();

                // re-add TableRows (with new averageRowHeight) and TableCells (with templateTableCell)
                for (int I = 0; I < numberOfDataRows; I++)
                {
                    D.TableRow tableRow = new D.TableRow() { Height = averageRowHeight };
                    for (int J = 0; J < numberOfDataColumns; J++)
                        tableRow.Append(templateTableCell.CloneNode(true));
                    table.Append(tableRow);
                }
            }
            else
            {
                // clean up all text values in TableCell (each cell could have multiple paragraphs and runs)
                var runs = from tRow in table.Elements<D.TableRow>()
                           from tCell in tRow.Elements<D.TableCell>()
                           from para in tCell.TextBody.Elements<D.Paragraph>()
                           from run in para.Elements<D.Run>()
                           select run;
                foreach (D.Run run in runs)
                    run.Text = new D.Text("");
            }

            // populate the cell value
            int indexRow = -1, indexColumn = -1, prevRow = -1;
            string prevColumn = "";
            foreach(myCell cell in tableData.OrderBy(td => td.Row).OrderBy(td => td.Column))
            {
                // advance row and column index
                if (prevRow != cell.Row)
                {
                    indexRow++;
                    indexColumn = 0;
                }
                else 
                    indexColumn++;

                // the places (rows) that need to add a blank row (skip the index)
                if (indexRow % (padding + 1) == padding)
                    indexRow++;

                // assing text value base on the row index of TableRow and column index of TableCell
                //  also, only the first (index=0) Paragraph and Run will be populated
                D.TableCell tCell = table.Elements<D.TableRow>().ElementAt(indexRow).Elements<D.TableCell>().ElementAt(indexColumn);
                tCell.TextBody.Elements<D.Paragraph>().ElementAt(0).Elements<D.Run>().ElementAt(0).Text = new D.Text(cell.Value);

                // prepare for the next iteration
                prevRow = cell.Row;
                prevColumn = cell.Column;
            }

        }

        // set data to Table -- override function
        public static void setTableData(SlidePart slidePart, string shapeName, List<myCell> tableData, myCellColumnIndicator indicator, int padding = 0)
        {
            // add data (GridColum and TableRow) to the Table
            setTableData(getTable(slidePart, shapeName), tableData, indicator, padding);
        }

        // clone SlidePart with Images and Charts
        //  if insertPosition=-1, insert at the end of the Presentation
        public static SlidePart cloneSlide(PresentationPart presentationPart, SlidePart slidePartTemplate, int insertPosition = -1)
        {
            int i = presentationPart.SlideParts.Count();
            // Create a new slide part in the presentation.
            SlidePart newSlidePart = presentationPart.AddNewPart<SlidePart>("newSlide" + i);
            //i++;

            //Add the source slide content into the new slide.
            newSlidePart.FeedData(slidePartTemplate.GetStream(FileMode.Open));
            // Make sure the new slide references the proper slide layout.
            newSlidePart.AddPart(slidePartTemplate.SlideLayoutPart, slidePartTemplate.GetIdOfPart(slidePartTemplate.SlideLayoutPart));

            // copy the image parts
            foreach (ImagePart ipart in slidePartTemplate.ImageParts)
            {
                ImagePart newipart = newSlidePart.AddImagePart(ipart.ContentType, slidePartTemplate.GetIdOfPart(ipart));
                newipart.FeedData(ipart.GetStream());
            }

            // copy the chart parts
            foreach (ChartPart cpart in slidePartTemplate.ChartParts)
            {
                ChartPart newcpart = newSlidePart.AddNewPart<ChartPart>(slidePartTemplate.GetIdOfPart(cpart));
                newcpart.FeedData(cpart.GetStream());
                // copy the embedded excel file
                EmbeddedPackagePart epart = newcpart.AddEmbeddedPackagePart(cpart.EmbeddedPackagePart.ContentType);
                epart.FeedData(cpart.EmbeddedPackagePart.GetStream());
                // link the excel to the chart
                newcpart.ChartSpace.Elements<C.ExternalData>().First().Id = newcpart.GetIdOfPart(epart);
                newcpart.ChartSpace.Save();
            }

            // Get the list of slide ids.
            SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;

            // get the next SlideId (one more from the max SlideId)
            uint nextSlideId = slideIdList.Elements<SlideId>().Max(s => s.Id) ?? 0U;
            nextSlideId++;

            SlideId newSlideId = null;
            // Deternmine where to add the next slide (find max number of slides).
            //  Make sure the id and relid are set appropriately.
            if (insertPosition == -1 || insertPosition >= slideIdList.Count())
                newSlideId = slideIdList.InsertAfter(new SlideId() { Id = nextSlideId, RelationshipId = presentationPart.GetIdOfPart(newSlidePart) }, slideIdList.ElementAt(slideIdList.Count() - 1));
            else
                newSlideId = slideIdList.InsertAt(new SlideId() { Id = nextSlideId, RelationshipId = presentationPart.GetIdOfPart(newSlidePart) }, insertPosition);

            // save Slide & Presentation
            newSlidePart.Slide.Save();
            //presentationPart.Presentation.Save();

            // return new SlidePart
            return newSlidePart;
        }
        
        // reposition slide in a PowerPoint
        //  if insertPosition=-1, move to the end of the Presentation
        public static void moveSlide(PresentationPart presentationPart, SlidePart slidePart, int newPosition = -1)
        {
        }

        // delete slide - by slidePart
        public static void deleteSlide(PresentationPart presentationPart, SlidePart slidePart)
        {
            // delete SlidePart
            if (slidePart != null)
                presentationPart.DeletePart(slidePart);

            // find slideId from slidePart
            SlideId slideId = getSlideId(presentationPart, slidePart);
            if (slideId != null)
                slideId.Remove();
        }

        // delete slide - by slideIndex
        public static void deleteSlide(PresentationPart presentationPart, int slideIndex)
        {
            SlidePart slidePart = getSlidePart(presentationPart, slideIndex);
            deleteSlide(presentationPart, slidePart);
        }


    }

    
    static class XLSX
    {
        // transform column number to column alphabet
        public static string getColumnName(int column)
        {
            // This algorithm was found here:
            // http://stackoverflow.com/questions/181596/how-to-convert-a-column-number-eg-127-into-an-excel-column-eg-aa

            // Given a column number, retrieve the corresponding
            // column string name:
            int value = 0;
            int remainder = 0;
            string result = string.Empty;
            value = column;

            while (value > 0)
            {
                remainder = (value - 1) % 26;
                result = (char)(65 + remainder) + result;
                value = (int)(Math.Floor((double)((value - remainder) / 26)));
            }
            return result;
        }

        // get workSheet
        public static WorksheetPart getWorkSheetPart(WorkbookPart workBookPart, string sheetName)
        {
            WorksheetPart worksheetPart = null;

            // Find the sheet with the supplied name, and then use that Sheet object
            // to retrieve a reference to the appropriate worksheet.
            Sheet theSheet = workBookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();

            // Retrieve a reference to the worksheet part, and then use its Worksheet property to get 
            // a reference to the cell whose address matches the address you've supplied:
            if (theSheet != null)
                worksheetPart = (WorksheetPart)(workBookPart.GetPartById(theSheet.Id));

            return worksheetPart;
        }

        // get Cell Value via ColumnName and RowNumber
        public static string getCellValue(WorkbookPart workBookPart, WorksheetPart worksheetPart, string columnName, int rowNumber)
        {
            string value = null;
            
            // get Cell via row/column
            Cell theCell = worksheetPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == columnName + rowNumber.ToString()).FirstOrDefault();

            // If the cell doesn't exist, return an empty string.
            if (theCell != null)
            {
                value = theCell.InnerText;

                // If the cell represents an integer number, you're done. 
                // For dates, this code returns the serialized value that 
                // represents the date. The code handles strings and booleans
                // individually. For shared strings, the code looks up the corresponding
                // value in the shared string table. For booleans, the code converts 
                // the value into the words TRUE or FALSE.
                if (theCell.DataType != null)
                {
                    switch (theCell.DataType.Value)
                    {
                        case CellValues.SharedString:
                            // For shared strings, look up the value in the shared strings table.
                            var stringTable = workBookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                            // If the shared string table is missing, something's wrong.
                            // Just return the index that you found in the cell.
                            // Otherwise, look up the correct text in the table.
                            if (stringTable != null)
                            {
                                value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                            }
                            break;

                        case CellValues.Boolean:
                            switch (value)
                            {
                                case "0":
                                    value = "FALSE";
                                    break;
                                default:
                                    value = "TRUE";
                                    break;
                            }
                            break;
                    }
                }
            }

            return value;
        }

        // Given a workbook part, a sheet name, and row and column names, retrieve the value of the cell.
        // Call the function like this:
        // string value = xGetCellValue("Sample.xlsx", "Sheet1", "B", 3);
        public static string getCellValue(WorkbookPart workBookPart, string sheetName, string columnName, int rowNumber)
        {
            // get worksheetPart
            WorksheetPart worksheetPart = getWorkSheetPart(workBookPart, sheetName);
            
            // get CellValue from override function
            return getCellValue(workBookPart, worksheetPart, columnName, rowNumber);
        }

        // read all cells in a excel worksheet
        public static List<myCell> loadWorksheetData(string fileName, int worksheetNumber)
        {
            List<myCell> parsedCells = new List<myCell>();
            //string fileName = @"C:\Documents and Settings\awang\Desktop\Office XML\alex_test 64K rows.xlsx";
            using (Package xlsxPackage = Package.Open(fileName, FileMode.Open, FileAccess.Read))
            {
                try
                {
                    PackagePartCollection allParts = xlsxPackage.GetParts();

                    // prepare shared string dictionary
                    PackagePart sharedStringsPart = (from part in allParts
                                                     where part.ContentType.Equals("application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml")
                                                     select part).SingleOrDefault();

                    Dictionary<int, string> sharedStrings = new Dictionary<int, string>();
                    if (sharedStringsPart != null)
                    {
                        XElement sharedStringsElement = XElement.Load(XmlReader.Create(sharedStringsPart.GetStream()));
                        parseSharedStrings(sharedStringsElement, sharedStrings);
                    }

                    // get worksheet and prepare to get datat
                    XElement worksheetElement = getWorksheetXElement(allParts, worksheetNumber);

                    // get all cells (just cells)
                    IEnumerable<XElement> allCells = from c in worksheetElement.Descendants(OpenXMLNamespaces.excelNamespace + "c")
                                                     select c;

                    foreach (XElement cell in allCells)
                    {
                        string cellPosition = cell.Attribute("r").Value;

                        // check if shared string
                        bool isShared = (cell.Attribute("t") == null) ? false : cell.Attribute("t").Value == "s";
                        int index = indexOfNumber(cellPosition);
                        string column = cellPosition.Substring(0, index);
                        int row = Convert.ToInt32(cellPosition.Substring(index, cellPosition.Length - index));

                        // get value 
                        XElement xel = cell.Descendants(OpenXMLNamespaces.excelNamespace + "v").SingleOrDefault();
                        if (xel != null)
                        {
                            if (isShared)
                                parsedCells.Add(new myCell(column, row, sharedStrings[Convert.ToInt32(xel.Value)]));
                            else
                                parsedCells.Add(new myCell(column, row, xel.Value.ToString()));
                        }
                    }
                }
                finally
                {
                    xlsxPackage.Close();
                }
            }

            // for testing
            //Console.WriteLine(parsedCells.Count().ToString());
            //foreach (Cell cell in parsedCells)
            //{
            //    Console.WriteLine(cell);
            //}

            return parsedCells;
        }

        public static List<myCell> loadWorksheetData(string fileName, string worksheetName)
        {
            int worksheetNumber = getWorksheetIndex(fileName, worksheetName);
            return loadWorksheetData(fileName, worksheetNumber);
        }

        // return first numeric character in a string; if not found, return 0
        private static int indexOfNumber(string value)
        {
            for (int counter = 0; counter < value.Length; counter++)
            {
                if (char.IsNumber(value[counter]))
                {
                    return counter;
                }
            }

            return 0;
        }

        // populate shared string into a Dictionary object
        private static void parseSharedStrings(XElement SharedStringsElement, Dictionary<int, string> sharedStrings)
        {
            IEnumerable<XElement> sharedStringsElements = from s in SharedStringsElement.Descendants(OpenXMLNamespaces.excelNamespace + "t")
                                                          select s;

            int Counter = 0;
            foreach (XElement sharedString in sharedStringsElements)
            {
                sharedStrings.Add(Counter, sharedString.Value);
                Counter++;
            }
        }

        // get excel worksheet (as XEelement) from a PackagePartCollection
        private static XElement getWorksheetXElement(PackagePartCollection allParts, int worksheetID)
        {
            PackagePart worksheetPart = (from part in allParts
                                         where part.Uri.OriginalString.Equals(String.Format("/xl/worksheets/sheet{0}.xml", worksheetID))
                                         select part).Single();

            return XElement.Load(XmlReader.Create(worksheetPart.GetStream()));
        }

        // get worksheet index - by fileName
        public static int getWorksheetIndex(string fileName, string worksheetName)
        {
            int rtnIndex = -1;

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
                 rtnIndex = getWorksheetIndex(document.WorkbookPart, worksheetName);

            return rtnIndex;
        }

        // get worksheet index - by workbookPart
        public static int getWorksheetIndex(WorkbookPart workbookPart, string worksheetName)
        {
            int rtnIndex = -1, index = -1;
            
                foreach (Sheet sheet in workbookPart.Workbook.Sheets)
                {
                    index++;
                    if (sheet.Name == worksheetName)
                    {
                        rtnIndex = index;
                        break;
                    }
                }
 
            return rtnIndex;
        }
    }


    static class XLS
    {
        // read all cells in a excel worksheet
        public static List<myCell> loadWorksheetData(string fileName, string worksheetName)
        {
            string connectionString = @"Provider=Microsoft.Jet.
                                        OLEDB.4.0;Data Source=Book1.xls;Extended
                                        Properties=""Excel 8.0;HDR=YES;""";

            DbProviderFactory factory = DbProviderFactories.GetFactory("System.Data.OleDb");

            using (DbConnection connection = factory.CreateConnection())
            {
                connection.ConnectionString = connectionString;

                using (DbCommand command = connection.CreateCommand())
                {
                    // Cities$ comes from the name of the worksheet
                    command.CommandText = "SELECT * FROM [" + worksheetName + "$]";

                    connection.Open();

                    using (DbDataReader dr = command.ExecuteReader())
                    {
                        while (dr.Read())
                        {
                            Debug.WriteLine(dr["ID"].ToString());
                        }
                    }
                }
            }

        }


    }

    static class Util
    {

        // set barchart data - override function
        public static void PPT_setBarChartData(SlidePart slidePart, string chartName, List<myCell> chartData, myCellColumnIndicator indicator, int padding = 0)
        {
            PPT.setBarChartData(PPT.getChartPart(slidePart, chartName), chartData, indicator, padding);
        }

        /// <summary>
        /// Perform a deep Copy of the object.
        /// </summary>
        /// <typeparam name="T">The type of object being copied.</typeparam>
        /// <param name="source">The object instance to copy.</param>
        /// <returns>The copied object.</returns>
        public static T Clone<T>(T source)
        {
            if (!typeof(T).IsSerializable)
            {
                throw new ArgumentException("The type must be serializable.", "source");
            }

            // Don't serialize a null object, simply return the default for that object
            if (Object.ReferenceEquals(source, null))
            {
                return default(T);
            }

            IFormatter formatter = new BinaryFormatter();
            Stream stream = new MemoryStream();
            using (stream)
            {
                formatter.Serialize(stream, source);
                stream.Seek(0, SeekOrigin.Begin);
                return (T)formatter.Deserialize(stream);
            }
        }

        /// <summary>
        /// Extension method: return a new initiated object if query return empty result
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="query"></param>
        /// <returns></returns>
        public static T SingleOrNew<T>(this IEnumerable<T> query) where T : new()
        {
            try
            {
                return query.Single();
            }
            catch (InvalidOperationException)
            {
                return new T();
            }
        }

        /// <summary>
        /// Extension method: return a new initiated object if query return empty result
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="query"></param>
        /// <returns></returns>
        public static T FirstOrNew<T>(this IEnumerable<T> query) where T : new()
        {
            try
            {
                return query.First();
            }
            catch (InvalidOperationException)
            {
                return new T();
            }
        }

    }


    static class Deprecated
    {
        // addTableRows to PowerPoint Table
        public static D.TableRow[] addPPTTableRows(WorkbookPart wbPart, string sheetName, int colNumber, int rowNumber, int startCol = 0, int startRow = 0)
        {
            D.TableRow[] tableRows;
            tableRows = new D.TableRow[rowNumber-startRow+1];

            for (int I = startRow; I <= rowNumber; I++)
            {
                tableRows[I-startRow] = new D.TableRow() { Height = 170840L };

                D.TableCell[] tableCells;
                tableCells = new D.TableCell[colNumber-startCol+1];

                for (int J = startCol; J <= colNumber; J++)
                {
                    tableCells[J-startCol] = new D.TableCell();

                    D.TextBody textBody1 = new D.TextBody();
                    D.BodyProperties bodyProperties1 = new D.BodyProperties();
                    D.ListStyle listStyle1 = new D.ListStyle();

                    D.Paragraph paragraph1 = new D.Paragraph();

                    D.Run run1 = new D.Run();
                    D.RunProperties runProperties1 = new D.RunProperties() { Language = "en-US", Dirty = false, SmartTagClean = false, FontSize = 1000 };
                    D.LatinFont latinFont1 = new D.LatinFont() { Typeface = "Verdana", PitchFamily = 34, CharacterSet = 0 };
                    runProperties1.Append(latinFont1);

                    D.Text text1 = new D.Text();

                    text1.Text = getXLSCellValueRowCol(wbPart, sheetName, J+1, I+1);    // here, row/col are 1-base

                    run1.Append(runProperties1);
                    run1.Append(text1);
                    D.EndParagraphRunProperties endParagraphRunProperties1 = new D.EndParagraphRunProperties() { Language = "en-US", Dirty = false };

                    paragraph1.Append(run1);
                    paragraph1.Append(endParagraphRunProperties1);

                    textBody1.Append(bodyProperties1);
                    textBody1.Append(listStyle1);
                    textBody1.Append(paragraph1);
                    D.TableCellProperties tableCellProperties1 = new D.TableCellProperties();

                    tableCells[J-startCol].Append(textBody1);
                    tableCells[J-startCol].Append(tableCellProperties1);

                    tableRows[I-startRow].Append(tableCells[J-startCol]);
                }
            }
            return tableRows;
        }

        // Retrieve a list of slide titles, including hidden slides.
        // Some slide titles might be empty strings.
        public static List<string> getSlideTitles(string fileName)
        {
            return getSlideTitles(fileName, true);
        }

        // Retrieve a list of slide titles, optionally excluding hidden slides.
        public static List<string> getSlideTitles(string fileName, bool includeHidden)
        {
            List<string> titlesList = null;

            using (PresentationDocument doc = PresentationDocument.Open(fileName, false))
            {
                // Get a PresentationPart object from the PresentationDocument object.
                PresentationPart presentationPart = doc.PresentationPart;
                if (presentationPart != null && presentationPart.Presentation != null)
                {

                    // Get a Presentation object from the PresentationPart object.
                    Presentation presentation = presentationPart.Presentation;
                    if (presentation.SlideIdList != null)
                    {

                        titlesList = new List<string>();

                        // Get the title of each slide in the slide order.
                        // This requires investigating the actual slide IDs, rather 
                        // than just retrieving the slide parts.
                        int slideNumber = 1;
                        foreach (var slideId in presentation.SlideIdList.Elements<SlideId>())
                        {
                            SlidePart slidePart = (SlidePart)(presentationPart.GetPartById(slideId.RelationshipId.ToString()));

                            // Get the slide title.
                            string title = string.Empty;

                            if (slidePart == null)
                            {
                                throw new ArgumentNullException("presentationDocument");
                            }

                            // Declare a paragraph separator.
                            string paragraphSeparator = null;
                            if (slidePart.Slide != null)
                            {
                                Slide theSlide = slidePart.Slide;
                                bool slideVisible = false;

                                // If the Show doesn't exist, or it does exist and has a value of True,
                                // set slideVisible to True.
                                if (theSlide.Show == null || theSlide.Show.HasValue && theSlide.Show.Value)
                                {
                                    slideVisible = true;
                                }

                                // If includeHidden is false, and the slide isn't visible, skip to the next slide.
                                if (!(slideVisible || includeHidden))
                                {
                                    continue;
                                }

                                // Find all the title shapes.
                                var shapes = from shape in slidePart.Slide.Descendants<Shape>()
                                             where (isPPTTitleShape(shape))
                                             select shape;

                                StringBuilder paragraphText = new StringBuilder();

                                foreach (var shape in shapes)
                                {

                                    // Get the text in each paragraph in this shape.
                                    foreach (var paragraph in shape.TextBody.Descendants<D.Paragraph>())
                                    {

                                        // Add a line break.
                                        paragraphText.Append(paragraphSeparator);

                                        foreach (var text in paragraph.Descendants<D.Text>())
                                        {
                                            paragraphText.Append(text.Text);
                                        }
                                        // Now that you're past the first paragraph,
                                        // change the separator to a CR/LF.
                                        paragraphSeparator = Environment.NewLine;
                                    }
                                }
                                title = paragraphText.ToString();
                            }

                            // A slide's title might be empty, so you might
                            // end up with multiple empty strings in the list.
                            titlesList.Add(slideNumber.ToString() + ". " + title);
                            slideNumber++;
                        }
                    }
                }
            }
            return titlesList;
        }

        private static bool isPPTTitleShape(Shape shape)
        {
            bool isTitle = false;
            var placeholderShape = shape.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties.GetFirstChild<PlaceholderShape>();

            if (((placeholderShape) != null) && (((placeholderShape.Type) != null) && placeholderShape.Type.HasValue))
            {
                // Any title shape
                if (placeholderShape.Type.Value == PlaceholderValues.Title)
                {
                    isTitle = true;

                }
                // A centered title.
                else if (placeholderShape.Type.Value == PlaceholderValues.CenteredTitle)
                {
                    isTitle = true;
                }
            }
            return isTitle;
        }

        // get all slide tag values
        public static List<string> getPPTSlideTagValues(string fileName, string tagName)
        {
            List<string> tagList = new List<string>();

            using (PresentationDocument doc = PresentationDocument.Open(fileName, false))
            {
                // Hm... All 3 of these will produce different slide order from the last one
                //tagList = (from sldPart in doc.PresentationPart.SlideParts
                //           orderby sldPart
                //           select findPPTTagValue(sldPart, tagName) ?? "").ToList();

                //foreach (SlidePart sldPart in doc.PresentationPart.SlideParts)
                //{
                //    tagList.Add(findPPTTagValue(sldPart, tagName) ?? "");
                //}

                //for (int I = 0; I < doc.PresentationPart.SlideParts.Count(); I++)
                //{
                //    tagList.Add(findPPTTagValue(doc.PresentationPart.SlideParts.ElementAt(I), tagName) ?? "");
                //}

                foreach (var slideId in doc.PresentationPart.Presentation.SlideIdList.Elements<SlideId>())
                {
                    SlidePart slidePart = (SlidePart)(doc.PresentationPart.GetPartById(slideId.RelationshipId.ToString()));
                    tagList.Add(PPT.getTagValue(slidePart, tagName) ?? "");
                }
            }

            return tagList;
        }

        // find PPT SlideIndex that has a tag name the same as the selected XLS SheetName
        //  if not found, return -1
        public static int getPPTSlideIndex(string fileName, string tagName, string sheetName)
        {
            int index = -1, outIndex = -1;

            using (var doc = PresentationDocument.Open(fileName, false))
            {
                //foreach (SlidePart sldPart in doc.PresentationPart.SlideParts)
                //{
                //    index++;
                //    if (Util.findPPTTagValue(sldPart, tagName) == sheetName)
                //    {
                //        outIndex = index;
                //        break;
                //    }
                //}
                foreach (var slideId in doc.PresentationPart.Presentation.SlideIdList.Elements<SlideId>())
                {
                    SlidePart sldPart = (SlidePart)(doc.PresentationPart.GetPartById(slideId.RelationshipId.ToString()));
                    index++;
                    if (PPT.getTagValue(sldPart, tagName) == sheetName)
                    {
                        outIndex = index;
                        break;
                    }
                }
            }

            return outIndex;
        }

        // clone slide 
        public static SlidePart clonePPTSlidePart(PresentationPart presentationPart, SlidePart slidePart)
        {
            //Create a new slide part in the presentation. 

            SlidePart newSlidePart = presentationPart.AddNewPart<SlidePart>("Slide" + (presentationPart.SlideParts.Count() + 1).ToString());

            //Add the slide template content into the new slide. 
            newSlidePart.FeedData(slidePart.GetStream(FileMode.Open));
            //Make sure the new slide references the proper slide layout. 
            newSlidePart.AddPart(slidePart.SlideLayoutPart);
            //Get the list of slide ids. 
            SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;
            //Deternmine where to add the next slide (find max number of slides). 
            uint maxSlideId = 1;
            SlideId prevSlideId = null;
            foreach (SlideId slideId in slideIdList.ChildElements)
            {
                if (slideId.Id > maxSlideId)
                {
                    maxSlideId = slideId.Id;
                    prevSlideId = slideId;
                }
            }
            maxSlideId++;
            //Add the new slide at the end of the deck. 
            SlideId newSlideId = slideIdList.InsertAfter(new SlideId(), prevSlideId);
            //Make sure the id and relid are set appropriately. 
            newSlideId.Id = maxSlideId;
            newSlideId.RelationshipId = presentationPart.GetIdOfPart(newSlidePart);
            return newSlidePart;
        }

        // delete slide
        public static void deletePPTSlide(PresentationPart presentationPart, SlidePart slidePart)
        {
            //Get the list of slide ids. 
            SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;
            //Delete the template slide reference. 
            foreach (SlideId slideId in slideIdList.ChildElements)
            {
                if (slideId.RelationshipId.Value.Equals("rId3")) slideIdList.RemoveChild(slideId);
            }
            //Delete the template slide. 
            presentationPart.DeletePart(slidePart);
        }

        //
        public static void ReplaceValuesInChartInSlide(ChartPart chartPart, Dictionary<string, int> data, string categoryTitle)
        {
            C.ChartSpace chartSpace = chartPart.ChartSpace;
            List<int> indexOfUsedItems = new List<int>();

            C.BarChart barChart = chartSpace.Descendants<C.BarChart>().FirstOrDefault();

            barChart.RemoveAllChildren<C.BarChartSeries>();
            uint i = 0;
            foreach (string key in data.Keys)
            {
                C.BarChartSeries barChartSeries = barChart.AppendChild<C.BarChartSeries>(
                    new C.BarChartSeries(new C.Index() { Val = new UInt32Value(i) },
                                         new C.Order() { Val = new UInt32Value(i) },
                                         new C.SeriesText(new C.NumericValue() { Text = key })));

                C.StringLiteral strLit = barChartSeries.AppendChild<C.CategoryAxisData>(new C.CategoryAxisData()).AppendChild<C.StringLiteral>(new C.StringLiteral());
                strLit.Append(new C.PointCount() { Val = new UInt32Value(1U) });
                strLit.AppendChild<C.StringPoint>(new C.StringPoint() { Index = new UInt32Value(0U) }).Append(new C.NumericValue(categoryTitle));

                C.NumberLiteral numLit = barChartSeries.AppendChild<DocumentFormat.OpenXml.Drawing.Charts.Values>(
                    new DocumentFormat.OpenXml.Drawing.Charts.Values()).AppendChild<C.NumberLiteral>(new C.NumberLiteral());
                numLit.Append(new C.FormatCode("General"));
                numLit.Append(new C.PointCount() { Val = new UInt32Value(1U) });
                numLit.AppendChild<C.NumericPoint>(new C.NumericPoint() { Index = new UInt32Value(0u) }).Append(new C.NumericValue(data[key].ToString()));

                i++;
            }

            chartSpace.Save();
        }

        public static void replacePPTChartData(ChartPart chartPart, List<myCell> chartData)
        {
        }


        // ==========================================
        // Excel Functions
        // ==========================================

        // find all worksheets from an Excel file (workbookPart) - return as List
        public static List<string> getAllXLSWorksheets(string xFileName)
        {
            List<string> wsList = null;

            using (SpreadsheetDocument myDoc = SpreadsheetDocument.Open(xFileName, false))
            {
                WorkbookPart wbPart = myDoc.WorkbookPart;
                //System.Collections.ArrayList wsList 
                wsList = (from Sheet s in wbPart.Workbook.Sheets
                          select s.Name.ToString()).ToList();
            }

            return wsList;
        }

        // find all worksheets from an Excel file (workbookPart) - return as ArrayList
        public static ArrayList getAllXLSWorksheets2(string xFileName)
        {
            ArrayList wsArrayList = new ArrayList();

            List<string> wsList = getAllXLSWorksheets(xFileName);
            foreach (string wsName in wsList)
                wsArrayList.Add(wsName);

            return wsArrayList;
        }

        // create PPT table from an Excel Spreadsheet
        public static D.Table transferXLSSheet2PPTTable(string xFileName, string xSheetName, int startCol = 0, int startRow = 0)
        {
            D.Table table1 = new D.Table();

            D.TableProperties tableProperties1 = new D.TableProperties() { FirstRow = false, BandRow = true };

            D.TableStyleId tableStyleId1 = new D.TableStyleId();
            tableStyleId1.Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}";
            tableProperties1.Append(tableStyleId1);

            D.TableGrid tableGrid1 = new D.TableGrid();

            D.GridColumn[] tableCols;
            D.TableRow[] tableRows;

            using (SpreadsheetDocument myDoc = SpreadsheetDocument.Open(xFileName, false))
            {
                WorkbookPart wbPart = myDoc.WorkbookPart;
                // find worksheetPart from xSheetName
                WorksheetPart wsPart = (from Sheet s in wbPart.Workbook.Sheets
                                        where s.Name == xSheetName
                                        select (WorksheetPart)(wbPart.GetPartById(s.Id))).SingleOrDefault();

                // cannot count the number of columns because not all columns are recorded (some empty columns are skipped)
                // 3 different ways to get maxCol
                // #1
                //int maxCol = (from Column col in wsPart.Worksheet
                //                select (int)col.Max.Value).Max();
                // #2
                //Column lastCol = wsPart.Worksheet.Descendants<Column>().Last();
                //int maxCol = (int)lastCol.Max.Value;
                // #3
                int maxCol = wsPart.Worksheet.Descendants<Column>().Max(col => (int)col.Max.Value);
                int maxRow = wsPart.Worksheet.Descendants<Row>().Count();

                tableCols = new D.GridColumn[maxCol-startCol+1];

                for (int I = startCol; I <= maxCol; I++)
                {
                    tableCols[I-startCol] = new D.GridColumn() { Width = 524000L * (I < startCol+4 ? 2 : 1) };
                    tableGrid1.Append(tableCols[I-startCol]);
                }

                // generate tableRows (generate the exact number of rows)
                tableRows = addPPTTableRows(wbPart, xSheetName, maxCol, maxRow, startCol, startRow);

            }

            // no need to substract startRow, since the exact number of rows is returned
            table1.Append(tableProperties1);
            table1.Append(tableGrid1);
            for (int I = 0; I < tableRows.Count(); I++)
                table1.Append(tableRows[I]);

            return table1;
        }

        // Given a workbook part, a sheet name, and row and column integer values, retrieve the value of the cell.
        // Call the function like this to retrieve the value from cell B3:
        // string value = XLGetCellValueRowCol("Sample.xlsx", "Sheet1", 2, 3);
        public static string getXLSCellValueRowCol(WorkbookPart wbPart, string sheetName, int colNumber, int rowNumber)
        {
            // The challenge: Convert a cell number into a letter name.
            return getXLSCellValueRowCol(wbPart, sheetName, XLS.getColumnName(colNumber), rowNumber);
        }

        // Given a workbook part, a sheet name, and row and column names, retrieve the value of the cell.
        // Call the function like this:
        // string value = xGetCellValue("Sample.xlsx", "Sheet1", "B", 3);
        public static string getXLSCellValueRowCol(WorkbookPart wbPart, string sheetName, string colName, int rowNumber)
        {
            string value = null;

            //    WorkbookPart wbPart = document.WorkbookPart;

            // Find the sheet with the supplied name, and then use that Sheet object
            // to retrieve a reference to the appropriate worksheet.
            Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();

            if (theSheet == null)
            {
                throw new ArgumentException("sheetName");
            }

            // Retrieve a reference to the worksheet part, and then use its Worksheet property to get 
            // a reference to the cell whose address matches the address you've supplied:
            WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));
            Cell theCell = wsPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == colName + rowNumber.ToString()).FirstOrDefault();

            // If the cell doesn't exist, return an empty string.
            if (theCell != null)
            {
                value = theCell.InnerText;

                // If the cell represents an integer number, you're done. 
                // For dates, this code returns the serialized value that 
                // represents the date. The code handles strings and booleans
                // individually. For shared strings, the code looks up the corresponding
                // value in the shared string table. For booleans, the code converts 
                // the value into the words TRUE or FALSE.
                if (theCell.DataType != null)
                {
                    switch (theCell.DataType.Value)
                    {
                        case CellValues.SharedString:
                            // For shared strings, look up the value in the shared strings table.
                            var stringTable = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                            // If the shared string table is missing, something's wrong.
                            // Just return the index that you found in the cell.
                            // Otherwise, look up the correct text in the table.
                            if (stringTable != null)
                            {
                                value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                            }
                            break;

                        case CellValues.Boolean:
                            switch (value)
                            {
                                case "0":
                                    value = "FALSE";
                                    break;
                                default:
                                    value = "TRUE";
                                    break;
                            }
                            break;
                    }
                }
            }

            return value;
        }

        // read all cells in a excel worksheet
        public static List<myCell> readFromXLSSheet(string fileName, int sheetNumber)
        {
            List<myCell> parsedCells = new List<myCell>();
            //string fileName = @"C:\Documents and Settings\awang\Desktop\Office XML\alex_test 64K rows.xlsx";
            using (Package xlsxPackage = Package.Open(fileName, FileMode.Open, FileAccess.ReadWrite))
            {
                try
                {
                    PackagePartCollection allParts = xlsxPackage.GetParts();

                    // prepare shared string dictionary
                    PackagePart sharedStringsPart = (from part in allParts
                                                     where part.ContentType.Equals("application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml")
                                                     select part).SingleOrDefault();

                    Dictionary<int, string> sharedStrings = new Dictionary<int, string>();
                    if (sharedStringsPart != null)
                    {
                        XElement sharedStringsElement = XElement.Load(XmlReader.Create(sharedStringsPart.GetStream()));
                        parseXLSSharedStrings(sharedStringsElement, sharedStrings);
                    }

                    // get worksheet and prepare to get datat
                    XElement worksheetElement = getXLSWorksheet(sheetNumber, allParts);

                    // get all cells (just cells)
                    IEnumerable<XElement> allCells = from c in worksheetElement.Descendants(OpenXMLNamespaces.excelNamespace + "c")
                                                     select c;

                    foreach (XElement cell in allCells)
                    {
                        string cellPosition = cell.Attribute("r").Value;
                        // check if shared string
                        bool isShared = (cell.Attribute("t") == null) ? false : cell.Attribute("t").Value == "s";
                        int index = indexOfNumber(cellPosition);
                        string column = cellPosition.Substring(0, index);
                        int row = Convert.ToInt32(cellPosition.Substring(index, cellPosition.Length - index));
                        // get value 
                        XElement xel = cell.Descendants(OpenXMLNamespaces.excelNamespace + "v").SingleOrDefault();
                        if (xel != null)
                        {
                            if (isShared)
                                parsedCells.Add(new myCell(column, row, sharedStrings[Convert.ToInt32(xel.Value)]));
                            else
                                parsedCells.Add(new myCell(column, row, xel.Value.ToString()));
                        }
                    }
                }
                finally
                {
                    xlsxPackage.Close();
                }
            }
            //Console.WriteLine(parsedCells.Count().ToString());
            //From here is additional code not covered in the posts, just to show it works
            //foreach (Cell cell in parsedCells)
            //{
            //    Console.WriteLine(cell);
            //}

            return parsedCells;
        }

        // populate shared string into a Dictionary object
        private static void parseXLSSharedStrings(XElement SharedStringsElement, Dictionary<int, string> sharedStrings)
        {
            IEnumerable<XElement> sharedStringsElements = from s in SharedStringsElement.Descendants(OpenXMLNamespaces.excelNamespace + "t")
                                                          select s;

            int Counter = 0;
            foreach (XElement sharedString in sharedStringsElements)
            {
                sharedStrings.Add(Counter, sharedString.Value);
                Counter++;
            }
        }

        // get excel worksheet (as XEelement) from a PackagePartCollection
        private static XElement getXLSWorksheet(int worksheetID, PackagePartCollection allParts)
        {
            PackagePart worksheetPart = (from part in allParts
                                         where part.Uri.OriginalString.Equals(String.Format("/xl/worksheets/sheet{0}.xml", worksheetID))
                                         select part).Single();

            return XElement.Load(XmlReader.Create(worksheetPart.GetStream()));
        }

        // return first numeric character in a string; if not found, return 0
        private static int indexOfNumber(string value)
        {
            for (int counter = 0; counter < value.Length; counter++)
            {
                if (char.IsNumber(value[counter]))
                {
                    return counter;
                }
            }

            return 0;
        }

    }

    internal static class OpenXMLNamespaces
    {
        internal static XNamespace excelNamespace = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        internal static XNamespace relationshipsNamepace = XNamespace.Get("http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        internal static XNamespace drawingChartNamespace = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/chart");
    }

    // myCell to hold Excel Cell's row/column/value
    public class myCell
    {
        public myCell(string column, int row, string value)
        {
            this.Column = column;
            this.Row = row;
            this.Value = value;
        }

        public override string ToString()
        {
            return string.Format("{0}:{1} - {2}", Row, Column, Value);
        }

        public string Column { get; set; }
        public int Row { get; set; }
        public string Value { get; set; }
    }

    // structure for myCell Column Indicator
    public struct myCellColumnIndicator
    {
        public string slide;
        public string shape;
        public string banner;
        public string stub;
        public string firm;
        public string measure;

        public myCellColumnIndicator(List<myCell> myCells)
        {
            slide = myCells.Where(c => c.Row == PPTX.HEADER_ROW && c.Value == PPTX.SLIDE_COLUMN_NAME).Single().Column;
            shape = myCells.Where(c => c.Row == PPTX.HEADER_ROW && c.Value == PPTX.SHAPE_COLUMN_NAME).Single().Column;
            banner = myCells.Where(c => c.Row == PPTX.HEADER_ROW && c.Value == PPTX.BANNER_COLUMN_NAME).Single().Column;
            stub = myCells.Where(c => c.Row == PPTX.HEADER_ROW && c.Value == PPTX.STUB_COLUMN_NAME).Single().Column;
            firm = myCells.Where(c => c.Row == PPTX.HEADER_ROW && c.Value == PPTX.FIRM_COLUMN_NAME).Single().Column;
            measure = myCells.Where(c => c.Row == PPTX.HEADER_ROW && c.Value == PPTX.MEASURE_COLUMN_NAME).Single().Column;
        }
    }

}

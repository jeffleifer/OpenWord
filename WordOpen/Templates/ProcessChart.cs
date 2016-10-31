using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPowerTools;
using WordOpen.Model;
using DataTable = System.Data.DataTable;

namespace WordOpen.Templates
{
    internal class ProcessChart : BaseProcess
    {

        public static void Process(WordprocessingDocument docx,
            DataColumnInfo dcInfo,
            ReportDetailTemplate detail,
            string contentControlTag)
        {
            var data = GetChartData(detail);
            ChartUpdater.UpdateChart(docx, contentControlTag, data);
        }

        public static void Process(WordprocessingDocument docx,
            DataColumnInfo dcInfo,
            ReportDetailTemplate detail,
            SimpleField field)
        {
            var p = GetFirstParent<Paragraph>(field);
            var drawing = p?.Descendants<Drawing>().FirstOrDefault();
            var chartRef = drawing?.Descendants<ChartReference>().FirstOrDefault();
            if (chartRef == null) return;

            var firstOrDefault = docx.MainDocumentPart
                .Parts
                .FirstOrDefault(c => c.RelationshipId == chartRef.Id);
            if (firstOrDefault == null)
            {
                return;
            }
            var chartPart = (ChartPart)firstOrDefault
                .OpenXmlPart;
            var data = GetChartData(detail);
            ChartUpdater.UpdateChart(chartPart, data);
        }

   
        private static ChartData GetChartData(ReportDetailTemplate detail)
        {
           

            var data = new ChartData
            {
                SeriesNames = detail.SeriesFields.ToArray(),
                CategoryDataType = (ChartDataType)detail.DetailType
               
            };
            if (detail.DetailType == DetailType.DateTime)
                data.CategoryFormatCode = 14;
            var vw = new DataView(detail.Data);
            var rws = vw.ToTable(true, detail.CategoriesFieldName);
            var chartValues = new double[detail.SeriesFields.Count][];
            for(var c= 0; c<detail.SeriesFields.Count; c++)
                chartValues[c] = new double[rws.Rows.Count];
            var categoryNames = new List<string>();
            
            var i = 0;
            foreach (DataRow rw in rws.Rows )
            {
                var catValue = rw[0].ToString();
                categoryNames.Add(detail.DetailType == DetailType.DateTime
                    ? ToExcelInteger((DateTime) rw[0])
                    : rw[0].ToString());
                string appostroph = (detail.DetailType == DetailType.Number ? "" : "'");
                string filter = detail.CategoriesFieldName + "=" + appostroph + catValue + appostroph + "";
                
                var j = 0;
                foreach (var dr in detail.Data.Select(filter))
                {
                    foreach (var d in detail.SeriesFields.Select(field => double.Parse(dr[field].ToString())))
                    {
                        chartValues[j][i] = d;
                        j++;
                    }
                    break;
                }
                i++;
            }
            data.CategoryNames = categoryNames.ToArray();
            data.Values = chartValues;
            return data;
        }

       
    }
}
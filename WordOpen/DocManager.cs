using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordOpen.Model;
using WordOpen.Templates;

namespace WordOpen
{
    public class DocManager
    {
       
        public DocManager()
        {
            
        }

        private string _connectionString= @"Data Source=.\SQLEXPRESS;Initial Catalog=STEPS;Integrated Security=True";
        public DocManager(string connectionString)
        {
            _connectionString = connectionString;
        }

        private ReportTemplate GetReport(int id)
        {
            var rpt = new ReportTemplate
            {
                Id = id,
                Name = "Word Report"
            };
            rpt.Details.Add(new ReportDetailTemplate
            {
                Id = 1,
                ReportTemplateId = 1,
                DetailType = DetailType.Table,
                DetailName = "MonthlySummary",
                SourceSql = "SELECT  ReportDate, SamplesCompleted, ReportsCompleted, AnalyticalRequests  FROM  vMonthlySummary"
            });

            rpt.Details.Add(new ReportDetailTemplate
            {
                Id = 2,
                ReportTemplateId = 1,
                DetailName = "ESeries",
                DetailType = DetailType.Table,
                SourceSql = "SELECT  ReportDate, No, Days, avgDays FROM   vESeries"
            });

            rpt.Details.Add(new ReportDetailTemplate
            {
                Id = 3,
                ReportTemplateId = 1,
                DetailType =  DetailType.DateTime,
                SeriesField = "SamplesCompleted,ReportsCompleted",
                CategoriesFieldName = "ReportDate",
                DetailName = "Series1",
                SourceSql = "SELECT  ReportDate, SamplesCompleted, ReportsCompleted, AnalyticalRequests  FROM  vMonthlySummary Series1"
            });
            rpt.Details.Add(new ReportDetailTemplate
            {
                Id = 4,
                ReportTemplateId = 1,
                DetailType = DetailType.Number,
                DetailName = "Series2",
                SeriesField = "Days,avgDays",
                CategoriesFieldName = "No",
                SourceSql = "SELECT  No, Days, avgDays FROM vESeries Series2"
            });
            rpt.Details.Add(new ReportDetailTemplate
            {
                Id = 5,
                ReportTemplateId = 1,
                DetailType = DetailType.Number,
                DetailName = "Series3",
                SeriesField = "avgDays",
                CategoriesFieldName = "Quarter",
                SourceSql = "SELECT  Quarter, Avg(avgDays) as avgDays FROM vAvgPie Series3 GROUP BY Quarter"
            });
            return rpt;
        }

        private void Populate( ReportTemplate template)
        {
            template.DataSet = new DataSet();
            foreach (var detail in template.Details)
            {
                using (var da = new SqlDataAdapter(detail.SourceSql,_connectionString))
                {
                    detail.Data = new DataTable(detail.DetailName);
                    da.Fill(detail.Data);
                    template.DataSet.Tables.Add(detail.Data);
                }
            }
        }

       
        public void Process(string templateName, string fileName, string measured, float operatorValue, float inspectorValue, float sigma)
        {
            // first read document in as stream
            byte[] original = File.ReadAllBytes(templateName);
            string[] switches = new string[] {};
            var template = GetReport(1);
            Populate(template);
            
            using (var stream = new MemoryStream())
            {
                stream.Write(original, 0, original.Length);

                using (var docx = WordprocessingDocument.Open(stream, true))
                {
                    ConvertFieldCodes(docx.MainDocumentPart.Document);
                    foreach (var detail in template.Charts)
                    {
                        var fieldName = (detail.DetailName.ToUpper().StartsWith("CRT") ? detail.DetailName: "CRT_" + detail.DetailName);
                        var dcInfo = BaseProcess.GetColumnInfo(fieldName);
                        if (!GetTableName(template, dcInfo))
                            continue;
                        if (dcInfo.FieldType == FieldType.Chart)
                            ProcessChart.Process(docx, dcInfo, detail, detail.DetailName);
                    }
                    foreach (var field in docx.MainDocumentPart.Document.Descendants<SimpleField>())
                    {
                        var fieldname = BaseProcess.GetFieldName(field, out switches);
                        if (string.IsNullOrEmpty(fieldname))
                            continue;
                        var dcInfo = BaseProcess.GetColumnInfo(fieldname);
                        if(dcInfo.FieldType == FieldType.Chart) continue;
                        if (dcInfo.FieldType == FieldType.Table)
                        {
                            if (!GetTableName(template, dcInfo))
                                continue;
                            var table = template.DataSet.Tables[dcInfo.TableNameInDb];
                            if (dcInfo.FieldType == FieldType.Table)
                                ProcessTable.Process(docx, dcInfo, table, field, fieldname);
                        }
                        else if (dcInfo.FieldType == FieldType.Image)
                            ProcessImage.ShowIndicator(docx, field,measured,operatorValue,inspectorValue,sigma);
                    }
                    ProcessTable.RemoveEmptyTables(docx);


                    docx.MainDocumentPart.Document.Save();


                }

                stream.Seek(0, SeekOrigin.Begin);
                byte[] data = stream.ToArray();

                File.WriteAllBytes(fileName, data);
            }
        }

        internal static void ConvertFieldCodes(OpenXmlElement mainElement)
        {
            //  search for all the Run elements 
            var runs = mainElement.Descendants<Run>().ToArray();
            if (runs.Length == 0) return;

            var newfields = new Dictionary<Run, Run[]>();

            int cursor = 0;
            do
            {
                Run run = runs[cursor];

                if (run.HasChildren && run.Descendants<FieldChar>().Count() > 0
                    && (run.Descendants<FieldChar>().First().FieldCharType & FieldCharValues.Begin) == FieldCharValues.Begin)
                {
                    List<Run> innerRuns = new List<Run>();
                    innerRuns.Add(run);

                    //  loop until we find the 'end' FieldChar
                    bool found = false;
                    string instruction = null;
                    RunProperties runprop = null;
                    do
                    {
                        cursor++;
                        run = runs[cursor];

                        innerRuns.Add(run);
                        if (run.HasChildren && run.Descendants<FieldCode>().Count() > 0)
                            instruction += run.GetFirstChild<FieldCode>().Text;
                        if (run.HasChildren && run.Descendants<FieldChar>().Count() > 0
                            && (run.Descendants<FieldChar>().First().FieldCharType & FieldCharValues.End) == FieldCharValues.End)
                        {
                            found = true;
                        }
                        if (run.HasChildren && run.Descendants<RunProperties>().Count() > 0)
                            runprop = run.GetFirstChild<RunProperties>();
                    } while (found == false && cursor < runs.Length);

                    //  something went wrong : found Begin but no End. Throw exception
                    if (!found)
                        throw new Exception("Found a Begin FieldChar but no End !");

                    if (!string.IsNullOrEmpty(instruction))
                    {
                        //  build new Run containing a SimpleField
                        Run newrun = new Run();
                        if (runprop != null)
                            newrun.AppendChild(runprop.CloneNode(true));
                        SimpleField simplefield = new SimpleField();
                        simplefield.Instruction = instruction;
                        newrun.AppendChild(simplefield);

                        newfields.Add(newrun, innerRuns.ToArray());
                    }
                }

                cursor++;
            } while (cursor < runs.Length);

            //  replace all FieldCodes by old-style SimpleFields
            foreach (KeyValuePair<Run, Run[]> kvp in newfields)
            {
                kvp.Value[0].Parent.ReplaceChild(kvp.Key, kvp.Value[0]);
                for (int i = 1; i < kvp.Value.Length; i++)
                    kvp.Value[i].Remove();
            }
        }
        private static bool GetTableName(ReportTemplate template, DataColumnInfo dcInfo)
        {
            var name = dcInfo.TableName.ToUpper();
            dcInfo.TableNameInDb = string.Empty;
            foreach (DataTable t in template.DataSet.Tables)
            {
                
                if (t.TableName.ToUpper() != name)
                    continue;
                dcInfo.TableNameInDb = t.TableName;
                dcInfo.HasRows = t.Rows.Count > 0;
                return dcInfo.HasRows;
            }
            return false;
        }

    }
}

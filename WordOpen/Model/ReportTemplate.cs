using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data;
using System.Linq;

namespace WordOpen.Model
{
    
    public class ReportTemplate
    {
        public int Id { get; set; }

        public string Name { get; set; }

        public string WordTemplate { get; set; }
        public List<ReportDetailTemplate> Details { get; set; } = new List<ReportDetailTemplate>();

        private Dictionary<string, ReportDetailTemplate> _detailDictionary = new Dictionary<string, ReportDetailTemplate>();
        public Dictionary<string, ReportDetailTemplate> DetailDictionary
        {
            get
            {
                if (Details.Count != _detailDictionary.Count)
                {
                    _detailDictionary = Details.ToDictionary(d => d.DetailName);
                }
                return _detailDictionary;
            }
        } 
        public List<ReportDetailTemplate> Tables
        {
            get { return Details.Where(d => d.DetailType == DetailType.Table).ToList(); }
        }

        public List<ReportDetailTemplate> Charts
        {
            get { return Details.Where(d => d.DetailType != DetailType.Table ).ToList(); }
        }
        [NotMapped]
        public DataSet DataSet { get; set; }
    }

    public class ReportDetailTemplate
    {
        public int Id { get; set; }

        public DetailType DetailType { get; set; }
        public int ReportTemplateId { get; set; }

        public string DetailName { get; set; }

        public string SourceSql { get; set; }

        [NotMapped]
        public DataTable Data { get; set; }

        public string SeriesField { get; set; }

        public string CategoriesFieldName { get; set; }

        
        private List<string> _fieldList = new List<string>();
        [NotMapped]
        public List<string> SeriesFields
        {
            get
            {
                
                if (Data.Columns.Count == 0 || string.IsNullOrEmpty(SeriesField))
                {
                    _fieldList = new List<string>();
                    return _fieldList;
                }
                
                var sep = new[] {','};
                var fields = SeriesField.Split(sep, StringSplitOptions.RemoveEmptyEntries);
                if (fields.Length == _fieldList.Count) return _fieldList;
                _fieldList = new List<string>();
                _fieldList.AddRange(from field in fields from DataColumn column in Data.Columns where string.Equals(column.ColumnName, field, StringComparison.CurrentCultureIgnoreCase) select column.ColumnName);
                foreach (var column in Data.Columns.Cast<DataColumn>().Where(column => string.Equals(column.ColumnName, CategoriesFieldName, StringComparison.CurrentCultureIgnoreCase)))
                {
                    CategoriesFieldName = column.ColumnName;
                    break;
                }
                return _fieldList;
            }
        }

    }
}
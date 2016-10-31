using System;
using System.Collections.Generic;
using System.Data;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordOpen.Model;

namespace WordOpen.Templates
{
    internal  class ProcessTable : BaseProcess
    {
 
        public static void Process(WordprocessingDocument docx,
            DataColumnInfo dcInfo,
            DataTable dataTable,
            SimpleField field,
            string fieldName)
        {

            string[] switches;
            if (string.IsNullOrEmpty(fieldName) ||
                dcInfo.FieldType != FieldType.Table)
                return;
            var wordRow = GetFirstParent<TableRow>(field);
            if (wordRow == null)
                return;

            var wordTable = GetFirstParent<Table>(wordRow);


            if (!dcInfo.HasRows || wordTable == null)
                return;
            var props = new List<TableCellProperties>();
            var cellcolumnnames = new List<string>();
            var paragraphInfo = new List<string>();
            var cellfields = new List<SimpleField>();

            foreach (var cell in wordRow.Descendants<TableCell>())
            {
                props.Add(cell.GetFirstChild<TableCellProperties>());
                Paragraph p = cell.GetFirstChild<Paragraph>();
                if (p != null)
                {
                    var pp = p.GetFirstChild<ParagraphProperties>();
                    paragraphInfo.Add(pp?.OuterXml);
                }
                else
                {
                    paragraphInfo.Add(null);
                }

                var colname = string.Empty;

                SimpleField colfield = null;
                foreach (var cellfield in cell.Descendants<SimpleField>())
                {
                    colfield = cellfield;
                    colname = GetColumnInfo(GetFieldName(cellfield, out switches)).FieldName;
                    break;
                }

                cellcolumnnames.Add(colname);
                cellfields.Add(colfield);
            }

            // keep reference to row properties
            var rprops = wordRow.GetFirstChild<TableRowProperties>();

            foreach (DataRow row in dataTable.Rows)
            {
                TableRow nrow = new TableRow();

                if (rprops != null)
                {
                    nrow.Append(new TableRowProperties(rprops.OuterXml));
                }

                for (var i = 0; i < props.Count; i++)
                {
                    TableCellProperties cellproperties = new TableCellProperties(props[i].OuterXml);
                    TableCell cell = new TableCell();
                    cell.Append(cellproperties);
                    Paragraph p = new Paragraph(new ParagraphProperties(paragraphInfo[i]));
                    cell.Append(p);   // cell must contain at minimum a paragraph !

                    if (!string.IsNullOrEmpty(cellcolumnnames[i]))
                    {
                        if (!dataTable.Columns.Contains(cellcolumnnames[i]))
                        {
                            throw new Exception(string.Format("Unable to complete template: column name '{0}' is unknown in parameter tables !", cellcolumnnames[i]));
                        }

                        if (!row.IsNull(cellcolumnnames[i]))
                        {
                            string val = row[cellcolumnnames[i]].ToString();
                            p.Append(GetRunElementForText(val, cellfields[i]));
                        }
                    }

                    nrow.Append(cell);
                }

                wordTable.Append(nrow);
            }
        }

        internal static void RemoveEmptyTables(WordprocessingDocument docx)
        {
           
        }
      

        internal static Run GetRunElementForText(string text, SimpleField placeHolder)
        {
            string rpr = null;
            if (placeHolder != null)
            {
                foreach (RunProperties placeholderrpr in placeHolder.Descendants<RunProperties>())
                {
                    rpr = placeholderrpr.OuterXml;
                    break;  // break at first
                }
            }

            Run r = new Run();
            if (!string.IsNullOrEmpty(rpr))
            {
                r.Append(new RunProperties(rpr));
            }

            if (!string.IsNullOrEmpty(text))
            {
                // first process line breaks
                string[] split = text.Split(new[] { "\n" }, StringSplitOptions.None);
                bool first = true;
                foreach (string s in split)
                {
                    if (!first)
                    {
                        r.Append(new Break());
                    }

                    first = false;

                    // then process tabs
                    bool firsttab = true;
                    string[] tabsplit = s.Split(new[] { "\t" }, StringSplitOptions.None);
                    foreach (string tabtext in tabsplit)
                    {
                        if (!firsttab)
                        {
                            r.Append(new TabChar());
                        }

                        r.Append(new Text(tabtext));
                        firsttab = false;
                    }
                }
            }

            return r;
        }

        
    }

}
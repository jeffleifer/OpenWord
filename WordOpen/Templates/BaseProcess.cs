using System;
using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using WordOpen.Model;

namespace WordOpen.Templates
{
   
    internal  class BaseProcess
    {
        private static readonly Regex InstructionRegEx =
    new Regex(
        @"^[\s]*MERGEFIELD[\s]+(?<name>[#\w]*){1}               
                            [\s]*(\\\*[\s]+(?<Format>[\w]*){1})?                
                            [\s]*(\\b[\s]+[""]?(?<PreText>[^\\]*){1})?        
                                                                               
                            [\s]*(\\f[\s]+[""]?(?<PostText>[^\\]*){1})?",
        RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.ExplicitCapture | RegexOptions.IgnoreCase | RegexOptions.IgnorePatternWhitespace | RegexOptions.Singleline);

        protected static string ToExcelInteger(DateTime dateTime)
        {
            return dateTime.ToOADate().ToString(CultureInfo.InvariantCulture);
        }
        internal static string GetFieldName(SimpleField field, out string[] switches)
        {
            var a = field.GetAttribute("instr", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            switches = new string[0];
            string fieldname = string.Empty;
            string instruction = a.Value;

            if (!string.IsNullOrEmpty(instruction))
            {
                Match m = InstructionRegEx.Match(instruction);
                if (m.Success)
                {
                    fieldname = m.Groups["name"].ToString().Trim();
                    int pos = fieldname.IndexOf('#');
                    if (pos > 0)
                    {
                        // Process the switches, correct the fieldname.
                        switches = fieldname.Substring(pos + 1).ToLower().Split(new[] { '#' }, StringSplitOptions.RemoveEmptyEntries);
                        fieldname = fieldname.Substring(0, pos);
                    }
                }
            }

            return fieldname;
        }

        internal static DataColumnInfo GetColumnInfo(string fieldName)
        {

            if (fieldName == null)
                throw new ArgumentException("Error: table-MERGEFIELD should be formatted as follows: XYZ_TableName_ColumnName.");
            var sep = new[] { '_' };
            var splits = fieldName.Split(sep, StringSplitOptions.RemoveEmptyEntries);
            var cInfo = new DataColumnInfo
            {
                TableName = splits[1]

            };



            if (splits.Length == 3)
            {

                cInfo.FieldName = splits[2];
            }
            if (splits[0].ToUpper().Equals("TBL"))
                cInfo.FieldType = FieldType.Table;
            else if (splits[0].ToUpper().Equals("CRT"))
                cInfo.FieldType = FieldType.Chart;
            else if (splits[0].ToUpper().Equals("IMG"))
                cInfo.FieldType = FieldType.Image;
            else
            {
                cInfo.FieldType = FieldType.Invalid;
            }
            return cInfo;
        }


        internal static T GetFirstParent<T>(OpenXmlElement element)
          where T : OpenXmlElement
        {
            if (element.Parent == null)
            {
                return null;
            }
            if (element.Parent.GetType() == typeof(T))
            {
                return element.Parent as T;
            }
            return GetFirstParent<T>(element.Parent);
        }
    }

    
}
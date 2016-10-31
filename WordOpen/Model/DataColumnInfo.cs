namespace WordOpen.Model
{
    public class DataColumnInfo
    {

        public FieldType FieldType { get; set; }

       
        public string FieldName { get; set; }

        public string TableName { get; set; }

        public string TableNameInDb { get; set; }

        public bool HasRows { get; set; }
    }

    public enum FieldType
    {
        Invalid = 0,
        Table = 1,
        Chart = 2,
        Image =3
    }

    public enum DetailType
    {
        Number = 0,
        String = 1,
        DateTime = 2,
        Table =99


    }
}
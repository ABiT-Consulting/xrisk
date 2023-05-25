using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;  
using System.Web.Script.Serialization;
using System.Data;

namespace ExtensionMethods
{
    public static class JSONHelper
    {
        public static string ToJSON(this object obj)
        {
            JavaScriptSerializer serializer = new JavaScriptSerializer();
            return serializer.Serialize(obj);
        }
        public static Object ToObjectFromJson(this string obj)
        {
            JavaScriptSerializer serializer = new JavaScriptSerializer();
            return serializer.Deserialize(obj, typeof(object));
        }
        public static string ToJSON(this object obj, int recursionDepth)
        {
            JavaScriptSerializer serializer = new JavaScriptSerializer();
            serializer.RecursionLimit = recursionDepth;
            return serializer.Serialize(obj);
        }
        public static System .Data .DataTable GetDataTable(this SAPbobsCOM .Recordset recset )
        {

            var datatable = new System.Data.DataTable();
            for (int i = 0; i < recset .Fields .Count ; i++)
            {
                var columnName = recset.Fields.Item(i).Name;
                var type = recset.Fields.Item(i).Type;
                Type t = null ;
                switch (type)
                {
                    case SAPbobsCOM.BoFieldTypes.db_Alpha:
                    case SAPbobsCOM.BoFieldTypes.db_Memo:
                        t = typeof (string);
                        break;
                    case SAPbobsCOM.BoFieldTypes.db_Date:
                        t = typeof (DateTime);
                        break;
                    case SAPbobsCOM.BoFieldTypes.db_Float:
                        t = typeof (double );
                        break;
                        
                    case SAPbobsCOM.BoFieldTypes.db_Numeric:
                        
                        t = typeof (int);
                        break;
                    default:
                        t = typeof(string);
                        break;
                }
              datatable.Columns.Add(columnName,t);

            }
            var rowcount = 0;
            while(!recset.EoF)
            {
                datatable.Rows.Add();
                for (int i = 0; i < recset.Fields.Count; i++)
                {

                    datatable.Rows[rowcount][recset.Fields.Item(i).Name] = recset.Fields.Item(i).Value;
                }
                rowcount++;
                recset.MoveNext();
            }
            return datatable;
        }
     
  
public static string DataTableToJSONWithJSONNet(this DataTable table) {  
   string JSONString=string.Empty;  
   JSONString = JsonConvert.SerializeObject(table);  
   return JSONString;  
}  

    }
}

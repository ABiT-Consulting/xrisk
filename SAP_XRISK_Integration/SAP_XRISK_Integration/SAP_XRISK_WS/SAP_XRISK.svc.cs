using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;
using System.Web;
using System.Web.Hosting;
using System.Xml.Linq;
namespace SAP_XRISK_WS
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the class name "Service1" in code, svc and config file together.
    // NOTE: In order to launch WCF  Test Client for testing this service, please select Service1.svc or Service1.svc.cs at the Solution Explorer and start debugging.
    public class SAP_XRISK : ISAP_XRISK
    {
       
     public   void SaveConnection(string Connection,out string msg)
        {
         var path =   Path.Combine ( HttpContext.Current.Server.MapPath("~/App_Data"),"cnx.p");
         System.IO.File.WriteAllText(path, Connection);
         var c=    Initializer.Company;
            msg = Initializer._Company.GetLastErrorDescription();
        }
        public string GetConnection()
        {
            var path = Path.Combine(HttpContext.Current.Server.MapPath("~/App_Data"), "cnx.p");
            if (!System.IO.File.Exists(path)) return null;
            return System.IO.File.ReadAllText(path);
        }
       
        
        
        #region SAP TO X-RISK


        public int GetXmlandVersion(int version, string objectCode,string criteria, out string xmlobj)
        {
            string result = "";int  newversion = 0;XDocument resultdoc = new XDocument();
            SAPbobsCOM.Recordset rec = Initializer.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
            try {
                var query = string.Format(System.Configuration.ConfigurationManager.AppSettings["getkeys"], Initializer._Company.CompanyDB, version, objectCode);
                  
                if (!string.IsNullOrEmpty(criteria))
                {
                    var query1 = String.Format(System.Configuration.ConfigurationManager.AppSettings["getkeyandtable"], Initializer._Company.CompanyDB, objectCode);
                    rec.DoQuery(query1);
                    rec.MoveFirst();
                    var keyname = rec.Fields.Item("U_KeyName").Value.ToString();
                    var htbl = rec.Fields.Item("U_HTBL").Value.ToString();
                    var csv = getcsv(keyname, htbl, criteria);
                    if (!string.IsNullOrEmpty(csv))
                    {
                        //query = string.Format(System.Configuration.ConfigurationManager.AppSettings["getkeys"], Initializer._Company.CompanyDB, version, objectCode);
                        query = query + "AND(U_Key IN(" + csv + "))";

                    }
                    else
                    {
                        throw new System.Exception("Condition yielded no results");
                    }
                }
                if (objectCode == "157") resultdoc = GETPaymentRun(query, objectCode);
                else resultdoc = getXmlDoc(query, objectCode);

                var objType = resultdoc.Descendants("AdmInfo").First().Descendants("Object").First().Value.ToString();
                if(objType=="24" || objType=="46")
                {
                  //  var query2 = string.Format(System.Configuration.ConfigurationManager.AppSettings["getmaxkey"], );
                   // rec.DoQuery(query2);
                }


              var query2 = string.Format(System.Configuration.ConfigurationManager.AppSettings["getmaxkey"], Initializer._Company.CompanyDB, version, objectCode);
              rec.DoQuery(query2);
              newversion = Convert.ToInt32( rec.Fields.Item("maxnum").Value );





              result = resultdoc.ToString();  
            }
            catch(Exception ex) {
                insertError(version.ToString(), objectCode, DateTime.Now.ToString(), ex.Message, "S2X", "");
            }
            finally
            {
              System.Runtime.InteropServices.Marshal.ReleaseComObject(rec);
            }
            xmlobj = result;
            return newversion;
        }

        public string getcsv(string keyname,string tablename,string cond)
        {
            string result = "";
            SAPbobsCOM.Recordset rec = Initializer.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
            string query = String.Format(System.Configuration.ConfigurationManager.AppSettings["getcsv"], keyname,tablename,"");
            try
            {
                if (!string.IsNullOrEmpty(cond)) { query += " WHERE " + cond; }
                rec.DoQuery(query);
                for(int i = 0; i < rec.RecordCount; i++)
                {
                    if (i == rec.RecordCount - 1) result = result + "'" + rec.Fields.Item(keyname).Value.ToString() + "'";
                    else result = result + "'" + rec.Fields.Item(keyname).Value.ToString() + "',";
                    rec.MoveNext();
                }
            }
            catch(Exception ex)
            {
                insertError("", "", DateTime.Now.ToString(), ex.Message, "S2X", "");
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rec);
            }
            return result;
        }
       
        public XDocument getXmlDoc(string query,string objectCode)
        {
            SAPbobsCOM.Recordset rec = Initializer.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
            XDocument mainDoc = new XDocument();
            try
            {
                #region Dictionary
                IDictionary<string, string> d = new Dictionary<string, string>();
                rec.DoQuery(query);
                while (!rec.EoF)
                {
                    var key = rec.Fields.Item("U_KEY").Value.ToString();
                    var transType = rec.Fields.Item("U_TrType").Value.ToString();
                   
                    if (!d.ContainsKey(key))
                    {
                      d.Add(new KeyValuePair<string, string>(key, transType));
                    }
                    else
                    {
                        d[key] = transType;
                    }
                    
                    rec.MoveNext();
                }
                #endregion

                mainDoc.Add(new XElement("BOM"));
                var BOM = mainDoc.Element("BOM");

                foreach (KeyValuePair<string, string> el in d)
                {
                    var key = el.Key ;
                    var transType = el.Value;
                    Initializer._Company.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode;
                    dynamic obj = Initializer._Company.GetBusinessObject((SAPbobsCOM.BoObjectTypes)Convert.ToInt32(objectCode));
                    obj.GetByKey(key);
                    XDocument doc = XDocument.Parse(obj.GetAsXML().ToString());
                    var botag = doc.Element("BOM").Element("BO");
                    botag.Descendants("AdmInfo").First().Add(new XElement("TransectionType") { Value = Convert.ToString(transType) });
                    
                    if (objectCode=="24" || objectCode == "46")
                    {
                        string result1 = "";
                        string query1 = string.Format(System.Configuration.ConfigurationManager.AppSettings["getPRKey"], key); 
                        SAPbobsCOM.Recordset rec1 = Initializer.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                        rec1.DoQuery(query1);
                        if (rec1.RecordCount > 0)
                        {
                            result1 = Convert.ToString(rec1.Fields.Item("IdEntry").Value);
                        }
                        botag.Descendants("AdmInfo").First().Add(new XElement("PaymentRunId") { Value = Convert.ToString(result1) });
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(rec1);
                    }
                    BOM.Add(botag);
                }

                d.Clear();
                if (rec.RecordCount == 0) mainDoc = new XDocument();
            }
            catch(Exception ex)
            {
                insertError("", objectCode, DateTime.Now.ToString(), ex.Message, "S2X", "");
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rec);
            }
            return mainDoc;
        }

        public void insertError(string version,string objtype,string date,string error,string direction,string uxml)
        {
            SAPbobsCOM.Recordset rec = Initializer.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
            try
            {
                var query =String.Format(System.Configuration.ConfigurationManager.AppSettings["error"],Initializer.Company.CompanyDB,version,objtype,date,error,direction,uxml);
                rec.DoQuery(query);
            } catch(Exception ex) { }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rec);
            }
        }

        public XDocument GETPaymentRun(string query,string objectCode)
        {
            SAPbobsCOM.Recordset rec = Initializer.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
            SAPbobsCOM.Recordset rec2 = Initializer.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;

            XDocument mainDoc = new XDocument();
            try
            {
                rec.DoQuery(query);
                mainDoc.Add(new XElement("BOM"));
                var BOM = mainDoc.Element("BOM");
                while (!rec.EoF)
                {
                    var key = rec.Fields.Item("U_KEY").Value;
                    var query2 = String.Format(System.Configuration.ConfigurationManager.AppSettings["getpaymentrun"], key);
                    rec2.DoQuery(query2);
                    while (!rec2.EoF)
                    {
                        var objid = "";
                        var PymMeth = rec2.Fields.Item("PymMeth").Value.ToString();
                        var RctId = rec2.Fields.Item("RctId").Value.ToString();
                        if (PymMeth == "Outgoing BT") objid = "24";
                        else if(PymMeth == "Incoming BT") objid = "46";
                        Initializer._Company.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode;
                        dynamic obj = Initializer._Company.GetBusinessObject((SAPbobsCOM.BoObjectTypes)Convert.ToInt32(objid));
                        obj.GetByKey(key);
                        XDocument doc = XDocument.Parse(obj.GetAsXML().ToString());
                        var botag = doc.Element("BOM").Element("BO");
                        XName name = "Key";
                        botag.SetAttributeValue(name, key);
                        BOM.Add(botag);
                        rec2.MoveNext();
                    }
                    
                    rec.MoveNext();
                }
                if (rec.RecordCount == 0) mainDoc = new XDocument();
            }
            catch (Exception ex)
            {
                insertError("", objectCode, DateTime.Now.ToString(), ex.Message, "S2X", "");
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rec);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rec2);
            }
            return mainDoc;
        }

        public bool addChangeTracks(string objectCode, string start, string end,string trtype="A")
        {
            SAPbobsCOM.Recordset rec = Initializer.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
            SAPbobsCOM.Recordset rec2 = Initializer.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;

            var result = true;
            try
            {
                var query1 = String.Format(System.Configuration.ConfigurationManager.AppSettings["getkeyandtable"], Initializer._Company.CompanyDB, objectCode);
                rec.DoQuery(query1);
                if (rec.RecordCount == 0) {
                    insertError("", objectCode, DateTime.Now.ToString(), "Table and Key not found from [AB_TRKOBJ]", "S2X", "");
                    return false;
                }
                rec.MoveFirst();
                var keyname = rec.Fields.Item("U_KeyName").Value.ToString();
                var htbl = rec.Fields.Item("U_HTBL").Value.ToString();
                var query = String.Format(System.Configuration.ConfigurationManager.AppSettings["getrecords"], keyname, htbl, start, end);
                rec.DoQuery(query);
                if (rec.RecordCount == 0)
                {
                    insertError("", objectCode, DateTime.Now.ToString(), "No records found to insert in changetracking", "S2X", "");
                    return false;
                }
                while (!rec.EoF)
                {
                    var key = rec.Fields.Item(keyname).Value.ToString();
                    var query2 = string.Format(System.Configuration.ConfigurationManager.AppSettings["getmaxkey"], Initializer._Company.CompanyDB, "0", objectCode);
                    rec2.DoQuery(query2);
                    var newversion = Convert.ToInt32(rec2.Fields.Item("maxnum").Value);
                    var query3 = String.Format(System.Configuration.ConfigurationManager.AppSettings["insertChangeTrack"], Initializer.Company.CompanyDB, newversion+1, objectCode, key, DateTime.Now, trtype, keyname);
                    rec2.DoQuery(query3);
                    rec.MoveNext();
                }

            }
            catch(Exception ex)
            {
                insertError("", objectCode, DateTime.Now.ToString(), ex.Message, "S2X", "");
                result = false;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rec);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rec2);

            }

            return result;
        }


        public string getSchema(int VersionNumber)
        {
            Initializer.Company.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode;
            string res = "";
            List<string> objcts = new List<string>();
            objcts.Add("24");
            objcts.Add("46");
            objcts.Add("4");
            objcts.Add("2");
            objcts.Add("1");
            objcts.Add("13");
            objcts.Add("22");
            objcts.Add("20");
            objcts.Add("18");
            objcts.Add("17");
            objcts.Add("15");
            objcts.Add("157");
            try
            {
               
                if (objcts.Contains(VersionNumber.ToString()))
                {
                    var xml = "";
                    if (VersionNumber == 157)
                    {
                        var path1 = Path.Combine(HostingEnvironment.MapPath("~/App_Code"), "157.xsd");
                        if(File.Exists(path1))
                        xml = File.ReadAllText(path1);
                    }
                    else {

                        xml =Initializer.Company.GetBusinessObjectXmlSchema((SAPbobsCOM.BoObjectTypes)VersionNumber);
                    }
                    res =Convert.ToString(XDocument.Parse( xml));
                }
                else
                {
                    res = "Object Not Implemented Yet";
                }
            
            }catch(Exception ex)
            {
                res = "Error Occoured " + ex.Message;
            }
            finally
            {
              
            }
            
            return res;
        }
        #endregion


        #region X-RISK TO SAP

        public bool XRiskToSap(string xml, out string msg , out string KeyValue)
        {
            bool result = true;
            KeyValue = "";
            try
            {
                List<string> objcts = new List<string>();
                objcts.Add("24");
                objcts.Add("46");
                objcts.Add("4");
                objcts.Add("2");
                objcts.Add("1");
                objcts.Add("13");

                string filename = Path.GetTempFileName();
                XDocument doc1 = XDocument.Parse(xml);
                var objType= doc1.Descendants("AdmInfo").First().Descendants("Object").First().Value.ToString();
                doc1.Save(filename);
                dynamic doc = Initializer.Company.GetBusinessObjectFromXML(filename, 0);
               
                if(objcts.Contains(objType))
                {
                    var i = doc.Add();
                    if (i == 0)
                    {
                        msg = "Succes";
                        KeyValue = Initializer.Company.GetNewObjectKey();
                    }
                    else
                    {
                        result = false;
                        msg = Initializer.Company.GetLastErrorCode() + " " + Initializer.Company.GetLastErrorDescription();
                        insertError("", "", DateTime.Now.ToString(), msg, "X2S", xml);
                    }
                    File.Delete(filename);
                }
                else
                {
                    throw new ArgumentException("Object not implemented yet");
                }
                

            }catch(Exception ex){
                result = false;
                msg = Initializer.Company.GetLastErrorCode() + " " + Initializer.Company.GetLastErrorDescription();
                insertError("", "", DateTime.Now.ToString(), msg, "X2S", xml);
            }

            return result;
        }

        #endregion
    }
}

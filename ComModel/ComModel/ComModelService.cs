using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.Text;
using ExtensionMethods;
using System.Web.SessionState;
using System.Runtime.Remoting.Contexts;
using EncryptDecrypt;
using System.Web.Script.Serialization;
using System.Xml.Linq;
namespace ComModel
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the class name "ComModelService" in both code and config file together.
    public  class ComModelService : IComModelService
    {
        static Dictionary<string, SAPbobsCOM.Company> CompaniesSessions = new Dictionary<string, SAPbobsCOM.Company>();


        /// <summary>
        /// This function will return Session ID as per UserName and Password. This Session ID will be used for Transections
        /// </summary>
        /// <param name="username">SAP user name</param>
        /// <param name="password">SAP Password</param>
        /// <param name="Message">output message as returned by authentication function</param>
        ///  <param name="SessionID">SessionID as returned by authentication function</param>
        /// <returns></returns>
        public void Login(String UserName, String Password, out String Message, out String SessionID)
        {
            SessionID = "";
            Message = "";
            try
            {
                SAPbobsCOM.Company _Company = new SAPbobsCOM.Company();
                _Company.Server = System.Configuration.ConfigurationManager.AppSettings["Server"];
                _Company.DbServerType = (SAPbobsCOM.BoDataServerTypes)Enum.Parse(typeof(SAPbobsCOM.BoDataServerTypes), System.Configuration.ConfigurationManager.AppSettings["DbServerType"]);
                _Company.DbUserName = System.Configuration.ConfigurationManager.AppSettings["DbUserName"];
                _Company.DbPassword = System.Configuration.ConfigurationManager.AppSettings["DbPassword"];
                _Company.UserName = UserName;// System.Configuration.ConfigurationManager.AppSettings["UserName"];
                _Company.Password = Password;// System.Configuration.ConfigurationManager.AppSettings["Password"];
                _Company.CompanyDB = System.Configuration.ConfigurationManager.AppSettings["CompanyDB"];

                var i = _Company.Connect();
                if (i != 0)
                {
                    Message = _Company.GetLastErrorDescription();
                }
                else
                {
                    SessionIDManager manager = new SessionIDManager();
                    count++;
                    count = count + rand.Next();
                    SessionID = count.ToString().EncryptMe("asdiiie1244332434099r#$");
                    CompaniesSessions.Add(SessionID, _Company);
                }
            }
            catch (Exception ex)
            {
                Message = ex.Message + " " + ex.StackTrace;
            }

        }
        static Random rand = new Random();
        static int count = 0;
        public void LogOut(String SessionID)
        {
            CompaniesSessions[SessionID].Disconnect();
            CompaniesSessions.Remove(SessionID);
        }

        /// <summary>
        /// Get Object With The Key
        /// </summary>
        /// <param name="Key"></param>
        /// <param name="ObjType">Integer value of SAPbobsCOM.BoObjectTypes</param>
        /// <returns></returns>
        public JSON_OBJ GetObjectByKey(String Key, object ObjType, string SessionID)
        {
            var json_obj = new JSON_OBJ();
            var oCompany = CompaniesSessions[SessionID];
            //var r = oCompany.Connect();
            if (oCompany.Connected)
            {
                #region Company Connected
                if ((int)ObjType != 11)
                {

                    try
                    {
                        object obj = oCompany.GetBusinessObject((SAPbobsCOM.BoObjectTypes)ObjType);
                        try
                        {

                            if ((int)ObjType == (int)SAPbobsCOM.BoObjectTypes.oInvoices || (int)ObjType == (int)SAPbobsCOM.BoObjectTypes.oDownPayments)
                            {
                                #region Get Document
                                var invoice = obj as SAPbobsCOM.Documents;
                                invoice.GetByKey(Convert.ToInt32(Key));
                                var recset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                                recset.DoQuery(string.Format("SELECT \"Tel1\" from ocpr where \"CntctCode\" = '{0}' ", invoice.ContactPersonCode));

                                var contactpersoncode = recset.Fields.Item("Tel1").Value;

                                System.Runtime.InteropServices.Marshal.ReleaseComObject(recset);
                                GC.Collect();

                                var row = new List<object>();
                                var lines = invoice.Lines;
                                lines.SetCurrentLine(0);
                                var count = invoice.Lines.Count;
                                for (int i = 0; i < count; i++)
                                {
                                    lines.SetCurrentLine(i);
                                    var item = new { ItemCode = lines.ItemCode, ItemName = lines.ItemDescription, Quantity = lines.Quantity, UnitPrice = lines.Price, RowDiscountPerc = lines.DiscountPercent, RowTotal = lines.LineTotal };
                                    row.Add(item);
                                }
                                var data = new
                                {
                                    ObjectInfo = new
                                    {
                                        ObjType = ObjType
                                        ,
                                        ObjectName = ((SAPbobsCOM.BoObjectTypes)ObjType).ToString()
                                        ,
                                        Key = Key
                                    },

                                    Header = new
                                    {

                                        CustomerCode = invoice.CardCode,
                                        CustomerName = invoice.CardName,
                                        ContactNumber = contactpersoncode,
                                        DocumentDate = invoice.DocDate.ToString("yyyyMMdd"),
                                        DueDate = invoice.DocDueDate.ToString("yyyyMMdd"),
                                        Status = invoice.DocumentStatus
                                    }
                                    ,
                                    Rows = row
                                };
                                json_obj.JSON_OBJECT = data.ToJSON();
                                json_obj.OBJECT_Type = ((SAPbobsCOM.BoObjectTypes)ObjType).ToString();

                                #endregion
                            }
                            if ((int)ObjType == (int)SAPbobsCOM.BoObjectTypes.oItems)
                            {
                                #region Item Master Data
                                var item = obj as SAPbobsCOM.Items;
                                item.GetByKey(Key);

                                var data = new
                                {
                                    ObjectInfo = new
                                    {
                                        ObjType = ObjType
                                        ,
                                        ObjectName = ((SAPbobsCOM.BoObjectTypes)ObjType).ToString()
                                        ,
                                        Key = Key
                                    },

                                    Header = new
                                    {

                                        Code = item.ItemCode,
                                        Description = item.ItemName,
                                        Status = item.Frozen
                                    }
                                };
                                json_obj.JSON_OBJECT = data.ToJSON();
                                json_obj.OBJECT_Type = "Item Master Data";

                                #endregion
                            }
                            if ((int)ObjType == (int)SAPbobsCOM.BoObjectTypes.oBusinessPartners)
                            {
                                #region BP Master Data
                                var item = obj as SAPbobsCOM.BusinessPartners;
                                item.GetByKey(Key);

                                var data = new
                                {
                                    ObjectInfo = new
                                    {
                                        ObjType = ObjType
                                        ,
                                        ObjectName = ((SAPbobsCOM.BoObjectTypes)ObjType).ToString()
                                        ,
                                        Key = Key
                                    },

                                    Header = new
                                    {

                                        Code = item.CardCode,
                                        Description = item.CardName,
                                        Status = item.Frozen
                                    }
                                };
                                json_obj.JSON_OBJECT = data.ToJSON();
                                json_obj.OBJECT_Type = "BP Master Data";

                                #endregion
                            }
                        }
                        finally
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                            GC.Collect();
                            obj = null;
                        }
                    }
                    catch (Exception ex)
                    {
                        json_obj.Message = ex.Message;
                    }
                }
                if ((int)ObjType == 11)
                {
                    try
                    {

                        #region Contact Person
                        var recset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                        recset.DoQuery(string.Format("Select * from OCPR where \"CntctCode\"='{0}'", Key));
                        var contactCode = recset.Fields.Item("CntctCode").Value;
                        var data = new
                        {
                            ObjectInfo = new
                            {
                                ObjType = ObjType
                                ,
                                ObjectName = "ContactPerson"
                                ,
                                Key = Key
                            },

                            Header = new
                            {

                                ID = recset.Fields.Item("CntctCode").Value,
                                ContactPersonID = recset.Fields.Item("CntctCode").Value,
                                FirstName = recset.Fields.Item("FirstName").Value,
                                MiddleName = recset.Fields.Item("MiddleName").Value,
                                LastName = recset.Fields.Item("LastName").Value,
                                Status = recset.Fields.Item("Active").Value


                            }
                        };

                        json_obj.JSON_OBJECT = data.ToJSON();
                        json_obj.OBJECT_Type = "Contact Person";
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(recset);
                        GC.Collect();
                        #endregion
                    }

                    catch (Exception ex)
                    {

                        json_obj.Message = ex.Message;
                    }
                }

                #endregion

            }
            else
            {

                json_obj.Message = "Company Not Connected " + oCompany.GetLastErrorDescription();
            }

            #region get udo
            //var r = oCompany.Connect();
            if (oCompany.Connected && ObjType is string)
            {
                ComModel.UserDefinedObject udo = new UserDefinedObject(ObjType.ToString(), SAPbobsCOM.BoUDOObjType.boud_MasterData, oCompany);
                udo.GetByKey(Key);
                // udo.GetProperty ("U_ItemCode");

                var data = new
                {
                    ObjectInfo = new
                    {
                        ObjType = ObjType
                        ,
                        ObjectName = (ObjType).ToString()
                        ,
                        Key = Key
                    },

                    Header = new
                    {

                        U_ItemCode = udo.GetProperty("U_ItemCode"),
                        U_Itemcode1 = udo.GetProperty("U_Itemcode1"),
                        U_ItemCodename = udo.GetProperty("U_ItemCodename"),
                        U_Itemcode2 = udo.GetProperty("U_Itemcode2"),
                        U_itemCodename2 = udo.GetProperty("U_itemCodename2"),
                        U_Otherfees = udo.GetProperty("U_Otherfees"),

                        U_DoPayCode = udo.GetProperty("U_DoPayCode"),
                        U_Lockfee = udo.GetProperty("U_Lockfee"),
                        U_testfee = udo.GetProperty("U_testfee")
                    }
                };
                json_obj.JSON_OBJECT = data.ToJSON();
            }
            else
            {

                json_obj.Message = "Company Not Connected " + oCompany.GetLastErrorDescription();
            }
            #endregion

            // Write here the code that uses objType to get an Object and Key to get its reference.
            // Make json Object and return as string 



            return json_obj;
        }
        public Response AddUpdateObject(JSON_OBJ Object, string SessionID)
        {
            var response = new Response();

            var oCompany = CompaniesSessions[SessionID];
            JavaScriptSerializer serializer = new JavaScriptSerializer();
            dynamic obj = serializer.Deserialize(Object.JSON_OBJECT, typeof(object));
            var objtype = obj["ObjectInfo"]["ObjType"];
            var key = "";
            key = obj["ObjectInfo"]["Key"];


            int errCode = 0;
            string errMsg = "";
            if (objtype == 13)
            {
                try
                {
                    oCompany.StartTransaction();
                    var sum = 0.0;
                    #region add invoice
                    if (string.IsNullOrEmpty(key))
                    {
                        var dc = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices) as SAPbobsCOM.Documents;

                        dc.CardCode = obj["Header"]["CustomerCode"];
                        dc.DocDate = DateTime.ParseExact(obj["Header"]["DocumentDate"], "yyyyMMdd", null);
                        dc.DocDueDate = DateTime.ParseExact(obj["Header"]["DocDueDate"], "yyyyMMdd", null);
                        dc.TaxDate = DateTime.ParseExact(obj["Header"]["TaxDate"], "yyyyMMdd", null);

                        var rows = obj["Rows"];
                        foreach (var item in rows)
                        {
                            var ItemCode = item["ItemCode"];
                            var Quantity = item["Quantity"];
                            var UnitPrice = item["UnitPrice"];
                            var RowDiscountPerc = item["RowDiscountPerc"];
                            var RowTotal = item["RowTotal"];

                            dc.Lines.ItemCode = ItemCode;
                            dc.Lines.Quantity = Quantity;
                            dc.Lines.UnitPrice = UnitPrice;
                            dc.Lines.DiscountPercent = RowDiscountPerc;

                            dc.Lines.Add();

                        }


                        var ans = dc.Add();
                        if (ans != 0)
                        {
                            oCompany.GetLastError(out errCode, out errMsg);
                        }
                        else
                        {
                            errCode = 0;
                            errMsg = "successfully added";
                        }
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(dc);

                        dc = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices) as SAPbobsCOM.Documents;
                       
                        var docentry = Convert.ToInt32(oCompany.GetNewObjectKey());
                        dc.GetByKey(docentry);
                        sum = dc.DocTotal;
                    }
                    #endregion

                    #region if payment keu exist
                    if (obj["PaymentMeans"].Length > 0)
                    {
                        var docentry = Convert.ToInt32(oCompany.GetNewObjectKey());
                        var dc = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments) as SAPbobsCOM.Payments;

                        dc.CardCode = obj["Header"]["CustomerCode"];
                        dc.DocDate = DateTime.ParseExact(obj["Header"]["DocumentDate"], "yyyyMMdd", null);
                        dc.DueDate = DateTime.ParseExact(obj["Header"]["DocDueDate"], "yyyyMMdd", null);
                        dc.TaxDate = DateTime.ParseExact(obj["Header"]["TaxDate"], "yyyyMMdd", null);
                        //dc.CashSum = Convert.ToDouble(obj["Header"]["DocTotal"]);
                        dc.Invoices.DocEntry = docentry;
                        foreach (var item in obj["PaymentMeans"])
                        {

                            dc.CreditCards.CreditType = SAPbobsCOM.BoRcptCredTypes.cr_Regular;
                            dc.CreditCards.CardValidUntil =DateTime.ParseExact ( item["ExpiryDate"],"yyyyMMdd",null );
                            dc.CreditCards.CreditCard = Convert.ToInt32(item["CreditCard"]);
                            dc.CreditCards.CreditCardNumber = item["Last4digits"];
                            dc.CreditCards.CreditSum =  Convert.ToDouble(item["Amount"]);
                            dc.CreditCards.VoucherNum = item["VoucherNumber"];

                            dc.CreditCards.Add();

                        }
                        var i = dc.Add();
                        if (i != 0)
                        {

                            oCompany.GetLastError(out errCode, out errMsg);
                            if (oCompany.InTransaction)
                                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                        }
                        else
                        {

                            if (oCompany.InTransaction)
                                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                        }
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(dc);

                        GC.Collect();

                    }
                    else
                    {

                        if (oCompany.InTransaction)
                            oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    }

                    #endregion

                }
                catch (Exception ex)
                {
                    if (oCompany.InTransaction)
                        oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);

                }
            }
            if (objtype == 24)
            {
                if (string.IsNullOrEmpty(key))
                {
                    try
                    {
                        #region Incoming Payment
                        var dc = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments) as SAPbobsCOM.Payments;

                        dc.CardCode = obj["Header"]["CustomerCode"];
                        dc.DocDate = DateTime.ParseExact(obj["Header"]["DocumentDate"], "yyyyMMdd", null);
                        dc.TaxDate = DateTime.ParseExact(obj["Header"]["TaxDate"], "yyyyMMdd", null);
                        dc.CashSum = Convert.ToDouble(obj["Header"]["DocTotal"]);
                        try
                        {
                            foreach (var item in obj["BaseDocs"])
                            {
                                var DocEntry = item["DocEntry"];
                                dc.Invoices.DocEntry = Convert.ToInt32(DocEntry);
                                dc.Invoices.SumApplied = Convert.ToDouble(item["Amount"]);
                            }
                        }
                        catch (Exception ex) { }
                        if (obj["Header"]["PaymentOnAccount"].ToString().ToLower() == "y")
                        {
                            foreach (var item in obj["PaymentMeans"])
                            {
                                dc.CreditCards.CreditType = SAPbobsCOM.BoRcptCredTypes.cr_Regular;
                                dc.CreditCards.CardValidUntil = DateTime.ParseExact(item["ExpiryDate"], "yyyyMMdd", null);
                                dc.CreditCards.CreditCard = Convert.ToInt32(item["CreditCard"]);
                                dc.CreditCards.CreditCardNumber = item["Last4digits"];
                                dc.CreditCards.CreditSum = Convert.ToDouble(item["Amount"]);
                                dc.CreditCards.VoucherNum = item["VoucherNumber"];
                                dc.CreditCards.Add();

                            }

                        }
                        var path=System.Configuration.ConfigurationManager.AppSettings["Failure_path"];
                        oCompany .XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode ;
                        var xml = dc.GetAsXML();
                       
                        var ans = dc.Add();
                        if (ans != 0)
                        {
                            oCompany.GetLastError(out errCode, out errMsg);
                            try
                            {
                                if (!System.IO.Directory.Exists(path))
                                    System.IO.Directory.CreateDirectory(path);
                                System.IO.File.WriteAllText(System.IO.Path.Combine(path, objtype + "_" + key + ".txt"), Object.JSON_OBJECT);
                                XDocument.Parse(xml).Save(System.IO.Path.Combine(path, objtype + "_" + key + ".xml"));
                                System.IO.File.WriteAllText(System.IO.Path.Combine(path, objtype + "_" + key + "_Error" + ".txt"), errCode + Environment.NewLine + errMsg);
                                
                            }
                            catch { }
                         }
                        else
                        {
                            errCode = 0;
                            errMsg = "successfully added";
                        }
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(dc);
                        GC.Collect();
                        #endregion
                    }catch (Exception ex)
                    {
                        errMsg = ex.Message;
                        errCode = -1;
                    }
                }
            }
            var data = new { errCode = errCode, errMsg = errMsg, Key = oCompany.GetNewObjectKey () };
            response.JSON_Response = data.ToJSON();
            return response;

        }


        public void GetObjectKeys(object ObjType, string SessionID, out System.Collections.Generic.List<object> List)
        {
            List = new List<object>();

            var oCompany = CompaniesSessions[SessionID];
            //var r = oCompany.Connect();
            if (oCompany.Connected)
            {
                #region Company Connected
                if ((int)ObjType != 11)
                {

                    try
                    {
                        object obj = oCompany.GetBusinessObject((SAPbobsCOM.BoObjectTypes)ObjType);
                        try
                        {

                            if ((int)ObjType == (int)SAPbobsCOM.BoObjectTypes.oInvoices || (int)ObjType == (int)SAPbobsCOM.BoObjectTypes.oDownPayments)
                            {
                                #region Get Document
                                var invoice = obj as SAPbobsCOM.Documents;
                                invoice.Browser.MoveFirst();
                                for (int i = 0; i < invoice.Browser.RecordCount; i++)
                                {
                                    List.Add(invoice.DocEntry);

                                    invoice.Browser.MoveNext();
                                }
                                #endregion
                            }
                            if ((int)ObjType == (int)SAPbobsCOM.BoObjectTypes.oItems)
                            {
                                #region Item Master Data
                                var item = obj as SAPbobsCOM.Items;
                                item.Browser.MoveFirst();
                                for (int i = 0; i < item.Browser.RecordCount ; i++)
                                {
                                    List.Add(item.ItemCode);
                                }
                                #endregion
                            }
                            if ((int)ObjType == (int)SAPbobsCOM.BoObjectTypes.oBusinessPartners)
                            {
                                #region BP Master Data
                                var item = obj as SAPbobsCOM.BusinessPartners;
                                item.Browser.MoveFirst();
                                for (int i = 0; i < item.Browser .RecordCount ; i++)
                                {
                                    List.Add(item.CardCode);
                                }

                                #endregion
                            }
                        }
                        finally
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                            GC.Collect();
                            obj = null;
                        }
                    }
                    catch (Exception ex)
                    {
                        //json_obj.Message = ex.Message;
                    }
                }
                if ((int)ObjType == 11)
                {
                    try
                    {

                        #region Contact Person
                        var recset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                        recset.DoQuery(string.Format("Select * from OCPR"));
                        while(!recset.EoF)
                        {

                            var contactCode = recset.Fields.Item("CntctCode").Value;
                            List.Add(contactCode);
                            recset.MoveNext();
                        }
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(recset);
                        GC.Collect();
                        #endregion
                    }

                    catch (Exception ex)
                    {

                        //json_obj.Message = ex.Message;
                    }
                }

                #endregion

            }
            else
            {

                //json_obj.Message = "Company Not Connected " + oCompany.GetLastErrorDescription();
            }

            #region get udo
            //var r = oCompany.Connect();
            if (oCompany.Connected && ObjType is string)
            {
                ComModel.UserDefinedObject udo = new UserDefinedObject(ObjType.ToString(), SAPbobsCOM.BoUDOObjType.boud_MasterData, oCompany);
               
                //udo.GetByKey(Key);
                //// udo.GetProperty ("U_ItemCode");

                //var data = new
                //{
                //    ObjectInfo = new
                //    {
                //        ObjType = ObjType
                //        ,
                //        ObjectName = (ObjType).ToString()
                //        ,
                //        Key = Key
                //    },

                //    Header = new
                //    {

                //        U_ItemCode = udo.GetProperty("U_ItemCode"),
                //        U_Itemcode1 = udo.GetProperty("U_Itemcode1"),
                //        U_ItemCodename = udo.GetProperty("U_ItemCodename"),
                //        U_Itemcode2 = udo.GetProperty("U_Itemcode2"),
                //        U_itemCodename2 = udo.GetProperty("U_itemCodename2"),
                //        U_Otherfees = udo.GetProperty("U_Otherfees"),

                //        U_DoPayCode = udo.GetProperty("U_DoPayCode"),
                //        U_Lockfee = udo.GetProperty("U_Lockfee"),
                //        U_testfee = udo.GetProperty("U_testfee")
                //    }
                //};
                //json_obj.JSON_OBJECT = data.ToJSON();
            }
            else
            {

                //json_obj.Message = "Company Not Connected " + oCompany.GetLastErrorDescription();
            }
            #endregion

        }

        /// <param name="ObjKeys">List of all Keys associated to the object type</param>
        
        //public  void GetObjectKeysList(object ObjType, string SessionID, out List<object> ObjKeys)
        //{
        //    ObjKeys = new List<object>();

        //    var oCompany = CompaniesSessions[SessionID];
        //    //var r = oCompany.Connect();
        //    if (oCompany.Connected)
        //    {
        //        #region Company Connected
        //        if ((int)ObjType != 11)
        //        {

        //            try
        //            {
        //                object obj = oCompany.GetBusinessObject((SAPbobsCOM.BoObjectTypes)ObjType);
        //                try
        //                {

        //                    if ((int)ObjType == (int)SAPbobsCOM.BoObjectTypes.oInvoices || (int)ObjType == (int)SAPbobsCOM.BoObjectTypes.oDownPayments)
        //                    {
        //                        #region Get Document
        //                        var invoice = obj as SAPbobsCOM.Documents;
        //                        invoice.Browser.MoveFirst();
        //                        for (int i = 0; i < invoice.Browser.RecordCount; i++)
        //                        {


        //                            invoice.Browser.MoveNext();
        //                        }
        //                        #endregion
        //                    }
        //                    if ((int)ObjType == (int)SAPbobsCOM.BoObjectTypes.oItems)
        //                    {
        //                        #region Item Master Data
        //                        var item = obj as SAPbobsCOM.Items;
                            
        //                        #endregion
        //                    }
        //                    if ((int)ObjType == (int)SAPbobsCOM.BoObjectTypes.oBusinessPartners)
        //                    {
        //                        #region BP Master Data
        //                        var item = obj as SAPbobsCOM.BusinessPartners;
                              

        //                        #endregion
        //                    }
        //                }
        //                finally
        //                {
        //                    System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
        //                    GC.Collect();
        //                    obj = null;
        //                }
        //            }
        //            catch (Exception ex)
        //            {
        //                //json_obj.Message = ex.Message;
        //            }
        //        }
        //        if ((int)ObjType == 11)
        //        {
        //            try
        //            {

        //                #region Contact Person
        //                var recset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
        //                //recset.DoQuery(string.Format("Select * from OCPR where \"CntctCode\"='{0}'", Key));
        //                //var contactCode = recset.Fields.Item("CntctCode").Value;
        //                //var data = new
        //                //{
        //                //    ObjectInfo = new
        //                //    {
        //                //        ObjType = ObjType
        //                //        ,
        //                //        ObjectName = "ContactPerson"
        //                //        ,
        //                //        Key = Key
        //                //    },

        //                //    Header = new
        //                //    {

        //                //        ID = recset.Fields.Item("CntctCode").Value,
        //                //        ContactPersonID = recset.Fields.Item("CntctCode").Value,
        //                //        FirstName = recset.Fields.Item("FirstName").Value,
        //                //        MiddleName = recset.Fields.Item("MiddleName").Value,
        //                //        LastName = recset.Fields.Item("LastName").Value,
        //                //        Status = recset.Fields.Item("Active").Value


        //                //    }
        //                //};

        //                //json_obj.JSON_OBJECT = data.ToJSON();
        //                //json_obj.OBJECT_Type = "Contact Person";
        //                //System.Runtime.InteropServices.Marshal.ReleaseComObject(recset);
        //                //GC.Collect();
        //                #endregion
        //            }

        //            catch (Exception ex)
        //            {

        //                //json_obj.Message = ex.Message;
        //            }
        //        }

        //        #endregion

        //    }
        //    else
        //    {

        //        //json_obj.Message = "Company Not Connected " + oCompany.GetLastErrorDescription();
        //    }

        //    #region get udo
        //    //var r = oCompany.Connect();
        //    if (oCompany.Connected && ObjType is string)
        //    {
        //        ComModel.UserDefinedObject udo = new UserDefinedObject(ObjType.ToString(), SAPbobsCOM.BoUDOObjType.boud_MasterData, oCompany);
        //        //udo.GetByKey(Key);
        //        //// udo.GetProperty ("U_ItemCode");

        //        //var data = new
        //        //{
        //        //    ObjectInfo = new
        //        //    {
        //        //        ObjType = ObjType
        //        //        ,
        //        //        ObjectName = (ObjType).ToString()
        //        //        ,
        //        //        Key = Key
        //        //    },

        //        //    Header = new
        //        //    {

        //        //        U_ItemCode = udo.GetProperty("U_ItemCode"),
        //        //        U_Itemcode1 = udo.GetProperty("U_Itemcode1"),
        //        //        U_ItemCodename = udo.GetProperty("U_ItemCodename"),
        //        //        U_Itemcode2 = udo.GetProperty("U_Itemcode2"),
        //        //        U_itemCodename2 = udo.GetProperty("U_itemCodename2"),
        //        //        U_Otherfees = udo.GetProperty("U_Otherfees"),

        //        //        U_DoPayCode = udo.GetProperty("U_DoPayCode"),
        //        //        U_Lockfee = udo.GetProperty("U_Lockfee"),
        //        //        U_testfee = udo.GetProperty("U_testfee")
        //        //    }
        //        //};
        //        //json_obj.JSON_OBJECT = data.ToJSON();
        //    }
        //    else
        //    {

        //        //json_obj.Message = "Company Not Connected " + oCompany.GetLastErrorDescription();
        //    }
        //    #endregion

        //}
    }
}

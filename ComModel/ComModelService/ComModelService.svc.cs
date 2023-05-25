using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;
using System.Web.SessionState;
using ExtensionMethods;
using EncryptDecrypt;
using System.Web.Script.Serialization;
using System.Runtime.InteropServices;

namespace ComModelService
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the class name "Service1" in code, svc and config file together.
    // NOTE: In order to launch WCF Test Client for testing this service, please select Service1.svc or Service1.svc.cs at the Solution Explorer and start debugging.
    public class ComModelService : IComModelService
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
                // optmize the number of connection, if the company is connected return the current  SessionID
                if (Convert.ToBoolean(System.Configuration.ConfigurationManager.AppSettings["OptmizeConnection"]))
                {

                    var Compnay = CompaniesSessions.Select(x => new { x.Key, x.Value.UserName }).Where(y => y.UserName == UserName).ToList();
                    if (Compnay.Count > 0)
                    {
                        if (CompaniesSessions[Compnay[0].Key].Connected)
                        {
                            string TempSessionID = Compnay[0].Key.DecryptMe("asdiiie1244332434099r#$");
                            if (TempSessionID.Split(new string[] { "##" }, StringSplitOptions.None)[1] == Password)
                            {
                                SessionID = Compnay[0].Key;
                                return;
                            }
                        }
                    }
                } 
            }
            catch (Exception ex)
            {
                
            }
            try
            {

                SAPbobsCOM.Company _Company = new SAPbobsCOM.Company();
                _Company.Server = System.Configuration.ConfigurationManager.AppSettings["Server"];
                _Company.LicenseServer = System.Configuration.ConfigurationManager.AppSettings["LicenseServer"];
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
                    //SessionIDManager manager = new SessionIDManager();
                    count++;
                    var count1 = rand.Next().ToString() + count.ToString()+"##"+Password;
                    SessionID = count1.ToString().EncryptMe("asdiiie1244332434099r#$");
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
        public bool LogOut(String SessionID)
        {
            bool lout = true;
            CompaniesSessions[SessionID].Disconnect();
            lout = !CompaniesSessions[SessionID].Connected;
            CompaniesSessions.Remove(SessionID);
            return lout;
        }

        /// <summary>
        /// Get Object With The Key
        /// </summary>
        /// <param name="Key"></param>
        /// <param name="ObjType">Integer value of SAPbobsCOM.BoObjectTypes</param>
        /// <returns></returns>
        public JSON_OBJ GetObjectByKey(String Key, string  ObjType, string SessionID)
        {
            var json_obj = new JSON_OBJ();
            int value;
                var oCompany = CompaniesSessions[SessionID];
            if (int.TryParse(ObjType, out value))
            {


                var ObjType2 = value;
                //var r = oCompany.Connect();
                if (oCompany.Connected)
                {
                    #region Company Connected
                    //  if ((int)ObjType != 11)
                    {

                        try
                        {
                            object obj = oCompany.GetBusinessObject((SAPbobsCOM.BoObjectTypes)ObjType2);
                            try
                            {

                                if ((int)ObjType2 == (int)SAPbobsCOM.BoObjectTypes.oInvoices || (int)ObjType2 == (int)SAPbobsCOM.BoObjectTypes.oDownPayments)
                                {
                                    #region Get Document
                                    var invoice = obj as SAPbobsCOM.Documents;
                                    invoice.GetByKey(Convert.ToInt32(Key));
                                    var recset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                                    //recset.DoQuery(string.Format("SELECT \"Tel1\" , \"Name\", \"Address\",\"Fax\" from ocpr where \"CntctCode\" = '{0}' ", invoice.ContactPersonCode));
                                    recset.DoQuery(string.Format("SELECT     T2.\"Cellolar\",   T2.\"Tel1\",T2.\"Tel2\", T2.\"Name\",T4.\"Street\", T4.\"Building\",T4.\"ZipCode\",T5.\"Name\" Country,T4.\"City\"  , T2.\"Fax\", T3.\"GroupName\", T0.\"CreditLine\", T0.\"DebtLine\" FROM \"OCPR\" AS T2 INNER JOIN \"OCRD\" AS T0 ON T2.\"CardCode\" = T0.\"CardCode\" INNER JOIN \"OCRG\" AS T3 ON T0.\"GroupCode\" = T3.\"GroupCode\"  INNER JOIN \"CRD1\" T4 ON T0.\"CardCode\" = T4.\"CardCode\" INNER JOIN \"OCRY\" AS T5 ON T4.Country = T5.\"Code\" where \"CntctCode\" = '{0}' ", invoice.ContactPersonCode));
                                    var contactpersoncode = recset.Fields.Item("Cellolar").Value;
                                    var contactpersonName = recset.Fields.Item("Name").Value;
                                    var contactpersonAddr = string.Format("{0} , ZIP:{1} , {2} , {3}", recset.Fields.Item("Street").Value, recset.Fields.Item("ZipCode").Value, recset.Fields.Item("City").Value, recset.Fields.Item("Country").Value);
                                    var contactpersonFax = recset.Fields.Item("Tel1").Value == "" ? recset.Fields.Item("Tel2").Value : recset.Fields.Item("Tel1").Value;
                                    var CustGroupName = recset.Fields.Item("GroupName").Value;
                                    var CustCreditLimit = recset.Fields.Item("CreditLine").Value;
                                    var CustCommitmentLimit = recset.Fields.Item("DebtLine").Value;
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(recset);
                                    GC.Collect();

                                    var row = new List<object>();
                                    var lines = invoice.Lines;
                                    lines.SetCurrentLine(0);
                                    var count = invoice.Lines.Count;
                                    for (int i = 0; i < count; i++)
                                    {
                                        lines.SetCurrentLine(i);
                                        var item = new { ItemCode = lines.ItemCode, ItemName = lines.ItemDescription,
                                            //DocumentDate = invoice.DocDate.ToString("yyyyMMdd"),//
                                            // DueDate = invoice.DocDueDate.ToString("yyyyMMdd"),//
                                            //Details = lines.ItemDetails,//
                                            Quantity = lines.Quantity, UnitPrice = lines.Price, RowDiscountPerc = lines.DiscountPercent, RowTotal = lines.LineTotal,
                                            //Balance=lines.OpenAmount//
                                        };
                                        row.Add(item);
                                    }
                                    var data = new
                                    {
                                        ObjectInfo = new
                                        {
                                            ObjType = ObjType2
                                            ,
                                            ObjectName = ((SAPbobsCOM.BoObjectTypes)ObjType2).ToString()
                                            ,
                                            Key = Key
                                        },

                                        Header = new
                                        {

                                            CustomerCode = invoice.CardCode,
                                            CustomerName = invoice.CardName,
                                            CustomerGroup=CustGroupName,//
                                            ContactNumber = contactpersoncode,

                                            FaxNumber=contactpersonFax,//
                                            Address=contactpersonAddr,//
                                            ContactPerson=contactpersonName,//
                                            CreditLimit=CustCreditLimit,//
                                            CommitmentLimit=CustCommitmentLimit,//

                                            DocumentDate = invoice.DocDate.ToString("yyyyMMdd"),
                                            DueDate = invoice.DocDueDate.ToString("yyyyMMdd"),

                                            DueAmount = invoice.DocTotal, //

                                            Status = invoice.DocumentStatus
                                        }
                                        ,
                                        Rows = row
                                    };
                                    json_obj.JSON_OBJECT = data.ToJSON();
                                    json_obj.OBJECT_Type = ((SAPbobsCOM.BoObjectTypes)ObjType2).ToString();

                                    #endregion
                                }
                                if ((int)ObjType2 == (int)SAPbobsCOM.BoObjectTypes.oItems)
                                {
                                    #region Item Master Data
                                    var item = obj as SAPbobsCOM.Items;
                                    item.GetByKey(Key);
                                    var st = "";
                                    if (item.Frozen == SAPbobsCOM.BoYesNoEnum.tYES) st = "Active";
                                    if (item.Frozen == SAPbobsCOM.BoYesNoEnum.tNO) st = "Inctive";


                                    var data = new
                                    {
                                        ObjectInfo = new
                                        {
                                            ObjType = ObjType2
                                            ,
                                            ObjectName = ((SAPbobsCOM.BoObjectTypes)ObjType2).ToString()
                                            ,
                                            Key = Key
                                        },

                                        Header = new
                                        {

                                            Code = item.ItemCode,
                                            Description = item.ItemName,
                                            Status = st
                                        }
                                    };
                                    json_obj.JSON_OBJECT = data.ToJSON();
                                    json_obj.OBJECT_Type = "Item Master Data";

                                    #endregion
                                }
                                if ((int)ObjType2 == (int)SAPbobsCOM.BoObjectTypes.oBusinessPartners)
                                {
                                    #region BP Master Data
                                    var item = obj as SAPbobsCOM.BusinessPartners;
                                    item.GetByKey(Key);

                                    var data = new
                                    {
                                        ObjectInfo = new
                                        {
                                            ObjType = ObjType2
                                            ,
                                            ObjectName = ((SAPbobsCOM.BoObjectTypes)ObjType2).ToString()
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
                    //if (ObjType is int)
                    {
                        if ((int)ObjType2 == 11)
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
                                        ObjType = ObjType2
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
                        else
                            if ((int)ObjType2 == -203)//Inventory Transfer Request
                        {
                            #region Inventory Transfer Request
                            var row = new List<object>();

                            var recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                            recordset.DoQuery(string.Format("select * from odpi where \"CreateTran\"  = 'N' and \"DocEntry\" = '{0}'", Key));
                                                   
                            var recset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                            recset.DoQuery(string.Format("SELECT \"Tel1\" , \"Name\", \"Address\",\"Fax\" from ocpr where \"CntctCode\" = '{0}' ", recordset.Fields.Item("CntctCode").Value));

                            var contactpersoncode = recset.Fields.Item("Tel1").Value;
                            var contactpersonName = recset.Fields.Item("Name").Value;
                            var contactpersonAddr = recset.Fields.Item("Address").Value;
                            var contactpersonFax = recset.Fields.Item("Fax").Value;

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(recset);
                            
                            var recordset1 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                            recordset1.DoQuery(string.Format("Select * from dpi1 where \"DocEntry\" = '{0}'", Key));
                            while (!recordset1.EoF)
                            {
                                var item = new { ItemCode = recordset1.Fields.Item("ItemCode").Value, ItemName = recordset1.Fields.Item("Dscription").Value, Quantity = recordset1.Fields.Item("Quantity").Value, UnitPrice = recordset1.Fields.Item("Price").Value, RowDiscountPerc = recordset1.Fields.Item("DiscPrcnt").Value, RowTotal = recordset1.Fields.Item("LineTotal").Value };
                                row.Add(item);
                                recordset1.MoveNext();
                            }
                            Marshal.ReleaseComObject(recordset1);

                            var data = new
                            {
                                ObjectInfo = new
                                {
                                    ObjType = ObjType2
                                    ,
                                    ObjectName = "AR Invoice Payment Request"
                                    ,
                                    Key = Key
                                },

                                Header = new
                                {

                                    CustomerCode = recordset.Fields.Item("CardCode").Value,
                                    CustomerName = recordset.Fields.Item("CardName").Value,
                                    ContactNumber = contactpersoncode,
                                    DocumentDate = recordset.Fields.Item("DocDate").Value.ToString("yyyyMMdd"),
                                    DueDate = recordset.Fields.Item("DocDueDate").Value.ToString("yyyyMMdd"),
                                    Status = recordset.Fields.Item("DocStatus").Value
                                }
                                ,
                                Rows = row
                            };
                            json_obj.JSON_OBJECT = data.ToJSON();
                            json_obj.OBJECT_Type = ((SAPbobsCOM.BoObjectTypes)ObjType2).ToString();

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(recordset);

                            #endregion
                        }
                    }

                    #endregion

                }
                else
                {

                    json_obj.Message = "Company Not Connected " + oCompany.GetLastErrorDescription();
                }
            }
            else
            {


                #region get udo
                //var r = oCompany.Connect();
                if (oCompany.Connected )
                {
                    ComModel.UserDefinedObject udo = new ComModel.UserDefinedObject(ObjType.ToString(), SAPbobsCOM.BoUDOObjType.boud_MasterData, oCompany);
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
            }

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
                    #region UpdateInvoice
                    if (!string.IsNullOrEmpty(key))
                    {
                        var dc = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices) as SAPbobsCOM.Documents;
                        dc.GetByKey(Convert.ToInt32(key));
                        if (obj["ObjectInfo"]["Cancel"] == "Y")
                        {
                            var cd = dc.CreateCancellationDocument();
                            errCode = cd.Add();
                            if (errCode != 0)
                                errMsg = oCompany.GetLastErrorDescription();
                            else
                                errMsg = "Success";
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(cd);
                        }


                        System.Runtime.InteropServices.Marshal.ReleaseComObject(dc);
                        if (oCompany.InTransaction)
                            oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
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
                            dc.CreditCards.CardValidUntil = DateTime.ParseExact(item["ExpiryDate"], "yyyyMMdd", null);
                            dc.CreditCards.CreditCard = Convert.ToInt32(item["CreditCard"]);
                            dc.CreditCards.CreditCardNumber = item["Last4digits"];
                            dc.CreditCards.CreditSum = Convert.ToDouble(item["Amount"]);
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
                    errCode = oCompany.GetLastErrorCode();
                    errMsg =ex.Message + " " + ex.StackTrace ;
                }
            }
            else
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
                            //dc.CashSum = Convert.ToDouble(obj["Header"]["DocTotal"]); this put the invoice value as cach payment while they take Credit Card only
                            dc.CashSum = 0;
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
                            GC.Collect();
                            #endregion
                        }
                        catch (Exception ex)
                        {
                            errCode = oCompany.GetLastErrorCode();
                            errMsg = ex.Message + " " + ex.StackTrace;
                        }
                    }
                }
                else
                    if (objtype == 2)
                    {
                        if (!string.IsNullOrEmpty(key))
                        {
                            try
                            {
                                var customer = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners) as SAPbobsCOM.BusinessPartners;
                                customer.GetByKey(key);
                                try { customer.CardName = obj["Header"]["CustomerName"]; }
                                catch { }

                                try { customer.Frozen = (SAPbobsCOM.BoYesNoEnum)Convert.ToInt32(obj["Header"]["Status"]); }
                                catch { }
                                errCode = customer.Update();
                                if (errCode != 0)
                                {
                                    errMsg = oCompany.GetLastErrorDescription();
                                }
                                else
                                    errMsg = "Success";
                                Marshal.ReleaseComObject(customer);
                            }
                            catch (Exception ex)
                            {
                                errCode = oCompany.GetLastErrorCode();
                                errMsg = ex.Message + " " + ex.StackTrace;
                            }
                        }
                    }
                    else
                        if (objtype == 4)
                        {
                            if (!string.IsNullOrEmpty(key))
                            {
                                try
                                {
                                    var oitm = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems) as SAPbobsCOM.Items;
                                    oitm.GetByKey(key);
                                    try { oitm.ItemName = obj["Header"]["ItemName"]; }
                                    catch { }

                                    try { oitm.Frozen = (SAPbobsCOM.BoYesNoEnum)Convert.ToInt32(obj["Header"]["Status"]); }
                                    catch { }
                                    errCode = oitm.Update();
                                    if (errCode != 0)
                                    {
                                        errMsg = oCompany.GetLastErrorDescription();
                                    }
                                    else
                                        errMsg = "Success";
                                    Marshal.ReleaseComObject(oitm);
                                }
                                catch (Exception ex)
                                {
                                    errCode = oCompany.GetLastErrorCode();
                                    errMsg = ex.Message + " " + ex.StackTrace;
                                }
                            }
                        }
            var data = new { errCode = errCode, errMsg = errMsg, Key = oCompany.GetNewObjectKey() };
            response.JSON_Response = data.ToJSON();
            return response;

        }

        public JSON_OBJ GetObjectBulkonDate(string FieldName, string FieldValue, string StartDate, string EndDate,string status, string ObjType, string SessionID)
        {

            var json_obj = new JSON_OBJ();
            List<string> Keys = new List<string>();



            int value;
            var oCompany = CompaniesSessions[SessionID];
            if (int.TryParse(ObjType, out value))
            {

                var ObjType2 = value;
                Dictionary<string, dynamic> Data_Dictionary = new Dictionary<string, dynamic>();
                {
                    var TableName = "";
                    var statusName = "";
                    if ((int)ObjType2 == (int)SAPbobsCOM.BoObjectTypes.oInvoices)
                    {
                        TableName = "OINV";
                        statusName = "DocStatus";
                    }
                    if ((int)ObjType2 == (int)SAPbobsCOM.BoObjectTypes.oDownPayments)
                    {
                        TableName = "ODPI";
                        statusName = "DocStatus";
                    }
                    if ((int)ObjType2 == (int)SAPbobsCOM.BoObjectTypes.oItems)
                    {
                        TableName = "OITM";
                        statusName = "FrozenFor";
                    }
                    if ((int)ObjType2 == (int)SAPbobsCOM.BoObjectTypes.oBusinessPartners)
                    {
                        TableName = "OCRD";
                        statusName = "FrozenFor";
                    }
                    if ((int)ObjType2 == 11)
                    {
                        

                        TableName = "OCPR";
                        statusName = "Active";
                    }
                    var query = string.Format("Select * from {2} where {0}='{1}' and \"CreateDate\" >= '{3}' and \"CreateDate\" <= '{4}' and \"{5}\" = '{6}'", FieldName, FieldValue, TableName,StartDate ,EndDate,statusName,status);
                    var recset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                    recset.DoQuery(query);
                    while (!recset.EoF)
                    {
                        string docentry = ((object)recset.Fields.Item(0).Value).ToString();
                        Keys.Add(docentry);
                        recset.MoveNext();
                    }
                    Marshal.ReleaseComObject(recset);
                }

                // Define the Contact person data to avoid redundant sql call
                string contactpersoncodeID = "";
                string contactpersonName = "";
                string contactpersonAddr = "";
                string contactpersonFax = "";
                string CustGroupName = "";
                double CustCreditLimit = 0;
                double CustCommitmentLimit = 0;

                foreach (string Key in Keys)
                {


                    //var r = oCompany.Connect();
                    if (oCompany.Connected)
                    {
                        #region Company Connected
                        //  if ((int)ObjType != 11)
                        {

                            try
                            {
                                object obj = oCompany.GetBusinessObject((SAPbobsCOM.BoObjectTypes)ObjType2);
                                try
                                {

                                    if ((int)ObjType2 == (int)SAPbobsCOM.BoObjectTypes.oInvoices || (int)ObjType2 == (int)SAPbobsCOM.BoObjectTypes.oDownPayments)
                                    {
                                        #region Get Document
                                        var invoice = obj as SAPbobsCOM.Documents;
                                        invoice.GetByKey(Convert.ToInt32(Key));
                                        if (contactpersoncodeID == "")
                                        {
                                            var recset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                                            //recset.DoQuery(string.Format("SELECT \"Tel1\" from ocpr where \"CntctCode\" = '{0}' ", invoice.ContactPersonCode));
                                            //var contactpersoncode = recset.Fields.Item("Tel1").Value;
                                            recset.DoQuery(string.Format("SELECT     T2.\"Cellolar\",   T2.\"Tel1\",T2.\"Tel2\", T2.\"Name\",T4.\"Street\", T4.\"Building\",T4.\"ZipCode\",T5.\"Name\" Country,T4.\"City\"  , T2.\"Fax\", T3.\"GroupName\", T0.\"CreditLine\", T0.\"DebtLine\" FROM \"OCPR\" AS T2 INNER JOIN \"OCRD\" AS T0 ON T2.\"CardCode\" = T0.\"CardCode\" INNER JOIN \"OCRG\" AS T3 ON T0.\"GroupCode\" = T3.\"GroupCode\"  INNER JOIN \"CRD1\" T4 ON T0.\"CardCode\" = T4.\"CardCode\" INNER JOIN \"OCRY\" AS T5 ON T4.Country = T5.\"Code\" where \"CntctCode\" = '{0}' ", invoice.ContactPersonCode));
                                            contactpersoncodeID = recset.Fields.Item("Cellolar").Value;
                                            contactpersonName = recset.Fields.Item("Name").Value;
                                            contactpersonAddr = string.Format("{0} , ZIP:{1} , {2} , {3}", recset.Fields.Item("Street").Value, recset.Fields.Item("ZipCode").Value, recset.Fields.Item("City").Value, recset.Fields.Item("Country").Value);
                                            contactpersonFax = recset.Fields.Item("Tel1").Value == "" ? recset.Fields.Item("Tel2").Value : recset.Fields.Item("Tel1").Value;
                                            CustGroupName = recset.Fields.Item("GroupName").Value;
                                            CustCreditLimit = recset.Fields.Item("CreditLine").Value;
                                            CustCommitmentLimit = recset.Fields.Item("DebtLine").Value;

                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(recset);
                                            GC.Collect();
                                        }
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
                                                ObjType = ObjType2
                                                ,
                                                ObjectName = ((SAPbobsCOM.BoObjectTypes)ObjType2).ToString()
                                                ,
                                                Key = Key
                                            },

                                            Header = new
                                            {

                                                CustomerCode = invoice.CardCode,
                                                CustomerName = invoice.CardName,
                                                CustomerGroup = CustGroupName,//

                                                ContactNumber = contactpersoncodeID,

                                                FaxNumber = contactpersonFax,//
                                                Address = contactpersonAddr,//
                                                ContactPerson = contactpersonName,//
                                                CreditLimit = CustCreditLimit,//
                                                CommitmentLimit = CustCommitmentLimit,//

                                                DocumentDate = invoice.DocDate.ToString("yyyyMMdd"),
                                                DueDate = invoice.DocDueDate.ToString("yyyyMMdd"),

                                                DueAmount = invoice.DocTotal, //
                                                Status = invoice.DocumentStatus
                                            }
                                            ,
                                            Rows = row
                                        };
                                        //  json_obj.JSON_OBJECT = data.ToJSON();
                                        // json_obj.OBJECT_Type = ((SAPbobsCOM.BoObjectTypes)ObjType2).ToString();

                                        Data_Dictionary.Add(Key, data.ToJSON());
                                        #endregion
                                    }
                                    if ((int)ObjType2 == (int)SAPbobsCOM.BoObjectTypes.oItems)
                                    {
                                        #region Item Master Data
                                        var item = obj as SAPbobsCOM.Items;
                                        item.GetByKey(Key);
                                        var st = "";
                                        if (item.Frozen == SAPbobsCOM.BoYesNoEnum.tYES) st = "Active";
                                        if (item.Frozen == SAPbobsCOM.BoYesNoEnum.tNO) st = "Inctive";


                                        var data = new
                                        {
                                            ObjectInfo = new
                                            {
                                                ObjType = ObjType2
                                                ,
                                                ObjectName = ((SAPbobsCOM.BoObjectTypes)ObjType2).ToString()
                                                ,
                                                Key = Key
                                            },

                                            Header = new
                                            {

                                                Code = item.ItemCode,
                                                Description = item.ItemName,
                                                Status = st
                                            }
                                        };
                                        //   json_obj.JSON_OBJECT = data.ToJSON();
                                        json_obj.OBJECT_Type = "Item Master Data";

                                        Data_Dictionary.Add(Key, data.ToJSON());
                                        #endregion
                                    }
                                    if ((int)ObjType2 == (int)SAPbobsCOM.BoObjectTypes.oBusinessPartners)
                                    {
                                        #region BP Master Data
                                        var item = obj as SAPbobsCOM.BusinessPartners;
                                        item.GetByKey(Key);

                                        var data = new
                                        {
                                            ObjectInfo = new
                                            {
                                                ObjType = ObjType2
                                                ,
                                                ObjectName = ((SAPbobsCOM.BoObjectTypes)ObjType2).ToString()
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
                                        //   json_obj.JSON_OBJECT = data.ToJSON();
                                        json_obj.OBJECT_Type = "BP Master Data";

                                        Data_Dictionary.Add(Key, data.ToJSON());
                                        #endregion
                                    }

                                    json_obj.JSON_OBJECT = Data_Dictionary.ToJSON();
                                    //json_obj.OBJECT_Type = "Dictionary";
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
                        //if (ObjType is int)
                        {
                            if ((int)ObjType2 == 11)
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
                                            ObjType = ObjType2
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

                                    //json_obj.JSON_OBJECT = data.ToJSON();
                                    json_obj.OBJECT_Type = "Contact Person";

                                    Data_Dictionary.Add(Key, data.ToJSON());
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(recset);
                                    GC.Collect();
                                    #endregion
                                }

                                catch (Exception ex)
                                {
                                    json_obj.Message = ex.Message;
                                }
                            }
                            else
                                if ((int)ObjType2 == -203)//Inventory Transfer Request
                                {
                                    #region Inventory Transfer Request
                                    var row = new List<object>();

                                    var recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                                    recordset.DoQuery(string.Format("select * from odpi where \"CreateTran\"  = 'N' and \"DocEntry\" = '{0}'", Key));
                                    var contactpersoncode = "";
                                    {
                                        var recset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                                        recset.DoQuery(string.Format("SELECT \"Tel1\" from ocpr where \"CntctCode\" = '{0}' ", recordset.Fields.Item("CntctCode").Value));
                                        contactpersoncode = recset.Fields.Item("Tel1").Value;
                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(recset);
                                    }
                                    var recordset1 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                                    recordset1.DoQuery(string.Format("Select * from dpi1 where \"DocEntry\" = '{0}'", Key));
                                    while (!recordset1.EoF)
                                    {
                                        var item = new { ItemCode = recordset1.Fields.Item("ItemCode").Value, ItemName = recordset1.Fields.Item("Dscription").Value, Quantity = recordset1.Fields.Item("Quantity").Value, UnitPrice = recordset1.Fields.Item("Price").Value, RowDiscountPerc = recordset1.Fields.Item("DiscPrcnt").Value, RowTotal = recordset1.Fields.Item("LineTotal").Value };
                                        row.Add(item);
                                        recordset1.MoveNext();
                                    }
                                    Marshal.ReleaseComObject(recordset1);

                                    var data = new
                                    {
                                        ObjectInfo = new
                                        {
                                            ObjType = ObjType2
                                            ,
                                            ObjectName = "AR Invoice Payment Request"
                                            ,
                                            Key = Key
                                        },

                                        Header = new
                                        {

                                            CustomerCode = recordset.Fields.Item("CardCode").Value,
                                            CustomerName = recordset.Fields.Item("CardName").Value,
                                            ContactNumber = contactpersoncode,
                                            DocumentDate = recordset.Fields.Item("DocDate").Value.ToString("yyyyMMdd"),
                                            DueDate = recordset.Fields.Item("DocDueDate").Value.ToString("yyyyMMdd"),
                                            Status = recordset.Fields.Item("DocStatus").Value
                                        }
                                        ,
                                        Rows = row
                                    };
                                    // json_obj.JSON_OBJECT = data.ToJSON();


                                    Data_Dictionary.Add(Key, data.ToJSON());
                                    json_obj.OBJECT_Type = ((SAPbobsCOM.BoObjectTypes)ObjType2).ToString();

                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(recordset);

                                    #endregion
                                }
                        }
                        json_obj.JSON_OBJECT = Data_Dictionary.ToJSON();
                        #endregion

                    }
                    else
                    {

                        json_obj.Message = "Company Not Connected " + oCompany.GetLastErrorDescription();
                    }
                }
            }


            return json_obj;

        }
  
     public    JSON_OBJ GetObjectBulk(string FieldName, string FieldValue, string ObjType, string SessionID)
        {

            var json_obj = new JSON_OBJ();
            List<string> Keys = new List<string>();


         
            int value; 
            var oCompany = CompaniesSessions[SessionID];
            if (int.TryParse(ObjType, out value))
            {

                var ObjType2 = value;
                Dictionary<string,dynamic> Data_Dictionary = new Dictionary<string,dynamic>();
                {
                    var TableName = "";
                    if ((int)ObjType2 == (int)SAPbobsCOM.BoObjectTypes.oInvoices)
                    {
                        TableName = "OINV";
                    }
                    if ((int)ObjType2 == (int)SAPbobsCOM.BoObjectTypes.oDownPayments)
                    {
                        TableName = "ODPI";
                    }
                    if ((int)ObjType2 == (int)SAPbobsCOM.BoObjectTypes.oItems)
                    {
                        TableName = "OITM";
                    }
                    if ((int)ObjType2 == (int)SAPbobsCOM.BoObjectTypes.oBusinessPartners)
                    {
                        TableName = "OCRD";
                    }
                    if ((int)ObjType2 == 11)
                    {

                        TableName = "OCPR";
                    }
                        var query = string.Format("Select * from {2} where {0}='{1}'", FieldName, FieldValue, TableName);
                    var recset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                    recset.DoQuery(query);
                    while (!recset.EoF)
                    {
                        string  docentry = ((object)recset.Fields.Item(0).Value).ToString ();
                        Keys.Add(docentry);
                        recset.MoveNext();
                    }
                    Marshal.ReleaseComObject(recset);
                }

                // Define the Contact person data to avoid redundant sql call
                string contactpersoncodeID = "";
                string contactpersonName = "";
                string contactpersonAddr = "";
                string contactpersonFax = "";
                string CustGroupName = "";
                double CustCreditLimit = 0;
                double CustCommitmentLimit = 0 ;

                foreach (string Key in Keys)
                {


                    //var r = oCompany.Connect();
                    if (oCompany.Connected)
                    {
                        #region Company Connected
                        //  if ((int)ObjType != 11)
                        {

                            try
                            {
                                object obj = oCompany.GetBusinessObject((SAPbobsCOM.BoObjectTypes)ObjType2);
                                try
                                {

                                    if ((int)ObjType2 == (int)SAPbobsCOM.BoObjectTypes.oInvoices || (int)ObjType2 == (int)SAPbobsCOM.BoObjectTypes.oDownPayments)
                                    {
                                        #region Get Document
                                        var invoice = obj as SAPbobsCOM.Documents;
                                        invoice.GetByKey(Convert.ToInt32(Key));
                                        
                                        //recset.DoQuery(string.Format("SELECT \"Tel1\" from ocpr where \"CntctCode\" = '{0}' ", invoice.ContactPersonCode));
                                        //var contactpersoncode = recset.Fields.Item("Tel1").Value;
                                        if(contactpersoncodeID == "")
                                        {
                                            var recset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                                            recset.DoQuery(string.Format("SELECT     T2.\"Cellolar\",   T2.\"Tel1\",T2.\"Tel2\", T2.\"Name\",T4.\"Street\", T4.\"Building\",T4.\"ZipCode\",T5.\"Name\" Country,T4.\"City\"  , T2.\"Fax\", T3.\"GroupName\", T0.\"CreditLine\", T0.\"DebtLine\" FROM \"OCPR\" AS T2 INNER JOIN \"OCRD\" AS T0 ON T2.\"CardCode\" = T0.\"CardCode\" INNER JOIN \"OCRG\" AS T3 ON T0.\"GroupCode\" = T3.\"GroupCode\"  INNER JOIN \"CRD1\" T4 ON T0.\"CardCode\" = T4.\"CardCode\" INNER JOIN \"OCRY\" AS T5 ON T4.Country = T5.\"Code\" where \"CntctCode\" = '{0}' ", invoice.ContactPersonCode));
                                            contactpersoncodeID = recset.Fields.Item("Cellolar").Value;
                                            contactpersonName = recset.Fields.Item("Name").Value;
                                            contactpersonAddr = string.Format("{0} , ZIP:{1} , {2} , {3}", recset.Fields.Item("Street").Value, recset.Fields.Item("ZipCode").Value, recset.Fields.Item("City").Value, recset.Fields.Item("Country").Value);
                                            contactpersonFax = recset.Fields.Item("Tel1").Value == "" ? recset.Fields.Item("Tel2").Value : recset.Fields.Item("Tel1").Value;
                                            CustGroupName = recset.Fields.Item("GroupName").Value;
                                            CustCreditLimit = recset.Fields.Item("CreditLine").Value;
                                            CustCommitmentLimit = recset.Fields.Item("DebtLine").Value;

                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(recset);
                                            GC.Collect();
                                        }


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
                                                ObjType = ObjType2
                                                ,
                                                ObjectName = ((SAPbobsCOM.BoObjectTypes)ObjType2).ToString()
                                                ,
                                                Key = Key
                                            },

                                            Header = new
                                            {

                                                CustomerCode = invoice.CardCode,
                                                CustomerName = invoice.CardName,

                                                CustomerGroup = CustGroupName,//

                                                ContactNumber = contactpersoncodeID,

                                                FaxNumber = contactpersonFax,//
                                                Address = contactpersonAddr,//
                                                ContactPerson = contactpersonName,//
                                                CreditLimit = CustCreditLimit,//
                                                CommitmentLimit = CustCommitmentLimit,//

                                                DocumentDate = invoice.DocDate.ToString("yyyyMMdd"),
                                                DueDate = invoice.DocDueDate.ToString("yyyyMMdd"),

                                                DueAmount = invoice.DocTotal, //

                                                Status = invoice.DocumentStatus
                                            }
                                            ,
                                            Rows = row
                                        };
                                      //  json_obj.JSON_OBJECT = data.ToJSON();
                                       // json_obj.OBJECT_Type = ((SAPbobsCOM.BoObjectTypes)ObjType2).ToString();

                                        Data_Dictionary.Add(Key, data.ToJSON());
                                        #endregion
                                    }
                                    if ((int)ObjType2 == (int)SAPbobsCOM.BoObjectTypes.oItems)
                                    {
                                        #region Item Master Data
                                        var item = obj as SAPbobsCOM.Items;
                                        item.GetByKey(Key);
                                        var st = "";
                                        if (item.Frozen == SAPbobsCOM.BoYesNoEnum.tYES) st = "Active";
                                        if (item.Frozen == SAPbobsCOM.BoYesNoEnum.tNO) st = "Inctive";


                                        var data = new
                                        {
                                            ObjectInfo = new
                                            {
                                                ObjType = ObjType2
                                                ,
                                                ObjectName = ((SAPbobsCOM.BoObjectTypes)ObjType2).ToString()
                                                ,
                                                Key = Key
                                            },

                                            Header = new
                                            {

                                                Code = item.ItemCode,
                                                Description = item.ItemName,
                                                Status = st
                                            }
                                        };
                                     //   json_obj.JSON_OBJECT = data.ToJSON();
                                        json_obj.OBJECT_Type = "Item Master Data";

                                        Data_Dictionary.Add(Key, data.ToJSON());
                                        #endregion
                                    }
                                    if ((int)ObjType2 == (int)SAPbobsCOM.BoObjectTypes.oBusinessPartners)
                                    {
                                        #region BP Master Data
                                        var item = obj as SAPbobsCOM.BusinessPartners;
                                        item.GetByKey(Key);

                                        var data = new
                                        {
                                            ObjectInfo = new
                                            {
                                                ObjType = ObjType2
                                                ,
                                                ObjectName = ((SAPbobsCOM.BoObjectTypes)ObjType2).ToString()
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
                                     //   json_obj.JSON_OBJECT = data.ToJSON();
                                      json_obj.OBJECT_Type = "BP Master Data";

                                        Data_Dictionary.Add(Key, data.ToJSON());
                                        #endregion
                                    }
                                    
                                    json_obj.JSON_OBJECT = Data_Dictionary.ToJSON();
                                    //json_obj.OBJECT_Type = "Dictionary";
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
                        //if (ObjType is int)
                        {
                            if ((int)ObjType2 == 11)
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
                                            ObjType = ObjType2
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

                                    //json_obj.JSON_OBJECT = data.ToJSON();
                                    json_obj.OBJECT_Type = "Contact Person";

                                    Data_Dictionary.Add(Key, data.ToJSON());
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(recset);
                                    GC.Collect();
                                    #endregion
                                }

                                catch (Exception ex)
                                {
                                    json_obj.Message = ex.Message;
                                }
                            }
                            else
                                if ((int)ObjType2 == -203)//Inventory Transfer Request
                            {
                                #region Inventory Transfer Request
                                var row = new List<object>();

                                var recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                                recordset.DoQuery(string.Format("select * from odpi where \"CreateTran\"  = 'N' and \"DocEntry\" = '{0}'", Key));
                                var contactpersoncode = "";
                                {
                                    var recset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                                    recset.DoQuery(string.Format("SELECT \"Tel1\" from ocpr where \"CntctCode\" = '{0}' ", recordset.Fields.Item("CntctCode").Value));
                                    contactpersoncode = recset.Fields.Item("Tel1").Value;
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(recset);
                                }
                                var recordset1 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                                recordset1.DoQuery(string.Format("Select * from dpi1 where \"DocEntry\" = '{0}'", Key));
                                while (!recordset1.EoF)
                                {
                                    var item = new { ItemCode = recordset1.Fields.Item("ItemCode").Value, ItemName = recordset1.Fields.Item("Dscription").Value, Quantity = recordset1.Fields.Item("Quantity").Value, UnitPrice = recordset1.Fields.Item("Price").Value, RowDiscountPerc = recordset1.Fields.Item("DiscPrcnt").Value, RowTotal = recordset1.Fields.Item("LineTotal").Value };
                                    row.Add(item);
                                    recordset1.MoveNext();
                                }
                                Marshal.ReleaseComObject(recordset1);

                                var data = new
                                {
                                    ObjectInfo = new
                                    {
                                        ObjType = ObjType2
                                        ,
                                        ObjectName = "AR Invoice Payment Request"
                                        ,
                                        Key = Key
                                    },

                                    Header = new
                                    {

                                        CustomerCode = recordset.Fields.Item("CardCode").Value,
                                        CustomerName = recordset.Fields.Item("CardName").Value,
                                        ContactNumber = contactpersoncode,
                                        DocumentDate = recordset.Fields.Item("DocDate").Value.ToString("yyyyMMdd"),
                                        DueDate = recordset.Fields.Item("DocDueDate").Value.ToString("yyyyMMdd"),
                                        Status = recordset.Fields.Item("DocStatus").Value
                                    }
                                    ,
                                    Rows = row
                                };
                               // json_obj.JSON_OBJECT = data.ToJSON();

                                
                                Data_Dictionary.Add(Key, data.ToJSON());
                                json_obj.OBJECT_Type = ((SAPbobsCOM.BoObjectTypes)ObjType2).ToString();

                                System.Runtime.InteropServices.Marshal.ReleaseComObject(recordset);

                                #endregion
                            }
                        }
                        json_obj.JSON_OBJECT = Data_Dictionary.ToJSON();
                        #endregion

                    }
                    else
                    {

                        json_obj.Message = "Company Not Connected " + oCompany.GetLastErrorDescription();
                    }
                }
            }
          

            return json_obj;

        }
        public void GetObjectKeysFilter(string ObjType, string FieldName, string FieldValue, string SessionID, out Response Response)
        {
            var List = new List<object>();
            int ObjType2 = Convert.ToInt32(ObjType);
            var oCompany = CompaniesSessions[SessionID];
            Response = new global::ComModelService.Response();

            if (oCompany.Connected)
            {
                #region Company Connected

                try
                {

                    if ((int)ObjType2 == (int)SAPbobsCOM.BoObjectTypes.oInvoices || (int)ObjType2 == (int)SAPbobsCOM.BoObjectTypes.oDownPayments)
                    {
                        #region Get Document
                        // var invoice = obj as SAPbobsCOM.Documents;
                        var recset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                        recset.DoQuery(string .Format ("Select \"DocEntry\" from oinv where  {0} = '{1}'" ,FieldName,FieldValue));
                        while (!recset.EoF)
                        {
                            var docentry = recset.Fields.Item("DocEntry").Value;
                            List.Add(docentry);
                            recset.MoveNext();
                        }
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(recset);
                        GC.Collect();

                        #endregion
                    }
                    if ((int)ObjType2 == (int)SAPbobsCOM.BoObjectTypes.oItems)
                    {
                        #region Item Master Data
                        var recset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                        recset.DoQuery(string.Format ("Select \"ItemCode\" from OITM where  {0} = '{1}'" ,FieldName,FieldValue));
                        while (!recset.EoF)
                        {
                            var docentry = recset.Fields.Item("ItemCode").Value;
                            List.Add(docentry);
                            recset.MoveNext();
                        }

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(recset);
                        GC.Collect();

                        #endregion
                    }
                    if ((int)ObjType2 == (int)SAPbobsCOM.BoObjectTypes.oBusinessPartners)
                    {
                        #region Customer Master Data
                        var recset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                        recset.DoQuery(string.Format ("Select \"CardCode\" from OCRD where  {0} = '{1}'" ,FieldName,FieldValue));
                        while (!recset.EoF)
                        {
                            var docentry = recset.Fields.Item("CardCode").Value;
                            List.Add(docentry);
                            recset.MoveNext();
                        }

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(recset);
                        GC.Collect();
                        #endregion
                    }
                    if ((int)ObjType2 == 11)
                    {
                        #region Contact Person
                        var recset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                        recset.DoQuery(string.Format("Select * from OCPR where  {0} = '{1}'" ,FieldName,FieldValue));
                        while (!recset.EoF)
                        {
                            var contactCode = recset.Fields.Item("CntctCode").Value;
                            List.Add(contactCode);
                            recset.MoveNext();
                        }
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(recset);
                        GC.Collect();
                        #endregion
                    }

                    Response.JSON_Response = List.ToJSON();
                }
                catch (Exception ex)
                {

                    Response.JSON_Response = (new { errMessage = ex.Message }).ToJSON();
                    //json_obj.Message = ex.Message;
                }



                #endregion

            }
            else
            {

                Response.JSON_Response = "Company Not Connected " + oCompany.GetLastErrorDescription();
            }
        }
        public void GetObjectKeys(string ObjType, string SessionID, out Response Response)
        {
            var List = new List<object>();
            int ObjType2 = Convert .ToInt32 (ObjType);
            var oCompany = CompaniesSessions[SessionID];
            Response = new global::ComModelService.Response();
            
                if (oCompany.Connected)
                {
                    #region Company Connected

                    try
                    {

                        if ((int)ObjType2 == (int)SAPbobsCOM.BoObjectTypes.oInvoices || (int)ObjType2 == (int)SAPbobsCOM.BoObjectTypes.oDownPayments)
                        {
                            #region Get Document
                            // var invoice = obj as SAPbobsCOM.Documents;
                            var recset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                            recset.DoQuery("Select \"DocEntry\" from oinv ");
                            while (!recset.EoF)
                            {
                                var docentry = recset.Fields.Item("DocEntry").Value;
                                List.Add(docentry);
                                recset.MoveNext();
                            }
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(recset);
                            GC.Collect();

                            #endregion
                        }
                        if ((int)ObjType2 == (int)SAPbobsCOM.BoObjectTypes.oItems)
                        {
                            #region Item Master Data
                            var recset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                            recset.DoQuery("Select \"ItemCode\" from OITM");
                            while (!recset.EoF)
                            {
                                var docentry = recset.Fields.Item("ItemCode").Value;
                                List.Add(docentry);
                                recset.MoveNext();
                            }

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(recset);
                            GC.Collect();

                            #endregion
                        }
                        if ((int)ObjType2 == (int)SAPbobsCOM.BoObjectTypes.oBusinessPartners)
                        {
                            #region Customer Master Data
                            var recset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                            recset.DoQuery("Select \"CardCode\" from OCRD");
                            while (!recset.EoF)
                            {
                                var docentry = recset.Fields.Item("CardCode").Value;
                                List.Add(docentry);
                                recset.MoveNext();
                            }

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(recset);
                            GC.Collect();
                            #endregion
                        }
                        if ((int)ObjType2 == 11)
                        {
                            #region Contact Person
                            var recset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                            recset.DoQuery(string.Format("Select * from OCPR"));
                            while (!recset.EoF)
                            {
                                var contactCode = recset.Fields.Item("CntctCode").Value;
                                List.Add(contactCode);
                                recset.MoveNext();
                            }
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(recset);
                            GC.Collect();
                            #endregion
                        }

                        Response.JSON_Response = List.ToJSON();
                    }
                    catch (Exception ex)
                    {

                        Response.JSON_Response = (new { errMessage = ex.Message }).ToJSON();
                        //json_obj.Message = ex.Message;
                    }



                    #endregion

                }
                else
                {

                    Response.JSON_Response = "Company Not Connected " + oCompany.GetLastErrorDescription();
                }
                #region get udo
                //var r = oCompany.Connect();
                //if (oCompany.Connected )
                //{
                //    ComModel.UserDefinedObject udo = new ComModel.UserDefinedObject(ObjType2.ToString(), SAPbobsCOM.BoUDOObjType.boud_MasterData, oCompany);

                //    udo.GetByKey(Key);
                //    //// udo.GetProperty ("U_ItemCode");

                //    //var data = new
                //    //{
                //    //    ObjectInfo = new
                //    //    {
                //    //        ObjType = ObjType
                //    //        ,
                //    //        ObjectName = (ObjType).ToString()
                //    //        ,
                //    //        Key = Key
                //    //    },

                //    //    Header = new
                //    //    {

                //    //        U_ItemCode = udo.GetProperty("U_ItemCode"),
                //    //        U_Itemcode1 = udo.GetProperty("U_Itemcode1"),
                //    //        U_ItemCodename = udo.GetProperty("U_ItemCodename"),
                //    //        U_Itemcode2 = udo.GetProperty("U_Itemcode2"),
                //    //        U_itemCodename2 = udo.GetProperty("U_itemCodename2"),
                //    //        U_Otherfees = udo.GetProperty("U_Otherfees"),

                //    //        U_DoPayCode = udo.GetProperty("U_DoPayCode"),
                //    //        U_Lockfee = udo.GetProperty("U_Lockfee"),
                //    //        U_testfee = udo.GetProperty("U_testfee")
                //    //    }
                //    //};
                //    //json_obj.JSON_OBJECT = data.ToJSON();
                //}
                //else
                //{

                //    //json_obj.Message = "Company Not Connected " + oCompany.GetLastErrorDescription();
                //}
                #endregion
            
        }


        public void DoQuery(String Query, String SessionID, out Response responses)
        {

            var oCompany = CompaniesSessions[SessionID];
            responses = new Response();
            var recset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
            try
            {
                recset.DoQuery(Query);
                var dt = recset.GetDataTable();

                responses.JSON_Response = dt.DataTableToJSONWithJSONNet();

                dt.Dispose();
            }
            catch (Exception ex) { responses.JSON_Response = ex.Message + " " + ex.StackTrace; }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(recset);
                GC.Collect();
            }
        }

    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExtensionMethods;
namespace TestComModel
{
    class Program
    {
        static void Main(string[] args)
        {
            TestComModel.ServiceReference1.ComModelServiceClient client = new ServiceReference1.ComModelServiceClient();
            String Key = "1";
            String SessionID = "";
            string message = "";

            message = client.Login("manager", "1234", out SessionID);

            if (SessionID == "")
            {
                Console.WriteLine(message);
                Console.ReadKey();
                return;
            }
              //EXAMPLE1. GET OBJECT KEYS and GET all INVOICES
            {
                ServiceReference1.JSON_OBJ obj_ARINV = new ServiceReference1.JSON_OBJ();
         
                int ObjType = 13;//OARINVOICE= 13 ,AR Down Payments(Requst)= 203,ItemMaster = 4 ,Business Partner Master = 2,Contact Person = 11
                //var o = client.GetObjectKeysFilter(ObjType.ToString(), "CardCode", "C20000", SessionID);
                var objects = client.GetObjectBulkonDate("CardCode", "C20000", "20100101", "20180101","O", ObjType.ToString(), SessionID);
                //dynamic objs = o.JSON_Response.ToObjectFromJson();
            var list = objects.JSON_OBJECT.ToObjectFromJson() as dynamic ;
            foreach(var item in list)
            {
                Console.WriteLine(item.Key+ item.Value);
            }
                //foreach (var item in objs)
                //{
                //    obj_ARINV = client.GetObjectByKey(Key, ObjType.ToString(), SessionID);
                //    Console.WriteLine(obj_ARINV.JSON_OBJECT.ToString());
                //}
            }
            #region Commented Code
            // int ObjType = 13;// AR Invoice
            //  var obj_ARINV = client.GetObjectByKey(Key, ObjType, SessionID);
            //String ObjType = "uGrade";// UDO
            //Key = "01";
            //var objUDO = client.GetObjectByKey(Key, ObjType, SessionID);


            //  var obj_ARINV = client.GetObjectByKey(Key, ObjType, SessionID);
            //var ObjType = 203;// AR Down Payments
            //var obj_ARDNPM = client.GetObjectByKey(Key, ObjType, SessionID);
            // var ObjType = -203;// AR Down Payments Request
            // var obj_ARDNPM = client.GetObjectByKey(Key, ObjType, SessionID);

            // var ObjType = 4; // Item Master
            //  Key = "P10004";// Item Code
            // var obj_ItemMaster = client.GetObjectByKey(Key, ObjType, SessionID);
            //var ObjType = 2; // Business Master
            //           Key = "C23900";// Business Partner Code
            //         var obj_BPMaster = client.GetObjectByKey(Key, ObjType, SessionID);

            // var ObjType = 11;// Contact Person Code
            //   Key = "1";

            // var obj_contactperson = client.GetObjectByKey(Key, ObjType, SessionID);
            //  System.IO.File.WriteAllText("oinv.txt", obj_ARINV.JSON_OBJECT);

            // var data =   client.DoQuery(string.Format("Select * from \"@GRADE\" where \"Code\" = '{0}'", "01"),SessionID );

            // var table = data.JSON_Response.ToObjectFromJson() as System.Data.DataTable;
            //ServiceReference2.JSON_OBJ obj_ARINV = new ServiceReference2.JSON_OBJ();

            // obj_ARINV.JSON_OBJECT = System.IO.File.ReadAllText("oinv1.txt"); 
            #endregion

            //Cancel Invoice Example
            {
                ServiceReference1.JSON_OBJ obj_ARINV = new ServiceReference1.JSON_OBJ();

                obj_ARINV.JSON_OBJECT = System.IO.File.ReadAllText("Canceloinv1.txt");
                var response_obj_ARINV = client.AddUpdateObject(obj_ARINV, SessionID);
            }
            //Data Interface by Query
            {
                var data = client.DoQuery(string.Format("Select * from \"@GRADE\" where \"Code\" = '{0}'", "01"), SessionID);
            }


            //Update BP
            {
                ServiceReference1.JSON_OBJ obj_BP = new ServiceReference1.JSON_OBJ();
                obj_BP.JSON_OBJECT = System.IO.File.ReadAllText("BPUpdate.txt");
                var response_obj_BP = client.AddUpdateObject(obj_BP, SessionID);
            }
            //Update Item Master
            {

                ServiceReference1.JSON_OBJ obj_Item = new ServiceReference1.JSON_OBJ();
                 obj_Item.JSON_OBJECT = System.IO.File.ReadAllText("itemUpdate.txt");
                var response_obj_Item = client.AddUpdateObject(obj_Item, SessionID);

            }
            //Add Payment
            {
                ServiceReference1.JSON_OBJ obj_pay = new ServiceReference1.JSON_OBJ();
                obj_pay.JSON_OBJECT = System.IO.File.ReadAllText("payment.txt");
                var response_obj_pay = client.AddUpdateObject(obj_pay, SessionID);
            }
            // add invoice
            {
                ServiceReference1.JSON_OBJ obj_ARINV = new ServiceReference1.JSON_OBJ();
                obj_ARINV.JSON_OBJECT = System.IO.File.ReadAllText("oinv1.txt");
                var response_obj_ARINV = client.AddUpdateObject(obj_ARINV, SessionID);
            }
            client.LogOut(SessionID);
          //  Console.WriteLine(response_obj_Item.JSON_Response);
            //Console.WriteLine(response_obj_BP.JSON_Response);
            //Console.WriteLine(response_obj_ARINV.JSON_Response);
            //Console.WriteLine(response_obj_pay.JSON_Response);
            Console.ReadKey();

            //SAPbobsCOM.Company _Company = new SAPbobsCOM.Company();
            //if (_Company == null || !_Company.Connected)
            //{
            //    _Company = new SAPbobsCOM.Company();
            //    _Company.Server = System.Configuration.ConfigurationManager.AppSettings["Server"];
            //    _Company.DbServerType = (SAPbobsCOM.BoDataServerTypes)Enum.Parse(typeof(SAPbobsCOM.BoDataServerTypes), System.Configuration.ConfigurationManager.AppSettings["DbServerType"]);
            //    _Company.DbUserName = System.Configuration.ConfigurationManager.AppSettings["DbUserName"];
            //    _Company.DbPassword = System.Configuration.ConfigurationManager.AppSettings["DbPassword"];
            //    _Company.UserName = System.Configuration.ConfigurationManager.AppSettings["UserName"];
            //    _Company.Password = System.Configuration.ConfigurationManager.AppSettings["Password"];
            //    _Company.CompanyDB = System.Configuration.ConfigurationManager.AppSettings["CompanyDB"];

            //}
            //var oCompany = _Company;
            //var json_obj = new JSON_OBJ();
            //var r = oCompany.Connect();
            //if (r == 0)
            //{
            //    object obj = oCompany.GetBusinessObject((SAPbobsCOM.BoObjectTypes)ObjType);

            //    try
            //    {

            //        if (ObjType == (int)SAPbobsCOM.BoObjectTypes.oInvoices)
            //        {
            //            var invoice = obj as SAPbobsCOM.Documents;
            //            invoice.GetByKey(Convert.ToInt32(Key));
            //            var contactpersoncode = invoice.ContactPersonCode;
            //            var lst = new List<object>();
            //            var lines = invoice.Lines;
            //            lines.SetCurrentLine(0);
            //            var count = invoice.Lines.Count;
            //            for (int i = 0; i < count; i++)
            //            {
            //                lines.SetCurrentLine(i);
            //                var item = new { ItemCode = lines.ItemCode, ItemName = lines.ItemDescription, Quantity = lines.Quantity, UnitPrice = lines.Price, RowDiscountPerc = lines.DiscountPercent, RowTotal = lines.LineTotal };
            //                lst.Add(item);
            //            }
            //            var data = new
            //            {
            //                CustomerCode = invoice.CardCode,
            //                CustomerName = invoice.CardName,
            //                ContactNumber = contactpersoncode,
            //                DocumentDate = invoice.DocDate.ToString("yyyyMMdd"),
            //                DueDate = invoice.DocDueDate.ToString("yyyyMMdd"),
            //                Status = invoice.DocumentStatus
            //                ,
            //                lst
            //            };
            //            json_obj.JSON_OBJECT = data.ToJSON();
            //            json_obj.OBJECT_Type = "AR Invoice";

            //        }


            //    }
            //    catch (Exception ex)
            //    {
            //        json_obj.Message = ex.Message;
            //    }
            //    finally
            //    {
            //        System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
            //        GC.Collect();
            //        obj = null;
            //    }


            //}
            //else
            //{

            //    json_obj.Message = oCompany.GetLastErrorDescription();
            //}

            //Console.WriteLine(obj.JSON_OBJECT);
        }
    }
}

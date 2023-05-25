using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using ADDONBASE;
using System.Windows.Forms;
using EncryptDecrypt;
using System.IO;
using System.Xml.Linq;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;
using ADDONBASE.Extensions;

namespace SAP_XRISK_Addon.BusinessLogic
{
    [FormAttribute("SAP_XRISK_Addon.BusinessLogic.FRM_SETUP", "BusinessLogic/FRM_SETUP.b1f")]
    class FRM_SETUP : _UserFormBase
    {
        public FRM_SETUP()
        {
        }
        string password = "";
        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("Item_6").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_7").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_8").Specific));
            this.Button0.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button0_ClickBefore);
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_10").Specific));
            this.EditText5 = ((SAPbouiCOM.EditText)(this.GetItem("Item_11").Specific));
            this.StaticText5 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_12").Specific));
            this.EditText6 = ((SAPbouiCOM.EditText)(this.GetItem("Item_13").Specific));
            this.StaticText6 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_14").Specific));
            this.EditText8 = ((SAPbouiCOM.EditText)(this.GetItem("Item_17").Specific));
            this.StaticText8 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_18").Specific));
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_0").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_1").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.LoadAfter += new LoadAfterHandler(this.Form_LoadAfter);

        }
        private void OnCustomInitialize()
        {
            var SAPUser = this.CurrentForm.DataSources.UserDataSources.Item("SAPUser");
            var DBUser = this.CurrentForm.DataSources.UserDataSources.Item("DBUser");
            SAPUser.Value = this.Company.UserName;
            DBUser.Value = this.Company.DbUserName;
            //add new field
            var tbl = this.Company.UserTables.Item("AB_SerUrl");
            bool exist = tbl.GetByKey("01");
            if (!exist)
            {
                tbl.Code = "01";
                tbl.Name = "01";
                tbl.UserFields.Fields.Item("U_SerUrl").Value = "http://localhost:19541/SAP_XRISK.svc";
                tbl.Add();
            }
            var recset = Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
            recset.DoQuery("select U_SerUrl from [@AB_SERURL] where Code='01'");
            this.CurrentForm.DataSources.UserDataSources.Item("ServiceURL").Value=recset.Fields.Item(0).Value.ToString();
        }

        private string getStoredProced(string name,SAPbobsCOM.Company mycomp)
        {
            var conctionstring = string.Format("Data Source={0}; Initial Catalog={3};User ID={1};Password={2};Integrated Security = False", this.Company.Server, this.Company.DbUserName, password,this.Company.CompanyDB);
            string result = "";
            string query = string.Format(GetQuery("getSP"), name);
            SqlConnection conn = new SqlConnection(conctionstring);
            conn.Open();
            SqlCommand command = new SqlCommand(query,conn);
            using (SqlDataReader reader = command.ExecuteReader())
            {
                if (reader.Read())
                {
                    result =Convert.ToString( reader["definition"]);
                }
            }
            conn.Close();
            return result;
        }
        private bool insertrecords(string obj, string keyname, string active, string tablename ,SqlConnection myConnection)
        {
            bool result = true;
            string query = string.Format(GetQuery("insertdefaultdata"), obj,keyname,active,tablename); 
            SqlCommand myCommand = new SqlCommand(query, myConnection);
            int i= myCommand.ExecuteNonQuery();
            if (i == 0) result = false;
            return result;
        }

        private void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
        }
        private SAPbouiCOM.EditText EditText3;
        private SAPbouiCOM.StaticText StaticText3;
        private SAPbouiCOM.Button Button0;

        const string Connection = "Connection";

        const string _Server = "Server";
        const string _CompanyDB = "CompanyDB";
        const string _DbUserName = "DbUserName";
        const string _DbPassword = "DbPassword";
        const string _UserName = "UserName";
        const string _Password = "Password";
        const string _LicenseServer = "LicenseServer";
        const string _language = "language";
        const string _DbServerType = "DbServerType";
        const string _EncFileName = "cnx.p";
        const string psw = "!@#908!112aFgKts";
        private void Button0_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            this.Application.SendKeys("{TAB}");
            var SAPPass = this.CurrentForm.DataSources.UserDataSources.Item("SAPPass");
            var DBPass = this.CurrentForm.DataSources.UserDataSources.Item("DBPass");
            var xml = "<Connection> <Server>SAMIR</Server> <CompanyDB>SQL_CFM</CompanyDB> <DbUserName>sa</DbUserName><DbPassword>P@ssw0rd</DbPassword><UserName>manager</UserName>  <Password>manager</Password>  <LicenseServer>Samir:30000</LicenseServer>  <language>3</language>  <DbServerType>7</DbServerType></Connection>";
           
           XDocument doc = XDocument.Parse(xml);
             var connection = doc.Element(Connection);
           connection.Element(_Server).Value = this.Company.Server;
           connection.Element(_CompanyDB).Value =this.Company. CompanyDB;
           connection.Element(_DbUserName).Value =this.Company. DbUserName;
           connection.Element(_DbPassword).Value = DBPass.Value;
           connection.Element(_UserName).Value = this.Company.UserName;
           connection.Element(_Password).Value = SAPPass.Value;
           connection.Element(_LicenseServer).Value = this.Company.LicenseServer;
           connection.Element(_language).Value = this.Company.language.ToString();
           connection.Element(_DbServerType).Value = this.Company.DbServerType.ToString();
            password = DBPass.Value;
            //var connectionstring = string.Format("Data Source={0};Initial Catalog={1};User ID={2};Password={3}",StagingServer.Value, StagingDB.Value, DBUser.Value,DBPass.Value);
            //connectionstring = connectionstring.EncryptMe("ABIT@!@#$");
            //var path = Path.Combine( Path.GetTempPath(),"cnx.p");
            //doc.ToString().EncryptMeToFile(psw, path);
            #region service Url
            var tbl = this.Company.UserTables.Item("AB_SerUrl");
            tbl.GetByKey("01");
            var URL = this.CurrentForm.DataSources.UserDataSources.Item("ServiceURL").Value;
            tbl.UserFields.Fields.Item("U_SerUrl").Value =URL;
            tbl.Update();
            #endregion

            SAP_XRISK_WS.SAP_XRISKClient client = new SAP_XRISK_WS.SAP_XRISKClient();
            client.Endpoint.Address = new System.ServiceModel.EndpointAddress(URL);
            string msg=client.SaveConnection(doc.ToString().EncryptMe(psw));
            if (!string.IsNullOrEmpty(msg))
            {
                Application.StatusBar.SetText("Error saving credentials and connecting company " + msg, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            else
            {
                Application.StatusBar.SetText("Company Connectd Successfully " , SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

            }
          
            Task tsk = new Task(() => {
                var connectionstring = string.Format("Data Source={0};User ID={1};Password={2};Integrated security=False;database=master", this.Company.Server, this.Company.DbUserName, DBPass.Value);
                var path = System.Configuration.ConfigurationManager.AppSettings["pathdatabase"].ToString();
                SqlConnection myConn = new SqlConnection(connectionstring);
                string str =string.Format(GetQuery("Createdatabase"), path, this.Company.CompanyDB);
                SqlCommand myCommand = new SqlCommand(str, myConn);
                try
                {
                    this.Company.DoQuery(string.Format(GetQuery("dropChangeTracking")));
                    this.Company.DoQuery(string.Format(GetQuery("changeTracking"), this.Company.CompanyDB));
                  
                    myConn.Open();
                    myCommand.ExecuteNonQuery();
                    var conctionstring = string.Format("Data Source={0}; Initial Catalog={3}_CT;User ID={1};Password={2};Integrated Security = False", this.Company.Server, this.Company.DbUserName, DBPass.Value, this.Company.CompanyDB);
                    SqlConnection myConn1 = new SqlConnection(conctionstring);
                    myConn1.Open();
                    myCommand = new SqlCommand(GetQuery("Create@AB_CTRK"), myConn1);
                    myCommand.ExecuteNonQuery();
                    myCommand = new SqlCommand(GetQuery("Create@AB_TRKOBJ"), myConn1);
                    myCommand.ExecuteNonQuery();
                    myCommand = new SqlCommand(GetQuery("Create@Expiry"), myConn1);
                    myCommand.ExecuteNonQuery();
                    myCommand = new SqlCommand(GetQuery("Create@Log"), myConn1);
                    myCommand.ExecuteNonQuery();
                    string[] splitter = new string[] { "GO" };
                    string[] commandTexts = GetQuery("Create.SP.GET_CHANGES").Split(splitter, StringSplitOptions.RemoveEmptyEntries);
                    foreach (string commandText in commandTexts)
                    {
                            myCommand = new SqlCommand(commandText, myConn1);
                            myCommand.ExecuteNonQuery();
                    }
                    commandTexts = GetQuery("Create.SF.AB_VN").Split(splitter, StringSplitOptions.RemoveEmptyEntries);
                    foreach (string commandText in commandTexts)
                    {
                        myCommand = new SqlCommand(commandText, myConn1);
                        myCommand.ExecuteNonQuery();
                    }
                    //code to add docs data

                    insertrecords("24", "DocEntry", "Y", "ORCT", myConn1);
                    insertrecords("46", "DocEntry", "Y", "OVPM", myConn1);
                    insertrecords("4", "ItemCode", "Y", "OITM", myConn1);
                    insertrecords("2", "CardCode", "Y", "OCRD", myConn1);
                    insertrecords("1", "AcctCode", "Y", "OACT", myConn1);
                    insertrecords("13", "DocEntry", "Y", "OINV", myConn1);
                    insertrecords("13", "DocEntry", "Y", "OINV", myConn1);
                    insertrecords("157", "IdNumber", "Y", "OPWZ", myConn1);

                    Application.StatusBar.SetText("Database Created Successfully ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    if (myConn1.State == ConnectionState.Open)
                    {
                        myConn1.Close();
                    }
                    myConn1.Dispose();
                }
                catch (System.Exception ex)
                {
                    Application.MessageBox(ex.ToString());
                    ex.AppendInLogFile();
                }
                finally
                {
                    if (myConn.State == ConnectionState.Open)
                    {
                        myConn.Close();
                    }
                    myConn.Dispose();
                    myCommand.Dispose();
                } 
               
               

            });
            tsk.Start();
//      string argument = "/select, \"" + path + "\"";
  //    System.Diagnostics.Process.Start("explorer.exe", argument);
        }

        private SAPbouiCOM.StaticText StaticText4;
        private SAPbouiCOM.EditText EditText5;
        private SAPbouiCOM.StaticText StaticText5;
        private SAPbouiCOM.EditText EditText6;
        private SAPbouiCOM.StaticText StaticText6;
        private SAPbouiCOM.EditText EditText8;
        private SAPbouiCOM.StaticText StaticText8;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.EditText EditText0;
    }
}

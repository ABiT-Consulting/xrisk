using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using ADDONBASE;
using SAPbouiCOM;

namespace SAP_XRISK_Addon.BusinessLogic
{
    [FormAttribute("SAP_XRISK_Addon.BusinessLogic.FRM_Log", "BusinessLogic/FRM_Log.b1f")]
    class FRM_Log : _UserFormBase
    {
        public FRM_Log()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_0").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("search").Specific));
            this.Button0.PressedAfter += Button0_PressedAfter;
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_3").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("fromdate").Specific));
            this.EditText0.ValidateBefore += EditText0_ValidateBefore;
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_5").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("todate").Specific));
            this.EditText1.ValidateBefore += EditText1_ValidateBefore;
            this.OnCustomInitialize();

        }

        void EditText0_ValidateBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            var fromdate = (this.CurrentForm.Items.Item("fromdate").Specific as SAPbouiCOM.EditText).Value;
            if (string.IsNullOrEmpty(fromdate))
            {
                this.Application.SetStatusBarMessage("From Date cannot be Empty", BoMessageTime.bmt_Short, true);
                BubbleEvent = false;

            }
        }

        void EditText1_ValidateBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            var fromdate =convertDate( (this.CurrentForm.Items.Item("fromdate").Specific as SAPbouiCOM.EditText).Value);
            if (string.IsNullOrEmpty((this.CurrentForm.Items.Item("todate").Specific as SAPbouiCOM.EditText).Value))
            {
                BubbleEvent = false;
                this.Application.SetStatusBarMessage("To Date cannot be Empty", BoMessageTime.bmt_Short, true);
            }
            else
            {
                var todate = convertDate((this.CurrentForm.Items.Item("todate").Specific as SAPbouiCOM.EditText).Value);
                if (todate < fromdate)
                {
                    this.Application.SetStatusBarMessage("To Date should be greater than From Date", BoMessageTime.bmt_Short, true);
                    BubbleEvent = false;

                }
            }
        }
        public DateTime convertDate(string val)
        {
            string newDate = val.ToString().Substring(6, 2) + "/" + val.Substring(4, 2) + "/" + val.Substring(0, 4);
            DateTime date = DateTime.ParseExact(newDate, "dd/MM/yyyy", null);
            return date;

        }

        void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            var fromdate = (this.CurrentForm.Items.Item("fromdate").Specific as SAPbouiCOM.EditText).Value;
            var todate = (this.CurrentForm.Items.Item("todate").Specific as SAPbouiCOM.EditText).Value;
            fillmatrix(fromdate, todate);
        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            
        }

        private SAPbouiCOM.Matrix Matrix0;

        private void OnCustomInitialize()
        {
            var date=DateTime.Now;
            this.CurrentForm.DataSources.UserDataSources.Item("fromdate").Value=(date.AddDays(-10)).ToString("yyyyMMdd");
            this.CurrentForm.DataSources.UserDataSources.Item("todate").Value = date.ToString("yyyyMMdd");
            var fromdate = (this.CurrentForm.Items.Item("fromdate").Specific as SAPbouiCOM.EditText).Value;
            var todate = (this.CurrentForm.Items.Item("todate").Specific as SAPbouiCOM.EditText).Value;
            fillmatrix(fromdate,todate);
        }

        private void fillmatrix(string fromdate, string todate)
        {
            this.CurrentForm.DataSources.DataTables.Item("@log").Clear();
            //var fdate = Convert.ToDateTime(fromdate);
            //var tdate=Convert.ToDateTime(todate);
            this.CurrentForm.DataSources.DataTables.Item("@log").ExecuteQuery(string.Format(GetQuery("LogMatrix"), fromdate, todate, this.Company.CompanyDB));
            this.Matrix0.Columns.Item("DocEntry").DataBind.Bind("@log", "DocEntry");
            this.Matrix0.Columns.Item("U_version").DataBind.Bind("@log", "U_version");
            this.Matrix0.Columns.Item("U_ObjType").DataBind.Bind("@log", "U_ObjType");
            this.Matrix0.Columns.Item("U_DateTime").DataBind.Bind("@log", "U_DateTime");
            this.Matrix0.Columns.Item("U_Error").DataBind.Bind("@log", "U_Error");
            this.Matrix0.Columns.Item("U_Dirction").DataBind.Bind("@log", "U_Direction");
            this.Matrix0.Columns.Item("U_XML").DataBind.Bind("@log", "U_XML");
            this.Matrix0.LoadFromDataSource();

        }


        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.EditText EditText1;
    }
}

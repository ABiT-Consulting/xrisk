using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace TEST
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            //  lblcompany.Text = Program.company.CompanyDB + " " + Program.company.CompanyName;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //try
            //{
            //    lblcompany.Text = Program.company.CompanyDB + " " + Program.company.CompanyName;
            //}
            //catch { }
        }

        private void btnLoadXML_Click(object sender, EventArgs e)
        {
            var dlg = openFileDialog1.ShowDialog();
            if (dlg == DialogResult.OK)
            {
                tbPath.Text = openFileDialog1.FileName;

                tbXML.Text = System.IO.File.ReadAllText(tbPath.Text);
            }

        }

        private void btnGETObject_Click(object sender, EventArgs e)
        {

            //object objectid = tbObjectID.Text;
            //try
            //{
            //    objectid = Convert.ToInt32(tbObjectID.Text);
            //}
            //catch
            //{

            //}
            //if (objectid is int)
            //{
            //    var type = (SAPbobsCOM.BoXmlExportTypes)Enum.Parse(typeof(SAPbobsCOM.BoXmlExportTypes), cbxmlType.SelectedItem.ToString());
            //    Program.company.XmlExportType = type;

            //    dynamic documnet = Program.company.GetBusinessObject((SAPbobsCOM.BoObjectTypes)objectid);
            //    try
            //    {
            //        if (documnet.GetByKey(tbKey.Text))
            //        {
            //            XDocument xdoc = XDocument.Parse(documnet.GetAsXML());
            //            tbXML.Text = xdoc.ToString();
            //        }
            //    }catch
            //    {
            //    }

            //}

        }
        private void HighlightString(string stringToHighlight)
        {
            // Code here to search your text and highlight a string.
        }

        private void SearchDialog()
        {
            var findDialog = new Form { Width = 500, Height = 142, Text = "Find" };
            var textLabel = new Label() { Left = 10, Top = 20, Text = "Find Text:", Width = 100 };
            var inputBox = new TextBox() { Left = 150, Top = 20, Width = 300 };
            var search = new Button() { Text = "Find", Left = 350, Width = 100, Top = 70 };

            search.Click += (object sender, EventArgs e) =>
        {
            var t = inputBox.Text;
            var startindex = tbXML.Text.IndexOf(t);
            tbXML.BackColor = Color.White;
            try
            {
                while (startindex != -1)
                {

                    tbXML.Select(startindex, t.Length);
                    tbXML.SelectionBackColor = Color.Red;
                    startindex = tbXML.Text.Substring(startindex + t.Length).IndexOf(t) + startindex + t.Length;

                }
            }
            catch (Exception ex)
            { }
            findDialog.DialogResult = DialogResult.OK;
        };

            findDialog.Controls.Add(search);
            findDialog.Controls.Add(textLabel);
            findDialog.Controls.Add(inputBox);
            if (findDialog.ShowDialog() == DialogResult.OK)
            {

            }
        }



        private void btnAdd_Click(object sender, EventArgs e)
        {
            //dynamic doc = Program.company.GetBusinessObjectFromXML(tbPath.Text, 0);
            //var i = doc.Add();
            //if (i == 0)
            //{
            //    MessageBox.Show($"Business Object Added at {Program.company.GetNewObjectKey()}");
            //}
            //else
            //{
            //    MessageBox.Show(Program.company.GetLastErrorCode() + " " + Program.company.GetLastErrorDescription());

            //}
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SearchDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //var xml = XDocument.Parse(tbXML.Text);
            //var path = System.IO.Path.GetRandomFileName();
            //xml.Save(path);
            //tbPath.Text = path;
            //dynamic doc = Program.company.GetBusinessObjectFromXML(tbPath.Text, 0);
            //var i = doc.Update();
            //if (i == 0)
            //{
            //    MessageBox.Show($"Business Object Added at {Program.company.GetNewObjectKey()}");
            //}
            //else
            //{
            //    MessageBox.Show(Program.company.GetLastErrorCode() + " " + Program.company.GetLastErrorDescription());

            //}
        }

        private void btnCallService_Click(object sender, EventArgs e)
        {
            tbXML.Clear();
            tbmaxno.Text = "";
            ServiceReference1.SAP_XRISKClient client = new ServiceReference1.SAP_XRISKClient();
            int versionnumber = 0;
            string xml = "";
            versionnumber = client.GetXmlandVersion(Convert.ToInt32(tbversion.Text), tbobjid.Text,tbcriteria.Text, out xml);
            
            if (!string.IsNullOrEmpty(xml))
            {var doc = XDocument.Parse(xml);
                tbXML.Text = doc.ToString();
                tbmaxno.Text = versionnumber.ToString() + " --------- " + doc.Descendants("BO").Count().ToString();
            }
            else
            {
                tbXML.Text = "----No Result Found----";
                tbmaxno.Text = versionnumber.ToString();
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

      

        private void btnx2s_Click(object sender, EventArgs e)
        {
            ServiceReference1.SAP_XRISKClient client = new ServiceReference1.SAP_XRISKClient();
            string msg;
            string key = "";
            if(client.XRiskToSap(tbXML.Text,out msg,out key )) { MessageBox.Show("Successfully Added at "+ key); }
            else { MessageBox.Show("ERROR : " + msg); }
        }

        private void btnSchema_Click(object sender, EventArgs e)
        {
            tbXML.Clear();
            ServiceReference1.SAP_XRISKClient client = new ServiceReference1.SAP_XRISKClient();
            string msg;
            msg=client.getSchema(Convert.ToInt32(tbobjid.Text));
            tbXML.Text = msg;
        }
    }
}

namespace TEST
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.tbXML = new System.Windows.Forms.RichTextBox();
            this.lblcompany = new System.Windows.Forms.Label();
            this.btnLoadXML = new System.Windows.Forms.Button();
            this.tbPath = new System.Windows.Forms.TextBox();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.button1 = new System.Windows.Forms.Button();
            this.btnCallService = new System.Windows.Forms.Button();
            this.tbcriteria = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnSchema = new System.Windows.Forms.Button();
            this.label7 = new System.Windows.Forms.Label();
            this.tbmaxno = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.tbversion = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.tbobjid = new System.Windows.Forms.TextBox();
            this.btnx2s = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(15, 94);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(75, 17);
            this.label1.TabIndex = 0;
            this.label1.Text = "ImportXML";
            // 
            // tbXML
            // 
            this.tbXML.Location = new System.Drawing.Point(100, 90);
            this.tbXML.Margin = new System.Windows.Forms.Padding(4);
            this.tbXML.Name = "tbXML";
            this.tbXML.Size = new System.Drawing.Size(679, 247);
            this.tbXML.TabIndex = 1;
            this.tbXML.Text = "";
            // 
            // lblcompany
            // 
            this.lblcompany.AutoSize = true;
            this.lblcompany.Location = new System.Drawing.Point(15, 11);
            this.lblcompany.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblcompany.Name = "lblcompany";
            this.lblcompany.Size = new System.Drawing.Size(81, 17);
            this.lblcompany.TabIndex = 9;
            this.lblcompany.Text = "lblCompany";
            // 
            // btnLoadXML
            // 
            this.btnLoadXML.Location = new System.Drawing.Point(680, 54);
            this.btnLoadXML.Margin = new System.Windows.Forms.Padding(4);
            this.btnLoadXML.Name = "btnLoadXML";
            this.btnLoadXML.Size = new System.Drawing.Size(100, 25);
            this.btnLoadXML.TabIndex = 10;
            this.btnLoadXML.Text = "Load";
            this.btnLoadXML.UseVisualStyleBackColor = true;
            this.btnLoadXML.Click += new System.EventHandler(this.btnLoadXML_Click);
            // 
            // tbPath
            // 
            this.tbPath.Enabled = false;
            this.tbPath.Location = new System.Drawing.Point(100, 54);
            this.tbPath.Margin = new System.Windows.Forms.Padding(4);
            this.tbPath.Name = "tbPath";
            this.tbPath.Size = new System.Drawing.Size(571, 22);
            this.tbPath.TabIndex = 11;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            this.openFileDialog1.Filter = "XML|*.xml";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(786, 94);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(78, 23);
            this.button1.TabIndex = 14;
            this.button1.Text = "Search";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // btnCallService
            // 
            this.btnCallService.Location = new System.Drawing.Point(563, 39);
            this.btnCallService.Name = "btnCallService";
            this.btnCallService.Size = new System.Drawing.Size(110, 25);
            this.btnCallService.TabIndex = 16;
            this.btnCallService.Text = "Submit";
            this.btnCallService.UseVisualStyleBackColor = true;
            this.btnCallService.Click += new System.EventHandler(this.btnCallService_Click);
            // 
            // tbcriteria
            // 
            this.tbcriteria.AccessibleDescription = "";
            this.tbcriteria.Location = new System.Drawing.Point(297, 42);
            this.tbcriteria.Name = "tbcriteria";
            this.tbcriteria.Size = new System.Drawing.Size(260, 22);
            this.tbcriteria.TabIndex = 17;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(294, 21);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(53, 17);
            this.label4.TabIndex = 18;
            this.label4.Text = "Criteria";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnSchema);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.tbmaxno);
            this.groupBox1.Controls.Add(this.btnCallService);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.tbcriteria);
            this.groupBox1.Controls.Add(this.tbversion);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.tbobjid);
            this.groupBox1.Location = new System.Drawing.Point(100, 424);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(680, 71);
            this.groupBox1.TabIndex = 20;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Get XML and Version";
            // 
            // btnSchema
            // 
            this.btnSchema.Location = new System.Drawing.Point(563, 12);
            this.btnSchema.Name = "btnSchema";
            this.btnSchema.Size = new System.Drawing.Size(110, 25);
            this.btnSchema.TabIndex = 26;
            this.btnSchema.Text = "Get Schema";
            this.btnSchema.UseVisualStyleBackColor = true;
            this.btnSchema.Click += new System.EventHandler(this.btnSchema_Click);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(168, 21);
            this.label7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(83, 17);
            this.label7.TabIndex = 25;
            this.label7.Text = "new version";
            // 
            // tbmaxno
            // 
            this.tbmaxno.Location = new System.Drawing.Point(171, 42);
            this.tbmaxno.Name = "tbmaxno";
            this.tbmaxno.ReadOnly = true;
            this.tbmaxno.Size = new System.Drawing.Size(118, 22);
            this.tbmaxno.TabIndex = 24;
            this.tbmaxno.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(94, 21);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(54, 17);
            this.label5.TabIndex = 23;
            this.label5.Text = "version";
            // 
            // tbversion
            // 
            this.tbversion.Location = new System.Drawing.Point(95, 42);
            this.tbversion.Margin = new System.Windows.Forms.Padding(4);
            this.tbversion.Name = "tbversion";
            this.tbversion.Size = new System.Drawing.Size(68, 22);
            this.tbversion.TabIndex = 22;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(7, 21);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(66, 17);
            this.label6.TabIndex = 21;
            this.label6.Text = "Object ID";
            // 
            // tbobjid
            // 
            this.tbobjid.Location = new System.Drawing.Point(11, 42);
            this.tbobjid.Margin = new System.Windows.Forms.Padding(4);
            this.tbobjid.Name = "tbobjid";
            this.tbobjid.Size = new System.Drawing.Size(76, 22);
            this.tbobjid.TabIndex = 20;
            // 
            // groupBox2
            // 
            //this.groupBox2.Controls.Add(this.cbxmlType);
            //this.groupBox2.Controls.Add(this.tbObjectID);
            //this.groupBox2.Controls.Add(this.button2);
            //this.groupBox2.Controls.Add(this.lblObjectID);
            //this.groupBox2.Controls.Add(this.tbKey);
            //this.groupBox2.Controls.Add(this.label3);
            //this.groupBox2.Controls.Add(this.btnGETObject);
            //this.groupBox2.Controls.Add(this.label2);
            //this.groupBox2.Controls.Add(this.btnAdd);
            //this.groupBox2.Location = new System.Drawing.Point(100, 344);
            //this.groupBox2.Name = "groupBox2";
            //this.groupBox2.Size = new System.Drawing.Size(677, 74);
            //this.groupBox2.TabIndex = 21;
            //this.groupBox2.TabStop = false;
            //this.groupBox2.Text = "Basic Operations";
            // 
            // groupBox3
            // 
            //this.groupBox3.Controls.Add(this.label12);
            //this.groupBox3.Controls.Add(this.tbresult1);
            //this.groupBox3.Controls.Add(this.label8);
            //this.groupBox3.Controls.Add(this.tbend);
            //this.groupBox3.Controls.Add(this.button3);
            //this.groupBox3.Controls.Add(this.label9);
            //this.groupBox3.Controls.Add(this.label10);
            //this.groupBox3.Controls.Add(this.tbtype);
            //this.groupBox3.Controls.Add(this.tbstart);
            //this.groupBox3.Controls.Add(this.label11);
            //this.groupBox3.Controls.Add(this.tbobjid1);
            //this.groupBox3.Location = new System.Drawing.Point(97, 501);
            //this.groupBox3.Name = "groupBox3";
            //this.groupBox3.Size = new System.Drawing.Size(680, 71);
            //this.groupBox3.TabIndex = 22;
            //this.groupBox3.TabStop = false;
            //this.groupBox3.Text = "addChangeTracks";
            // 
            // label12
            // 
            //this.label12.AutoSize = true;
            //this.label12.Location = new System.Drawing.Point(375, 22);
            //this.label12.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            //this.label12.Name = "label12";
            //this.label12.Size = new System.Drawing.Size(46, 16);
            //this.label12.TabIndex = 27;
            //this.label12.Text = "Result";
            // 
            // tbresult1
            // 
            //this.tbresult1.AccessibleDescription = "";
            //this.tbresult1.Location = new System.Drawing.Point(378, 43);
            //this.tbresult1.Name = "tbresult1";
            //this.tbresult1.Size = new System.Drawing.Size(56, 22);
            //this.tbresult1.TabIndex = 26;
            // 
            // label8
            // 
            //this.label8.AutoSize = true;
            //this.label8.Location = new System.Drawing.Point(193, 21);
            //this.label8.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            //this.label8.Name = "label8";
            //this.label8.Size = new System.Drawing.Size(32, 16);
            //this.label8.TabIndex = 25;
            //this.label8.Text = "End";
            // 
            // tbend
            // 
            //this.tbend.Location = new System.Drawing.Point(195, 42);
            //this.tbend.Name = "tbend";
            //this.tbend.Size = new System.Drawing.Size(94, 22);
            //this.tbend.TabIndex = 24;
            // 
            // button3
            // 
            //this.button3.Location = new System.Drawing.Point(564, 41);
            //this.button3.Name = "button3";
            //this.button3.Size = new System.Drawing.Size(110, 25);
            //this.button3.TabIndex = 16;
            //this.button3.Text = "Submit";
            //this.button3.UseVisualStyleBackColor = true;
            //this.button3.Click += new System.EventHandler(this.button3_Click);
            //// 
            //// label9
            //// 
            //this.label9.AutoSize = true;
            //this.label9.Location = new System.Drawing.Point(313, 22);
            //this.label9.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            //this.label9.Name = "label9";
            //this.label9.Size = new System.Drawing.Size(40, 16);
            //this.label9.TabIndex = 18;
            //this.label9.Text = "Type";
            //// 
            //// label10
            //// 
            //this.label10.AutoSize = true;
            //this.label10.Location = new System.Drawing.Point(94, 21);
            //this.label10.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            //this.label10.Name = "label10";
            //this.label10.Size = new System.Drawing.Size(35, 16);
            //this.label10.TabIndex = 23;
            //this.label10.Text = "Start";
            //// 
            //// tbtype
            //// 
            //this.tbtype.AccessibleDescription = "";
            //this.tbtype.Location = new System.Drawing.Point(316, 43);
            //this.tbtype.Name = "tbtype";
            //this.tbtype.Size = new System.Drawing.Size(56, 22);
            //this.tbtype.TabIndex = 17;
            //this.tbtype.Text = "U";
            //// 
            //// tbstart
            //// 
            //this.tbstart.Location = new System.Drawing.Point(95, 42);
            //this.tbstart.Margin = new System.Windows.Forms.Padding(4);
            //this.tbstart.Name = "tbstart";
            //this.tbstart.Size = new System.Drawing.Size(93, 22);
            //this.tbstart.TabIndex = 22;
            //// 
            //// label11
            //// 
            //this.label11.AutoSize = true;
            //this.label11.Location = new System.Drawing.Point(7, 21);
            //this.label11.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            //this.label11.Name = "label11";
            //this.label11.Size = new System.Drawing.Size(63, 16);
            //this.label11.TabIndex = 21;
            //this.label11.Text = "Object ID";
            //// 
            //// tbobjid1
            //// 
            //this.tbobjid1.Location = new System.Drawing.Point(11, 42);
            //this.tbobjid1.Margin = new System.Windows.Forms.Padding(4);
            //this.tbobjid1.Name = "tbobjid1";
            //this.tbobjid1.Size = new System.Drawing.Size(76, 22);
            //this.tbobjid1.TabIndex = 20;
            // 
            // btnx2s
            // 
            this.btnx2s.Location = new System.Drawing.Point(786, 123);
            this.btnx2s.Name = "btnx2s";
            this.btnx2s.Size = new System.Drawing.Size(99, 27);
            this.btnx2s.TabIndex = 23;
            this.btnx2s.Text = "XRisk-2-SAP";
            this.btnx2s.UseVisualStyleBackColor = true;
            this.btnx2s.Click += new System.EventHandler(this.btnx2s_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(931, 581);
            this.Controls.Add(this.btnx2s);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.tbPath);
            this.Controls.Add(this.btnLoadXML);
            this.Controls.Add(this.lblcompany);
            this.Controls.Add(this.tbXML);
            this.Controls.Add(this.label1);
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.Name = "Form1";
            this.Text = "Input";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.RichTextBox tbXML;
        private System.Windows.Forms.Label lblcompany;
        private System.Windows.Forms.Button btnLoadXML;
        private System.Windows.Forms.TextBox tbPath;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button btnCallService;
        private System.Windows.Forms.TextBox tbcriteria;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox tbmaxno;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox tbversion;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox tbobjid;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Button btnx2s;
        private System.Windows.Forms.Button btnSchema;
    }
}


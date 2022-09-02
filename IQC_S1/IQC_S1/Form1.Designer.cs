
namespace IQC_S1
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.label6 = new System.Windows.Forms.Label();
            this.listView1 = new System.Windows.Forms.ListView();
            this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader2 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader3 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader4 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader5 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader6 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.lblName = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.txtID = new System.Windows.Forms.TextBox();
            this.txtSSN = new System.Windows.Forms.TextBox();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.txtRL = new System.Windows.Forms.TextBox();
            this.txtuploadlink = new System.Windows.Forms.TextBox();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.label11 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label18 = new System.Windows.Forms.Label();
            this.label17 = new System.Windows.Forms.Label();
            this.label16 = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.totalReceivedQty = new System.Windows.Forms.Label();
            this.grn_text = new System.Windows.Forms.TextBox();
            this.total_received_Qty = new System.Windows.Forms.TextBox();
            this.txtUrgent = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Maiandra GD", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(90, 286);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(111, 18);
            this.label6.TabIndex = 80;
            this.label6.Text = "Share Drive Link";
            // 
            // listView1
            // 
            this.listView1.BackColor = System.Drawing.Color.Silver;
            this.listView1.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1,
            this.columnHeader2,
            this.columnHeader3,
            this.columnHeader4,
            this.columnHeader5,
            this.columnHeader6});
            this.listView1.Font = new System.Drawing.Font("Malgun Gothic", 6.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.listView1.HideSelection = false;
            this.listView1.Location = new System.Drawing.Point(13, 591);
            this.listView1.Name = "listView1";
            this.listView1.Size = new System.Drawing.Size(519, 238);
            this.listView1.TabIndex = 79;
            this.listView1.UseCompatibleStateImageBehavior = false;
            this.listView1.View = System.Windows.Forms.View.Details;
            // 
            // columnHeader1
            // 
            this.columnHeader1.Text = "Employee ID";
            this.columnHeader1.Width = 70;
            // 
            // columnHeader2
            // 
            this.columnHeader2.Text = "Serial Number";
            this.columnHeader2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.columnHeader2.Width = 80;
            // 
            // columnHeader3
            // 
            this.columnHeader3.Text = "Volex P/N";
            this.columnHeader3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.columnHeader3.Width = 80;
            // 
            // columnHeader4
            // 
            this.columnHeader4.Text = "Reel Length";
            this.columnHeader4.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.columnHeader4.Width = 80;
            // 
            // columnHeader5
            // 
            this.columnHeader5.Text = "Reel Incoming Date";
            this.columnHeader5.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.columnHeader5.Width = 130;
            // 
            // columnHeader6
            // 
            this.columnHeader6.Text = "Urgent";
            this.columnHeader6.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.columnHeader6.Width = 70;
            // 
            // lblName
            // 
            this.lblName.AutoSize = true;
            this.lblName.Font = new System.Drawing.Font("Myanmar Text", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblName.Location = new System.Drawing.Point(263, 148);
            this.lblName.Name = "lblName";
            this.lblName.Size = new System.Drawing.Size(145, 20);
            this.lblName.TabIndex = 76;
            this.lblName.Text = "Requester Employee Name";
            this.lblName.Click += new System.EventHandler(this.lblName_Click);
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Maiandra GD", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.Location = new System.Drawing.Point(90, 175);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(98, 18);
            this.label10.TabIndex = 72;
            this.label10.Text = "Serial Number";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Maiandra GD", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.Location = new System.Drawing.Point(90, 248);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(83, 18);
            this.label9.TabIndex = 71;
            this.label9.Text = "Reel Length";
            // 
            // label
            // 
            this.label.AutoSize = true;
            this.label.Font = new System.Drawing.Font("Maiandra GD", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label.Location = new System.Drawing.Point(90, 360);
            this.label.Name = "label";
            this.label.Size = new System.Drawing.Size(41, 18);
            this.label.TabIndex = 70;
            this.label.Text = "GRN";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Maiandra GD", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(90, 324);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(134, 18);
            this.label5.TabIndex = 69;
            this.label5.Text = "Reel Incoming Date";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Maiandra GD", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(90, 211);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(76, 18);
            this.label4.TabIndex = 68;
            this.label4.Text = "Volex P/N";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Maiandra GD", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(90, 124);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(159, 18);
            this.label3.TabIndex = 67;
            this.label3.Text = "Requester Employee ID";
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.Teal;
            this.button1.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Font = new System.Drawing.Font("Myanmar Text", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.ForeColor = System.Drawing.SystemColors.ButtonFace;
            this.button1.Location = new System.Drawing.Point(220, 483);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(98, 34);
            this.button1.TabIndex = 63;
            this.button1.Text = "Submit";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Palatino Linotype", 50.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(133, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(273, 90);
            this.label1.TabIndex = 62;
            this.label1.Text = "VOLEX";
            this.label1.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // timer1
            // 
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // txtID
            // 
            this.txtID.Location = new System.Drawing.Point(266, 122);
            this.txtID.Name = "txtID";
            this.txtID.Size = new System.Drawing.Size(199, 20);
            this.txtID.TabIndex = 81;
            this.txtID.TextChanged += new System.EventHandler(this.txtID_TextChanged_1);
            // 
            // txtSSN
            // 
            this.txtSSN.Location = new System.Drawing.Point(266, 173);
            this.txtSSN.Name = "txtSSN";
            this.txtSSN.Size = new System.Drawing.Size(199, 20);
            this.txtSSN.TabIndex = 82;
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(266, 207);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(199, 21);
            this.comboBox1.TabIndex = 83;
            this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // txtRL
            // 
            this.txtRL.Location = new System.Drawing.Point(266, 248);
            this.txtRL.Name = "txtRL";
            this.txtRL.Size = new System.Drawing.Size(199, 20);
            this.txtRL.TabIndex = 84;
            // 
            // txtuploadlink
            // 
            this.txtuploadlink.Location = new System.Drawing.Point(266, 283);
            this.txtuploadlink.Name = "txtuploadlink";
            this.txtuploadlink.Size = new System.Drawing.Size(199, 20);
            this.txtuploadlink.TabIndex = 85;
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Location = new System.Drawing.Point(266, 321);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(199, 20);
            this.dateTimePicker1.TabIndex = 86;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Myanmar Text", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label11.Location = new System.Drawing.Point(405, 565);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(34, 21);
            this.label11.TabIndex = 90;
            this.label11.Text = "time";
            this.label11.Click += new System.EventHandler(this.label11_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Rockwell", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(263, 563);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(143, 20);
            this.label2.TabIndex = 89;
            this.label2.Text = "Submission Time";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.ForeColor = System.Drawing.SystemColors.GrayText;
            this.label7.Location = new System.Drawing.Point(486, 835);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(53, 13);
            this.label7.TabIndex = 175;
            this.label7.Text = "@Rev1.1";
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Font = new System.Drawing.Font("Yu Gothic", 6.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label18.Location = new System.Drawing.Point(78, 553);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(88, 12);
            this.label18.TabIndex = 191;
            this.label18.Text = "(<2 working days)";
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Font = new System.Drawing.Font("Yu Gothic", 6.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label17.Location = new System.Drawing.Point(55, 537);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(90, 12);
            this.label17.TabIndex = 190;
            this.label17.Text = "(2-3 working days)";
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Font = new System.Drawing.Font("Yu Gothic", 6.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label16.Location = new System.Drawing.Point(62, 521);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(81, 12);
            this.label16.TabIndex = 189;
            this.label16.Text = "(4 working days)";
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Font = new System.Drawing.Font("Yu Gothic", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label15.Location = new System.Drawing.Point(13, 535);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(42, 14);
            this.label15.TabIndex = 188;
            this.label15.Text = "Urgent";
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Font = new System.Drawing.Font("Yu Gothic", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label14.Location = new System.Drawing.Point(13, 551);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(65, 14);
            this.label14.TabIndex = 187;
            this.label14.Text = "Top Urgent";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Font = new System.Drawing.Font("Yu Gothic", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label13.Location = new System.Drawing.Point(13, 520);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(47, 14);
            this.label13.TabIndex = 186;
            this.label13.Text = "Regular";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("Maiandra GD", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label12.Location = new System.Drawing.Point(90, 430);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(52, 18);
            this.label12.TabIndex = 192;
            this.label12.Text = "Urgent";
            // 
            // totalReceivedQty
            // 
            this.totalReceivedQty.AutoSize = true;
            this.totalReceivedQty.Font = new System.Drawing.Font("Maiandra GD", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.totalReceivedQty.Location = new System.Drawing.Point(88, 396);
            this.totalReceivedQty.Name = "totalReceivedQty";
            this.totalReceivedQty.Size = new System.Drawing.Size(176, 18);
            this.totalReceivedQty.TabIndex = 193;
            this.totalReceivedQty.Text = "Total Received Qty (Reel)";
            // 
            // grn_text
            // 
            this.grn_text.Location = new System.Drawing.Point(266, 360);
            this.grn_text.Name = "grn_text";
            this.grn_text.Size = new System.Drawing.Size(199, 20);
            this.grn_text.TabIndex = 194;
            // 
            // total_received_Qty
            // 
            this.total_received_Qty.Location = new System.Drawing.Point(267, 395);
            this.total_received_Qty.Name = "total_received_Qty";
            this.total_received_Qty.Size = new System.Drawing.Size(198, 20);
            this.total_received_Qty.TabIndex = 195;
            // 
            // txtUrgent
            // 
            this.txtUrgent.FormattingEnabled = true;
            this.txtUrgent.Items.AddRange(new object[] {
            "Regular",
            "Urgent",
            "Top Urgent"});
            this.txtUrgent.Location = new System.Drawing.Point(266, 430);
            this.txtUrgent.Name = "txtUrgent";
            this.txtUrgent.Size = new System.Drawing.Size(198, 21);
            this.txtUrgent.TabIndex = 196;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LightGray;
            this.ClientSize = new System.Drawing.Size(552, 851);
            this.Controls.Add(this.txtUrgent);
            this.Controls.Add(this.total_received_Qty);
            this.Controls.Add(this.grn_text);
            this.Controls.Add(this.totalReceivedQty);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.label18);
            this.Controls.Add(this.label17);
            this.Controls.Add(this.label16);
            this.Controls.Add(this.label15);
            this.Controls.Add(this.label14);
            this.Controls.Add(this.label13);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.dateTimePicker1);
            this.Controls.Add(this.txtuploadlink);
            this.Controls.Add(this.txtRL);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.txtSSN);
            this.Controls.Add(this.txtID);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.listView1);
            this.Controls.Add(this.lblName);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "IQC";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ListView listView1;
        private System.Windows.Forms.ColumnHeader columnHeader1;
        private System.Windows.Forms.ColumnHeader columnHeader2;
        private System.Windows.Forms.ColumnHeader columnHeader3;
        private System.Windows.Forms.ColumnHeader columnHeader4;
        private System.Windows.Forms.ColumnHeader columnHeader5;
        private System.Windows.Forms.ColumnHeader columnHeader6;
        private System.Windows.Forms.Label lblName;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.TextBox txtID;
        private System.Windows.Forms.TextBox txtSSN;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.TextBox txtRL;
        private System.Windows.Forms.TextBox txtuploadlink;
        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label totalReceivedQty;
        private System.Windows.Forms.TextBox grn_text;
        private System.Windows.Forms.TextBox total_received_Qty;
        private System.Windows.Forms.ComboBox txtUrgent;
    }
}


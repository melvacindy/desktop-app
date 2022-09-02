using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using System.IO;
using Outlook = Microsoft.Office.Interop.Outlook; //library for sending outlook email

namespace FA_S1
{
    public partial class Form1 : Form
    {
        
        public Form1()
        {
            InitializeComponent();
            timer1.Start();
        }

       
        public void CreateOutlookEmail()
        {

            var m = 0;

            List<string> email = new List<string>();
            List<string> pemail = new List<string>();
            List<string> volex_pn = new List<string>();

            try
            {
                Outlook.Application outlookApp = new Outlook.Application();
                Outlook.MailItem mailItem = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);
                //this.CreateItem(Outlook.OlItemType.olMailItem);
                mailItem.Subject = "Incoming Test Request-FA";

                using (var connection = new MySqlConnection("datasource=BAT-VM-PRD-001;port=3305;username=VB_Lab;password=Volex123"))
                using (var command = connection.CreateCommand())
                {
                    connection.Open();
                    command.CommandText = "select Email_FA from iqc.iqc_receiver";

                    using (var reader = command.ExecuteReader())

                        while (reader.Read())
                        {
                             
                            string output = reader.GetString(0);
                            string[] aaa = output.Split(';');
                            email.Add(aaa[0]);
                            string aaaa = email[m] + ";";
                            pemail.Add(aaaa);
                            m++;

                        }

                    string result = String.Join("; ", email);
                    mailItem.To = result;

                }
                if (checkedListBox1.CheckedItems.Count != 0)
                {
                    // If so, loop through all checked items and print results.  
                    string s = "";
                   
                    for (int x = 0; x < checkedListBox1.CheckedItems.Count; x++)
                    {
                        s = s.ToString() + checkedListBox1.CheckedItems[x].ToString() + ", ";
                    }

                    //condition for lane
                    string lane = "";
                    for (int x = 0; x < checkedListBox2.CheckedItems.Count; x++)
                    {
                        lane = lane.ToString() + checkedListBox2.CheckedItems[x].ToString() + " ";
                    }

                    mailItem.Body = "Hi Team!" + "\n" + "Please assist for the measurement below:" + "\n"
                                + "\n" + "∙ Sample S/N" + ": " + this.txtSSN.Text
                                + "\n" + "∙ FG P/N" + ": " + this.txtFG_pn.Text
                                + "\n" + "∙ Lot.No/Reel.No" + ": " + this.txtLotNo_ReelNo.Text 
                                + "\n" + "∙ Client" + ": " + this.comboBox1.Text
                                + "\n" + "∙ Device Under Test (DUT)" + ": " + this.comboBox2.Text
                                + "\n" + "∙ Measurement Parameter" + ": " + s
                                + "\n" + "∙ Lane" + ": " + lane
                                + "\n" + "∙ Urgent" + ": " + this.urgent.Text
                                + "\n" + "∙ Remarks" + ": " + this.txtRemarks.Text
                                + " \n" + "\n"
                                + "Thanks," + "\n"
                                + this.lblEname.Text + "\n" + this.txtEID.Text;

                    mailItem.Importance = Outlook.OlImportance.olImportanceLow;
                    mailItem.Display(false);
                    mailItem.Send();

                }
            }

            catch (Exception eX)
            {
                throw new Exception("Error Create an Outlook Email" + "\n" + "Please ensure your outlook mail already open/log in. " + "\n" + eX.Message);
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {

            if (txtEID.Text.Trim() == string.Empty)
            {
                MessageBox.Show("Please enter Employee ID", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return; // return because we don't want to run normal code of button click
            }
            if (txtSSN.Text.Trim() == string.Empty)
            {
                MessageBox.Show("Please enter Sample Serial Number", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (comboBox1.Text.Trim() == string.Empty)
            {
                MessageBox.Show("Please enter Client", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (txtFG_pn.Text.Trim() == string.Empty)
            {
                MessageBox.Show("Please enter FG P/N", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (txtLotNo_ReelNo.Text.Trim() == string.Empty)
            {
                MessageBox.Show("Please enter Lot No./Reel No.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (comboBox2.Text.Trim() == string.Empty)
            {
                MessageBox.Show("Please enter Device Under Test (DUT)", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (txtRemarks.Text.Trim() == string.Empty)
            {
                MessageBox.Show("Please enter Remarks", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (urgent.Text.Trim() == string.Empty)
            {
                MessageBox.Show("Please enter Urgent", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (txtEID.Text.Length < 8)
            {
                MessageBox.Show("Employee ID Min characters reached", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;

            }

            //reading csv file (employee id, name)
            using (var reader = new StreamReader(@".\EmployeeID.csv"))
            {
                List<string> EmployeeID = new List<string>();
                List<string> EmployeeName = new List<string>();

                //new variabel for contain string employee id and employee name
                var Employeelist = 0;

                while (!reader.EndOfStream)
                {

                    var line = reader.ReadLine(); //reading line
                    var values = line.Split(','); //storing data and split into 2 (employee id and name)
                    EmployeeID.Add(values[0]); // get [0] array = input type
                    EmployeeName.Add(values[1]);

                    // condition for calling employee name if employee id (input) equal to employee name
                    if (txtEID.Text == EmployeeID[Employeelist])
                    {
                        lblEname.Text = EmployeeName[Employeelist];
                        break;
                    }
                    else
                    {
                        Employeelist++;
                    }
                }
                if (txtEID.Text != EmployeeID[Employeelist])
                {
                    MessageBox.Show("EmployeeID is not registed", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;

                }
            }

            // Determine if there are any items checked.  
            if (checkedListBox1.CheckedItems.Count != 0)
            {
                string s = "";
                for (int x = 0; x < checkedListBox1.CheckedItems.Count; x++)
                {
                    s =  s.ToString() + checkedListBox1.CheckedItems[x].ToString() + ", " ;
                }

                
            //condition for lane
                string lane = "";
                for (int x = 0; x < checkedListBox2.CheckedItems.Count; x++)
                {
                    lane = lane.ToString() + checkedListBox2.CheckedItems[x].ToString() + " ";
                }

                
                //if (checkBox1.Checked)
                //{
                //    result = "Insertion Loss Sdd21";
                //}
                //string lane = string.Join(",", result.ToString()); 

                string message = "Please confirm before proceed." + "\n" + "\n" +
                                      "Sample Serial Number: " + txtSSN.Text + "\n" +
                                      "FG P/N: " + txtFG_pn.Text + "\n" +
                                      "LotNo./ReelNo.: " + txtLotNo_ReelNo.Text + "\n" +
                                      "Client: " + comboBox1.Text + "\n" +
                                      "Device Under Test (DUT): " + comboBox2.Text + "\n" +
                                      "Measurement Parameter: " + s + "\n" +
                                      "Lane: " + lane + "\n" +
                                      "Urgent: " + urgent.Text + "\n" +
                                      "Remark: " + txtRemarks.Text + "\n" + "\n" +
                                      "Do you want to Continue?" + "\n";

                string title = "Confirm";
                MessageBoxButtons buttons = MessageBoxButtons.YesNo;
                DialogResult result = MessageBox.Show(message, title, buttons);
                if (result == DialogResult.Yes)
                {
                    //string RLMeter = this.txtRL.Text + " Meter";
                    try
                    {
                        string MyConnection2 = "datasource=BAT-VM-PRD-001;port=3305;username=VB_Lab;password=Volex123";
                        string Query = "insert into iqc.fa(Employee_ID,Employee_Name,submission_time,sample_serial_number,FG_pn,LotNo_ReelNo,client,device_under_test,measurement_parameter,lane,urgent,remarks) " +
                        "values('" + this.txtEID.Text + "','" + this.lblEname.Text + "','" + this.time.Text + "','" + this.txtSSN.Text + "','" + this.txtFG_pn.Text + "','" +
                        this.txtLotNo_ReelNo.Text + "','" + this.comboBox1.Text + "','" + this.comboBox2.Text + "','" + s + "','" + lane + "','" + this.urgent.Text + "','" + this.txtRemarks.Text + "')";

                        MySqlConnection MyConn2 = new MySqlConnection(MyConnection2);
                        MySqlCommand MyCommand2 = new MySqlCommand(Query, MyConn2);
                        MySqlDataReader MyReader2;
                        MyConn2.Open();
                        MyReader2 = MyCommand2.ExecuteReader();
                        MessageBox.Show("Data Save.");
                        CreateOutlookEmail();

                        txtSSN.Text = "";
                        urgent.Text = "";
                        txtFG_pn.Text = "";
                        txtLotNo_ReelNo.Text = "";
                        comboBox1.Text = "";
                        comboBox2.Text = "";
                        
                        for (int i = 0; i < checkedListBox1.Items.Count; i++)
                        {
                            checkedListBox1.SetItemChecked(i, false);
                        }
                        for (int i = 0; i < checkedListBox2.Items.Count; i++)
                        {
                            checkedListBox2.SetItemChecked(i, false);
                        }
                        txtRemarks.Text = "";

                        while (MyReader2.Read()) { }

                        MyConn2.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                else
                { }

            }
        }

        private void time_Click(object sender, EventArgs e)
        {
           
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            //display submission time
            DateTime dateTime = DateTime.Now;
            this.time.Text = dateTime.ToString();
        }

        private void txtEID_TextChanged(object sender, EventArgs e)
        {
            //max 8 digit limit for employee id
            txtEID.MaxLength = 8;

            if (txtEID.Text.Length > 8)
            {
                MessageBox.Show("Employee ID Max characters reached", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
           
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                for (int i = 0; i < checkedListBox2.Items.Count; i++)
                {
                    checkedListBox2.SetItemChecked(i, true);
                }
            }
            else
            {
                for (int i = 0; i < checkedListBox2.Items.Count; i++)
                {
                    checkedListBox2.SetItemChecked(i, false);
                }
                return;
            }
            
        }
        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
           
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            
        }

        private void checkBox9_CheckedChanged(object sender, EventArgs e)
        {
            
        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {
          
        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            
        }

        private void checkBox8_CheckedChanged(object sender, EventArgs e)
        {

        }


    }
}

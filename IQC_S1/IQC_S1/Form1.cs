using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using MySql.Data.MySqlClient;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace IQC_S1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            timer1.Start();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //reading data from SE visual studio
            StreamReader sr = new StreamReader(@".\RawCable_Info.csv");
            string x = sr.ReadToEnd();
            string[] y = x.Split('\n');


            foreach (string s in y)
            {
                comboBox1.Items.Add(s);
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            //display submission time
            DateTime dateTime = DateTime.Now;
            this.label11.Text = dateTime.ToString();
        }

        private void txtID_TextChanged(object sender, EventArgs e)
        {
            //max 8 digit limit for employee id
            txtID.MaxLength = 8;

            if (txtID.Text.Length > 8)
            {
                MessageBox.Show("Employee ID Max characters reached", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //condition if-else input textbox cannot empty
            if (txtID.Text.Trim() == string.Empty)
            {
                MessageBox.Show("Please enter Employee ID", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return; // return because we don't want to run normal code of button click
            }
            if (txtSSN.Text.Trim() == string.Empty)
            {
                MessageBox.Show("Please enter Serial Number", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (comboBox1.Text.Trim() == string.Empty)
            {
                MessageBox.Show("Please enter Volex P/N", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (txtRL.Text.Trim() == string.Empty)
            {
                MessageBox.Show("Please enter Reel Length", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (dateTimePicker1.Text.Trim() == string.Empty)
            {
                MessageBox.Show("Please enter Reel Incoming Date", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (txtUrgent.Text.Trim() == string.Empty)
            {
                MessageBox.Show("Please enter Urgent", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (grn_text.Text.Trim() == string.Empty)
            {
                MessageBox.Show("Please enter GRN", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (total_received_Qty.Text.Trim() == string.Empty)
            {
                MessageBox.Show("Please enter Total Received QTY (Reel)", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            //min digit limit for employee id
            if (txtID.Text.Length < 8)
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
                    if (txtID.Text == EmployeeID[Employeelist])
                    {
                        lblName.Text = EmployeeName[Employeelist];
                    }
                    else
                    {
                        Employeelist++;
                    }
                }
                //if (txtID.Text != EmployeeID[Employeelist])
                //{
                //    MessageBox.Show("User is not registed", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //    return;

                //}
            }

            using (var reader1 = new StreamReader(@".\Raw_cable_parameter.csv"))
            {
                List<string> volex_pn = new List<string>();
                List<string> supplier_pn = new List<string>();
                List<string> awg = new List<string>();
                List<string> vendor_rc = new List<string>();
                List<string> pair = new List<string>();
                List<string> client = new List<string>();
                List<string> sample_length = new List<string>();
                List<string> irt = new List<string>();
                List<string> ts = new List<string>();
                List<string> dut = new List<string>();
                List<string> pn = new List<string>();
                List<string> a = new List<string>();


                //var x = (int)y;
                var x = 0;
                //var y = 0;

                while (!reader1.EndOfStream)
                {

                    var line = reader1.ReadLine();
                    var values = line.Split(',');
                    volex_pn.Add(values[0]);
                    supplier_pn.Add(values[1]);
                    awg.Add(values[2]);
                    vendor_rc.Add(values[3]);
                    pair.Add(values[4]);
                    client.Add(values[5]);
                    sample_length.Add(values[6]);
                    irt.Add(values[7]);
                    ts.Add(values[8]);
                    dut.Add(values[9]);

                    //new var for comboBox1 (input) and delete new line '\r' or '\n'
                    string volexpn = comboBox1.Text;
                    string volexpn2 = volexpn.ToString().TrimEnd('\r', '\n');


                    if (volexpn2 == volex_pn[x])
                    {
                        //y.Add(x);
                        break;
                    }

                    else
                    {
                        x++;
                    }

                }

                // make a message box before enter the data to database mysql                
                string message = "Please confirm before proceed." + "\n" + "\n" +
                                 "Employee ID: " + txtID.Text + "\n" +
                                 "Reel Income Date: " + dateTimePicker1.Text + "\n" +
                                 "Reel Length: " + txtRL.Text + "\n" +
                                 "Volex P/N: " + comboBox1.Text + "\n" +
                                 "Urgent: " + txtUrgent.Text + "\n" +
                                 "GRN: " + grn_text.Text + "\n" +
                                 "Total Received Quantity (Reel): " + total_received_Qty.Text + "\n" +
                                 "Sample Serial Number: " + txtSSN.Text + "\n" + "\n" +
                                 "Do you want to Continue?" + "\n";

                string title = "Confirm";
                MessageBoxButtons buttons = MessageBoxButtons.YesNo;
                DialogResult result = MessageBox.Show(message, title, buttons);
                if (result == DialogResult.Yes)
                {
                    string RLMeter = this.txtRL.Text + " Meter";
                    try
                    {
                        string MyConnection2 = "datasource=BAT-VM-PRD-001;port=3305;username=VB_Lab;password=Volex123";
                        string Query = "insert into iqc.user(requester_employeeid,requester_employeename,volex_pn,supplier_pn,awg,vendor_raw_cable,pair_no,client,sample_length,impedance_rise_time,test_spesification,device_under_test,submission_time,reel_length,urgent,GRN,total_received_qty,sample_serial_number,reel_incomingdate,sharedrive_link) " +
                        "values('" + this.txtID.Text + "','" + this.lblName.Text + "','" + volex_pn[x] + "','" + supplier_pn[x] + "','" + awg[x] + "','" +
                        vendor_rc[x] + "','" + pair[x] + "','" + client[x] + "','" + sample_length[x] + "','" + irt[x] + "', '" + ts[x] + "','" + dut[x]
                        + "','" + DateTime.Now + "','" + RLMeter + "','" + this.txtUrgent.Text + "','" + this.grn_text.Text + "','" + this.total_received_Qty.Text + "','" + this.txtSSN.Text + "','" + dateTimePicker1.Value.ToShortDateString() + "','" + txtuploadlink.Text + "')";


                        string y = sample_length[x];
                        //MessageBox.Show("Please ensure your outlook mail already open.", "WARNING", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
                        CreateOutlookEmail(x, y);


                        MySqlConnection MyConn2 = new MySqlConnection(MyConnection2);
                        MySqlCommand MyCommand2 = new MySqlCommand(Query, MyConn2);
                        MySqlDataReader MyReader2;
                        MyConn2.Open();
                        MyReader2 = MyCommand2.ExecuteReader();
                        MessageBox.Show("Data Save.");

                        //make list view for data input
                        ListViewItem item = new ListViewItem(txtID.Text);
                        item.SubItems.Add(txtSSN.Text);
                        item.SubItems.Add(comboBox1.Text);
                        item.SubItems.Add(txtRL.Text);
                        item.SubItems.Add(dateTimePicker1.Text);
                        item.SubItems.Add(txtUrgent.Text);
                        listView1.Items.Add(item);

                        txtSSN.Text = "";
                        comboBox1.Text = "";
                        txtRL.Text = "";
                        txtuploadlink.Text = "";
                        txtUrgent.Text = "";
                        grn_text.Text = "";
                        total_received_Qty.Text = "";

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

        private void txtRL_TextChanged(object sender, EventArgs e)
        {
            //4 digit limit for reel length textbox
            txtRL.MaxLength = 4;

            if (txtRL.Text.Length > txtRL.MaxLength)
            {
                MessageBox.Show("Reel Length Max characters reached", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void CreateOutlookEmail(int x, string y)
        {

            var s = 0;

            List<string> email = new List<string>();
            List<string> pemail = new List<string>();
            //List<string> sample_length = new List<string>();
            List<string> volex_pn = new List<string>();

            try
            {
                Outlook.Application outlookApp = new Outlook.Application();
                Outlook.MailItem mailItem = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);
                //this.CreateItem(Outlook.OlItemType.olMailItem);
                mailItem.Subject = "Incoming Test Request-IQC";

                using (var connection = new MySqlConnection("datasource=BAT-VM-PRD-001;port=3305;username=VB_Lab;password=Volex123"))
                using (var command = connection.CreateCommand())
                {
                    connection.Open();
                    command.CommandText = "select Email from iqc.iqc_receiver";

                    using (var reader = command.ExecuteReader())

                        while (reader.Read())
                        {
                            string output = reader.GetString(0);
                            string[] aaa = output.Split(';');
                            email.Add(aaa[0]);
                            string aaaa = email[s] + ";";
                            pemail.Add(aaaa);
                            s++;

                        }

                    string result = String.Join("; ", email);
                    mailItem.To = result;

                }

                mailItem.Body = "Hi Team!" + "\n" + "Please assist for the measurement below:" + "\n"
                                + "\n" + "∙ Sample S/N" + ": " + this.txtSSN.Text
                                + "\n" + "∙ Volex P/N" + ": " + this.comboBox1.Text
                                + "\n" + "∙ Reel Length" + ": " + this.txtRL.Text + " Meter"
                                + "\n" + "∙ Sample Length" + ": " + y
                                + "\n" + "∙ Share Drive Link" + ": " + this.txtuploadlink.Text
                                + "\n" + "∙ GRN" + ": " + this.grn_text.Text
                                + "\n" + "∙ Total Received Qty (Reel)" + ": " + this.total_received_Qty.Text
                                + "\n" + "∙ Urgent" + ": " + this.txtUrgent.Text
                                + " \n" + "\n"
                                + "Thanks," + "\n"
                                + this.lblName.Text + "\n" + this.txtID.Text;

                mailItem.Importance = Outlook.OlImportance.olImportanceLow;
                mailItem.Display(false);
                mailItem.Send();

            }

            catch (Exception eX)
            {
                throw new Exception("Error Create an Outlook Email" + "\n" + "Please ensure your outlook mail already open/log in. " + "\n" + eX.Message);
            }
        }

        private void lblName_Click(object sender, EventArgs e)
        {

        }

        private void txtID_TextChanged_1(object sender, EventArgs e)
        {
            //max 8 digit limit for employee id
            txtID.MaxLength = 8;

            if (txtID.Text.Length > 8)
            {
                MessageBox.Show("Employee ID Max characters reached", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}

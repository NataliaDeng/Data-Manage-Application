using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using MySql.Data.MySqlClient;  

namespace New
{
    public partial class Upload : Form
    {
        public Upload()
        {
            InitializeComponent();
        }
        // Date, TripID, PatientLastName,PatientFirstName, PickupTime,ArriveTime, AddressType, Address1, ApartmentNo, City1, State1,Zip1, PatientPhone, Facility, Address2, City2, State2, Zip2, FacilityPhone, AppTime, PatientDOB, SSN, Age,Gender, AdditionalInfo 
        public void ExcelToDS1(string path)  //insert into patient table
        {
            //label4.Text = path;
            String fileType = System.IO.Path.GetExtension(path);
            String strConn;
            String sheetNo = comboBox1.Text;
            if (fileType == ".xls")
            {
                strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + path + ";" + ";Extended Properties=\"Excel 8.0;HDR=YES;IMEX=2\"";
            }
            else
            {
                strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + path + ";" + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=2\"";
            }
            OleDbConnection conn = new OleDbConnection(strConn);
            conn.Open();
            string strExcel = "";
            OleDbDataAdapter myCommand = null;
            DataSet ds = null;
            strExcel = "select * from [Sheet1$];";//[Sheet1$]  [Sheet" + sheetNo[0] + "$]  select PatientLastName,PatientFirstName, AddressType, Address1, ApartmentNo, City1, State1,Zip1, PatientPhone, from [Sheet1$];";//[Sheet1$]  [Sheet" + sheetNo[0] + "$]
            
            myCommand = new OleDbDataAdapter(strExcel, strConn);
            DataTable table1 = new DataTable();
            ds = new DataSet();
            myCommand.Fill(table1);
            myCommand.Fill(ds);
            //DataSet DS = new DataSet();
            //myCommand.Fill(DS);
            //dataGridView1.DataSource = ds.Tables[0];
            conn.Close();

            DataRow[] dr = ds.Tables[0].Select();
            //label2.Text = dr.Length.ToString();
            List<string> list2 = (from DataRow row in dr select String.Format("insert into asap.appointment values('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}')", row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8])).ToList(); //,'{13} 

            String mysqlStr = "server=localhost;user id=root;password=Happy111$@ ;persistsecurityinfo=True;database=asap";
            MySqlConnection conn2 = new MySqlConnection(mysqlStr);
            conn2.Open();

            //label2.Text = list.Count.ToString();
            foreach (String item in list2)
            {
                MySqlCommand comn = new MySqlCommand(item, conn2);
                comn.ExecuteNonQuery();
                //if (comn.ExecuteNonQuery() > 0)
                // 
            }
            conn2.Close();
        }

        public void ExcelToDS2(string path) //DataSet  insert into Appointment Table
        {
            String fileType = System.IO.Path.GetExtension(path);
            String strConn2;
           // String sheetNo = comboBox1.Text;
            if (fileType == ".xls")
            {
                strConn2 = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + path + ";" + ";Extended Properties=\"Excel 8.0;HDR=YES;IMEX=2\"";
            }
            else
            {
                strConn2 = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + path + ";" + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=2\"";
            }
            OleDbConnection con = new OleDbConnection(strConn2);
            con.Open();
            string strExcel2 = "";
            OleDbDataAdapter myCom = null;
            DataSet ds = null;
            strExcel2 = "select * from [Sheet1$];";// [Sheet" + sheetNo[0] + "$]
            myCom = new OleDbDataAdapter(strExcel2, strConn2);
            DataTable table = new DataTable();
            ds = new DataSet();
            myCom.Fill(table);
            myCom.Fill(ds);
            //dataGridView1.DataSource = ds.Tables[0];
            con.Close();
            
            DataRow[] dr = ds.Tables[0].Select();
            //label2.Text = dr.Length.ToString();
            List<string> list = (from DataRow row in dr select String.Format("insert into asap.appointment values('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}')", row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9], row[10])).ToList(); //,'{13} 
            String mysqlStr = "server=localhost;user id=root;password=Happy111$@ ;persistsecurityinfo=True;database=asap";
            MySqlConnection con2 = new MySqlConnection(mysqlStr);
            con2.Open();

            //label2.Text = list.Count.ToString();
            foreach (String item in list)
            {
                MySqlCommand com = new MySqlCommand(item, con2);
                com.ExecuteNonQuery();
                //if (comn.ExecuteNonQuery() > 0)
            }
            con2.Close();
            
            //return ds;
        }  
        private void button1_Click(object sender, EventArgs e)
        {
            //upload一张excel，split成Patient和Appointment两张表

            if(comboBox1.Text=="patient")
                ExcelToDS1(textBox1.Text); // DataSet ds = 
            else 
                ExcelToDS2(textBox1.Text); 
            
            MessageBox.Show("Success!");
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openfile = new OpenFileDialog();
            openfile.Filter = "所有文件(*.*)|*.*";
            if (openfile.FilterIndex == 1 && openfile.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openfile.FileName;
            }       
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Hide();
            this.Owner.Show();
        }
    }
}

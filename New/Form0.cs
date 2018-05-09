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
    public partial class Form0 : Form
    {
        public Form0()
        {
            InitializeComponent();
        }
        String mysqlStr = "Server=localhost;Port=3306;Database=asap;Uid=root;Pwd=Happy111$@;"; //Happy111$@
        private void button3_Click(object sender, EventArgs e)
        {        
            this.Hide();
            this.Owner.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openfile = new OpenFileDialog();
            openfile.Filter = "All files(*.*)|*.*";
            if (openfile.FilterIndex == 1 && openfile.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openfile.FileName;
            }      
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
                ExcelToDS1(textBox1.Text); 
            
                ExcelToDS2(textBox1.Text);

            MessageBox.Show("Success!","Success");
        }
        public void ExcelToDS1(string path)  //insert into patient table
        {
            String fileType = System.IO.Path.GetExtension(path);
            String strConn;            
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
            
            //获取Excel的第一个Sheet名称
            System.Data.DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            String sheetName = schemaTable.Rows[0]["TABLE_NAME"].ToString().Trim();
            string strExcel = "";
            OleDbDataAdapter myCommand = null;
            DataSet ds = null;
            strExcel = "select PatientLastName, PatientFirstName, Gender,Address1, ApartmentNo, PatientPhone, PatientDOB, SSN from [" + sheetName + "]";//[Sheet1$]  [Sheet" + sheetNo[0] + "$]  select PatientLastName,PatientFirstName, AddressType, Address1, ApartmentNo, City1, State1,Zip1, PatientPhone, from [Sheet1$];";//[Sheet1$]  [Sheet" + sheetNo[0] + "$]

            myCommand = new OleDbDataAdapter(strExcel, strConn);
            DataTable table1 = new DataTable();
            ds = new DataSet();
            myCommand.Fill(table1);
            myCommand.Fill(ds);
            //dataGridView1.DataSource = ds.Tables[0];
            conn.Close();

            DataRow[] dr = ds.Tables[0].Select();
            //label2.Text = dr.Length.ToString();
            List<string> list2 = (from DataRow row in dr select String.Format("insert into asap.patient (LastName, FirstName, Gender,Address1, ApartmentNo, Phone,DateOfBirth, SSN) values('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}')", row[0], row[1], row[2], row[3], row[4], row[5], row[6].ToString(), row[7])).ToList(); //,'{13} 

            MySqlConnection conn2 = new MySqlConnection(mysqlStr);
                                            
            try
            {
                conn2.Open();

                foreach (String item in list2)
                {
                    MySqlCommand comn = new MySqlCommand(item, conn2);
                    comn.ExecuteNonQuery();
                    //if (comn.ExecuteNonQuery() > 0)
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Are you sure you want to upload this file?", "Confirm");
            }
            finally
            {
                conn2.Close();
            }  
      
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
            //获取Excel的第一个Sheet名称
            System.Data.DataTable schemaTable = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            String sheetName = schemaTable.Rows[0]["TABLE_NAME"].ToString().Trim(); 
            string strExcel2 = "";
            OleDbDataAdapter myCom = null;
            DataSet ds = null;
            strExcel2 = "select DoctorName,AppDate, TripID, PatientLastName, PatientFirstName, AppTime,FinishTime, Facility,Address2, PhoneNumber, AdditionalInfo from [" + sheetName + "]";// [Sheet" + sheetNo[0] + "$]
            myCom = new OleDbDataAdapter(strExcel2, strConn2);
            DataTable table = new DataTable();
            ds = new DataSet();
            myCom.Fill(table);
            myCom.Fill(ds);

            con.Close();

            DataRow[] dr = ds.Tables[0].Select();
            //label2.Text = dr.Length.ToString();
            List<string> list = (from DataRow row in dr select String.Format("insert into asap.appointment (DoctorName,AppDate, TripID, PatientLastName, PatientFirstName, AppointmentTime,FinishTime, Facility,Address2, PhoneNumber, AdditionalInfo) values('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}')", row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9], row[10])).ToList(); //,'{8}','{9}','{10}' ,row[9], row[10]
            //String mysqlStr = "server=localhost;user id=root;password=Happy123$@ ;persistsecurityinfo=True;database=asap";
            MySqlConnection con2 = new MySqlConnection(mysqlStr);
            try
            {
                con2.Open();

                foreach (String item in list)
                {
                    MySqlCommand com = new MySqlCommand(item, con2);
                    com.ExecuteNonQuery();
                    //if (comn.ExecuteNonQuery() > 0)
                }
            }
            catch (MySql.Data.MySqlClient.MySqlException e)
            {
                //MessageBox.Show(e.ToString(),"Confirm");
            }
            finally
            {
                con2.Close();
            }
            

        }  
    }
}

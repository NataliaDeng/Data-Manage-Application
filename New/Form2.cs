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

namespace New
{
    public partial class Form2 : Form
    {
        static String mysqlStr = "Server=localhost;Port=3306;Database=asap;Uid=root;Pwd=Happy111$@;";//"server=localhost;user id=root;password=mysql_djz;persistsecurityinfo=True;database=asap"; //Happy111$@ 
        MySqlConnection con = new MySqlConnection(mysqlStr);
        //String add,zip;
        public Form2(String lname,String fname)
        {
            InitializeComponent();
            String last=lname, first=fname;
            con.Open();
            //找病人的信息
            String sqlSel;
            sqlSel = "select * from patient where LastName= '"+last+ "' and FirstName ='"+ first +"';";
            MySqlCommand com = new MySqlCommand(sqlSel, con);
            MySqlDataAdapter da = new MySqlDataAdapter(com);

            DataSet DS = new DataSet();
            da.Fill(DS);
            lastname.Text = DS.Tables[0].Rows[0]["LastName"].ToString();// DS.Tables[0].Rows[0][4];
            firstname.Text = DS.Tables[0].Rows[0]["FirstName"].ToString();
           
            //找Dr的信息，然后按行遍历添加到combobox中
            String sqlSel2;
            sqlSel2 = "SELECT distinct DoctorName, facility, Address2,PhoneNumber FROM asap.appointment;";//"select distinct DoctorName FROM asap.appointment;";  
            MySqlCommand com2 = new MySqlCommand(sqlSel2, con);
            MySqlDataAdapter da2 = new MySqlDataAdapter(com2);
            DataSet DS2 = new DataSet();
            da2.Fill(DS2);
            
            DataRowCollection DC = DS2.Tables[0].Rows;
            //Fill comboBox and Label//
            comboBox1.Text = DS2.Tables[0].Rows[0]["DoctorName"].ToString();
            DrphoneNo.Text = DS2.Tables[0].Rows[0]["PhoneNumber"].ToString();
            facility.Text = DS2.Tables[0].Rows[0]["Facility"].ToString();
            //add = DS2.Tables[0].Rows[0]["Address2"].ToString();
            //zip = DS2.Tables[0].Rows[0]["Zip"].ToString();
            facilityaddress.Text = DS2.Tables[0].Rows[0]["Address2"].ToString();
            //DrFirstname = DS2.Tables[0].Rows[0]["FirstName"].ToString();

            for (int i = 0; i < DC.Count; i++)
            {
                comboBox1.Items.Add(DS2.Tables[0].Rows[i]["DoctorName"].ToString()); //DoctorName
            }

            con.Close();
            
        }

        private void button2_Click(object sender, EventArgs e)
        {          
            con.Open();
            String sqlSel;
            String addInfo = textBox4.Text;
            String TripID = textBox1.Text;
            String Drname = comboBox1.Text; //Dr's last name
            //String year, month, day;
            String date = dateTimePicker1.Text;
            String AppDate = "";
            for (int i = 0; i < date.Length; i++)
            {
                if (Char.IsDigit(date[i]))
                {
                    AppDate += date[i];
                }
                else
                {
                    if (i != date.Length - 1)
                        AppDate += "-";
                }
            }
            
            //insert info
            //sqlSel = "insert into asap.appointment (DoctorName,AppDate, TripID,PatientLastName, PatientFirstName, AppointmentTime,FinishTime, Facility,Address2, PhoneNumber, AdditionalInfo) values('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}')";
            sqlSel = "insert into appointment(DoctorName,AppDate, TripID,PatientLastName, PatientFirstName, AppointmentTime,FinishTime, Facility,Address2, PhoneNumber, AdditionalInfo) values ('" + comboBox1.Text + "','"+AppDate+ "','" + TripID + "','" + lastname.Text + "','" + firstname.Text + "','" + textBox2.Text + "','" + textBox3.Text + "','" + facility.Text + "','" + facilityaddress + "','" + DrphoneNo.Text +"','" + addInfo +"');";
            //facilityaddress.Text = "values (" + PatientID + ",'" + lastname.Text + "','" + firstname.Text + "','" + textBox2.Text + "','" + textBox3.Text + "','" + facility.Text + "','" + facilityaddress.Text + "','" + comboBox1.Text + "','" + DrFirstname + "','" + DrphoneNo.Text + "';";
            MySqlCommand com2 = new MySqlCommand(sqlSel, con);
            //MySqlDataAdapter da2 = new MySqlDataAdapter(com2);
            if (com2.ExecuteNonQuery() > 0)
                MessageBox.Show("Success!");
            con.Close();
        }

        private void button1_Click(object sender, EventArgs e) // Add new Dr
        {

            con.Open();
                  

            con.Close();

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {        
            con.Open();
            String sql;
            sql = "select * FROM asap.appointment where DoctorName = '" + comboBox1.Text + "';"; //appointment表格里，Address是FacilityAddress
            MySqlCommand com = new MySqlCommand(sql, con);
            MySqlDataAdapter da = new MySqlDataAdapter(com);
            DataSet DS = new DataSet();
            da.Fill(DS);
            DrphoneNo.Text = DS.Tables[0].Rows[0]["PhoneNumber"].ToString();
            facility.Text = DS.Tables[0].Rows[0]["Facility"].ToString();
            
            //zip = DS.Tables[0].Rows[0]["Zip"].ToString();
            facilityaddress.Text = DS.Tables[0].Rows[0]["Address2"].ToString(); ;//appointment表格里叫FacilityAddress
            //DrFirstname = DS.Tables[0].Rows[0]["FirstName"].ToString(); 
            con.Close();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            this.Hide();
            this.Owner.Show();
        }


    }
}

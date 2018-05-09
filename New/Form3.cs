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
    public partial class Form3 : Form
    {
        static String mysqlStr = "Server=localhost;Port=3306;Database=asap;Uid=root;Pwd=Happy111$@;";//"server=localhost;user id=root;password=Happy111$@;persistsecurityinfo=True;database=asap"; // mysql_djz
        MySqlConnection con = new MySqlConnection(mysqlStr);
        int btnclick = 0;
        int add = 0;
        public Form3()
        {
            InitializeComponent();
        }
        private void Form3_Load(object sender, EventArgs e)
        {
            con.Open();
            String sqlSel;
            sqlSel = "SELECT * FROM asap.appointment order by PatientLastName;";
            MySqlCommand com = new MySqlCommand(sqlSel, con);
            MySqlDataAdapter da = new MySqlDataAdapter(com);

            DataSet DS = new DataSet();
            da.Fill(DS);
            dataGridView1.DataSource = DS.Tables[0];

            con.Close();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //String mysqlStr = "server=localhost;user id=root;password=Happy111$@ ;persistsecurityinfo=True;database=asap";
            //MySqlConnection con = new MySqlConnection(mysqlStr); 
            String sqlSel="";
            if (add == 0)
            {
                if (comboBox1.Text == "Patient's last name")
                    sqlSel = "select * from asap.appointment where PatientLastName = '" + textBox1.Text + "';";
                else if (comboBox1.Text == "Patient's first name")
                    sqlSel = "select * from asap.appointment where PatientFirstName = '" + textBox1.Text + "';";
                else //(comboBox1.Text == "Trip ID")
                    sqlSel = "select * from asap.appointment where TripID = '" + textBox1.Text + "';";

            }

            else if (add == 1) //两个搜索条件
            {
                if (comboBox1.Text == "Patient's last name")
                    sqlSel += " and LastName = '" + textBox1.Text + "'";
                else if (comboBox1.Text == "Patient's first name")
                    sqlSel += " and FirstName = '" + textBox1.Text + "'";
                else
                    sqlSel += " and TripID = '" + textBox1.Text + "'";

            }

            //创建SqlCommand对象
            MySqlCommand com = new MySqlCommand(sqlSel, con);
            MySqlDataAdapter da = new MySqlDataAdapter(com);

            DataSet DS = new DataSet();
            da.Fill(DS);
            dataGridView1.DataSource = DS.Tables[0];
           
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Hide();
            this.Owner.Show();
        }

        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            if (btnclick != 0)
                add++;
        }
        private void dataGridView1_RowValidated(object sender, DataGridViewCellEventArgs e)
        {
            DataTable changes = ((DataTable)dataGridView1.DataSource).GetChanges();
            try
            {
                if (changes != null)
                {
                    MySqlDataAdapter mySqlDataAdapter = new MySqlDataAdapter("select * from appointment", con);
                    MySqlCommandBuilder mcb = new MySqlCommandBuilder(mySqlDataAdapter);
                    mySqlDataAdapter.UpdateCommand = mcb.GetUpdateCommand();
                    mySqlDataAdapter.Update(changes);
                    ((DataTable)dataGridView1.DataSource).AcceptChanges();
                }
            }
            catch (MySqlException ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
    }
}

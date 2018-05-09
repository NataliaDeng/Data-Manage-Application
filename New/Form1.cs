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
    public partial class Form1 : Form
    {
        static String mysqlStr = "Server=localhost;Port=3306;Database=asap;Uid=root;Pwd=Happy111$@;";//"server=localhost;user id=root;password=mysql_djz;persistsecurityinfo=True;database=asap"; //mysql_djz
        MySqlConnection con = new MySqlConnection(mysqlStr);
        String sqlSel;    
        String lastname,firstname;
        int btnclick = 0;
        int add = 0;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            con.Open();
            String sql;
            sql = "select * from patient order by LastName;";
            MySqlCommand com = new MySqlCommand(sql, con);
            MySqlDataAdapter da = new MySqlDataAdapter(com);

            DataSet DS = new DataSet();
            da.Fill(DS);
            dataGridView1.DataSource = DS.Tables[0];
            dataGridView1.Visible = true;

            con.Close();

            AutoSizeColumn(dataGridView1);    
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (btnclick!=0)
                add++;          
        }

        private void button1_Click(object sender, EventArgs e)
        {                
            //打开数据库
            btnclick++;
            con.Open();
            if (add == 0) 
            {
                if (comboBox1.Text == "Patient's last name") 
                    sqlSel = "select * from asap.patient where LastName = '" + textBox1.Text + "'";
                else if (comboBox1.Text == "Patient's first name")
                    sqlSel = "select * from asap.patient where FirstName = '" + textBox1.Text + "'";
                else if (comboBox1.Text == "Patient's birthday")
                    sqlSel = "select * from asap.patient where DateOfBirth = '" + textBox1.Text + "'";
                else
                    sqlSel = "select * from asap.patient where SSN ='" + textBox1.Text + "'";
                
            }
            
            if (add == 1) //两个搜索条件
            {
                if (comboBox1.Text == "Patient's last name")
                    sqlSel+= " and LastName = '" + textBox1.Text + "'";
                else if (comboBox1.Text == "Patient's first name")
                    sqlSel += " and FirstName = '" + textBox1.Text + "'";
                else if (comboBox1.Text == "Patient's birthday")
                    sqlSel += " and DateOfBirth = '" + textBox1.Text + "'";
                else
                    sqlSel += " and SSN = '" + textBox1.Text + "'";
                
            }
            else if (add == 2) //3个搜索条件
            {
                if (comboBox1.Text == "Patient's last name")
                    sqlSel += " and LastName = '" + textBox1.Text + "'";
                else if (comboBox1.Text == "Patient's first name")
                    sqlSel += " and FirstName = '" + textBox1.Text + "'";
                else if (comboBox1.Text == "Patient's birthday")
                    sqlSel += " and DateOfBirth = '" + textBox1.Text + "'";
                else
                    sqlSel += " and SSN = '" + textBox1.Text + "'";
            }
            
            //创建SqlCommand对象
            MySqlCommand com = new MySqlCommand(sqlSel, con);
            MySqlDataAdapter da = new MySqlDataAdapter(com);

            DataSet DS = new DataSet();
            da.Fill(DS);
            dataGridView1.DataSource = DS.Tables[0];
            
            if (DS.Tables[0].Rows.Count == 1)
            {
                lastname = DS.Tables[0].Rows[0]["LastName"].ToString();
                firstname = DS.Tables[0].Rows[0]["FirstName"].ToString();
                button2.Visible = true;
            }
            
            con.Close();
            AutoSizeColumn(dataGridView1);   
        }
        private void dataGridView1_RowValidated(object sender, DataGridViewCellEventArgs e)
        {
            DataTable changes = ((DataTable)dataGridView1.DataSource).GetChanges();
            try
            {
                if (changes != null)
                {
                    MySqlDataAdapter mySqlDataAdapter = new MySqlDataAdapter("select * from patient order by LastName", con);
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
        /*
        private void DataGridView1_UserDeletedRow(Object sender, DataGridViewRowEventArgs e)
        {

            System.Text.StringBuilder messageBoxCS = new System.Text.StringBuilder();
            messageBoxCS.AppendFormat("{0} = {1}", "Row", e.Row);
            messageBoxCS.AppendLine();
            MessageBox.Show(messageBoxCS.ToString(), "UserDeletedRow Event");
        }
        */
        private void button2_Click(object sender, EventArgs e)
        {
            //MessageBox.Show(pID.ToString());
            this.Hide();
            Form2 f2 = new Form2(lastname,firstname);
            f2.Owner = this;
            f2.Show();
        }
        /// <summary>
        /// 使DataGridView的列自适应宽度
        /// </summary>
        /// <param name="dgViewFiles"></param>
        private void AutoSizeColumn(DataGridView dgViewFiles)
        {
            int width = 0;
            //使列自使用宽度
            //对于DataGridView的每一个列都调整
            for (int i = 0; i < dgViewFiles.Columns.Count; i++)
            {
                //将每一列都调整为自动适应模式
                dgViewFiles.AutoResizeColumn(i, DataGridViewAutoSizeColumnMode.AllCells);
                //记录整个DataGridView的宽度
                width += dgViewFiles.Columns[i].Width;
            }
            //判断调整后的宽度与原来设定的宽度的关系，如果是调整后的宽度大于原来设定的宽度，
            //则将DataGridView的列自动调整模式设置为显示的列即可，
            //如果是小于原来设定的宽度，将模式改为填充。
            if (width < dgViewFiles.Size.Width)
            {
                dgViewFiles.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
            }
            else
            {
                dgViewFiles.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            }
            //冻结某列 从左开始 0，1，2
            dgViewFiles.Columns[1].Frozen = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Hide();
            this.Owner.Show();
        }

    }
}

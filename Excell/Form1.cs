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
 
namespace Excell
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
        OleDbConnection OleDbcon;
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx";

            openFileDialog.ShowDialog();

            if (!string.IsNullOrEmpty(openFileDialog.FileName))

            {
                textBox1.Text = openFileDialog.FileName.ToString();

                OleDbcon = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + openFileDialog.FileName + ";Extended Properties=Excel 12.0;");

                OleDbcon.Open();

                DataTable dt = OleDbcon.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                OleDbcon.Close();

                comboBox1.Items.Clear();

                for (int i = 0; i < dt.Rows.Count; i++)

                {

                    String sheetName = dt.Rows[i]["TABLE_NAME"].ToString();

                    sheetName = sheetName.Substring(0, sheetName.Length - 1);

                    comboBox1.Items.Add(sheetName);

                }

            }

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            OleDbDataAdapter oledbDa = new OleDbDataAdapter("Select * from [" + comboBox1.Text + "$]", OleDbcon);

            DataTable dt = new DataTable();

            oledbDa.Fill(dt);

            dataGridView1.DataSource = dt;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialog = new OpenFileDialog();
            OpenFileDialog.ShowDialog();
            string path = OpenFileDialog.FileName;
        }
    }
}

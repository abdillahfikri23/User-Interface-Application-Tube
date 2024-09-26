using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace WindowsFormsApplication2
{
    public partial class DataSql : Form
    {
        SqlConn data = new SqlConn();
        public DataSql()
        {
            InitializeComponent();
        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void DataSql_Load(object sender, EventArgs e)
        {
            Calibration calibration = new Calibration();
            calibration.DataGridViewReference = dataGridAppTube;
            dataGridAppTube.CellFormatting += dataGridView1_CellFormatting;
            DataTable dt = data.Select();
            dataGridAppTube.DataSource = dt;
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            foreach (DataGridViewRow r in dataGridAppTube.Rows)
            {
                if (e.ColumnIndex == 19 && e.Value != null && !DBNull.Value.Equals(e.Value))
                {
                    string cellValue = e.Value.ToString();
                    if (cellValue == "PASS      ")
                    {
                        e.CellStyle.BackColor = Color.Green; // Warna hijau untuk PASS
                    }
                    else if (cellValue == "FAIL      ")
                    {
                        e.CellStyle.BackColor = Color.Red; // Warna merah untuk FAIL
                    }
                }
            }
        }

        void showTipeDataSql()
        {
            string connectionString = "Data Source=etcbtmsql01\\digitmont;Initial Catalog=dbDigitmontDetection;User ID=Metis;Password=Metis123"; ;
            string tableName = "TESTER_APPLICATION_TUBE"; // Ganti dengan nama tabel yang ingin Anda lihat
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    // Memastikan bahwa nama tabel valid dan menggunakan tanda kutip dengan benar
                    string query = "SELECT * FROM TESTER_APPLICATION_TUBE WHERE 1 = 0";
                    SqlDataAdapter adapter = new SqlDataAdapter(query, connection);
                    DataTable schemaTable = new DataTable();
                    adapter.FillSchema(schemaTable, SchemaType.Source);

                    DataTable dt = new DataTable();
                    dt.Columns.Add("ColumnName");
                    dt.Columns.Add("DataType");

                    foreach (DataColumn column in schemaTable.Columns)
                    {
                        DataRow row = dt.NewRow();
                        row["ColumnName"] = column.ColumnName;
                        row["DataType"] = column.DataType.ToString();
                        dt.Rows.Add(row);
                    }

                    dataGridAppTube.DataSource = dt;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message);
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            /*
            string connectionString = "Data Source=etcbtmsql01\\digitmont;Initial Catalog=dbDigitmontDetection;User ID=Metis;Password=Metis123";
            string keyword1 = textBoxSearch.Text;
            string keyword2 = textBoxSearch.Text;
            SqlConnection conn = new SqlConnection(connectionString);
            SqlDataAdapter sda = new SqlDataAdapter("SELECT * FROM TESTER_APPLICATION_TUBE WHERE MACHINE_NUMBER LIKE '%" + keyword1 + "%' AND DEVICE_TYPE LIKE '%" + keyword2 + "%'", conn);
            DataTable dt = new DataTable();
            sda.Fill(dt);

            // Memeriksa jika tidak ada data yang ditemukan
            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("Data not found!!!");
            }
            else
            {
                dataGridAppTube.DataSource = dt;
            }
            
            DialogResult delete = new DialogResult();
            delete = MessageBox.Show("Do You Want to Delete Machine Number: " + CB_Machine.Text + " and Device Type: " + CB_DeviceType.Text + " ? ", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (delete == DialogResult.Yes)
            {
                data.machineNumber = CB_Machine.Text;
                data.deviceType = CB_DeviceType.Text;

                bool success = data.Delete(data);
                if (success == true)
                {
                    MessageBox.Show("Data has been Delete!!!");
                    DataTable dt = data.Select();
                    dataGridAppTube.DataSource = dt;
                }
                else
                {
                    MessageBox.Show("Failed to Delete Data!!!", "Try Again");
                }
            }
            else
            {
                this.Show();
            }
             */
        }

        private void tableLayoutPanel2_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}

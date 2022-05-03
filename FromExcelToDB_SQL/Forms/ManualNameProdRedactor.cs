using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace MetalBase
{
    public partial class ManualNameProdRedactor : Form
    {
        string sqlConnString = @"Server=maks-pc\SQLEXPRESS;Database=MetalBase;User ID=metuser;Password=metuser";
        public ManualNameProdRedactor(string connectionString)
        {
            InitializeComponent();
            if (connectionString != "")
                sqlConnString = connectionString;
        }

        SqlConnection conn;
        bool isChange = false;


        private void ManualNameProdRedactor_Load(object sender, EventArgs e)
        {
            this.Icon = MetalBase.Properties.Resources.prokat_sortovoy_ico;
            update();
        }

        private void update()
        {
            try
            {
                conn = new SqlConnection(sqlConnString);
                if (conn.State == ConnectionState.Closed) conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    DataTable dt = new DataTable();
                    SqlDataAdapter sda = new SqlDataAdapter("select NameOrg as 'Организация', NameProd as 'Продукция' from ManualNameProd", conn);
                    sda.Fill(dt);
                    dataGridView1.DataSource = dt;
                }
            }
            catch (Exception ex) { ex.ToString(); }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (isChange)
            {
                if (MessageBox.Show("Были внесены изменения. Если нажать ОК, изменения будут утеряны!\n\nВыйти без сохранений?", "Внимание", MessageBoxButtons.OKCancel) == System.Windows.Forms.DialogResult.OK)
                {
                    this.Close();
                }
            }
            else this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 0)
            {
                try
                {
                    SqlCommand comm = new SqlCommand("dbo.delManProd", conn);
                    comm.CommandType = CommandType.StoredProcedure;
                    comm.ExecuteNonQuery();
                    comm = new SqlCommand("dbo.insManProd", conn);
                    comm.CommandType = CommandType.StoredProcedure;
                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        comm.Parameters.Clear();
                        comm.Parameters.AddWithValue("@NameOrg", dataGridView1.Rows[i].Cells["Организация"].Value);
                        comm.Parameters.AddWithValue("@NameProd", dataGridView1.Rows[i].Cells["Продукция"].Value);
                        comm.ExecuteNonQuery();
                    }
                    update();
                    isChange = false;
                }
                catch (Exception ex) { ex.ToString(); }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void ManualNameProdRedactor_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (isChange)
            {
                if (MessageBox.Show("Были внесены изменения. Если нажать ОК, изменения будут утеряны!\n\nВыйти без сохранений?", "Внимание", MessageBoxButtons.OKCancel) == System.Windows.Forms.DialogResult.OK)
                {
                    e.Cancel = false;
                }
                else e.Cancel = true;
            }
            else e.Cancel = false;
            
        }
    }
}

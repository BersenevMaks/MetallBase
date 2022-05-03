using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace MetalBase
{
    public partial class RenameProductForm : Form
    {
        public RenameProductForm(string sqlConnectionString)
        {
            InitializeComponent();
            sqlConnString = sqlConnectionString;
        }

        public RenameProductForm(string sqlConnectionString, string NameProd)
        {
            InitializeComponent();
            sqlConnString = sqlConnectionString;
            textBoxBefore.Text = NameProd;
        }

        string sqlConnString;
        
        private void RenameProductForm_Load(object sender, EventArgs e)
        {
            

        }

        private void btnRename_Click(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(sqlConnString);
            try
            {
                if (textBoxBefore.Text.Trim() != "" && textBoxBefore.Text.Trim() != textBoxAfter.Text.Trim())
                {
                    if (conn.State == ConnectionState.Closed) conn.Open();
                    SqlCommand comm = new SqlCommand("update dbo.product set [Name] = '"+textBoxAfter.Text.Trim()+"' where [Name]='"+textBoxBefore.Text.Trim()+"'", conn);
                    comm.ExecuteNonQuery();
                    MessageBox.Show("Название изменено");
                }
            }
            catch (Exception ex) { MessageBox.Show("Что-то пошло не так\n\n" + ex.ToString()); }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}

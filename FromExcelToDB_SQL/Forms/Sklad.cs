using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Text.RegularExpressions;

namespace MetalBase
{
    public partial class Sklad : Form
    {
        public Sklad()
        {
            InitializeComponent();
        }

        private void Sklad_Load(object sender, EventArgs e)
        {
            this.Icon = MetalBase.Properties.Resources.prokat_sortovoy_ico;
            btnAddDel.Visible = false;
            refresh();
        }

        public string sqlConnectionString;

        private void refresh()
        {
            string nameOrg = "";
            DataTable dtOrgname = new DataTable();
            DataTable dtSklad = new DataTable();
            SqlConnection conn = new SqlConnection(sqlConnectionString);
            treeView1.Nodes.Clear();
            comboBoxOrgName.Items.Clear();
            try
            {
                if (conn.State == ConnectionState.Closed) conn.Open();
                string query = @"select [Name] from Organization";
                SqlDataAdapter adapter;
                adapter = new SqlDataAdapter(query, conn);
                adapter.Fill(dtOrgname);
                foreach (DataRow row in dtOrgname.Rows)
                {
                    dtSklad.Clear();
                    nameOrg = row["Name"].ToString();
                    treeView1.Nodes.Add(nameOrg);
                    query = @"select Adress from Sklad where id_organization=(select id_organization from organization where [Name]='" + nameOrg + "')";
                    adapter = new SqlDataAdapter(query, conn);
                    adapter.Fill(dtSklad);
                    foreach (DataRow rowManager in dtSklad.Rows)
                    {
                        treeView1.Nodes[treeView1.Nodes.Count - 1].Nodes.Add(rowManager["Adress"].ToString());
                    }
                    comboBoxOrgName.Items.Add(nameOrg);
                    comboBoxOrgName.Text = "выберите организацию";
                }
                if (conn.State == ConnectionState.Open) conn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void btnAddDel_Click(object sender, EventArgs e)
        {
            if (comboBoxOrgName.Items.Contains(comboBoxOrgName.Text) && textBoxSkladAdress.Text.Trim() != "")
            {
                SqlConnection conn = new SqlConnection(sqlConnectionString);
                try
                {
                    if (conn.State == ConnectionState.Closed) conn.Open();
                    SqlCommand query = new SqlCommand(@"delete from sklad where [Adress] = '" + textBoxSkladAdress.Text + @"' and id_organization=(select id_organization
from organization where [name]='" + comboBoxOrgName.Text + "')", conn);
                    query.ExecuteNonQuery();
                    if (conn.State == ConnectionState.Open) conn.Close();

                    comboBoxOrgName.Text = "выберите организацию";
                    textBoxSkladAdress.Text = "";
                    btnClear.Visible = false;
                    btnAddDel.Visible = false;
                    if (conn.State == ConnectionState.Open) conn.Close();
                    refresh();
                    MessageBox.Show("Успешно удалено");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(sqlConnectionString);

            try
            {
                if (conn.State == ConnectionState.Closed) conn.Open();
                SqlCommand sqlCmd = new SqlCommand("dbo.insSklad", conn);

                sqlCmd.CommandType = CommandType.StoredProcedure;

                sqlCmd.Parameters.AddWithValue("@Name", /* Значение параметра */ comboBoxOrgName.SelectedItem.ToString());

                sqlCmd.Parameters.AddWithValue("@AdressSklad", /* Значение параметра */ textBoxSkladAdress.Text);

                sqlCmd.ExecuteNonQuery();

                comboBoxOrgName.Text = "выберите организацию";
                textBoxSkladAdress.Text = "";
                btnClear.Visible = false;
                btnAddDel.Visible = false;
                if (conn.State == ConnectionState.Open) conn.Close();
                refresh();
                MessageBox.Show("Успешно добавлено");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            comboBoxOrgName.Text = "выберите организацию";
            textBoxSkladAdress.Text = "";
        }

        private void treeView1_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            DataTable dtSklad = new DataTable();
            DataTable dtInfo = new DataTable();
            SqlConnection conn = new SqlConnection(sqlConnectionString);
            try
            {
                if (conn.State == ConnectionState.Closed) conn.Open();

                if (e.Node.Parent != null)
                {
                    string query = @"select o.[Name] as 'org', s.Adress as 'Adress' from organization o join Sklad s
on o.id_organization=s.id_organization where s.[adress]='" + e.Node.Text + "' and o.[name]='" + e.Node.Parent.Text + "'";
                    SqlDataAdapter adapter;
                    adapter = new SqlDataAdapter(query, conn);
                    adapter.Fill(dtSklad);
                    foreach (DataRow row in dtSklad.Rows)
                    {
                        comboBoxOrgName.SelectedItem = row["org"].ToString();
                        textBoxSkladAdress.Text = row["Adress"].ToString();
                    }
                }
                if (conn.State == ConnectionState.Open) conn.Close();
                btnClear.Visible = true;
                btnAddDel.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
    }
}

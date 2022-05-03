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
    public partial class Managers : Form
    {
        public Managers()
        {
            InitializeComponent();
        }

        public string sqlConnectionString;
        string nameManager;
        string telManager;

        private void Managers_Load(object sender, EventArgs e)
        {
            this.Icon = MetalBase.Properties.Resources.prokat_sortovoy_ico;
            btnAddDel.Visible = false;
            Refresh();
        }

        private void Refresh()
        {
            string nameOrg = "";
            DataTable dtOrgname = new DataTable();
            DataTable dtManager = new DataTable();
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
                    dtManager.Clear();
                    nameOrg = row["Name"].ToString();
                    treeView1.Nodes.Add(nameOrg);
                    query = @"select Name from Manager where id_organization=(select id_organization from organization where [Name]='"+nameOrg+"')";
                    adapter = new SqlDataAdapter(query, conn);
                    adapter.Fill(dtManager);
                    foreach (DataRow rowManager in dtManager.Rows)
                    {
                        treeView1.Nodes[treeView1.Nodes.Count - 1].Nodes.Add(rowManager["Name"].ToString());
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

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void treeView1_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            DataTable dtManager = new DataTable();
            DataTable dtInfo = new DataTable();
            SqlConnection conn = new SqlConnection(sqlConnectionString);
            try
            {
                if (conn.State == ConnectionState.Closed) conn.Open();

                if (e.Node.Parent != null)
                {
                    string query = @"select o.[Name] as 'org', m.name as 'manager', m.telefon as 'tel', o.email as 'em' from organization o join Manager m
on o.id_organization=m.id_organization where m.[name]='"+e.Node.Text+"' and o.[name]='"+e.Node.Parent.Text+"'";
                    SqlDataAdapter adapter;
                    adapter = new SqlDataAdapter(query, conn);
                    adapter.Fill(dtManager);
                    foreach (DataRow row in dtManager.Rows)
                    {
                        comboBoxOrgName.SelectedItem = row["org"].ToString();
                        textBoxManagerName.Text = row["manager"].ToString();
                        nameManager = row["manager"].ToString();
                        textBoxManagerTelefon.Text = row["tel"].ToString();
                        telManager = row["tel"].ToString();
                        textBoxManagerEmail.Text = row["em"].ToString();
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

        private void btnClear_Click(object sender, EventArgs e)
        {
            comboBoxOrgName.Text = "выберите организацию";
            textBoxManagerName.Text = "";
            textBoxManagerTelefon.Text = "";
            textBoxManagerEmail.Text = "";
            btnClear.Visible = false;
            btnAddDel.Visible = false;
        }

        private void btnAddDel_Click(object sender, EventArgs e)
        {
            if (comboBoxOrgName.Items.Contains(comboBoxOrgName.Text) && textBoxManagerName.Text.Trim() != "")
            {
                SqlConnection conn = new SqlConnection(sqlConnectionString);
                try
                {
                    if (conn.State == ConnectionState.Closed) conn.Open();
                    SqlCommand query = new SqlCommand(@"delete from manager where [Name] = '" + textBoxManagerName.Text + @"' and id_organization=(select id_organization
from organization where [name]='" + comboBoxOrgName.Text + "')", conn);
                    query.ExecuteNonQuery();
                    if (conn.State == ConnectionState.Open) conn.Close();

                    comboBoxOrgName.Text = "выберите организацию";
                    textBoxManagerName.Text = "";
                    textBoxManagerTelefon.Text = "";
                    textBoxManagerEmail.Text = "";
                    btnClear.Visible = false;
                    btnAddDel.Visible = false;
                    if (conn.State == ConnectionState.Open) conn.Close();
                    Refresh();
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
            if (comboBoxOrgName.Items.Contains(comboBoxOrgName.Text) && textBoxManagerName.Text.Trim() != "")
            {
                SqlConnection conn = new SqlConnection(sqlConnectionString);
                
                try
                {
                    if (conn.State == ConnectionState.Closed) conn.Open();
                    SqlCommand sqlCmd = new SqlCommand("dbo.insManager", conn);
                    if (treeView1.SelectedNode.Parent != null)
                    {
                        if(comboBoxOrgName.SelectedItem.ToString()==treeView1.SelectedNode.Parent.Text &&
                            (nameManager==textBoxManagerName.Text || telManager==textBoxManagerTelefon.Text))
                            sqlCmd = new SqlCommand("dbo.updateManager", conn);
                    }
                    else sqlCmd = new SqlCommand("dbo.insManager", conn);

                    sqlCmd.CommandType = CommandType.StoredProcedure;
                    
                    sqlCmd.Parameters.AddWithValue("@Name", /* Значение параметра */ comboBoxOrgName.SelectedItem.ToString());
                    
                    sqlCmd.Parameters.AddWithValue("@NameManager", /* Значение параметра */ textBoxManagerName.Text);
                    
                    if (textBoxManagerTelefon.Text.Trim() != "")
                        sqlCmd.Parameters.AddWithValue("@TelefonManager", /* Значение параметра */ textBoxManagerTelefon.Text);
                    else sqlCmd.Parameters.AddWithValue("@TelefonManager", /* Значение параметра */ " ");
                    
                    if (textBoxManagerEmail.Text.Trim() != "")
                        sqlCmd.Parameters.AddWithValue("@Email", /* Значение параметра */ textBoxManagerEmail.Text);
                    else sqlCmd.Parameters.AddWithValue("@Email", " ");
                    sqlCmd.ExecuteNonQuery();

                    comboBoxOrgName.Text = "выберите организацию";
                    textBoxManagerName.Text = "";
                    textBoxManagerTelefon.Text = "";
                    textBoxManagerEmail.Text = "";
                    btnClear.Visible = false;
                    btnAddDel.Visible = false;
                    if (conn.State == ConnectionState.Open) conn.Close();
                    Refresh();
                    MessageBox.Show("Успешно добавлено");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }
    }
}

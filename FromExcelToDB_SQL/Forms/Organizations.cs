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
    public partial class Organizations : Form
    {
        public Organizations()
        {
            InitializeComponent();
        }

        public string sqlConnectionString;
        string orgId;
        string priceDate;
        bool isAdd = false;

        private void Organizations_Load(object sender, EventArgs e)
        {
            this.Icon = MetalBase.Properties.Resources.prokat_sortovoy_ico;
            btnSave.Text = "Добавить и сохранить";
            isAdd = true;
            btnAddDel.Visible = false;
            refresh();
            
        }

        private void refresh()
        {
            string nameOrg = "";
            DataTable dtOrgname = new DataTable();
            SqlConnection conn = new SqlConnection(sqlConnectionString);
            treeView1.Nodes.Clear();
            
            try
            {
                if (conn.State == ConnectionState.Closed) conn.Open();
                string query = @"select [Name] from Organization";
                SqlDataAdapter adapter;
                adapter = new SqlDataAdapter(query, conn);
                adapter.Fill(dtOrgname);
                foreach (DataRow row in dtOrgname.Rows)
                {
                    nameOrg = row["Name"].ToString();
                    treeView1.Nodes.Add(nameOrg);
                }
                label10.Text = "Всего: " + treeView1.Nodes.Count;
                if (conn.State == ConnectionState.Open) conn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void btnAddDel_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("При удалении организии произойдет удаление всей связанной\nс данной организацией продукции\nВы действительно хотите удалить \""+textBoxOrgName.Text+"\"?", "Внимание", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.OK)
            {
                SqlConnection conn = new SqlConnection(sqlConnectionString);
                try
                {
                    if (conn.State == ConnectionState.Closed) conn.Open();
                    var sqlCmd = new SqlCommand("dbo.delOrg", conn);
                    sqlCmd.CommandType = CommandType.StoredProcedure;
                    
                    sqlCmd.Parameters.AddWithValue("@Name", /* Значение параметра */ textBoxOrgName.Text);

                    sqlCmd.ExecuteNonQuery();
                    refresh();
                    btnAddDel.Visible = false;
                    btnSave.Text = "Добавить и сохранить";
                    isAdd = true;
                    btnAddDel.Visible = false;
                    textBoxOrgName.Text = "";
                    textBoxOrgAdress.Text = "";
                    textBoxOrgEmail.Text = "";
                    textBoxOrgINN.Text = "";
                    textBoxOrgKS.Text = "";
                    textBoxOrgRS.Text = "";
                    textBoxOrgSite.Text = "";
                    textBoxOrgTelefon.Text = "";
                    textBoxBIK.Text = "";
                }
                catch (Exception ex) { MessageBox.Show(ex.ToString()); }
            }
        }

        private void treeView1_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            btnSave.Text = "Применить изменения";
            isAdd = false;
            SqlConnection conn = new SqlConnection(sqlConnectionString);
            var dtOrg = new DataTable();
            try
            {
                if (conn.State == ConnectionState.Closed) conn.Open();

                string query = @"select ID_Organization as 'id', [name] as 'org', adress as 'adr', telefon as 'tel', 
Email as 'mail', [site] as 'site', INNKPP as 'inn', RasSchet as 'rs', KorSchet as 'ks', BIK as 'bik', [datePriceList] as 'priceDate'
from organization where [name]='" + e.Node.Text + "'";
                SqlDataAdapter adapter;
                adapter = new SqlDataAdapter(query, conn);
                adapter.Fill(dtOrg);
                foreach (DataRow row in dtOrg.Rows)
                {
                    textBoxOrgName.Text = row["org"].ToString();
                    textBoxOrgAdress.Text = row["adr"].ToString();
                    textBoxOrgTelefon.Text = row["tel"].ToString();
                    textBoxOrgEmail.Text = row["mail"].ToString();
                    textBoxOrgSite.Text = row["site"].ToString();
                    textBoxOrgINN.Text = row["inn"].ToString();
                    textBoxOrgRS.Text = row["rs"].ToString();
                    textBoxOrgKS.Text = row["ks"].ToString();
                    textBoxBIK.Text = row["bik"].ToString();
                    orgId = row["id"].ToString();
                    priceDate = row["priceDate"].ToString();
                    txtBoxPriceDate.Text = priceDate;
                }
                btnAddDel.Visible = true;
                if (conn.State == ConnectionState.Open) conn.Close();
            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            btnSave.Text = "Добавить и сохранить";
            isAdd = true;
            btnAddDel.Visible = false;
            textBoxOrgName.Text = "";
            textBoxOrgAdress.Text = "";
            textBoxOrgEmail.Text = "";
            textBoxOrgINN.Text = "";
            textBoxOrgKS.Text = "";
            textBoxOrgRS.Text = "";
            textBoxOrgSite.Text = "";
            textBoxOrgTelefon.Text = "";
            textBoxBIK.Text = "";
            txtBoxPriceDate.Text = "";
            refresh();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(sqlConnectionString);
            try
            {
                if (conn.State == ConnectionState.Closed) conn.Open();
                SqlCommand sqlCmd;
                if (isAdd)
                {
                    sqlCmd = new SqlCommand("dbo.insOrg", conn);
                    sqlCmd.CommandType = CommandType.StoredProcedure;
                }
                else 
                {
                    sqlCmd = new SqlCommand("dbo.updateOrganization", conn);
                    sqlCmd.CommandType = CommandType.StoredProcedure;
                    sqlCmd.Parameters.AddWithValue("@ID_Organization", orgId);
                }

                sqlCmd.Parameters.AddWithValue("@Name", /* Значение параметра */ textBoxOrgName.Text);

                if (textBoxOrgAdress.Text == "") textBoxOrgAdress.Text = " ";
                sqlCmd.Parameters.AddWithValue("@Adress", /* Значение параметра */ textBoxOrgAdress.Text);

                if (textBoxOrgTelefon.Text == "") textBoxOrgTelefon.Text = " ";
                sqlCmd.Parameters.AddWithValue("@Telefon", /* Значение параметра */ textBoxOrgTelefon.Text);

                if (textBoxOrgEmail.Text == "") textBoxOrgEmail.Text = " ";
                sqlCmd.Parameters.AddWithValue("@Email", /* Значение параметра */ textBoxOrgEmail.Text);

                if (textBoxOrgSite.Text == "") textBoxOrgSite.Text = " ";
                sqlCmd.Parameters.AddWithValue("@Site", /* Значение параметра */ textBoxOrgSite.Text);

                if (textBoxOrgINN.Text == "") textBoxOrgINN.Text = " ";
                sqlCmd.Parameters.AddWithValue("@INNKPP", /* Значение параметра */ textBoxOrgINN.Text);

                if (textBoxOrgRS.Text == "") textBoxOrgRS.Text = " ";
                sqlCmd.Parameters.AddWithValue("@RasSchet", /* Значение параметра */ textBoxOrgRS.Text);

                if (textBoxOrgKS.Text == "") textBoxOrgKS.Text = " ";
                sqlCmd.Parameters.AddWithValue("@KorSchet", /* Значение параметра */ textBoxOrgKS.Text);

                if (textBoxBIK.Text == "") textBoxBIK.Text = " ";
                sqlCmd.Parameters.AddWithValue("@BIK", /* Значение параметра */ textBoxBIK.Text);

                if (isAdd) sqlCmd.Parameters.AddWithValue("@datePrice", DateTime.Now.Year.ToString() + "." + DateTime.Now.Month.ToString() + "." + DateTime.Now.Day.ToString());
                else sqlCmd.Parameters.AddWithValue("@datePrice", priceDate);

                sqlCmd.ExecuteNonQuery();

                btnAddDel.Visible = false;
                btnSave.Text = "Добавить и сохранить";
                refresh();
                isAdd = true;
                btnAddDel.Visible = false;
                textBoxOrgName.Text = "";
                textBoxOrgAdress.Text = "";
                textBoxOrgEmail.Text = "";
                textBoxOrgINN.Text = "";
                textBoxOrgKS.Text = "";
                textBoxOrgRS.Text = "";
                textBoxOrgSite.Text = "";
                textBoxOrgTelefon.Text = "";
                textBoxBIK.Text = "";
                if (conn.State == ConnectionState.Open) conn.Close();
            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
        }

    }
}

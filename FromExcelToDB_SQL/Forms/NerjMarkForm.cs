using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Text.RegularExpressions;
using System.IO;

namespace MetalBase.Forms
{
    public partial class NerjMarkForm : Form
    {
        public NerjMarkForm(string sqlConnString)
        {
            InitializeComponent();
            sqlConnectionString = sqlConnString;
        }
        private string sqlConnectionString;

        private void NerjMarkForm_Load(object sender, EventArgs e)
        {
            this.Icon = MetalBase.Properties.Resources.prokat_sortovoy_ico;
            //dataGridView1.Columns.Add("Марка");
            UpdateForm();
        }

        private void UpdateForm()
        {
            DataTable dtCheckTable = new DataTable();
            DataTable dtMarks = new DataTable();
            SqlConnection conn = new SqlConnection(sqlConnectionString);
            
            try
            {
                if (conn.State == ConnectionState.Closed) conn.Open();
                string query = @"IF EXISTS (SELECT * FROM SYSOBJECTS WHERE NAME='MarksNerj' AND xtype='U') select name from sysobjects where name = 'MarksNerj'";
                SqlDataAdapter adapter;
                adapter = new SqlDataAdapter(query, conn);
                adapter.Fill(dtCheckTable);
                if (dtCheckTable.Rows.Count > 0)
                {
                    query = @"SELECT Mark FROM dbo.MarksNerj";
                    adapter = new SqlDataAdapter(query, conn);
                    adapter.Fill(dtMarks);
                    if (dtMarks.Rows.Count > 0)
                    {
                        dataGridView1.DataSource = dtMarks;
                        if (dataGridView1.Columns.Count > 0)
                            dataGridView1.Columns[0].Width = 200;
                    }
                }
                else {
                    if (MessageBox.Show("Таблицы нет\nСоздать?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        query = @"use MetalBase create table dbo.MarksNerj (MarkID int PRIMARY KEY NOT NULL identity(1,1), Mark varchar(25) NOT NULL)";
                        SqlCommand sqlCommand = new SqlCommand(query, conn);
                        int sqlCount = sqlCommand.ExecuteNonQuery();
                        MessageBox.Show("Создана таблица марок нержавеющей стали в базе");
                    };
                    
                }
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
        string MarkName = "";
        private void btnAdd_Click(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(sqlConnectionString);
            AddMarkName amn = new AddMarkName(new MyDelegate(setMarkName));
            amn.ShowDialog();
            if (MarkName != "")
            {
                if (conn.State == ConnectionState.Closed) conn.Open();
                string query = @"INSERT INTO dbo.marksNerj (mark) VALUES ('" + MarkName + "');";
                SqlCommand comm = new SqlCommand(query, conn);
                int k = comm.ExecuteNonQuery();
                if (k > -1)
                {
                    MessageBox.Show("Марка " + MarkName + " успешно добавлена");
                    MarkName = "";
                }
                else
                {
                    MessageBox.Show("Ошибка, марка не добавлена");
                    MarkName = "";
                }
                UpdateForm();
            }
        }
        private void setMarkName (string markname)
        { MarkName = markname; }

        private void btnDel_Click(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(sqlConnectionString);
            if (conn.State == ConnectionState.Closed) conn.Open();
            string deletingMark = dataGridView1.SelectedCells[0].Value.ToString();
            if (MessageBox.Show("Вы действительно хотите удалить марку: "+deletingMark+"?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                SqlCommand comm = new SqlCommand("use MetalBase delete from dbo.MarksNerj where mark='"+deletingMark+"'", conn);
                int k = comm.ExecuteNonQuery();
                if (k > -1)
                {
                    MessageBox.Show("Успешно удалено");
                }
                else MessageBox.Show("Ошибка при удалении");
                UpdateForm();
            }
        }

        private void btn_Click(object sender, EventArgs e)
        {
            List<string> result = new List<string>();
            SqlConnection conn = new SqlConnection(sqlConnectionString);
            SqlCommand comm;
            string[] str;
            string query;
            int count = 0;
            string message = "Будут добавлены следующие марки:\n";
            string PathFile = "";
            if (conn.State == ConnectionState.Closed) conn.Open();
            try
            {
                OpenFileDialog ofd = new OpenFileDialog();
                if (ofd.ShowDialog() == DialogResult.OK)
                    PathFile = ofd.FileName;
                if (!string.IsNullOrEmpty(PathFile))
                {
                    using (StreamReader sr = new StreamReader(PathFile, System.Text.Encoding.Default))
                    {
                        str = sr.ReadToEnd().Split(';');
                        foreach (string s in str)
                        {
                            result.Add(s.Replace("\r", "").Replace("\n", ""));
                            message += result[result.Count-1] + "\n";
                        }
                    }
                }
                message += "\nДобавить их в базу?";
                if (MessageBox.Show(message, "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    foreach (string AddingMark in result)
                    {
                        query = @"if not exists (select mark from dbo.marksnerj where mark = '" + AddingMark + "') INSERT INTO dbo.marksNerj (mark) VALUES ('" + AddingMark + "');";
                        comm = new SqlCommand(query, conn);
                        int k = comm.ExecuteNonQuery();
                        if (k > -1)
                            count++;
                    }
                    UpdateForm();
                    if (count == result.Count)
                    {
                        MessageBox.Show("Успешно добавлено");
                    }
                    else MessageBox.Show("Не все марки были добавлены, проверьте список.");
                }
            }
            catch (Exception ex) { MessageBox.Show("Ошибка при чтении файла\n\n" + ex.ToString()); }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Эта операция необратима! Любые типы продукции имеющие перечисленные марки будут изменены на нержавеющий тип!\nВы уверены, что хотите продолжить???", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                SqlConnection conn = new SqlConnection(sqlConnectionString);
                if (conn.State == ConnectionState.Closed) conn.Open();
                SqlCommand comm;
                string query = @"update dbo.Product set type = 'нержавеющий' where [name]='Лист' and Marka in (select mark from dbo.marksnerj);";
                comm = new SqlCommand(query, conn);
                int k = comm.ExecuteNonQuery();
                query = @"update dbo.Product set type = 'нержавеющая' where [name] <>'Лист' and Marka in (select mark from dbo.marksnerj);";
                comm = new SqlCommand(query, conn);
                k += comm.ExecuteNonQuery();
                if (k > -1)
                {
                    MessageBox.Show("Обновлено " + k + " типов");
                }
            }
        }
    }
}

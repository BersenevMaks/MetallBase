using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace MetalBase
{
    public partial class settingForm : Form
    {
        settingDelegate sd;
        public settingForm(settingDelegate sender)
        {
            InitializeComponent();
            sd = sender;
        }
        string fileName; 
        private void settingForm_Load(object sender, EventArgs e)
        {
            panel1.Visible = false;
            fileName = Path.GetDirectoryName(Application.ExecutablePath) + "\\settings.set";
            this.Icon = MetalBase.Properties.Resources.prokat_sortovoy_ico;
            dataGridView1.Rows.Add("Имя сервера", "");
            dataGridView1.Rows.Add("Имя экземпляра", "");
            dataGridView1.Rows.Add("Номер порта", "");
            dataGridView1.Rows.Add("Database", "");
            dataGridView1.Rows.Add("User ID", "");
            dataGridView1.Rows.Add("Password", "");
            dataGridView1.Rows.Add("Количество потоков чтения", "");

            if (File.Exists(fileName) == true)
            {
                try
                {                                  //чтение файла
                    string[] allText = File.ReadAllLines(fileName);         //чтение всех строк файла в массив строк

                    for (int i = 0; i < allText.Length;i++ )
                    {
                        dataGridView1.Rows[i].Cells[1].Value=allText[i];
                    }
                }
                catch (FileNotFoundException ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
        }
        List<string> set = new List<string>();
        private void button1_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if(row.Cells["Values"].Value!=null)
                set.Add(row.Cells["Values"].Value.ToString());
                else set.Add("");
            }

            using (StreamWriter sw = new StreamWriter(new FileStream(fileName, FileMode.Create, FileAccess.Write)))
            {
                foreach (string s in set)
                {
                    sw.WriteLine(s);
                }
            }
            sd(set);
            this.Close();
        }
        
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                DataTable dt = System.Data.Sql.SqlDataSourceEnumerator.Instance.GetDataSources();
                dataGridView2.DataSource = dt;
                panel1.Visible = true;
            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
        }

        private void dataGridView2_MouseClick(object sender, MouseEventArgs e)
        {
            //panel1.Visible = false;
        }

        private void dataGridView2_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            dataGridView1.Rows[0].Cells[1].Value = dataGridView2.SelectedRows[0].Cells[0].Value;
            panel1.Visible = false;
        }

        private void settingForm_MouseClick(object sender, MouseEventArgs e)
        {
            panel1.Visible = false;
        }

        private void dataGridView1_MouseClick(object sender, MouseEventArgs e)
        {
            panel1.Visible = false;
        }

    }
}

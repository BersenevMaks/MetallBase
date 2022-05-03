using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace MetalBase
{
    public partial class AddNameProduct : Form
    {
        MyDelegate d;
        public AddNameProduct(MyDelegate sender)
        {
            InitializeComponent();
            d = sender;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.Trim() != "")
            {
                okFunc();
            }
        }

        private void okFunc()
        {
            d("");
            if (textBox1.Text != "") d(textBox1.Text);
            this.Close();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            d("");
            this.Close();
        }

        private void AddNameProduct_Load(object sender, EventArgs e)
        {
            label1.Text = "В файле " + this.Tag + " \nне найдено ни одного наименования продукции";
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (textBox1.Text.Trim() != "")
            {
                if (e.KeyCode == Keys.Enter)
                {
                    okFunc();
                }
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MetalBase.Forms
{
    public partial class AddMarkName : Form
    {
        public AddMarkName(MyDelegate sender)
        {
            InitializeComponent();
            d = sender;
        }
        MyDelegate d;

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
            //label1.Text = "В файле " + this.Tag + " \nне найдено ни одного наименования продукции";
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

        private void AddMarkName_Load(object sender, EventArgs e)
        {

        }
    }
}

namespace MetalBase
{
    partial class Organizations
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label10 = new System.Windows.Forms.Label();
            this.treeView1 = new System.Windows.Forms.TreeView();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.btnClear = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnAddDel = new System.Windows.Forms.Button();
            this.textBoxBIK = new System.Windows.Forms.TextBox();
            this.textBoxOrgKS = new System.Windows.Forms.TextBox();
            this.textBoxOrgRS = new System.Windows.Forms.TextBox();
            this.textBoxOrgINN = new System.Windows.Forms.TextBox();
            this.textBoxOrgSite = new System.Windows.Forms.TextBox();
            this.textBoxOrgEmail = new System.Windows.Forms.TextBox();
            this.textBoxOrgTelefon = new System.Windows.Forms.TextBox();
            this.textBoxOrgAdress = new System.Windows.Forms.TextBox();
            this.textBoxOrgName = new System.Windows.Forms.TextBox();
            this.txtBoxPriceDate = new System.Windows.Forms.TextBox();
            this.label11 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label10);
            this.groupBox1.Controls.Add(this.treeView1);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(199, 420);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Список организаций";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(6, 397);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(43, 13);
            this.label10.TabIndex = 1;
            this.label10.Text = "Всего: ";
            // 
            // treeView1
            // 
            this.treeView1.HideSelection = false;
            this.treeView1.Location = new System.Drawing.Point(7, 20);
            this.treeView1.Name = "treeView1";
            this.treeView1.ShowRootLines = false;
            this.treeView1.Size = new System.Drawing.Size(186, 371);
            this.treeView1.TabIndex = 0;
            this.treeView1.NodeMouseDoubleClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.treeView1_NodeMouseDoubleClick);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(228, 28);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(57, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Название";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(228, 61);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(38, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Адрес";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(228, 94);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(52, 13);
            this.label3.TabIndex = 1;
            this.label3.Text = "Телефон";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(228, 127);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(35, 13);
            this.label4.TabIndex = 1;
            this.label4.Text = "E-mail";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(228, 160);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(31, 13);
            this.label5.TabIndex = 1;
            this.label5.Text = "Сайт";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(228, 193);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(59, 13);
            this.label6.TabIndex = 1;
            this.label6.Text = "ИНН/КПП";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(228, 226);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(41, 13);
            this.label7.TabIndex = 1;
            this.label7.Text = "Р/счет";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(228, 259);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(53, 13);
            this.label8.TabIndex = 1;
            this.label8.Text = "Кор/счет";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(228, 292);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(29, 13);
            this.label9.TabIndex = 1;
            this.label9.Text = "БИК";
            // 
            // btnClear
            // 
            this.btnClear.Location = new System.Drawing.Point(484, 365);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(75, 23);
            this.btnClear.TabIndex = 16;
            this.btnClear.Text = "Очистить";
            this.btnClear.UseVisualStyleBackColor = true;
            this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(484, 409);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 15;
            this.btnClose.Text = "Закрыть";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(312, 409);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(166, 23);
            this.btnSave.TabIndex = 14;
            this.btnSave.Text = "Применить изменения";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnAddDel
            // 
            this.btnAddDel.Location = new System.Drawing.Point(231, 409);
            this.btnAddDel.Name = "btnAddDel";
            this.btnAddDel.Size = new System.Drawing.Size(75, 23);
            this.btnAddDel.TabIndex = 13;
            this.btnAddDel.Text = "Удалить";
            this.btnAddDel.UseVisualStyleBackColor = true;
            this.btnAddDel.Click += new System.EventHandler(this.btnAddDel_Click);
            // 
            // textBoxBIK
            // 
            this.textBoxBIK.Location = new System.Drawing.Point(318, 289);
            this.textBoxBIK.Name = "textBoxBIK";
            this.textBoxBIK.Size = new System.Drawing.Size(241, 20);
            this.textBoxBIK.TabIndex = 30;
            // 
            // textBoxOrgKS
            // 
            this.textBoxOrgKS.Location = new System.Drawing.Point(318, 256);
            this.textBoxOrgKS.Name = "textBoxOrgKS";
            this.textBoxOrgKS.Size = new System.Drawing.Size(241, 20);
            this.textBoxOrgKS.TabIndex = 29;
            // 
            // textBoxOrgRS
            // 
            this.textBoxOrgRS.Location = new System.Drawing.Point(318, 223);
            this.textBoxOrgRS.Name = "textBoxOrgRS";
            this.textBoxOrgRS.Size = new System.Drawing.Size(241, 20);
            this.textBoxOrgRS.TabIndex = 32;
            // 
            // textBoxOrgINN
            // 
            this.textBoxOrgINN.Location = new System.Drawing.Point(318, 190);
            this.textBoxOrgINN.Name = "textBoxOrgINN";
            this.textBoxOrgINN.Size = new System.Drawing.Size(241, 20);
            this.textBoxOrgINN.TabIndex = 31;
            // 
            // textBoxOrgSite
            // 
            this.textBoxOrgSite.Location = new System.Drawing.Point(318, 157);
            this.textBoxOrgSite.Name = "textBoxOrgSite";
            this.textBoxOrgSite.Size = new System.Drawing.Size(241, 20);
            this.textBoxOrgSite.TabIndex = 28;
            // 
            // textBoxOrgEmail
            // 
            this.textBoxOrgEmail.Location = new System.Drawing.Point(318, 124);
            this.textBoxOrgEmail.Name = "textBoxOrgEmail";
            this.textBoxOrgEmail.Size = new System.Drawing.Size(241, 20);
            this.textBoxOrgEmail.TabIndex = 27;
            // 
            // textBoxOrgTelefon
            // 
            this.textBoxOrgTelefon.Location = new System.Drawing.Point(318, 91);
            this.textBoxOrgTelefon.Name = "textBoxOrgTelefon";
            this.textBoxOrgTelefon.Size = new System.Drawing.Size(241, 20);
            this.textBoxOrgTelefon.TabIndex = 26;
            // 
            // textBoxOrgAdress
            // 
            this.textBoxOrgAdress.Location = new System.Drawing.Point(318, 58);
            this.textBoxOrgAdress.Name = "textBoxOrgAdress";
            this.textBoxOrgAdress.Size = new System.Drawing.Size(241, 20);
            this.textBoxOrgAdress.TabIndex = 25;
            // 
            // textBoxOrgName
            // 
            this.textBoxOrgName.Location = new System.Drawing.Point(318, 25);
            this.textBoxOrgName.Name = "textBoxOrgName";
            this.textBoxOrgName.Size = new System.Drawing.Size(241, 20);
            this.textBoxOrgName.TabIndex = 24;
            // 
            // txtBoxPriceDate
            // 
            this.txtBoxPriceDate.Location = new System.Drawing.Point(318, 322);
            this.txtBoxPriceDate.Name = "txtBoxPriceDate";
            this.txtBoxPriceDate.Size = new System.Drawing.Size(241, 20);
            this.txtBoxPriceDate.TabIndex = 34;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(228, 325);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(72, 13);
            this.label11.TabIndex = 33;
            this.label11.Text = "Дата прайса";
            // 
            // Organizations
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(576, 448);
            this.Controls.Add(this.txtBoxPriceDate);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.textBoxBIK);
            this.Controls.Add(this.textBoxOrgKS);
            this.Controls.Add(this.textBoxOrgRS);
            this.Controls.Add(this.textBoxOrgINN);
            this.Controls.Add(this.textBoxOrgSite);
            this.Controls.Add(this.textBoxOrgEmail);
            this.Controls.Add(this.textBoxOrgTelefon);
            this.Controls.Add(this.textBoxOrgAdress);
            this.Controls.Add(this.textBoxOrgName);
            this.Controls.Add(this.btnClear);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnAddDel);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.groupBox1);
            this.Name = "Organizations";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Организации";
            this.Load += new System.EventHandler(this.Organizations_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Button btnClear;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnAddDel;
        private System.Windows.Forms.TextBox textBoxBIK;
        private System.Windows.Forms.TextBox textBoxOrgKS;
        private System.Windows.Forms.TextBox textBoxOrgRS;
        private System.Windows.Forms.TextBox textBoxOrgINN;
        private System.Windows.Forms.TextBox textBoxOrgSite;
        private System.Windows.Forms.TextBox textBoxOrgEmail;
        private System.Windows.Forms.TextBox textBoxOrgTelefon;
        private System.Windows.Forms.TextBox textBoxOrgAdress;
        private System.Windows.Forms.TextBox textBoxOrgName;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.TextBox txtBoxPriceDate;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TreeView treeView1;
    }
}
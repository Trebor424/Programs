namespace AutoBOMmaker
{
    partial class Form1
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
            this.components = new System.ComponentModel.Container();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.txtFilename = new System.Windows.Forms.TextBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.label2 = new System.Windows.Forms.Label();
            this.cboSheet = new System.Windows.Forms.ComboBox();
            this.BomMake = new System.Windows.Forms.Button();
            this.dataGridView2 = new System.Windows.Forms.DataGridView();
            this.nullLogFactoryBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.Modify = new System.Windows.Forms.Button();
            this.Exit_btn = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.textBox5 = new System.Windows.Forms.TextBox();
            this.textBox6 = new System.Windows.Forms.TextBox();
            this.textBox7 = new System.Windows.Forms.TextBox();
            this.textBox8 = new System.Windows.Forms.TextBox();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.colors = new System.Windows.Forms.Button();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.Diodeinsbtn = new System.Windows.Forms.Button();
            this.Tolerancebtn = new System.Windows.Forms.Button();
            this.comboBox3 = new System.Windows.Forms.ComboBox();
            this.Markbtn = new System.Windows.Forms.Button();
            this.headerRowDeletebtn = new System.Windows.Forms.Button();
            this.Insert_ManufacturerNumberbtn = new System.Windows.Forms.Button();
            this.Insert_Descriptionbtn = new System.Windows.Forms.Button();
            this.Insert_CircuitReferencebtn = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.NecesseryColumnbtn = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nullLogFactoryBindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // btnBrowse
            // 
            this.btnBrowse.Cursor = System.Windows.Forms.Cursors.SizeAll;
            this.btnBrowse.Location = new System.Drawing.Point(812, 4);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(152, 24);
            this.btnBrowse.TabIndex = 1;
            this.btnBrowse.Text = "Load Excel Files";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(99, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "File Name Directory";
            // 
            // txtFilename
            // 
            this.txtFilename.Location = new System.Drawing.Point(117, 6);
            this.txtFilename.Name = "txtFilename";
            this.txtFilename.Size = new System.Drawing.Size(685, 20);
            this.txtFilename.TabIndex = 3;
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowDrop = true;
            this.dataGridView1.AllowUserToOrderColumns = true;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnKeystroke;
            this.dataGridView1.GridColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.dataGridView1.Location = new System.Drawing.Point(7, 63);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(795, 902);
            this.dataGridView1.StandardTab = true;
            this.dataGridView1.TabIndex = 4;
            this.dataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(970, 9);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(104, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "Insert Excel Script ->";
            // 
            // cboSheet
            // 
            this.cboSheet.FormattingEnabled = true;
            this.cboSheet.Location = new System.Drawing.Point(1080, 6);
            this.cboSheet.Name = "cboSheet";
            this.cboSheet.Size = new System.Drawing.Size(151, 21);
            this.cboSheet.TabIndex = 6;
            this.cboSheet.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // BomMake
            // 
            this.BomMake.Cursor = System.Windows.Forms.Cursors.SizeAll;
            this.BomMake.Location = new System.Drawing.Point(812, 34);
            this.BomMake.Name = "BomMake";
            this.BomMake.Size = new System.Drawing.Size(151, 23);
            this.BomMake.TabIndex = 7;
            this.BomMake.Text = "Create Bom";
            this.BomMake.UseVisualStyleBackColor = true;
            this.BomMake.Click += new System.EventHandler(this.BomMake_Click);
            // 
            // dataGridView2
            // 
            this.dataGridView2.AllowDrop = true;
            this.dataGridView2.AllowUserToOrderColumns = true;
            this.dataGridView2.BackgroundColor = System.Drawing.SystemColors.ActiveBorder;
            this.dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView2.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnKeystroke;
            this.dataGridView2.GridColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.dataGridView2.Location = new System.Drawing.Point(813, 63);
            this.dataGridView2.Name = "dataGridView2";
            this.dataGridView2.Size = new System.Drawing.Size(893, 907);
            this.dataGridView2.StandardTab = true;
            this.dataGridView2.TabIndex = 8;
            this.dataGridView2.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView2_CellContentClick);
            // 
            // nullLogFactoryBindingSource
            // 
            this.nullLogFactoryBindingSource.DataSource = typeof(ExcelDataReader.Log.Logger.NullLogFactory);
            // 
            // Modify
            // 
            this.Modify.Location = new System.Drawing.Point(969, 34);
            this.Modify.Name = "Modify";
            this.Modify.Size = new System.Drawing.Size(258, 23);
            this.Modify.TabIndex = 9;
            this.Modify.Text = "Make Excel From Data2";
            this.Modify.UseVisualStyleBackColor = true;
            this.Modify.Click += new System.EventHandler(this.Modify_Click);
            // 
            // Exit_btn
            // 
            this.Exit_btn.Location = new System.Drawing.Point(1734, 12);
            this.Exit_btn.Name = "Exit_btn";
            this.Exit_btn.Size = new System.Drawing.Size(86, 22);
            this.Exit_btn.TabIndex = 10;
            this.Exit_btn.Text = "Exit";
            this.Exit_btn.UseVisualStyleBackColor = true;
            this.Exit_btn.Click += new System.EventHandler(this.Exit_btn_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(1712, 98);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(106, 260);
            this.textBox1.TabIndex = 11;
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(1713, 393);
            this.textBox2.Multiline = true;
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(106, 121);
            this.textBox2.TabIndex = 12;
            // 
            // textBox3
            // 
            this.textBox3.Location = new System.Drawing.Point(1714, 520);
            this.textBox3.Multiline = true;
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(106, 108);
            this.textBox3.TabIndex = 13;
            // 
            // textBox4
            // 
            this.textBox4.Location = new System.Drawing.Point(1714, 634);
            this.textBox4.Multiline = true;
            this.textBox4.Name = "textBox4";
            this.textBox4.Size = new System.Drawing.Size(106, 108);
            this.textBox4.TabIndex = 14;
            // 
            // textBox5
            // 
            this.textBox5.Location = new System.Drawing.Point(1712, 72);
            this.textBox5.Name = "textBox5";
            this.textBox5.Size = new System.Drawing.Size(106, 20);
            this.textBox5.TabIndex = 15;
            // 
            // textBox6
            // 
            this.textBox6.Location = new System.Drawing.Point(1714, 367);
            this.textBox6.Name = "textBox6";
            this.textBox6.Size = new System.Drawing.Size(105, 20);
            this.textBox6.TabIndex = 16;
            // 
            // textBox7
            // 
            this.textBox7.Location = new System.Drawing.Point(1714, 748);
            this.textBox7.Multiline = true;
            this.textBox7.Name = "textBox7";
            this.textBox7.Size = new System.Drawing.Size(106, 108);
            this.textBox7.TabIndex = 17;
            // 
            // textBox8
            // 
            this.textBox8.Location = new System.Drawing.Point(1714, 862);
            this.textBox8.Multiline = true;
            this.textBox8.Name = "textBox8";
            this.textBox8.Size = new System.Drawing.Size(106, 108);
            this.textBox8.TabIndex = 18;
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(1624, 38);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(104, 21);
            this.comboBox1.TabIndex = 19;
            this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged_1);
            // 
            // colors
            // 
            this.colors.Location = new System.Drawing.Point(1734, 38);
            this.colors.Name = "colors";
            this.colors.Size = new System.Drawing.Size(84, 23);
            this.colors.TabIndex = 20;
            this.colors.Text = "<- Color";
            this.colors.UseVisualStyleBackColor = true;
            this.colors.Click += new System.EventHandler(this.colors_Click);
            // 
            // comboBox2
            // 
            this.comboBox2.FormattingEnabled = true;
            this.comboBox2.Location = new System.Drawing.Point(1624, 12);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(104, 21);
            this.comboBox2.TabIndex = 21;
            this.comboBox2.SelectedIndexChanged += new System.EventHandler(this.comboBox2_SelectedIndexChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(1561, 40);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(57, 13);
            this.label3.TabIndex = 22;
            this.label3.Text = "Odd Rows";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(1561, 15);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(62, 13);
            this.label5.TabIndex = 24;
            this.label5.Text = "Even Rows";
            // 
            // Diodeinsbtn
            // 
            this.Diodeinsbtn.Location = new System.Drawing.Point(1237, 4);
            this.Diodeinsbtn.Name = "Diodeinsbtn";
            this.Diodeinsbtn.Size = new System.Drawing.Size(165, 22);
            this.Diodeinsbtn.TabIndex = 25;
            this.Diodeinsbtn.Text = "Insert 0.7V/+-10% for diodes";
            this.Diodeinsbtn.UseVisualStyleBackColor = true;
            this.Diodeinsbtn.Click += new System.EventHandler(this.Diodeinsbtn_Click);
            // 
            // Tolerancebtn
            // 
            this.Tolerancebtn.Location = new System.Drawing.Point(1237, 35);
            this.Tolerancebtn.Name = "Tolerancebtn";
            this.Tolerancebtn.Size = new System.Drawing.Size(127, 22);
            this.Tolerancebtn.TabIndex = 26;
            this.Tolerancebtn.Text = "Tolerance less than ->";
            this.Tolerancebtn.UseVisualStyleBackColor = true;
            this.Tolerancebtn.Click += new System.EventHandler(this.Tolerancebtn_Click);
            // 
            // comboBox3
            // 
            this.comboBox3.FormattingEnabled = true;
            this.comboBox3.Items.AddRange(new object[] {
            "1%",
            "2%",
            "3%",
            "4%",
            "5%",
            "6%",
            "7%",
            "8%",
            "9%",
            "10%",
            "11%",
            "12%",
            "13%",
            "14%",
            "15%"});
            this.comboBox3.Location = new System.Drawing.Point(1370, 36);
            this.comboBox3.Name = "comboBox3";
            this.comboBox3.Size = new System.Drawing.Size(60, 21);
            this.comboBox3.TabIndex = 27;
            this.comboBox3.SelectedIndexChanged += new System.EventHandler(this.comboBox3_SelectedIndexChanged);
            // 
            // Markbtn
            // 
            this.Markbtn.Location = new System.Drawing.Point(1408, 4);
            this.Markbtn.Name = "Markbtn";
            this.Markbtn.Size = new System.Drawing.Size(142, 23);
            this.Markbtn.TabIndex = 28;
            this.Markbtn.Text = "Mark unnecessary element";
            this.Markbtn.UseVisualStyleBackColor = true;
            this.Markbtn.Click += new System.EventHandler(this.Markbtn_Click);
            // 
            // headerRowDeletebtn
            // 
            this.headerRowDeletebtn.Cursor = System.Windows.Forms.Cursors.SizeAll;
            this.headerRowDeletebtn.Location = new System.Drawing.Point(695, 36);
            this.headerRowDeletebtn.Name = "headerRowDeletebtn";
            this.headerRowDeletebtn.Size = new System.Drawing.Size(107, 22);
            this.headerRowDeletebtn.TabIndex = 29;
            this.headerRowDeletebtn.Text = "Delete Header Line";
            this.headerRowDeletebtn.UseVisualStyleBackColor = true;
            this.headerRowDeletebtn.Click += new System.EventHandler(this.headerRowDeletebtn_Click);
            // 
            // Insert_ManufacturerNumberbtn
            // 
            this.Insert_ManufacturerNumberbtn.Location = new System.Drawing.Point(194, 34);
            this.Insert_ManufacturerNumberbtn.Name = "Insert_ManufacturerNumberbtn";
            this.Insert_ManufacturerNumberbtn.Size = new System.Drawing.Size(116, 23);
            this.Insert_ManufacturerNumberbtn.TabIndex = 31;
            this.Insert_ManufacturerNumberbtn.Text = "Manufacture Number";
            this.Insert_ManufacturerNumberbtn.UseVisualStyleBackColor = true;
            this.Insert_ManufacturerNumberbtn.Click += new System.EventHandler(this.Insert_ManufacturerNumberbtn_Click);
            // 
            // Insert_Descriptionbtn
            // 
            this.Insert_Descriptionbtn.Location = new System.Drawing.Point(316, 34);
            this.Insert_Descriptionbtn.Name = "Insert_Descriptionbtn";
            this.Insert_Descriptionbtn.Size = new System.Drawing.Size(102, 23);
            this.Insert_Descriptionbtn.TabIndex = 32;
            this.Insert_Descriptionbtn.Text = "Description";
            this.Insert_Descriptionbtn.UseVisualStyleBackColor = true;
            this.Insert_Descriptionbtn.Click += new System.EventHandler(this.Insert_Descriptionbtn_Click);
            // 
            // Insert_CircuitReferencebtn
            // 
            this.Insert_CircuitReferencebtn.Location = new System.Drawing.Point(86, 34);
            this.Insert_CircuitReferencebtn.Name = "Insert_CircuitReferencebtn";
            this.Insert_CircuitReferencebtn.Size = new System.Drawing.Size(102, 23);
            this.Insert_CircuitReferencebtn.TabIndex = 33;
            this.Insert_CircuitReferencebtn.Text = "Circuit Reference";
            this.Insert_CircuitReferencebtn.UseVisualStyleBackColor = true;
            this.Insert_CircuitReferencebtn.Click += new System.EventHandler(this.Insert_CircuitReferencebtn_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(12, 39);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(74, 13);
            this.label4.TabIndex = 34;
            this.label4.Text = "Insert Header:";
            // 
            // NecesseryColumnbtn
            // 
            this.NecesseryColumnbtn.Location = new System.Drawing.Point(424, 34);
            this.NecesseryColumnbtn.Name = "NecesseryColumnbtn";
            this.NecesseryColumnbtn.Size = new System.Drawing.Size(117, 23);
            this.NecesseryColumnbtn.TabIndex = 35;
            this.NecesseryColumnbtn.Text = "Unnecessary Column";
            this.NecesseryColumnbtn.UseVisualStyleBackColor = true;
            this.NecesseryColumnbtn.Click += new System.EventHandler(this.NecesseryColumnbtn_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(1830, 982);
            this.Controls.Add(this.NecesseryColumnbtn);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.Insert_CircuitReferencebtn);
            this.Controls.Add(this.Insert_Descriptionbtn);
            this.Controls.Add(this.Insert_ManufacturerNumberbtn);
            this.Controls.Add(this.headerRowDeletebtn);
            this.Controls.Add(this.Markbtn);
            this.Controls.Add(this.comboBox3);
            this.Controls.Add(this.Tolerancebtn);
            this.Controls.Add(this.Diodeinsbtn);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.comboBox2);
            this.Controls.Add(this.colors);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.textBox8);
            this.Controls.Add(this.textBox7);
            this.Controls.Add(this.textBox6);
            this.Controls.Add(this.textBox5);
            this.Controls.Add(this.textBox4);
            this.Controls.Add(this.textBox3);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.Exit_btn);
            this.Controls.Add(this.Modify);
            this.Controls.Add(this.dataGridView2);
            this.Controls.Add(this.BomMake);
            this.Controls.Add(this.cboSheet);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtFilename);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnBrowse);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Form1";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nullLogFactoryBindingSource)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtFilename;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cboSheet;
        private System.Windows.Forms.Button BomMake;
        private System.Windows.Forms.DataGridView dataGridView2;
        private System.Windows.Forms.BindingSource nullLogFactoryBindingSource;
        private System.Windows.Forms.Button Modify;
        private System.Windows.Forms.Button Exit_btn;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.TextBox textBox4;
        private System.Windows.Forms.TextBox textBox5;
        private System.Windows.Forms.TextBox textBox6;
        private System.Windows.Forms.TextBox textBox7;
        private System.Windows.Forms.TextBox textBox8;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Button colors;
        private System.Windows.Forms.ComboBox comboBox2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button Diodeinsbtn;
        private System.Windows.Forms.Button Tolerancebtn;
        private System.Windows.Forms.ComboBox comboBox3;
        private System.Windows.Forms.Button Markbtn;
        private System.Windows.Forms.Button headerRowDeletebtn;
        private System.Windows.Forms.Button Insert_ManufacturerNumberbtn;
        private System.Windows.Forms.Button Insert_Descriptionbtn;
        private System.Windows.Forms.Button Insert_CircuitReferencebtn;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button NecesseryColumnbtn;
    }
}


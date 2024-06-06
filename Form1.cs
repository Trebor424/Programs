using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDataReader;
using System.IO;
using Microsoft.SqlServer;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;


namespace AutoBOMmaker
{
    public partial class Form1 : Form
    {
        DataTableCollection tableCollection;

        public Form1()
        {

            InitializeComponent();

            //Block buttons
            Modify.Enabled = false;
            BomMake.Enabled = false;
            Diodeinsbtn.Enabled = false;
            Tolerancebtn.Enabled = false;
            comboBox3.Enabled = false;
            Markbtn.Enabled = false;
            headerRowDeletebtn.Enabled = false;
            Insert_CircuitReferencebtn.Enabled = false;
            Insert_Descriptionbtn.Enabled = false;
            Insert_ManufacturerNumberbtn.Enabled = false;
            NecesseryColumnbtn.Enabled = false;


            //note
            textBox6.Text = "UWAGI";
            textBox5.Text = "INSTRUKCJA";
            textBox4.Text = "Można wczytać zrobiony plik jeżeli chcemy operować na polach discription";
            textBox3.Text = "Plik Excel musi posiadać rozszerzenie .xlsx";
            textBox2.Text = "W przypadklu błędu program otwiera zadania Excel w tle, należy zamknąć ich zadanie. A program poprawnie będzie działać";
            textBox1.Text = "EXCEL powinien być przygotowany w ten sposób, aby górne linie zawierały elementy Circut Elements, Disription, manufacture part number. Jeżeli aplikacja nie bedzie w stanie wyszukać elementów, należy podmienić nazwy kolumn";

            Excel_Function.Excel_init();

            //Begin Values of colors textbox
            comboBox1.DataSource = typeof(Color).GetProperties()
            .Where(x => x.PropertyType == typeof(Color))
            .Select(x => x.GetValue(null)).ToList();
            comboBox1.SelectedItem = Color.LightGreen;

            comboBox2.DataSource = typeof(Color).GetProperties()
            .Where(x => x.PropertyType == typeof(Color))
            .Select(x => x.GetValue(null)).ToList();
            comboBox2.SelectedItem = Color.Bisque;
        }

        #region Forms1_Components

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Save data to excel in dataGridView
            System.Data.DataTable dt = tableCollection[cboSheet.SelectedItem.ToString()];
            dataGridView1.DataSource = dt;

            BomMake.Enabled = true;
            headerRowDeletebtn.Enabled = true;
            Insert_CircuitReferencebtn.Enabled = true;
            Insert_Descriptionbtn.Enabled = true;
            Insert_ManufacturerNumberbtn.Enabled = true;
            NecesseryColumnbtn.Enabled = true;
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            //Connect cbosheet to dataViewGrid and textbox

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtFilename.Text = openFileDialog.FileName;
                    tableCollection = Excel_Function.LoadTables(openFileDialog.FileName).Tables;

                    cboSheet.Items.Clear();
                    foreach (System.Data.DataTable table in tableCollection)
                    {
                        cboSheet.Items.Add(table.TableName);
                    }
                }
            }
        }

        private void BomMake_Click(object sender, EventArgs e)
        {
            try
            {

                Excel_Function.Excel_Clear();

                //Excel_Function.Excel_init();
                //Excel_Function.Excel_close();

                Excel_Function.Bom_make_func_2(dataGridView1);

                Excel_Function.Excel_save();

                this.dataGridView2.DataSource = null;
                dataGridView2.Rows.Clear();

                //Excel_Function.Excel_close();

                Excel_Function.ShowDataGridView(dataGridView2);


                dataGridView2.Columns[3].DefaultCellStyle.Format = "##.##";
                dataGridView2.Columns[4].DefaultCellStyle.Format = "##.##%";
                dataGridView2.Columns[5].DefaultCellStyle.Format = "##.##%";

                Excel_Function.colourDatasource(dataGridView1);
                Excel_Function.colourDatasource(dataGridView2);

                Modify.Enabled = true;
                Diodeinsbtn.Enabled = true;
                comboBox3.Enabled = true;
                Markbtn.Enabled = true;


            }
            catch (Exception)
            {
                MessageBox.Show("PLIK NIE ZOSTAŁ ZAPISANY, Sprawdz czy jest tylko jeden proces Excela, oraz czy nie masz otwartego pliku którego wgrywasz do programu");
            }
            MessageBox.Show("Wartości należy przejżeć, Aplikacja nie gwarantuje 100% poprawnych wyników :) \n Pliki zapisane na Pulpicie");
        }

        private void Modify_Click(object sender, EventArgs e)
        {
            //Excel_Function.Excel_init();


            Excel_Function.ToGrid_func(dataGridView2);

            Excel_Function.Insert_first_row(dataGridView2);
            Excel_Function.Excel_save();

            //Excel_Function.Excel_close();

            Excel_Function.colourDatasource(dataGridView1);
            Excel_Function.colourDatasource(dataGridView2);


        }

        private void Exit_btn_Click(object sender, EventArgs e)
        {
            //Excel_Function.Excel_close();
            this.Close();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (!(dataGridView1.DefaultCellStyle.BackColor == Excel_Function.Row_Colors_P) | !(dataGridView1.DefaultCellStyle.BackColor == Excel_Function.Row_Colors_N))
            {
                Excel_Function.colourDatasource(dataGridView1);
            }
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (!(dataGridView2.DefaultCellStyle.BackColor == Excel_Function.Row_Colors_P) | !(dataGridView2.DefaultCellStyle.BackColor == Excel_Function.Row_Colors_N))
            {
                Excel_Function.colourDatasource(dataGridView2);
            }
        }

        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            //Create a colour box in combobox
            comboBox1.MaxDropDownItems = 10;
            comboBox1.IntegralHeight = false;
            comboBox1.DrawMode = DrawMode.OwnerDrawFixed;
            comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox1.DrawItem += comboBox1_DrawItem;
            //define the begin colors in combobox
            if (comboBox1.SelectedIndex >= 0)
                Excel_Function.Row_Colors_N = (Color)comboBox1.SelectedValue;
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Create a colour box in combobox  
            comboBox2.MaxDropDownItems = 10;
            comboBox2.IntegralHeight = false;
            comboBox2.DrawMode = DrawMode.OwnerDrawFixed;
            comboBox2.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox2.DrawItem += comboBox1_DrawItem;
            //define the begin colors in combobox
            if (comboBox2.SelectedIndex >= 0)
                Excel_Function.Row_Colors_P = (Color)comboBox2.SelectedValue;
        }

        private void Diodeinsbtn_Click(object sender, EventArgs e)
        {
            dataGridView2.Sort(dataGridView2.Columns["Circuit Reference"], ListSortDirection.Ascending);
            Excel_Function.Diode_fct(dataGridView2);
            Excel_Function.colourDatasource(dataGridView2);
        }

        private void Tolerancebtn_Click(object sender, EventArgs e)
        {
            dataGridView2.Sort(dataGridView2.Columns["Circuit Reference"], ListSortDirection.Ascending);
            Excel_Function.Combobox3_Values = comboBox3.SelectedItem.ToString();
            Excel_Function.Tolerance_fcn(dataGridView2);
            Excel_Function.colourDatasource(dataGridView2);

        }

        private void Markbtn_Click(object sender, EventArgs e)
        {
            Excel_Function.Mark_fct(dataGridView2, Excel_Function.Mark_List);
            Excel_Function.colourDatasource(dataGridView2);
        }

        private void headerRowDeletebtn_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt = tableCollection[cboSheet.SelectedItem.ToString()];

            for (int i = 0; i < dataGridView1.ColumnCount; i++)
            {
                //MessageBox.Show(dataGridView1[i, 0].Value.ToString());
                dataGridView1.Columns[i].HeaderText = dataGridView1[i, 0].Value.ToString();
            }

            dt.Rows.RemoveAt(0);
            dataGridView1.DataSource = dt;

            Excel_Function.colourDatasource(dataGridView1);

        }

        private void Insert_Descriptionbtn_Click(object sender, EventArgs e)
        {
            Excel_Function.InsertHeader(dataGridView1, "Description");
        }

        private void Insert_ManufacturerNumberbtn_Click(object sender, EventArgs e)
        {
            Excel_Function.InsertHeader(dataGridView1, "Manufacture Number");
        }

        private void Insert_CircuitReferencebtn_Click(object sender, EventArgs e)
        {
            Excel_Function.InsertHeader(dataGridView1, "Circuit Reference");
        }

        private void NecesseryColumnbtn_Click(object sender, EventArgs e)
        {
            Excel_Function.InsertHeader(dataGridView1, "Uneccesery Column");
        }

        #endregion

        #region Form1_Functions

        private void comboBox1_DrawItem(object sender, DrawItemEventArgs e)
        {
            //e.DrawBackground();
            if (e.Index >= 0)
            {
                var txt = comboBox1.GetItemText(comboBox1.Items[e.Index]);
                var color = (Color)comboBox1.Items[e.Index];
                var r1 = new System.Drawing.Rectangle(e.Bounds.Left + 1, e.Bounds.Top + 1,
                    2 * (e.Bounds.Height - 2), e.Bounds.Height - 2);
                var r2 = System.Drawing.Rectangle.FromLTRB(r1.Right + 2, e.Bounds.Top,
                    e.Bounds.Right, e.Bounds.Bottom);
                using (var b = new SolidBrush(color))
                    e.Graphics.FillRectangle(b, r1);
                e.Graphics.DrawRectangle(Pens.Black, r1);
                TextRenderer.DrawText(e.Graphics, txt, comboBox1.Font, r2,
                    comboBox1.ForeColor, TextFormatFlags.Left | TextFormatFlags.VerticalCenter);
            }
        }

        private void colors_Click(object sender, EventArgs e)
        {
            //Color functions
            Excel_Function.colourDatasource(dataGridView1);
            Excel_Function.colourDatasource(dataGridView2);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);

            if (e.CloseReason == CloseReason.WindowsShutDown) return;

            // Confirm user wants to close
            switch (MessageBox.Show(this, "Jesteś pewien ,że chcesz wyjść ?", "Closing", MessageBoxButtons.YesNo))
            {
                case DialogResult.No:
                    e.Cancel = true;
                    break;
                default:
                    Excel_Function.Excel_close();
                    break;
            }
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            Tolerancebtn.Enabled = true;
            //Excel_Function.Combobox3_Values = comboBox3.SelectedValue
        }

        #endregion


    }
}

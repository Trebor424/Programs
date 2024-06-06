using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using _Excel = Microsoft.Office.Interop.Excel;

namespace AutoBOMmaker
{
    static class Excel_Function
    {
        #region Global variables

        //Variables whose take the value index of label column
        internal static int Circut_Reference_Column_Values;
        

        //List of elements whose should be mark for remove from BOM
        public static List<string> Mark_List = new List<string>() { "MAINBOARD","LABEL", "PCB", "FLUE", "RIB", "EPOXY", "GLUE", "PASTE", "FLUX", "BAG", "BOX", "CARTON"
        , "FOAM", "DO NOT POPULATE", "BARCODE", "BAR CODE", "BARE-BOARD" ,"ASSEMBLY","SCREW"};

        //List of elements whose should be deleted
        public static List<string> Delete_List = new List<string>() { "PIN", "LABEL", "CON", "PCB", "DIODE", "FLUE", "RIB", "FERRIT", "CRYSTAL", "CHOKE", "EPOXY" };
        //Elements whose should be deleted
        public static List<string> Remove_Element = new List<string>() { "RST", "CHIP", "RES", "CAP", "PPM", "T/R" };

        //List of elements for resistor/capacitor/inductor/numbers search function
        private static List<string> resistor_list = new List<string> { "KOHM", "KOHMS", "OHMS", "OHM", "R", "M", "K" };
        private static List<string> capacitor_list = new List<string> { "P", "N", "µ", "U", "M", "F", };
        private static List<string> inductor_list = new List<string> { "P", "N", "µ", "M", "H" };
        private static List<string> diode_list = new List<string> { "DIODE", "DIO", "ZEN", "ZNR", "ZENER", "SCHOT", "LED", "SCHOTTKY" };

        private static List<char> numbers_list = new List<char> { '1', '2', '3', '4', '5', '6', '7', '8', '9', '0' };

        //Excel Functions variables
        private static _Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
        private static _Excel.Workbook xlWorkBook;
        private static _Excel.Worksheet xlWorkSheet;
        private static object misValue = System.Reflection.Missing.Value;

        //File save and location variables
        private static string filePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        internal static string fileName1 = filePath + @"\Bom_Make-Excel";
        internal static string fileName2 = filePath + @"\Bom_Make-txt-for-merlin";
        internal static bool flag = false;
        internal static bool flag2 = true;

        //Begin colour of datagrid 1 and 2 values
        internal static Color Row_Colors_P = Color.LightGreen;
        internal static Color Row_Colors_N = Color.Bisque;

        internal static string Combobox3_Values;
        internal static DataSet Remember_results;


        #endregion


        internal static DataSet LoadTables(string fileName)
        {
            //Load data to datagridview1 from combobox
            using (var stream = File.Open(fileName, FileMode.Open, FileAccess.Read))
            {
                using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                {
                    DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                    });

                    return result;
                }
            }
        }

        public static void InsertHeader(DataGridView dataGridViewX, string NewHeader)
        {
            flag2 = true;
            for (int i = 0; i < dataGridViewX.ColumnCount; i++)
            {
                if (NewHeader == dataGridViewX.Columns[i].HeaderText & NewHeader != "Uneccesery Column")
                {
                    flag2 = false;
                    MessageBox.Show("You have this column name");
                    break;
                }
            }

            if (flag2==true)
            {
                dataGridViewX.Columns[dataGridViewX.CurrentCell.ColumnIndex].HeaderText = NewHeader;
                colourDatasource(dataGridViewX);
            }
        }

        public static void colourDatasource(DataGridView dataGridViewX)
        {
            //Colour data rows
            foreach (DataGridViewRow row in dataGridViewX.Rows)
            {

                if (row.Index % 2 == 0)
                {
                    if (Row_Colors_P != null)
                        dataGridViewX.Rows[row.Index].DefaultCellStyle.BackColor = Row_Colors_P;
                }

                if (row.Index % 2 == 1)
                {
                    if (Row_Colors_N != null)
                        dataGridViewX.Rows[row.Index].DefaultCellStyle.BackColor = Row_Colors_N;
                }
            }
            dataGridViewX.Refresh();
        }

        //Function from closing Excel files 
        public static void releaseObject(object obj)
        {
            try
            {
                //MessageBox.Show(obj.ToString());
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        //Init of Excel in background
        public static void Excel_init()
        {
            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
        }

        //Close of Excel in background
        public static void Excel_Clear()
        {
            {
                foreach (dynamic worksheet in xlWorkBook.Worksheets)
                {
                    worksheet.Cells.ClearContents();
                }
                xlWorkBook.Save();
            }
        }

        //Save Excel Files
        public static void Excel_save()
        {
            File.Delete(fileName1);
            File.Delete(fileName2);

            try
            {

                //save as Excel format
                xlWorkBook.SaveAs(Excel_Function.fileName1, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

                //Save as txt format
                xlWorkBook.SaveAs(Excel_Function.fileName2, Microsoft.Office.Interop.Excel.XlFileFormat.xlTextWindows, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

            }
            catch (Exception)
            { }
        }

        //Replace value and tolerance of diode to 0.7V and +-10%
        public static void Diode_fct(DataGridView dataGridViewx)
        {
            foreach (DataGridViewRow row in dataGridViewx.Rows)
            {
                if (!(dataGridViewx[2, row.Index].Value == null))
                {

                    string Diode_Help_Variable = "";
                    Diode_Help_Variable = dataGridViewx[2, row.Index].Value.ToString().ToUpper();

                    Diode_Help_Variable = Diode_Help_Variable.Replace(";", " ");
                    Diode_Help_Variable = Diode_Help_Variable.Replace(".", " ");
                    Diode_Help_Variable = Diode_Help_Variable.Replace(",", " ");
                    Diode_Help_Variable = Diode_Help_Variable.Replace("x", " ");
                    Diode_Help_Variable = Diode_Help_Variable.Replace("/", " ");

                    List<string> Diode_Help_Variable_list = Diode_Help_Variable.ToUpper().Split(' ').ToList();

                    foreach (string List_item in Diode_Help_Variable_list)
                    {
                        foreach (string diode in diode_list)
                        {
                            if (List_item.ToUpper() == diode.ToUpper())
                            {
                                dataGridViewx[3, row.Index].Value = "0.7";
                                dataGridViewx[4, row.Index].Value = "10%";
                                dataGridViewx[5, row.Index].Value = "10%";
                            }

                        }
                    }

                }
            }
            dataGridViewx.Refresh();
        }

        //replace value <5%
        public static void Tolerance_fcn(DataGridView dataGridViewx)
        {
            foreach (DataGridViewRow row in dataGridViewx.Rows)
            {
                try
                {
                    if (!(dataGridViewx[4, row.Index].Value == null))
                    {
                        string HelpVariable = (dataGridViewx[4, row.Index].Value.ToString().Replace("%", ""));
                        HelpVariable = HelpVariable.Replace(".", ",");
                        double Operate_value = double.Parse(HelpVariable);

                        //MessageBox.Show(Operate_value.ToString());

                        double Combobox3_Values_doubble = double.Parse(Combobox3_Values.ToString().Replace("%", ""));

                        if (Operate_value < Combobox3_Values_doubble)
                        {
                            //MessageBox.Show(dataGridViewx[4, row.Index].Value.ToString()+ "<-4      " + dataGridViewx[5, row.Index].Value + "<-5");
                            //dataGridViewx[4, row.Index].Value = 0.05;
                            //dataGridViewx[5, row.Index].Value = 0.05;

                            dataGridViewx[4, row.Index].Value = Combobox3_Values_doubble.ToString() + "%";
                            dataGridViewx[5, row.Index].Value = Combobox3_Values_doubble.ToString() + "%";
                            //MessageBox.Show(dataGridViewx[4, row.Index].Value.ToString() + "<-4     " + dataGridViewx[5, row.Index].Value + "<-5");
                        }
                    }
                }
                catch (Exception)
                {
                    continue;
                }
            }

            dataGridViewx.Refresh();
        }

        //Mark Values for Delete
        public static void Mark_fct(DataGridView dataGridViewx, List<string> Values_list)
        {
            foreach (DataGridViewRow row in dataGridViewx.Rows)
            {
                foreach (string Mark in Mark_List)
                {
                    try
                    {
                        //MessageBox.Show(dataGridViewx[3, row.Index].Value.ToString() + "-> "+Mark);

                        if (dataGridViewx[2, row.Index].Value.ToString().Contains(Mark.ToUpper()) | dataGridViewx[0, row.Index].Value.ToString() == "")
                        {
                            dataGridViewx[3, row.Index].Value = "PROBABLY FOR DELETE!";
                        }
                    }
                    catch (Exception)
                    { }
                }
            }
            dataGridViewx.Refresh();
        }

        //Close The Excel files
        public static void Excel_close()
        {
            try
            {
                //Close the exel Sheets
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                Excel_Function.releaseObject(xlWorkSheet);
                Excel_Function.releaseObject(xlWorkBook);
                Excel_Function.releaseObject(xlApp);
            }
            catch (Exception)
            { }
        }

        //Main function whose make Datagrid2 from datagrid 1
        public static void Bom_make_func_2(DataGridView dataGridView1)
        {
            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                //---------------------------------------- Circut Reference------------------------------------------
                if (
                column.HeaderText.ToLower().Contains("circuit reference") |
                column.HeaderText.ToLower().Contains("locations") |
                column.HeaderText.ToLower().Contains("reference"))
                {
                    xlWorkSheet.Cells[1, 1] = "Circuit Reference";
                    xlWorkSheet.Cells[1, 2] = "Manfacturer Part Number ";
                    xlWorkSheet.Cells[1, 3] = "Description";
                    xlWorkSheet.Cells[1, 4] = "Value";
                    xlWorkSheet.Cells[1, 5] = "Tolerance POS";
                    xlWorkSheet.Cells[1, 6] = "Tolerance NEG";

                    //Save Column index for Values Excemptions 
                    Circut_Reference_Column_Values = column.Index;

                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        xlWorkSheet.Cells[row.Index + 2, 1] = dataGridView1[column.Index, row.Index].Value;
                    }
                }
            }

            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                //------------------------------------- Manfacturer Part # -------------------------------------------
                if (
                column.HeaderText.ToLower().Contains("ampl") |
                column.HeaderText.ToLower().Contains("manufacture") |
                column.HeaderText.ToLower().Contains("manfacturer"))
                {
                    try
                    {
                        string AMPL_help_Variable;

                        foreach (DataGridViewRow row in dataGridView1.Rows)
                        {
                            try
                            {
                                AMPL_help_Variable = (String)dataGridView1[column.Index, row.Index].Value;
                            }
                            catch (Exception)
                            {
                                continue;
                            }
                            List<string> AMPL_help_list = new List<string>();

                            AMPL_help_list = AMPL_help_Variable.Split(' ').ToList();

                            xlWorkSheet.Cells[row.Index + 2, 2] = AMPL_help_list[0];
                        }
                    }
                    catch (Exception)
                    {
                        continue;
                    }
                }

                //---------------------------------------Description -------------------------------------------------
                if (column.HeaderText.ToLower().Contains("description"))
                {
                    //Description save
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        xlWorkSheet.Cells[row.Index + 2, 3] = dataGridView1[column.Index, row.Index].Value;
                    }

                    string Help_Variables;

                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {

                        List<string> LIST_Help_Variables = new List<string>();
                        try
                        {
                            //Split Description 
                            xlWorkSheet.Cells[row.Index + 2, 3] = (String)dataGridView1[column.Index, row.Index].Value;
                            Help_Variables = (String)dataGridView1[column.Index, row.Index].Value;
                            Help_Variables = Help_Variables.Replace(";", " ");
                            Help_Variables = Help_Variables.Replace("@", " ");
                            Help_Variables = Help_Variables.Replace("%", "% ");
                            Help_Variables = Help_Variables.Replace("X", " ");
                            Help_Variables = Help_Variables.Replace("OHMS", "OHMS ");
                            Help_Variables = Help_Variables.Replace("OHM", "OHM ");
                            Help_Variables = Help_Variables.Replace("#%", " #%");

                            //Delete items from Delete_List
                            foreach (var delete_item in Delete_List)
                            {
                                if (Help_Variables.ToUpper().Contains(delete_item))
                                {
                                    Help_Variables = " ";
                                }
                            }

                            LIST_Help_Variables = Help_Variables.ToUpper().Split(' ').ToList();
                        }
                        catch (Exception)
                        {
                            continue;
                        }

                        //----------------------------Value Exceptions---------------------
                        List<string> LIST_Help_Variables0 = new List<string>();
                        int i = 0;

                        //--------Handling Excemption Values
                        foreach (string item in LIST_Help_Variables)
                        {

                            string item0;
                            item0 = item.ToUpper();

                            foreach (string Remove in Remove_Element)
                            {
                                if (item0.ToUpper().Contains(Remove))
                                {
                                    item0 = " ";
                                }
                            }
                            if (item0.Contains("+-") | item0.Contains("-+"))
                            {
                                item0 = " " + item0;
                            }
                            if (item.ToUpper().Contains("ZERO"))
                            {
                                item0 = item.ToUpper().Replace("ZERO", "0");
                            }

                            if (item.ToUpper().Contains("µ"))
                            {
                                item0 = item0.Replace("µ", "U");
                            }

                            if (item0.ToUpper().Contains("MM"))
                            {
                                item0 = item0.Replace("MM", "");
                            }
                            if (item0.EndsWith(",") | item0.EndsWith("."))
                            {
                                try
                                {
                                    item0 = item0.Remove(item.Count() - 1, 1);
                                }
                                catch (Exception)
                                { }
                            }
                            if (item0.StartsWith(",") | item0.EndsWith("."))
                            {
                                item0 = item0.Remove(0, 1);
                            }
                            if (item0 == "%")
                            {
                                item0 = LIST_Help_Variables[i - 1] + "%";
                            }
                            if ((item0.Contains("+/-") | item0.Contains("-/+")) & item0.EndsWith("%"))
                            {
                                item0 = item0.Remove(0, 3);
                                LIST_Help_Variables0.Add(item0 + "%");
                            }
                            if (item.Contains('/'))
                            {
                                try
                                {
                                    List<string> Split_help_list = new List<string>();
                                    Split_help_list = item0.Split('/').ToList();

                                    foreach (string SplitItem in Split_help_list)
                                    {
                                        LIST_Help_Variables0.Add(SplitItem);
                                        if (Split_help_list[1].IsNumeric() == true)
                                        {
                                            string SplitItem1 = Split_help_list[1] + "%";
                                            LIST_Help_Variables0.Add(SplitItem1);
                                        }
                                    }
                                }
                                catch (Exception)
                                { }
                                continue;
                            }

                            LIST_Help_Variables0.Add(item0);
                            i++;
                        }
                        //MAIN values functions
                        //Resistor Values
                        if (dataGridView1[Circut_Reference_Column_Values, row.Index].Value.ToString().ToUpper().Contains("R"))
                        {
                            Values_Find_Function(LIST_Help_Variables0, row, resistor_list);
                            //Tolerance for resistor
                            Tolerance_Find_Function(LIST_Help_Variables0, row);
                        }
                        //Capacitor Values
                        if (dataGridView1[Circut_Reference_Column_Values, row.Index].Value.ToString().ToUpper().Contains("C"))
                        {
                            Values_Find_Function(LIST_Help_Variables0, row, capacitor_list);
                            //Tolerance for capacitor
                            Tolerance_Find_Function(LIST_Help_Variables0, row);
                        }
                        //Inductor Values
                        if (dataGridView1[Circut_Reference_Column_Values, row.Index].Value.ToString().ToUpper().Contains("L"))
                        {
                            Values_Find_Function(LIST_Help_Variables0, row, inductor_list);
                            //Tolerance for inductor
                            Tolerance_Find_Function(LIST_Help_Variables0, row);
                        }
                    }
                }
            }
        }

        public static void Insert_first_row(DataGridView dataGridView2)
        {

            //Insert a first row
            xlWorkSheet.Rows["1"].Insert();
            //Insert a name (header row) into excel
            xlWorkSheet.Cells[1, 1] = "Circuit Reference";
            xlWorkSheet.Cells[1, 2] = "Manfacturer Part Number ";
            xlWorkSheet.Cells[1, 3] = "Description";
            xlWorkSheet.Cells[1, 4] = "Value";
            xlWorkSheet.Cells[1, 5] = "Tolerance POS";
            xlWorkSheet.Cells[1, 6] = "Tolerance NEG";
        }

        public static void ShowDataGridView(DataGridView dataGridViewX)
        {
            try
            {


                FileStream stream = File.Open(Excel_Function.fileName1 + ".xlsx", FileMode.Open, FileAccess.Read);
                IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                DataSet result = excelReader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = true
                    }
                });

                dataGridViewX.DataSource = result.Tables[0];
            }
            catch (Exception)
            {
                //MessageBox.Show("Dane nie zostały odczytane, sprawdz czy dane zostały zapisane");
            }
        }

        public static void ToGrid_func(DataGridView dataGridViewX)
        {
            Excel_Clear();

            //Take the elements from exel sheets to dataGridView2
            for (int i = 0; i <= dataGridViewX.RowCount - 1; i++)
            {
                for (int j = 0; j <= dataGridViewX.ColumnCount - 1; j++)
                {
                    DataGridViewCell cell = dataGridViewX[j, i];
                    xlWorkSheet.Cells[i + 1, j + 1] = cell.Value;
                }
            }
        }

        public static bool IsNumeric(this string text)
        {
            double test;
            return double.TryParse(text, out test);
        }

        //------------------------------------------List HELPVARIABLES, actual row, List of strings list of rez/cap/cew
        public static void Values_Find_Function(List<string> LIST_Help_Variables0, DataGridViewRow row, List<string> Values_String)
        {
            bool flaga = false;
            //--------Excemption for 2.123K
            foreach (var item0 in LIST_Help_Variables0)
            {
                foreach (char number in numbers_list)
                {
                    foreach (string value in Values_String)
                    {
                        if (item0.ToUpper().StartsWith(number.ToString()) & item0.ToUpper().EndsWith(value))
                        {
                            foreach (string values2 in Values_String)
                            {
                                if (item0.ToUpper().Contains(values2) & item0.ToUpper().EndsWith(value))
                                {
                                    string item1 = item0.ToUpper().Replace(values2, ".");
                                    item1 = item1.Replace(value, "");
                                    item1 = item1 + values2.ToString();
                                    item1 = item1.Replace("." + values2, values2);
                                    xlWorkSheet.Cells[row.Index + 2, 4] = item1.ToString();
                                    flaga = true;
                                    break;
                                }
                            }
                            if (flaga == false)
                            {
                                xlWorkSheet.Cells[row.Index + 2, 4] = item0.ToString();
                                flaga = true;
                            }

                        }
                    }
                }
            }
            //--------Excemption for 2k2
            if (flaga == false)
            {
                foreach (var item0 in LIST_Help_Variables0)
                {
                    foreach (char number in numbers_list)
                    {
                        foreach (char number2 in numbers_list)
                        {
                            if (item0.ToUpper().StartsWith(number.ToString()) & item0.ToUpper().EndsWith(number2.ToString()) & item0.Count() < 8)
                            {
                                foreach (string value in Values_String)
                                {
                                    if (item0.ToUpper().Contains(value.ToString()))
                                    {
                                        string item1 = item0.ToUpper().Replace(value, ".");
                                        item1 = item1 + value.ToString();
                                        xlWorkSheet.Cells[row.Index + 2, 4] = item1.ToString();
                                        flaga = true;
                                    }
                                }
                            }
                        }
                    }
                }
            }

            //------------------50 OHM-----------------
            if (flaga == false)
            {
                int i = 0;
                foreach (var item0 in LIST_Help_Variables0)
                {
                    foreach (string item in Values_String)
                    {
                        if (item0.ToUpper() == item.ToUpper())
                        {
                            try
                            {
                                xlWorkSheet.Cells[row.Index + 2, 4] = LIST_Help_Variables0[i - 1].ToString();
                            }
                            catch (Exception)
                            { }
                        }

                    }
                    i++;
                }
            }
        }

        public static void Tolerance_Find_Function(List<string> LIST_Help_Variables0, DataGridViewRow row)
        {
            //-----------------------------tolerance exceptions--------------------- 
            foreach (string percent in LIST_Help_Variables0)
            {
                //Handling of percent List
                string percent1;
                percent1 = percent.ToUpper();

                if (percent1.ToUpper().StartsWith("PMIC"))
                {
                    percent1 = percent.Remove(0, 4);
                }
                if (percent1.ToUpper().StartsWith("DF<"))
                {
                    percent1 = percent.Remove(0, 3);
                }
                if (percent.ToUpper().StartsWith("DF:"))
                {
                    percent1 = percent.Remove(0, 3);
                }
                if (percent.ToUpper().StartsWith("<"))
                {
                    percent1 = percent.Remove(0, 1);
                }



                // values type +-80%
                if ((percent1.ToUpper().Contains("+-") | percent1.ToUpper().Contains("-+")) & percent1.ToUpper().Contains("%"))
                {
                    percent1 = percent.Replace("-", "");
                    percent1 = percent1.Replace("+", "");

                    percent1 = percent1.Replace("%", "");
                    xlWorkSheet.Cells[row.Index + 2, 5].NumberFormat = "@";
                    xlWorkSheet.Cells[row.Index + 2, 6].NumberFormat = "@";
                    xlWorkSheet.Cells[row.Index + 2, 5] = percent1.ToString() + "%";
                    xlWorkSheet.Cells[row.Index + 2, 6] = percent1.ToString() + "%";
                }
                // values type +80-20%
                else if (percent1.ToUpper().StartsWith("+") & percent1.ToUpper().Contains("-"))
                {
                    List<string> list_percent = new List<string>();
                    percent1 = percent1.Replace("+", "");
                    list_percent = percent1.Split('-').ToList();
                    percent1 = percent1.Replace("-", "");

                    percent1 = percent1.Replace("%", "");
                    xlWorkSheet.Cells[row.Index + 2, 5].NumberFormat = "@";
                    xlWorkSheet.Cells[row.Index + 2, 6].NumberFormat = "@";
                    xlWorkSheet.Cells[row.Index + 2, 5] = percent1.ToString() + "%";
                    xlWorkSheet.Cells[row.Index + 2, 6] = percent1.ToString() + "%";
                }
                // values type -80+20%
                else if (percent1.ToUpper().StartsWith("-") & percent1.ToUpper().Contains("+"))
                {
                    List<string> list_percent = new List<string>();
                    percent1 = percent1.Replace("-", "");
                    list_percent = percent1.Split('+').ToList();
                    percent1 = percent1.Replace("+", "");

                    percent1 = percent1.Replace("%", "");
                    xlWorkSheet.Cells[row.Index + 2, 5].NumberFormat = "@";
                    xlWorkSheet.Cells[row.Index + 2, 6].NumberFormat = "@";
                    xlWorkSheet.Cells[row.Index + 2, 5] = percent1.ToString() + "%";
                    xlWorkSheet.Cells[row.Index + 2, 6] = percent1.ToString() + "%";
                }
                //values type 2.5~3.5%
                else if (percent1.Contains("~") & percent1.EndsWith("%"))
                {
                    percent1 = percent1.Replace("%", "");
                    List<string> list_percent = new List<string>();
                    list_percent = percent1.Split('~').ToList();

                    xlWorkSheet.Cells[row.Index + 2, 5].NumberFormat = "@";
                    xlWorkSheet.Cells[row.Index + 2, 6].NumberFormat = "@";
                    xlWorkSheet.Cells[row.Index + 2, 5] = list_percent[1].ToString() + "%";
                    xlWorkSheet.Cells[row.Index + 2, 6] = list_percent[1].ToString() + "%";
                }
                // values type +80%
                else if (percent1.ToUpper().Contains("+") & percent1.ToUpper().Contains("%"))
                {
                    percent1 = percent.Replace("+", "");

                    percent1 = percent1.Replace("%", "");
                    xlWorkSheet.Cells[row.Index + 2, 5].NumberFormat = "@";
                    xlWorkSheet.Cells[row.Index + 2, 5] = percent1.ToString() + "%";
                }
                // values type-20%
                else if (percent1.ToUpper().Contains("-") & percent1.ToUpper().Contains("%"))
                {
                    percent1 = percent.Replace("-", "");

                    percent1 = percent1.Replace("%", "");
                    xlWorkSheet.Cells[row.Index + 2, 6].NumberFormat = "@";
                    xlWorkSheet.Cells[row.Index + 2, 6] = percent1.ToUpper().ToString() + "%";
                }
                //values type pm1
                else if (percent1.ToUpper().StartsWith("PM"))
                {
                    percent1 = percent1.Replace("PM", "");

                    percent1 = percent1.Replace("%", "");
                    xlWorkSheet.Cells[row.Index + 2, 5].NumberFormat = "@";
                    xlWorkSheet.Cells[row.Index + 2, 6].NumberFormat = "@";
                    xlWorkSheet.Cells[row.Index + 2, 5] = percent1.ToString() + "%";
                    xlWorkSheet.Cells[row.Index + 2, 6] = percent1.ToString() + "%";
                }
                // values type 20%
                else if (percent1.ToUpper().EndsWith("%"))
                {
                    percent1 = percent1.Replace("%", "");
                    xlWorkSheet.Cells[row.Index + 2, 5].NumberFormat = "@";
                    xlWorkSheet.Cells[row.Index + 2, 6].NumberFormat = "@";
                    xlWorkSheet.Cells[row.Index + 2, 5] = percent1.ToString() + "%";
                    xlWorkSheet.Cells[row.Index + 2, 6] = percent1.ToString() + "%";
                }
            }
        }
    }
}



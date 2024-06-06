using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.IO;

namespace AutoMiniTabMaker
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            # region Initialize Values 
            int j = 2;


            string MSA_info_string = "";
            string filePath = "";
            string CapaValue = "5";
            double HelpVariable = 0.1;
            string HelpVariable2 = "";
            string HelpVariable3 = "";
            string IfElse_Condition_String;

            string Measuring_instrument;
            string Reported_by;
            string Date;
            string Miscellaneous;
            int OperatorValue;

            string fileName = "Minitab_output_Capability_Normal.txt";
            string filename0 = "Minitab_ALL_method_CMD.txt";
            string fileName1 = "Minitab_output_Sixpack_Normal.txt";
            string fileName2 = "Minitab_output.xlsx";
            string fileName3 = "MiniTabValues_Output_string_Anova.txt";


            string Output_Files = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"AutoMiniTab_Output_Files");
            string filePathOutputSave = Path.Combine(Output_Files, fileName);
            string filePathOutputSave0 = Path.Combine(Output_Files, filename0);
            string filePathOutputSave1 = Path.Combine(Output_Files, fileName1);
            string filePathOutputSave2 = Path.Combine(Output_Files, fileName2);
            string filePathOutputSave3 = Path.Combine(Output_Files, fileName3);

            List<string> MiniTabValues_Output_string_Sixpack_Normal = new List<string> { };
            List<string> MiniTabValues_Output_string_Capability_Normal = new List<string> { };
            List<string> MiniTabValues_Output_string_Anova = new List<string> { };
            var namelist = new List<string>();


            #endregion

            #region Init output excel

            // Create an Excel application instance
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = false; // Make Excel application visible

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "All Files (*.*)|*.*"; // Set the file filter if needed

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    filePath = openFileDialog.FileName;
                    Console.WriteLine("Selected File Path: " + filePath);
                }
            }

            // Open an existing Excel file
            Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);
            Excel.Workbook Output_Workbook = excelApp.Workbooks.Add();

            // Get the first worksheet
            Excel.Worksheet excelSheet = (Excel.Worksheet)workbook.Worksheets[1];
            Excel.Worksheet Output_WorkSheet = Output_Workbook.Sheets.Add();
            Output_WorkSheet.Name = "Output_Sheet";

            #endregion

            #region Main functions

            int numberOfSheets = workbook.Sheets.Count;

            //Here we take excemption for whose Excel format we want
            excelSheet = (Excel.Worksheet)workbook.Worksheets.get_Item(1);
            IfElse_Condition_String = excelSheet.Name.ToString();

            //Qlab MSA exception
            if ("SLOWNIK" == IfElse_Condition_String)
            {
                for (int i = 1; i <= numberOfSheets; i++)
                {
                    excelSheet = (Excel.Worksheet)workbook.Worksheets.get_Item(i);
                    namelist.Add(excelSheet.Name);
                }

                excelSheet = (Excel.Worksheet)workbook.Worksheets.get_Item(namelist.IndexOf("Information") + 1);
                excelSheet.Activate();

                Excel.Range rangeInfo = excelSheet.UsedRange;
                int rowCountInfo = rangeInfo.Rows.Count;
                int colCount2Info = rangeInfo.Columns.Count;


                MSA_info_string = rangeInfo.Cells[1, 2].Value.ToString();
                MSA_info_string = MSA_info_string.ToString().Replace("MSA ", "");

                // Split the string into an array using space as the delimiter
                string[] MSA_info_array = MSA_info_string.Split('x');


                Measuring_instrument = excelSheet.Cells[2, 2].Value.ToString();
                Reported_by = excelSheet.Cells[6, 2].Value.ToString();
                Miscellaneous = excelSheet.Cells[3, 2].Value.ToString();
                Date = excelSheet.Cells[7, 2].Value.ToString();

                CapaValue = MSA_info_array[1];// Trim any leading/trailing whitespace
                OperatorValue = int.Parse(MSA_info_array[0]);// Trim any leading/trailing whitespace

                // Read data from cells
                excelSheet = (Excel.Worksheet)workbook.Worksheets.get_Item(namelist.IndexOf("Data") + 1);
                excelSheet.Activate();

                Excel.Range range = excelSheet.UsedRange;
                int rowCount = range.Rows.Count;
                int colCount = range.Columns.Count;




                //For ANOVA
                Output_WorkSheet.Cells[1, 1] = "Operator";
                Output_WorkSheet.Cells[1, 2] = "Serialnumber";
                for (int i = 2; i < rowCount - 7; i++)
                {
                    string[] HelpVariable3_list = range.Cells[i + 2, 1].Value2.ToString().Split('_');
                    HelpVariable3 = HelpVariable3_list[0];
                    Output_WorkSheet.Cells[i, 2] = HelpVariable3.ToString();
                    Output_WorkSheet.Cells[i, 1] = "1";

                    //Console.WriteLine(range.Cells[i + 2, col].Value2.ToString() +"   "+HelpVariable3 +"     " +  Output_WorkSheet.Cells[i ,j].Value2.ToString());
                }

                for (int col = 7; col <= colCount; col++) //7
                {
                    try
                    {
                        HelpVariable = double.TryParse(range.Cells[rowCount - 1, col].Value2.ToString().Replace("%", ""), out HelpVariable);
                    }
                    catch (Exception)
                    {

                    }


                    try
                    {
                        if (string.IsNullOrEmpty(range.Cells[3, col].Value.ToString()))
                        {
                            if (string.IsNullOrEmpty(range.Cells[3, col].ToString()))
                            {
                                if (string.IsNullOrEmpty(range.Cells[3, col].ToString()))
                                {
                                    break;
                                }
                            }
                        }
                    }
                    catch (Exception)
                    {

                        break;
                    }



                    if ((!range.Cells[rowCount - 1, col].Value.ToString().Contains("Attribute")
                        & (HelpVariable > 0.1)))
                    {
                        j++;
                        HelpVariable2 = range.Cells[3, col].Value.ToString().ToLower();

                        if (range.Cells[3, col].Value.ToString().ToLower().StartsWith("c"))
                        {
                            HelpVariable2 = HelpVariable2.Remove(0, 1);
                            HelpVariable2 = "_C" + HelpVariable2;

                        }
                        Output_WorkSheet.Cells[1, j] = HelpVariable2;

                        Console.WriteLine(HelpVariable2);

                        for (int i = 2; i < rowCount - 7; i++)
                        {
                            HelpVariable3 = range.Cells[i + 2, col].Value2.ToString();
                            Output_WorkSheet.Cells[i, j] = double.Parse(HelpVariable3);
                            //Console.WriteLine(range.Cells[i + 2, col].Value2.ToString() +"   "+HelpVariable3 +"     " +  Output_WorkSheet.Cells[i ,j].Value2.ToString());
                        }

                        //Script command line finction
                        Capability_Normal_Function(HelpVariable2.ToString(), range.Cells[1, col].Value2.ToString(), range.Cells[2, col].Value2.ToString(),CapaValue, MiniTabValues_Output_string_Capability_Normal);
                        Sixpack_Normal_Function(HelpVariable2.ToString(), range.Cells[1, col].Value2.ToString(), range.Cells[2, col].Value2.ToString(), MiniTabValues_Output_string_Sixpack_Normal);
                        Anova_Normal_Function(HelpVariable2.ToString(), range.Cells[1, col].Value2.ToString(), range.Cells[2, col].Value2.ToString(), MiniTabValues_Output_string_Anova);

                    }
                }
            }
            else if (IfElse_Condition_String == "Info")
            {
                //industry MSA exception 30x1
                for (int i = 2; i <= numberOfSheets; i++)
                {
                    excelSheet = (Excel.Worksheet)workbook.Worksheets.get_Item(i);
                    namelist.Add(excelSheet.Name);
                    excelSheet.Activate();

                    Excel.Range range = excelSheet.UsedRange;
                    int rowCount = range.Rows.Count;
                    int colCount = range.Columns.Count;


                    HelpVariable2 = excelSheet.Name.ToString();
                    Output_WorkSheet.Cells[1, i - 1] = HelpVariable2;

                    if (excelSheet.Name.ToString().ToLower().StartsWith("c"))
                    {
                        HelpVariable2 = HelpVariable2.Remove(0, 1);
                        HelpVariable2 = "_C" + HelpVariable2;

                        Output_WorkSheet.Cells[1, i - 1] = HelpVariable2;
                    }

                    for (int z = 1; z <= 30; z++)
                    {
                        HelpVariable3 = range.Cells[z + 6, 10].Value2.ToString();
                        Output_WorkSheet.Cells[z + 1, i - 1] = double.Parse(HelpVariable3);
                    }

                    //Script command line finction
                    Capability_Normal_Function(HelpVariable2, range.Cells[12, 3].Value2, range.Cells[11, 3].Value2,CapaValue, MiniTabValues_Output_string_Capability_Normal);
                    Sixpack_Normal_Function(HelpVariable2, range.Cells[12, 3].Value2, range.Cells[11, 3].Value2, MiniTabValues_Output_string_Sixpack_Normal);
                    //Anova_Normal_Function(HelpVariable2, range.Cells[12, 3].Value2, range.Cells[12, 3].Value2, MiniTabValues_Output_string_Anova);
                }
            }
            else
            {
                MessageBox.Show("Nie znaleziono odpowiedniego arkusza");

            }

            #endregion

            #region Save Files

            // Utwórz folder na podstawie daty, jeśli nie istnieje
           if (!Directory.Exists(Output_Files))
            {
                Directory.CreateDirectory(Output_Files);
            }


            //Delete files
            if (System.IO.File.Exists(filePathOutputSave)) { File.Delete(filePathOutputSave); }
            if (System.IO.File.Exists(filePathOutputSave0)) { File.Delete(filePathOutputSave0); }
            if (System.IO.File.Exists(filePathOutputSave1)) { File.Delete(filePathOutputSave1); }
            if (System.IO.File.Exists(filePathOutputSave2)) { File.Delete(filePathOutputSave2); }
            if (System.IO.File.Exists(filePathOutputSave3)) { File.Delete(filePathOutputSave3); }

            //save as Excel format
            Output_Workbook.SaveAs(
            filePathOutputSave2,                                                        // 1. The file path where the workbook will be saved.
            Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault     // 2. The file format (default Excel workbook format).
            );

            // Close and release resources
            workbook.Close(false);
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

            //Replace . for , (.-makes errors in program)

           

            //Console txt appendline
            File.AppendAllLines(filePathOutputSave, MiniTabValues_Output_string_Capability_Normal);
            File.AppendAllLines(filePathOutputSave1, MiniTabValues_Output_string_Sixpack_Normal);
            File.AppendAllLines(filePathOutputSave3, MiniTabValues_Output_string_Anova);

            File.AppendAllLines(filePathOutputSave0, MiniTabValues_Output_string_Capability_Normal);
            File.AppendAllLines(filePathOutputSave0, MiniTabValues_Output_string_Sixpack_Normal);
            File.AppendAllLines(filePathOutputSave0, MiniTabValues_Output_string_Anova);


        }

        #endregion

        #region function

        async static void Function_Async(string Rsub, string Lspec, string Uspec, string CapaValue, List<string> MiniTabValues_Output_string_Capability_Normal, List<string> MiniTabValues_Output_string_Sixpack_Normal, List<string> MiniTabValues_Output_string_Anova)
        {
            await Task.Run(() => Capability_Normal_Function(Rsub, Lspec, Uspec, CapaValue, MiniTabValues_Output_string_Capability_Normal));
            await Task.Run(() => Sixpack_Normal_Function(Rsub, Lspec, Uspec, MiniTabValues_Output_string_Sixpack_Normal));
            await Task.Run(() => Anova_Normal_Function(Rsub, Lspec, Uspec, MiniTabValues_Output_string_Anova));
        }

        
        
        static void Capability_Normal_Function(string Rsub, string Lspec, string Uspec,string CapaValue, List<string> MiniTabValues_Output_string_Capability_Normal)
        {
           
            //Here is a cmd lines for Capability Statistical Normal
            MiniTabValues_Output_string_Capability_Normal.Add("Capa " + "'" + Rsub + "'" + " " + CapaValue + ";");
            MiniTabValues_Output_string_Capability_Normal.Add("Lspec " + Lspec.Replace('.', ',') + ";");
            MiniTabValues_Output_string_Capability_Normal.Add("Uspec " + Uspec.Replace('.', ',') + ";");
            MiniTabValues_Output_string_Capability_Normal.Add("Pooled;");
            MiniTabValues_Output_string_Capability_Normal.Add("AMR;");
            MiniTabValues_Output_string_Capability_Normal.Add("UnBiased;");
            MiniTabValues_Output_string_Capability_Normal.Add("OBiased;");
            MiniTabValues_Output_string_Capability_Normal.Add("Toler 6;");
            MiniTabValues_Output_string_Capability_Normal.Add("Within;");
            MiniTabValues_Output_string_Capability_Normal.Add("Overall; ");
            MiniTabValues_Output_string_Capability_Normal.Add("NoCI;");
            MiniTabValues_Output_string_Capability_Normal.Add("PPM;");
            MiniTabValues_Output_string_Capability_Normal.Add("CStat.");
            MiniTabValues_Output_string_Capability_Normal.Add("");
        }

        static void Sixpack_Normal_Function(string Rsub, string Lspec, string Uspec, List<string> MiniTabValues_Output_string_Sixpack_Normal)
        {
            //Here is a cmd lines for Sicpack Statistical Normal
            MiniTabValues_Output_string_Sixpack_Normal.Add("Sixpack;");
            MiniTabValues_Output_string_Sixpack_Normal.Add("Rsub " + "'" + Rsub + "'" + ";");
            MiniTabValues_Output_string_Sixpack_Normal.Add("Lspec " + Lspec.Replace('.', ',') + ";");
            MiniTabValues_Output_string_Sixpack_Normal.Add("Uspec " + Uspec.Replace('.', ',') + ";");
            MiniTabValues_Output_string_Sixpack_Normal.Add("Pooled;");
            MiniTabValues_Output_string_Sixpack_Normal.Add("AMR;");
            MiniTabValues_Output_string_Sixpack_Normal.Add("CCRbar;");
            MiniTabValues_Output_string_Sixpack_Normal.Add("CCSbar;");
            MiniTabValues_Output_string_Sixpack_Normal.Add("CCAMR;");
            MiniTabValues_Output_string_Sixpack_Normal.Add("UnBiased;");
            MiniTabValues_Output_string_Sixpack_Normal.Add("OBiased;");
            MiniTabValues_Output_string_Sixpack_Normal.Add("Breakout 25;");
            MiniTabValues_Output_string_Sixpack_Normal.Add("Toler 6;");
            MiniTabValues_Output_string_Sixpack_Normal.Add("CStat.");
            MiniTabValues_Output_string_Sixpack_Normal.Add("");
            ;

        }

        static void Anova_Normal_Function(string Rsub, string Lspec, string Uspec, List<string> MiniTabValues_Output_string_Anova)
        {
            //Wersja 2
            MiniTabValues_Output_string_Anova.Add("NOTE*** " + Rsub + " ***");
            MiniTabValues_Output_string_Anova.Add("GageRR;");
            MiniTabValues_Output_string_Anova.Add("Parts 'Serialnumber';");
            MiniTabValues_Output_string_Anova.Add("Opers 'Operator';");
            MiniTabValues_Output_string_Anova.Add("Response " + "'" + Rsub + "'" + ";");
            MiniTabValues_Output_string_Anova.Add("Studyvar 6;");
            MiniTabValues_Output_string_Anova.Add("LSL " + Lspec.Replace('.', ',') + ";");
            MiniTabValues_Output_string_Anova.Add("USL " + Uspec.Replace('.', ',') + ";");
            MiniTabValues_Output_string_Anova.Add("Risk.");
            MiniTabValues_Output_string_Anova.Add("");
        }



        #endregion

    }
}

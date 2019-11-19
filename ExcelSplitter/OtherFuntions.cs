using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSplitter
{
    class OtherFuntions
    {
        //private void read_excel_worksheet()
        //{
        //    string fname = "";
        //    OpenFileDialog fdlg = new OpenFileDialog();
        //    fdlg.Title = "Excel File Dialog";
        //    fdlg.InitialDirectory = @"c:\";
        //    fdlg.Filter = "All files (*.*)|*.*|All files (*.*)|*.*";
        //    fdlg.FilterIndex = 2;
        //    fdlg.RestoreDirectory = true;

        //    if (fdlg.ShowDialog() == DialogResult.OK)
        //        fname = fdlg.FileName;

        //    if (fname.Length > 0)
        //    {
        //        tab_control.TabPages.Clear();

        //        //Instance reference for Excel Application
        //        Excel.Application objXL = null;

        //        //Workbook refrence
        //        Excel.Workbook objWB = null;

        //        DataSet ds = new DataSet();

        //        try
        //        {
        //            //Instancing Excel using COM services
        //            objXL = new Excel.Application();
        //            //Adding WorkBook
        //            objWB = objXL.Workbooks.Open(fname);

        //            lbl_filename.Text = "Filename: " + fname.Substring(fname.LastIndexOf("\\") + 1);

        //            foreach (Microsoft.Office.Interop.Excel.Worksheet objSHT in objWB.Worksheets)
        //            {
        //                int rows = objSHT.UsedRange.Rows.Count; // total rows
        //                int cols = objSHT.UsedRange.Columns.Count; // total columns

        //                DataTable dt = new DataTable();
        //                int noofrow = 1;

        //                //If 1st Row Contains unique Headers for datatable include this part else remove it
        //                //Start
        //                for (int c = 1; c <= cols; c++)
        //                {
        //                    string colname = objSHT.Cells[1, c].Text;
        //                    dt.Columns.Add(colname);
        //                    noofrow = 2;
        //                }

        //                //End                                       
        //                for (int r = noofrow; r <= rows; r++)
        //                {
        //                    lbl_status.Text = "Row " + r + "/" + rows;
        //                    DataRow dr = dt.NewRow();
        //                    for (int c = 1; c <= cols; c++)
        //                    {
        //                        dr[c - 1] = objSHT.Cells[r, c].Text;
        //                    }
        //                    dt.Rows.Add(dr);
        //                }
        //                ds.Tables.Add(dt);
        //            }

        //            for (int q = 0; q < ds.Tables.Count; q++)
        //            {
        //                lbl_error.Text += "Table " + (q + 1) + " = " + ds.Tables[q].Rows.Count + ", ";

        //                DataGridView dtv = new DataGridView();
        //                dtv.Name = "" + q;
        //                dtv.AutoGenerateColumns = true;
        //                dtv.DataSource = ds.Tables[q].DefaultView;
        //                dtv.Height = 131;
        //                dtv.Width = 755;

        //                TabPage tb = new TabPage();
        //                tb.Name = "Book " + (q + 1);
        //                tb.Text = "Book " + (q + 1);
        //                tb.Controls.Add(dtv);
        //                tab_control.TabPages.Add(tb);

        //                tab_control.Refresh();
        //                //export_dataset_to_excel(ds, objXL);
        //            }

        //            //Closing workbook
        //            objWB.Close();

        //            //Closing excel application
        //            objXL.Quit();

        //            lbl_error.Text = "Success";
        //        }
        //        catch (Exception ex)
        //        {
        //            objWB.Saved = true;

        //            //Closing work book
        //            objWB.Close();

        //            //Closing excel application
        //            objXL.Quit();

        //            lbl_error.Text = "Failed";

        //            //Response.Write("Illegal permission");
        //            Console.WriteLine(ex.Message);
        //        }
        //    }
        //}

        //private void do_read_excel()
        //{
        //    string fname = "";
        //    OpenFileDialog fdlg = new OpenFileDialog();
        //    fdlg.Title = "Excel File Dialog";
        //    fdlg.InitialDirectory = @"c:\";
        //    fdlg.Filter = "All files (*.*)|*.*|All files (*.*)|*.*";
        //    fdlg.FilterIndex = 2;
        //    fdlg.RestoreDirectory = true;
        //    if (fdlg.ShowDialog() == DialogResult.OK)
        //    {
        //        fname = fdlg.FileName;
        //    }

        //    try
        //    {
        //        using (ExcelEngine excelEngine = new ExcelEngine())
        //        {
        //            //Initialize application
        //            IApplication application = excelEngine.Excel;

        //            //Open existing workbook with data entered
        //            Assembly assembly = typeof(Program).GetTypeInfo().Assembly;
        //            Stream fileStream = assembly.GetManifestResourceStream(fname);
        //            IWorkbook workbook = application.Workbooks.Open(fileStream);

        //            //Access first worksheet from the workbook instance
        //            IWorksheet worksheet = workbook.Worksheets[0];

        //            //Export Excel to DataTable
        //            DataTable dataTable = worksheet.ExportDataTable(worksheet.UsedRange, ExcelExportDataTableOptions.ColumnNames);

        //            DataGridView dtv = new DataGridView();
        //            dtv.Name = "" + 1;
        //            dtv.DataSource = dataTable;
        //            dtv.AutoGenerateColumns = true;
        //            dtv.Height = 231;
        //            dtv.Width = 755;

        //            TabPage tb = new TabPage();
        //            tb.Name = "Book 1";
        //            tb.Controls.Add(dtv);
        //            tab_control.TabPages.Add(tb);

        //            tab_control.Refresh();

        //            //Save the workbook to disk in xlsx format
        //            workbook.SaveAs("Output.xlsx");
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine(ex.Message);
        //    }
        //}

        //private void btn_read_excel()
        //{
        //    string filePath = string.Empty;
        //    string fileExt = string.Empty;
        //    OpenFileDialog file = new OpenFileDialog(); //open dialog to choose file  
        //    if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK) //if there is a file choosen by the user  
        //    {
        //        filePath = file.FileName; //get the path of the file  
        //        fileExt = Path.GetExtension(filePath); //get the file extension  
        //        if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
        //        {
        //            try
        //            {
        //                DataTable dtExcel = new DataTable();
        //                dtExcel = ReadExcel(filePath, fileExt); //read excel file  

        //                DataGridView dtv = new DataGridView();
        //                dtv.Name = "" + 1;
        //                dtv.DataSource = dtExcel;
        //                dtv.AutoGenerateColumns = true;
        //                dtv.Height = 231;
        //                dtv.Width = 755;

        //                TabPage tb = new TabPage();
        //                tb.Name = "Book 1";
        //                tb.Controls.Add(dtv);
        //                tab_control.TabPages.Add(tb);

        //                tab_control.Refresh();
        //            }
        //            catch (Exception ex)
        //            {
        //                MessageBox.Show(ex.Message.ToString());
        //            }
        //        }
        //        else
        //        {
        //            MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error  
        //        }
        //    }
        //}

        //private DataTable ReadExcel(string fileName, string fileExt)
        //{
        //    string conn = string.Empty;
        //    DataTable dtexcel = new DataTable();
        //    if (fileExt.CompareTo(".xls") == 0)
        //        conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
        //    else
        //        conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=NO';"; //for above excel 2007  
        //    using (OleDbConnection con = new OleDbConnection(conn))
        //    {
        //        try
        //        {
        //            OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [Sheet1$]", con); //here we read data from sheet1  
        //            oleAdpt.Fill(dtexcel); //fill excel data into dataTable  
        //        }
        //        catch (Exception ex)
        //        {
        //            Console.WriteLine(ex.Message);
        //        }
        //    }
        //    return dtexcel;
        //}

        //    private List<string> set_headers;

        //    private void read_excel_old()
        //    {
        //        try
        //        {
        //            set_headers = new List<string>();

        //            //datagrid.Rows.Clear();
        //            //datagrid.Columns.Clear();
        //            //datagrid.Refresh();

        //            string fname = "";
        //            OpenFileDialog fdlg = new OpenFileDialog();
        //            fdlg.Title = "Excel File Dialog";
        //            fdlg.InitialDirectory = @"c:\";
        //            fdlg.Filter = "All files (*.*)|*.*|All files (*.*)|*.*";
        //            fdlg.FilterIndex = 2;
        //            fdlg.RestoreDirectory = true;
        //            if (fdlg.ShowDialog() == DialogResult.OK)
        //            {
        //                fname = fdlg.FileName;
        //            }

        //            int n = 150;

        //            Excel.Application xlApp = new Excel.Application();
        //            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fname);
        //            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
        //            //Excel.Range xlRange = xlWorksheet.UsedRange;

        //            Excel.Range usedRange = xlWorksheet.UsedRange;
        //            object[,] values = usedRange.Value2;

        //            DateTime dtStart = DateTime.Now;
        //            int nColumnsMax = 0;

        //            if (usedRange.Rows.Count > 0)
        //            {
        //                //----< Read_Header >----

        //                dtStart = DateTime.Now;

        //                DataGridView dtv = new DataGridView();
        //                dtv.Name = "" + 1;
        //                dtv.AutoGenerateColumns = true;
        //                dtv.Height = 231;
        //                dtv.Width = 755;

        //                for (int iColumn = 1; iColumn <= usedRange.Columns.Count; iColumn++)
        //                {
        //                    //*slow Excel: 
        //                    //Excel.Range cell = usedRange.Cells[1, iColumn] as Excel.Range;
        //                    //String sValue = Convert.ToString(cell.Value2);
        //                    //*fast Excel: 

        //                    string sValue = Convert.ToString(values[1, iColumn]);
        //                    if (sValue == "" || sValue == null) break;
        //                    dtv.Columns.Add("column_" + iColumn, sValue);
        //                    set_headers.Add(sValue);
        //                    nColumnsMax = iColumn;
        //                }

        //                string string_headers = "";
        //                for (int q = 0; q < set_headers.Count; q++)
        //                    string_headers += set_headers[q][0].ToString().ToUpper() + set_headers[q].Substring(1) + " ";

        //                //label1.Text = string_headers;
        //                double no_of_files = Math.Ceiling(usedRange.Rows.Count / (double)n);
        //                int counter = 1, file_no = 0, get_last_counter_mod = (usedRange.Rows.Count - 1) % n, real_total = 1;
        //                lbl_filename.Text = "No. files: " + file_no + "/" + no_of_files + ", Total row added: " + real_total + "/" + usedRange.Rows.Count;

        //                dtStart = DateTime.Now;


        //                for (int iRow = 2; iRow <= usedRange.Rows.Count; iRow++)
        //                {
        //                    lbl_status.Text = "row " + iRow + "/" + usedRange.Rows.Count;

        //                    //< add_Row >
        //                    int iNewRow = dtv.Rows.Add(new DataGridViewRow());
        //                    DataGridViewRow newRow = dtv.Rows[iNewRow];

        //                    //</ add_Row >

        //                    //if (iRow > 20) break;

        //                    for (int iColumn = 1; iColumn <= nColumnsMax; iColumn++)
        //                    {
        //                        //< read column >
        //                        //*fast Excel: 

        //                        string sValue = Convert.ToString(values[iRow, iColumn]);

        //                        if (iColumn == 1)
        //                        {
        //                            if (sValue == "" || sValue == null)
        //                            {
        //                                //log_with_Date("rows= ", dtStart);
        //                                return; // break;
        //                            }
        //                        }

        //                        //</ read column >

        //                        //< write >

        //                        newRow.Cells[iColumn - 1].Value = sValue;

        //                        //</ write >

        //                        if (counter == n || file_no == no_of_files - 1 && counter == get_last_counter_mod - 1)
        //                        {
        //                            real_total += counter;
        //                            file_no++;
        //                            counter = 0;
        //                            lbl_filename.Text = "No. files: " + file_no + "/" + no_of_files + ", Total row added: " + real_total + "/" + usedRange.Rows.Count;
        //                        }
        //                    }
        //                    counter++;
        //                }


        //                TabPage tb = new TabPage();
        //                tb.Name = "Book 1";
        //                tb.Controls.Add(dtv);
        //                tab_control.TabPages.Add(tb);

        //                tab_control.Refresh();
        //            }

        //        }
        //        catch (Exception ex)
        //        { Console.WriteLine(ex.Message); }
        //    }

        //    private void test()
        //    {
        //        int size = -1;
        //        DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.
        //        if (result == DialogResult.OK) // Test result.
        //        {
        //            string file = openFileDialog1.FileName;
        //            try
        //            {
        //                string text = File.ReadAllText(file);
        //                size = text.Length;

        //                lbl_filename.Text = "File: " + file.Substring(file.LastIndexOf('\\') + 1) + " (Size: " + size + ")";

        //                Console.WriteLine(size); // <-- Shows file size in debugging mode.
        //                Console.WriteLine(file);
        //            }
        //            catch (IOException)
        //            {
        //            }
        //        }
        //        Console.WriteLine(result); // <-- For debugging use.
        //    }

        //    private void do_this()
        //    {
        //        try
        //        {
        //            string fname = "";
        //            OpenFileDialog fdlg = new OpenFileDialog();
        //            fdlg.Title = "Excel File Dialog";
        //            fdlg.InitialDirectory = @"c:\";
        //            fdlg.Filter = "All files (*.*)|*.*|All files (*.*)|*.*";
        //            fdlg.FilterIndex = 2;
        //            fdlg.RestoreDirectory = true;
        //            if (fdlg.ShowDialog() == DialogResult.OK)
        //            {
        //                fname = fdlg.FileName;
        //            }

        //            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
        //            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fname);
        //            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
        //            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

        //            int rowCount = xlRange.Rows.Count;
        //            int colCount = xlRange.Columns.Count;

        //            DataGridView dtv = new DataGridView();
        //            dtv.Name = "" + 1;
        //            dtv.AutoGenerateColumns = true;
        //            dtv.Height = 231;
        //            dtv.Width = 755;

        //            // dt.Column = colCount;  
        //            dtv.ColumnCount = colCount;
        //            dtv.RowCount = rowCount;

        //            string test = "";

        //            for (int i = 1; i <= rowCount; i++)
        //            {
        //                for (int j = 1; j <= colCount; j++)
        //                {
        //                    //write the value to the Grid  

        //                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
        //                    {
        //                        //datagrid.Rows[i - 1].Cells[j - 1].Value = xlRange.Cells[i, j].Value2.ToString();
        //                        //Console.WriteLine(xlRange.Cells[i, j].Value2.ToString() + "\t");
        //                        test += xlRange.Cells[i, j].Value2.ToString() + " || ";
        //                    }
        //                    //add useful things here!     
        //                }
        //            }

        //            TabPage tb = new TabPage();
        //            tb.Name = "Book 1";
        //            tb.Controls.Add(dtv);
        //            tab_control.TabPages.Add(tb);

        //            tab_control.Refresh();

        //            if (test.Length > 0)
        //                Console.Write(test);
        //            else
        //                Console.WriteLine("Empty");

        //            //cleanup  
        //            GC.Collect();
        //            GC.WaitForPendingFinalizers();

        //            //rule of thumb for releasing com objects:  
        //            //  never use two dots, all COM objects must be referenced and released individually  
        //            //  ex: [somthing].[something].[something] is bad  

        //            //release com objects to fully kill excel process from running in the background  
        //            Marshal.ReleaseComObject(xlRange);
        //            Marshal.ReleaseComObject(xlWorksheet);

        //            //close and release  
        //            xlWorkbook.Close();
        //            Marshal.ReleaseComObject(xlWorkbook);

        //            //quit and release  
        //            xlApp.Quit();
        //            Marshal.ReleaseComObject(xlApp);
        //        }
        //        catch (Exception ex)
        //        { Console.Write(ex.Message); }
        //    }
        //}
    }
}

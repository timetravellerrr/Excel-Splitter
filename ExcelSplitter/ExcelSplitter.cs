using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using ClosedXML.Excel;

namespace ExcelSplitter
{
    public partial class ExcelSplitter : Form
    {
        ExcelReader er;

        public ExcelSplitter()
        {
            InitializeComponent();
            initialize_objects();
            load_default_values();
            set_tooltip();
        }

        private void initialize_objects()
        {
            er = new ExcelReader();
            er.save_path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
        }

        private void load_default_values()
        {
            lbl_filename.Text = string.Empty;
            lbl_status.Text = "Row(s): 0/0";
            lbl_log.Text = "Total no. of tables processed: 0/0";
            lbl_n.Text = string.Empty;
            lbl_error.Text = string.Empty;
            lbl_savepath.Text = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
            txt_n.Text = 150.ToString();
            rb_xlsx.Checked = true;
            cb_header.Checked = true;
            rb_h_yes.Checked = true;
            pnl_headers.Visible = false;

            int filewidth = lbl_savepath.Width;
            btn_openfolder.Location = new System.Drawing.Point((204 + lbl_savepath.Text.Length + filewidth) - 65, 58);
            btn_openfolder.Click += new System.EventHandler(this.btn_openfolder_Click);
        }

        private void set_tooltip()
        {
            toolTip1.SetToolTip(btn_openfolder, "Open saved directory");
            toolTip1.SetToolTip(btn_file, "Select excel file");
            toolTip1.SetToolTip(btn_savepath, "Select saved directory");
            toolTip1.SetToolTip(btn_split, "Run");
            toolTip1.SetToolTip(txt_n, "Enter number of rows to split");
        }

        private void btn_file_Click(object sender, EventArgs e)
        {
            string fname = "";
            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Title = "Excel File Dialog";
            fdlg.InitialDirectory = @"c:\";
            fdlg.Filter = "All files (*.*)|*.*|All files (*.*)|*.*";
            fdlg.FilterIndex = 2;
            fdlg.RestoreDirectory = true;

            if (er.directory_name != null && er.directory_name.Length > 0)
                fdlg.InitialDirectory = er.directory_name;

            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                fname = fdlg.FileName;

                string file_ext = Path.GetExtension(fname); //get the file extension  
                if (file_ext.CompareTo(".xls") == 0 || file_ext.CompareTo(".xlsx") == 0)
                {
                    er.filename = fname.Substring(fname.LastIndexOf("\\") + 1);
                    er.directory_name = Path.GetDirectoryName(fname);
                    err_messages(lbl_filename, Color.Black, er.filename);
                }
                else
                {
                    err_messages(lbl_filename, Color.Red, StaticMessages.ERR_FILENAME_INVALID_FORMAT);
                }
            }
        }

        private void btn_savepath_Click(object sender, EventArgs e)
        {
            string fname = "";
            FolderBrowserDialog fdlg = new FolderBrowserDialog();
            fdlg.RootFolder = Environment.SpecialFolder.Desktop;
            fdlg.Description = "Save File Dialog";

            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                fname = fdlg.SelectedPath;

                string file_ext = Path.GetExtension(fname); //get the file extension  
                if (file_ext.Length > 0)
                {
                    err_messages(lbl_savepath, Color.Red, StaticMessages.ERR_SAVEPATH_INVALID);
                }
                else
                {
                    er.save_path = fname;
                    err_messages(lbl_savepath, Color.Black, fname);
                    int filewidth = lbl_savepath.Width;
                    btn_openfolder.Location = new System.Drawing.Point((204 + fname.Length + filewidth) - 65, 58);
                }
            }
        }

        private void btn_split_Click(object sender, EventArgs e)
        {
            if (validate_fields())
                read_excel(er);
        }

        private void btn_openfolder_Click(object sender, EventArgs e)
        {
            if (er.save_path != null && er.save_path.Length > 0)
            {
                // check if directory
                FileAttributes attr = File.GetAttributes(er.save_path);

                if (attr.HasFlag(FileAttributes.Directory))
                {
                    System.Diagnostics.Process.Start(er.save_path);
                }
            }
        }

        public void err_messages(Label label, Color color, string msg)
        {
            label.Text = msg;
            label.ForeColor = color;
        }

        private bool validate_fields()
        {
            bool ret_val = false;

            // check if n number is filled
            if (txt_n.Text.Length > 0)
            {
                int n = 0;
                bool try_n = Int32.TryParse(txt_n.Text, out n);
                if (try_n)
                {
                    er.n = n;
                    ret_val = true;
                    lbl_n.Text = string.Empty;
                }
                else
                {
                    err_messages(lbl_n, Color.Red, StaticMessages.ERR_N_INVALID);
                    return ret_val = false;
                }
            }
            else
            {
                txt_n.Text = 150.ToString();
                er.n = 150;
                ret_val = true;
            }

            // check if file is chosen
            if (er.filename != null && er.filename.Length > 0)
            {
                // check if file ext is xls
                string file_ext = Path.GetExtension(er.filename); //get the file extension
                if (file_ext.CompareTo(".xls") == 0 || file_ext.CompareTo(".xlsx") == 0)
                {
                    ret_val = true;
                }
                else
                {
                    err_messages(lbl_filename, Color.Red, StaticMessages.ERR_FILENAME_INVALID_FORMAT);
                    return ret_val = false;
                }
            }
            else
            {
                err_messages(lbl_filename, Color.Red, StaticMessages.ERR_FILENAME_EMPTY);
                return ret_val = false;
            }

            // check if save destination path is set
            if (er.save_path != null && er.save_path.Length > 0)
            {
                // check if directory
                FileAttributes attr = File.GetAttributes(er.save_path);

                if (attr.HasFlag(FileAttributes.Directory))
                    ret_val = true;
                else
                {
                    err_messages(lbl_savepath, Color.Red, StaticMessages.ERR_SAVEPATH_INVALID);
                    return ret_val = false;
                }
            }
            else
            {
                err_messages(lbl_savepath, Color.Red, StaticMessages.ERR_SAVEPATH_EMPTY);
                return ret_val = false;
            }

            return ret_val;
        }

        private void read_excel(ExcelReader er)
        {
            string fname = er.directory_name + @"\" + er.filename;
            string save_path = er.save_path;
            int n = er.n;

            if (fname.Length > 0)
            {
                tab_control.TabPages.Clear();

                //Instance reference for Excel Application
                Excel.Application objXL = null;

                //Workbook refrence
                Excel.Workbook objWB = null;

                DataSet ds = new DataSet();

                try
                {
                    //Instancing Excel using COM services
                    objXL = new Excel.Application();
                    //Adding WorkBook
                    objWB = objXL.Workbooks.Open(fname);

                    Excel.Worksheet objSHT = objWB.Worksheets[1];

                    int rows = objSHT.UsedRange.Rows.Count; // total rows
                    int cols = objSHT.UsedRange.Columns.Count; // total columns

                    double no_of_files = Math.Ceiling(rows / (double)n);
                    int get_last_counter_mod = (rows - 1) % n;

                    lbl_log.Text = "Total no. of tables processed: 0/" + no_of_files;
                    err_messages(lbl_error, ColorTranslator.FromHtml("#007bff"), StaticMessages.ERR_PROC);

                    DataTable dt = new DataTable();
                    int no_of_row = 1, counter = 0;
                    bool header = false, skip = false;

                    if (cb_header.Checked)
                        if (rb_h_yes.Checked)
                            header = true;
                        else if (rb_h_no.Checked)
                            skip = true;

                    for (int q = 0; q < no_of_files; q++)
                    {
                        dt = new DataTable();
                        dt.TableName = "Book " + (q + 1);

                        if (header || skip)
                        {
                            if(!skip)
                            for (int c = 1; c <= cols; c++)
                            {
                                string colname = objSHT.Cells[1, c].Text;
                                dt.Columns.Add(colname);
                            }
                            else if(skip)
                                for (int c = 1; c <= cols; c++)
                                {
                                    string colname = objSHT.Cells[1, c].Text;
                                    dt.Columns.Add("");
                                }
                            if (q == 0)
                                no_of_row++;
                        }

                        for (int r = no_of_row; r <= rows; r++)
                        {
                            lbl_status.Text = "Row(s): " + r + "/" + rows;
                            DataRow dr = dt.NewRow();
                            for (int c = 1; c <= cols; c++)
                            {
                                dr[c - 1] = objSHT.Cells[r, c].Text;
                            }
                            dt.Rows.Add(dr);
                            counter++;
                            if (counter % n == 0)
                                break;
                        }
                        ds.Tables.Add(dt);

                        if (q == 0)
                            no_of_row = n + 2;
                        else if (q < no_of_files - 1)
                            no_of_row += n;
                        else
                            no_of_row += get_last_counter_mod;
                        counter = 0;
                        lbl_log.Text = "Total no. of tables processed: " + ds.Tables.Count + "/" + no_of_files;
                    }

                    //Export to File
                    if(rb_xlsx.Checked)
                        export_to_closedxml(ds);
                    else if(rb_csv.Checked)
                        export_to_csv(ds);
                    //export_dataset_to_excel(ds, objXL);

                    //Closing workbook
                    objWB.Close();
                    Marshal.ReleaseComObject(objWB);

                    //Closing excel application
                    objXL.Quit();
                    Marshal.ReleaseComObject(objXL);

                    err_messages(lbl_error, Color.Green, StaticMessages.ERR_SUCC);
                }
                catch (Exception ex)
                {
                    objWB.Saved = true;

                    //Closing work book
                    objWB.Close();
                    Marshal.ReleaseComObject(objWB);

                    //Closing excel application
                    objXL.Quit();
                    Marshal.ReleaseComObject(objXL);

                    err_messages(lbl_error, Color.Red, ex.Message);

                    //Response.Write("Illegal permission");
                    Console.WriteLine(ex.Message);
                }
            }
        }

        private void update_grid_view(int q, DataTable dt)
        {
            DataGridView dtv = new DataGridView();
            dtv.Name = "" + q;
            dtv.AutoGenerateColumns = true;
            dtv.DataSource = dt.DefaultView;
            dtv.Height = 131;
            dtv.Width = 755;

            TabPage tb = new TabPage();
            tb.Name = dt.TableName;
            tb.Text = dt.TableName + " (" + dt.Rows.Count + ")";
            tb.Controls.Add(dtv);
            tab_control.TabPages.Add(tb);

            tab_control.Refresh();
        }

        private void export_to_closedxml(DataSet ds)
        {
            if (ds.Tables.Count > 0)
            { 
                for (int q = 0; q < ds.Tables.Count; q++)
                {
                    var wb = new XLWorkbook();
                    DataTable dt = ds.Tables[q];

                    // Add all DataTables in the DataSet as a worksheets
                    //var ws = wb.Worksheets.Add(dt);
                    var ws = wb.Worksheets.Add(dt.TableName);
                    ws.FirstCell().InsertTable(dt).Theme = XLTableTheme.None;
                    ws.Columns().Width = 11;
                    //ws.Row(1).Style = XLWorkbook.DefaultStyle;
                    ws.Row(1).Style.Font.Bold = true;
                    ws.Column(1).Width = 50;

                    string new_filename = er.filename.Substring(0, er.filename.LastIndexOf(".")) + " " + (q + 1);

                    wb.SaveAs(Path.Combine(er.save_path, new_filename + ".xlsx"));

                    update_grid_view(q, dt);

                    lbl_log.Text = "Total no. of tables processed: " + ds.Tables.Count + ", No. of files: " + (q + 1) + "/" + ds.Tables.Count + " file(s) created.";
                }
            }
        }

        private void export_to_csv(DataSet ds)
        {
            if (ds.Tables.Count > 0)
            {
                for (int q = 0; q < ds.Tables.Count; q++)
                {
                    StringBuilder content = new StringBuilder();

                    DataTable dt = ds.Tables[q];

                    DataRow dr1 = (DataRow)dt.Rows[0];
                    int intColumnCount = dr1.Table.Columns.Count;
                    int index = 1;

                    foreach (DataColumn item in dr1.Table.Columns)
                    {
                        content.Append(String.Format("\"{0}\"", item.ColumnName));
                        if (index < intColumnCount)
                            content.Append(",");
                        else
                            content.Append("\r\n");
                        index++;
                    }

                    foreach (DataRow currentRow in dt.Rows)
                    {
                        string strRow = string.Empty;
                        for (int y = 0; y <= intColumnCount - 1; y++)
                        {
                            strRow += "\"" + currentRow[y].ToString() + "\"";

                            if (y < intColumnCount - 1 && y >= 0)
                                strRow += ",";
                        }
                        content.Append(strRow + "\r\n");
                    }

                    string new_filename = er.filename.Substring(0, er.filename.LastIndexOf(".")) + " " + (q + 1);

                    File.WriteAllText(Path.Combine(er.save_path, new_filename + ".csv"), content.ToString());

                    update_grid_view(q, dt);

                    lbl_log.Text = "Total no. of tables processed: " + ds.Tables.Count + ", No. of files: " + (q + 1) + "/" + ds.Tables.Count + " file(s) created.";
                }
            }
        }

        private bool export_dataset_to_excel(DataSet ds, Excel.Application objXL)
        {
            //Create an Excel application instance
            //Excel.Application excelApp = new Excel.Application();

            //Create an Excel workbook instance and open it from the predefined location
            //Excel.Workbook excelWorkBook = excelApp.Workbooks.Open(@"C:\Users\hanif\Desktop\Org.xlsx");
            object mis_value = System.Reflection.Missing.Value;
            Excel.Workbook excelWorkBook = objXL.Workbooks.Add(mis_value);
            Excel.Worksheet excelWorkSheet = new Excel.Worksheet();
            int q = 1;

            try
            {
                foreach (DataTable table in ds.Tables)
                {
                    //Add a new worksheet to workbook with the Datatable name
                    excelWorkSheet = excelWorkBook.Sheets.Add();
                    excelWorkSheet.Name = table.TableName;

                    for (int i = 1; i < table.Columns.Count + 1; i++)
                    {
                        excelWorkSheet.Cells[1, i] = table.Columns[i - 1].ColumnName;
                    }

                    for (int j = 0; j < table.Rows.Count; j++)
                    {
                        for (int k = 0; k < table.Columns.Count; k++)
                        {
                            excelWorkSheet.Cells[j + 2, k + 1] = table.Rows[j].ItemArray[k].ToString();
                        }
                    }
                }
                string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Excel " + q + ".xlsx");
                excelWorkBook.SaveAs(path, Excel.XlFileFormat.xlOpenXMLWorkbook, mis_value, mis_value, mis_value, mis_value, Excel.XlSaveAsAccessMode.xlExclusive, mis_value, mis_value, mis_value, mis_value, mis_value);
                //excelWorkBook.Save();
                excelWorkBook.Close();
                //excelApp.Quit();
                return true;
            }
            catch (Exception ex)
            {
                excelWorkBook.Close();
                Console.WriteLine(ex.Message);
                return false;
            }
        }
    }        
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace JurisUtilityBase
{
    public partial class ReportDisplay : Form
    {
        public ReportDisplay(DataSet ds)
        {
            InitializeComponent();
            dataGridView1.DataSource = ds.Tables[0];
        }

        private void buttonBack_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void buttonPrint_Click(object sender, EventArgs e)
        {
            PrinterDialog pd = new PrinterDialog();
            pd.ShowDialog();
            string printer = pd.printerName;
            if (!string.IsNullOrEmpty(printer))
            {
                Cursor.Current = Cursors.WaitCursor;

                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

                Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlApp = new Microsoft.Office.Interop.Excel.Application();
                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                int StartCol = 1;
                int StartRow = 1;
                int j = 0, i = 0;

                //Write Headers
                for (j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[StartRow, StartCol + j];
                    myRange.Value2 = dataGridView1.Columns[j].HeaderText;
                    if (dataGridView1.Columns[j].HeaderText.ToString().ToLower().Contains("date"))
                    {
                        for (int a = 1; a < dataGridView1.Rows.Count + 1; a++)
                        {
                            myRange = xlWorkSheet.Cells[StartRow + a, StartCol + j];
                            myRange.EntireColumn.NumberFormat = "MM/DD/YYYY";
                        }

                    }
                }

                StartRow++;

                //Write datagridview content
                for (i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        try
                        {
                            Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[StartRow + i, StartCol + j];
                            myRange.Value2 = dataGridView1[j, i].Value == null ? "" : dataGridView1[j, i].Value;
                        }
                        catch
                        {
                            ;
                        }
                    }
                }

                Microsoft.Office.Interop.Excel.Range usedrange = xlWorkSheet.UsedRange;
                usedrange.Columns.AutoFit();
                xlApp.Visible = false;
                var _with1 = xlWorkSheet.PageSetup;
                _with1.Zoom = false;
                _with1.PrintGridlines = true;
                _with1.PaperSize = Microsoft.Office.Interop.Excel.XlPaperSize.xlPaperA4;
                _with1.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape;
                _with1.FitToPagesWide = 1;
                _with1.FitToPagesTall = false;

                _with1.PrintTitleRows = "$1:$" + dataGridView1.Columns.Count.ToString();

                string Defprinter = null;
                Defprinter = xlApp.ActivePrinter;
                xlApp.ActivePrinter = printer;

                // Print the range
                usedrange.PrintOutEx(misValue, misValue, misValue, misValue,
                misValue, misValue, misValue, misValue);
                // }
                xlApp.ActivePrinter = Defprinter;

                // Cleanup:
                GC.Collect();
                GC.WaitForPendingFinalizers();

                Marshal.FinalReleaseComObject(xlWorkSheet);

                xlWorkBook.Close(false, Type.Missing, Type.Missing);
                Marshal.FinalReleaseComObject(xlWorkBook);

                xlApp.Quit();
                Marshal.FinalReleaseComObject(xlApp);
                Cursor.Current = Cursors.Default;
            }
        }




    }
}

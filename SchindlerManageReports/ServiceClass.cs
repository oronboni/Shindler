using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SchindlerManageReports
{
    class ServiceClass
    {
        public void BuildFolders()
        {
            if (!Directory.Exists(Environment.CurrentDirectory + "\\success"))
                Directory.CreateDirectory(Environment.CurrentDirectory + "\\success");
            if (!Directory.Exists(Environment.CurrentDirectory + "\\fail"))
                Directory.CreateDirectory(Environment.CurrentDirectory + "\\fail");
        }

        public void createExcelFile(Dictionary<string, Information> all_files, int allcount)
        {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }


            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Cells[1, 1] = "SAIS";
            xlWorkSheet.Cells[1, 2] = "PQC";
            xlWorkSheet.Cells[1, 3] = "COUNT";
            xlWorkSheet.Cells[1, 4] = "PRECENTAGE";
            xlWorkSheet.Cells[1, 5] = "TOTAL-REPORTS";
            xlWorkSheet.Cells.NumberFormat = "@";
            int row = 2;


            foreach (KeyValuePair<string, Information> keyval in all_files)
            {
                if (!keyval.Key.Equals(""))
                {
                    xlWorkSheet.Cells[row, 3] = keyval.Value.Count;
                    xlWorkSheet.Cells[row, 4] = (((float)keyval.Value.Count / (float)allcount) * 100).ToString() + "%";
                    xlWorkSheet.Cells[row, 5] = allcount;
                    if (keyval.Value.Checktype.Equals("sais"))
                    {
                        xlWorkSheet.Cells[row, 1] = keyval.Key.ToString();
                    }
                    else
                    {
                        xlWorkSheet.Cells[row, 2] = keyval.Key.ToString();
                    }
                    row++;
                }
            }

            if (File.Exists(Environment.CurrentDirectory + "\\result.xls"))
                File.Delete(Environment.CurrentDirectory + "\\result.xls");

            xlWorkBook.SaveAs(Environment.CurrentDirectory + "\\result.xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }
    }
}

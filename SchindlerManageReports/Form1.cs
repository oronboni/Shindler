
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;


namespace SchindlerManageReports
{
    public partial class Form1 : MembraneUserControl
    {
        public Form1()
        {
            InitializeComponent();
        }

       

        private void button1_Click(object sender, EventArgs e)
        {
            ServiceClass sc = new ServiceClass();
            sc.BuildFolders();

            label1.Visible = true;
            int allcount = 0;

            Dictionary<string, Information> all_files = new Dictionary<string, Information>();

            allcount += ExcelHandler(all_files);
            allcount += WordHandler(all_files);

            sc.createExcelFile(all_files, allcount);

            label1.Visible = false;
            lblInfo.Visible = false;
            MessageBox.Show(" החישוב הסתיים בהצלחה");

            string strCmdText;
            strCmdText = "TASKKILL / IM WINWORD.EXE";
            System.Diagnostics.Process.Start("cmd.exe", "/k " + strCmdText + " & exit");

            strCmdText = "TASKKILL / IM EXCEL.EXE";
            System.Diagnostics.Process.Start("cmd.exe", "/k " + strCmdText + " & exit");

            Thread.Sleep(2000);

            foreach (string file in Directory.EnumerateFiles(Environment.CurrentDirectory, "*.docx"))
                File.Delete(file);

            foreach (string file in Directory.EnumerateFiles(Environment.CurrentDirectory, "*.xlsx"))
                File.Delete(file);

        }

        private int ExcelHandler(Dictionary<string, Information> all_files)
        {
            int allcount = 0;
            foreach (string file in Directory.EnumerateFiles(Environment.CurrentDirectory, "*.xlsx"))
            {
                try
                {
                    FileInfo fi = new FileInfo(file);
                    label1.Text = fi.Name;
                    lblInfo.Visible = true;

                    allcount++;
                    Microsoft.Office.Interop.Excel.Application xlApp;
                    Workbook xlWorkBook;
                    Worksheet xlWorkSheet;
                    Range range;

                    string str;
                    int rCnt;
                    int cCnt;

                    xlApp = new Microsoft.Office.Interop.Excel.Application();
                    xlWorkBook = xlApp.Workbooks.Open(file, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);

                    range = xlWorkSheet.UsedRange;
                    int rw = range.Rows.Count;
                    int cl = range.Columns.Count;


                    for (rCnt = 1; rCnt <= rw; rCnt++)
                    {
                        for (cCnt = 1; cCnt <= cl; cCnt++)
                        {
                            if (range.Cells[rCnt, cCnt] as Range != null)
                            {
                                if ((range.Cells[rCnt, cCnt] as Range).Value2 != null)
                                {
                                    int result;
                                    str = getCellValue(xlWorkSheet, rCnt, cCnt);
                                    if (int.TryParse(str, out result) && cCnt == 1)
                                    {
                                        string sais = getCellValue(xlWorkSheet, rCnt, 5);
                                        string pqc = getCellValue(xlWorkSheet, rCnt, 6);

                                        if (!sais.Equals("") && !sais.Equals("¾"))
                                        {
                                            TempInformation keyval = new TempInformation(sais, "sais");
                                            if (!all_files.ContainsKey(keyval.Checknum))
                                            {
                                                all_files.Add(keyval.Checknum, new Information(keyval.Checktype));
                                            }
                                            else
                                            {
                                                all_files[keyval.Checknum].Count++;
                                            }
                                        }
                                        if (!pqc.Equals("") && !pqc.Equals("¾"))
                                        {
                                            TempInformation keyval = new TempInformation(pqc, "pqc");
                                            if (!all_files.ContainsKey(keyval.Checknum))
                                            {
                                                all_files.Add(keyval.Checknum, new Information(keyval.Checktype));
                                            }
                                            else
                                            {
                                                all_files[keyval.Checknum].Count++;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    xlWorkBook.Close(true, null, null);
                    xlApp.Quit();

                    Marshal.ReleaseComObject(xlWorkSheet);
                    Marshal.ReleaseComObject(xlWorkBook);
                    Marshal.ReleaseComObject(xlApp);
                    copyexcel(file, Environment.CurrentDirectory + "\\success");
                }
                catch (Exception ex)
                {
                    copyexcel(file, Environment.CurrentDirectory + "\\fail");
                }
            }
      

            return allcount;
        }

        private string getCellValue(Worksheet Sheet, int Row, int Column)
        {
            string cellValue = Sheet.Cells[Row, Column].Text.ToString();
            return cellValue;
        }

        private int WordHandler(Dictionary<string, Information> all_files)
        {
            int allcount = 0;
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            foreach (string file in Directory.EnumerateFiles(Environment.CurrentDirectory, "*.docx"))
                allcount = WordMethod(allcount, all_files, word, file);
            word.Quit();
            return allcount;
        }

        #region Word

        private int WordMethod(int allcount, Dictionary<string, Information> all_files, Microsoft.Office.Interop.Word.Application word, string file)
        {

            FileInfo fi = new FileInfo(file);
            label1.Text = fi.Name;
            lblInfo.Visible = true;

            string contents = File.ReadAllText(file);
            object miss = System.Reflection.Missing.Value;
            object path = file;
            object readOnly = true;
            Microsoft.Office.Interop.Word.Document docs = word.Documents.Open(ref path, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);
            string totaltext = "";
            for (int i = 0; i < docs.Paragraphs.Count; i++)
            {
                totaltext += docs.Paragraphs[i + 1].Range.Text.ToString().Replace("\r\a", "|").Replace("\a", "\r\n");
            }


            string[] split = totaltext.Split(new string[] { "|||(" }, StringSplitOptions.None);

            if (!split[1].Contains("|(||SAIS|PQC||||||||"))
            {
                docs.Close();
                copyword(file, Environment.CurrentDirectory + "\\fail");
            }
            else
            {
                allcount++;
                split[split.Length - 1] = split[split.Length - 1].Split(new string[] { "||" }, StringSplitOptions.None)[0];

                for (int i = 2; i < split.Length; i++)
                {
                    split[i] = split[i].Replace("\r\n", "");
                    string[] att = split[i].Split(new string[] { "|" }, StringSplitOptions.None);

                    if (att.Length == 9 || att.Length == 7)
                    {
                        string[] btt = new string[3];
                        btt[0] = att[0];
                        btt[1] = att[1];
                        btt[2] = att[2];
                        if (att.Length == 7)
                        {
                            if (btt[1].EndsWith("((") && btt[2].Equals(""))
                            {
                                btt[1] = btt[1].Replace("(", "");
                                btt[2] = "(";
                            }
                        }
                        TempInformation keyval;
                        if ((keyval = checkstatus(btt)) != null)
                        {
                            if (!all_files.ContainsKey(keyval.Checknum))
                            {
                                all_files.Add(keyval.Checknum, new Information(keyval.Checktype));
                            }
                            else
                            {
                                all_files[keyval.Checknum].Count++;
                            }

                        }
                        if (att.Length == 9)
                        {
                            btt = new string[3];
                            btt[0] = att[6];
                            btt[1] = att[7];
                            btt[2] = att[8];
                            if ((keyval = checkstatus(btt)) != null)
                            {
                                if (!all_files.ContainsKey(keyval.Checknum))
                                {
                                    all_files.Add(keyval.Checknum, new Information(keyval.Checktype));
                                }
                                else
                                {
                                    all_files[keyval.Checknum].Count++;
                                }

                            }
                        }
                        if (att.Length == 7)
                        {
                            btt = new string[3];
                            btt[0] = att[4];
                            btt[1] = att[5];
                            btt[2] = att[6];
                            if (btt[2].Equals(""))
                            {
                                btt[2] = "(";
                            }
                            if ((keyval = checkstatus(btt)) != null)
                            {
                                if (!all_files.ContainsKey(keyval.Checknum))
                                {
                                    all_files.Add(keyval.Checknum, new Information(keyval.Checktype));
                                }
                                else
                                {
                                    all_files[keyval.Checknum].Count++;
                                }

                            }
                        }
                    }
                    else
                    {
                        TempInformation keyval;
                        if ((keyval = checkstatus(att)) != null)
                        {
                            if (!all_files.ContainsKey(keyval.Checknum))
                            {
                                all_files.Add(keyval.Checknum, new Information(keyval.Checktype));
                            }
                            else
                            {
                                all_files[keyval.Checknum].Count++;
                            }
                        }
                    }
                }

                docs.Close();
                copyword(file, Environment.CurrentDirectory + "\\success");
            }

            return allcount;
        }

        #endregion

        private string validatekey(string key)
        {
            string stop=string.Empty;
            if (key.StartsWith("."))
            {
                for (int i = 2; i < key.Length; i++)
                    stop += key[i];

                stop += key[0];
                stop += key[1];


                key = stop;
            }
            return key;
        }

        private void copyword(string file_path, string file_location)
        {

            try
            {
                string strCmdText;
                strCmdText = "TASKKILL / IM WINWORD.EXE";
                System.Diagnostics.Process.Start("cmd.exe", "/k " + strCmdText + " & exit");
  

                FileInfo fi = new FileInfo(file_path);
                File.Copy(file_path, file_location + "\\" + fi.Name);
            }
            catch (Exception ex)
            {
             
            }
        }
        private void copyexcel(string file_path, string file_location)
        {

            try
            {
                string strCmdText;
                strCmdText = "TASKKILL / IM EXCEL.EXE";
                System.Diagnostics.Process.Start("cmd.exe", "/k " + strCmdText + " & exit");


                FileInfo fi = new FileInfo(file_path);
                File.Copy(file_path, file_location + "\\" + fi.Name);
            }
            catch (Exception ex)
            {

            }
        }



        private TempInformation checkstatus(string[] att)
        {
            TempInformation ti;
            if (att.Length==2 && att[1].Contains("(") && !att[1].Contains("\r") && !att[1].Contains("((") && !att[1].StartsWith("("))
            {
                att[1] = att[1].Replace("(", "");
                ti = new TempInformation(validatekey(att[1]), "sais");
                return ti;
            }
            if (att.Length == 2 && att[1].Contains("(") && !att[1].Contains("\r") && !att[1].Contains("((") && att[1].StartsWith("("))
            {
                att[1] = att[1].Replace("(", "");
                ti = new TempInformation(validatekey(att[1]), "pqc");
                return ti;
            }
            if (att.Length == 4 && att[3].Contains("(") && att[2].Contains("\r"))
            {
                att[2] = att[2].Replace("\r", "");
                ti = new TempInformation(validatekey(att[2]), "sais");
                return ti;
            }
            if (att.Length == 3 && att[2].Equals("(") && !att[0].Equals(""))
            {
                ti = new TempInformation(validatekey(att[1]), "sais");
                return ti;
            }
            if (att.Length == 3 && att[2].Equals("(") && att[0].Equals(""))
            {
                ti = new TempInformation(validatekey(att[1]), "pqc");
                return ti;
            }
            if (att.Length == 3 && !att[2].Equals("(") && att[2].EndsWith("("))
            {
                att[2] = att[2].Replace("(", "");
                ti = new TempInformation(validatekey(att[2]), "sais");
                return ti;
            }
            if (att.Length == 4 && att[3].Contains("("))
            {
                ti = new TempInformation(validatekey(att[2]), "sais");
                return ti;
            }

            if (att.Length == 4 && att[3].Equals("") && att[2].Equals(""))
            {
                ti = new TempInformation(validatekey(att[1]), "sais");
                return ti;
            }

            if (att.Length == 4 && att[3].Equals("") && att[1].Equals(""))
            {
                ti = new TempInformation(validatekey(att[2]), "pqc");
                return ti;
            }

            if (att.Length == 5 && att[4].Equals("") && att[3].Equals(""))
            {
                ti = new TempInformation(validatekey(att[2]), "sais");
                return ti;
            }

            return null;
        }

        private void עזרהToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void אודותToolStripMenuItem1_Click(object sender, EventArgs e)
        {

        }

        private void אודוצToolStripMenuItem_Click(object sender, EventArgs e)
        {
            About frm = new SchindlerManageReports.About();
            frm.ShowDialog();

        }
    } 
}

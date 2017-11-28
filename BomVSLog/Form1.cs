using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace BomVSLog
{
    public partial class Form1 : Form
    {
        List<string> BOMtests = new List<string>();
        List<string> TOtests = new List<string>();
        List<string> TPtests = new List<string>();
        List<string> LOGtests = new List<string>();
        Excel.Application xlApp;
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;

        public Form1()
        {
            InitializeComponent();
            
        }

        private void SelBomB_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "Select BOM File..";
            openFileDialog1.Filter = "TextFiles|*.txt";
            DialogResult res = openFileDialog1.ShowDialog();
            if (res == DialogResult.OK)
                Select_BomTB.Text = openFileDialog1.FileName;
        }

        private void SelLogB_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "Select LOG File..";
            openFileDialog1.Filter = "AllFiles|*.*";
            DialogResult res = openFileDialog1.ShowDialog();
            if (res == DialogResult.OK)
                Select_LogTB.Text = openFileDialog1.FileName;
        }

        private void ExitB_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        //private void CompareB_Click(object sender, EventArgs e)
        //{
        //    Regex reg = new Regex("version*\"*\"", RegexOptions.IgnoreCase);
        //    //if (File.Exists(Select_BomTB.Text) && File.Exists(Select_LogTB.Text))
        //    //{
        //    //    string[] bomfile = File.ReadAllLines(Select_BomTB.Text);
        //    //    //string[] testorder = File.ReadAllLines(Sel_TOTB.Text);


        //    //}

        //    if (File.Exists(Select_BomTB.Text) && File.Exists(Select_LogTB.Text))
        //    {
        //        string logfile = File.ReadAllText(Select_LogTB.Text).ToLower();
        //        string[] bomfile = File.ReadAllLines(Select_BomTB.Text);

        //        foreach (string bomcomp in bomfile)
        //        {
        //            if (bomcomp.Trim() != "")
        //            {
        //                string bomcomptemp = "{@block|1%" + bomcomp.Trim().ToLower();
        //                string logcomptest = "";
        //                int logind = 0;
        //                int findcnt = 0;
        //                while (logfile.IndexOf(bomcomptemp,logind+1) != -1)
        //                {
        //                    logind = logfile.IndexOf(bomcomptemp, logind + 1);
        //                    //if (bomcomptemp == "ic701")
        //                    //{
        //                    //    MessageBox.Show((logfile.IndexOf("{", logind + 1) + 1).ToString());
        //                    //    MessageBox.Show(logfile.IndexOf("|", logfile.IndexOf("{", logind + 1)).ToString());
        //                    //    MessageBox.Show((logfile.IndexOf("|", logfile.IndexOf("{", logind + 1)) - logfile.IndexOf("{", logind + 1) - 1).ToString());
        //                    //}
        //                    logcomptest = logcomptest + logfile.Substring(logfile.IndexOf("{", logind + 1) + 1, logfile.IndexOf("|", logfile.IndexOf("{", logind + 1)) - logfile.IndexOf("{", logind + 1) - 1);
        //                    findcnt++;
        //                }
        //                if (findcnt > 0 || (findcnt == 0 && logfile.IndexOf(bomcomp.Trim().ToLower(), logind + 1) != -1))
        //                {
        //                    FileStream myWriteStream = new FileStream(Path.GetFullPath(@"C:\Users\uidw8018\Desktop\LOG_C.txt"), FileMode.Append, FileAccess.Write);
        //                    byte[] newLine = Encoding.Default.GetBytes(Environment.NewLine);

        //                    myWriteStream.Write(Encoding.ASCII.GetBytes(bomcomp.Trim().ToLower() + "_" + logcomptest), 0, Encoding.ASCII.GetByteCount(bomcomp.Trim().ToLower() + "_" + logcomptest));
        //                    myWriteStream.Write(newLine, 0, newLine.Length);
        //                    myWriteStream.Close();
        //                }
        //                else
        //                {
        //                    FileStream myWriteStream = new FileStream(Path.GetFullPath(@"C:\Users\uidw8018\Desktop\LOG_C.txt"), FileMode.Append, FileAccess.Write);
        //                    byte[] newLine = Encoding.Default.GetBytes(Environment.NewLine);

        //                    myWriteStream.Write(Encoding.ASCII.GetBytes(bomcomp.Trim().ToLower() + "_" + "Not Placed"), 0, Encoding.ASCII.GetByteCount(bomcomp.Trim().ToLower() + "_" + "Not Placed"));
        //                    myWriteStream.Write(newLine, 0, newLine.Length);
        //                    myWriteStream.Close();
        //                }
        //            }
        //        }
        //    }
        //    else
        //    {
        //        if (!File.Exists(Select_BomTB.Text))
        //            MessageBox.Show("BOM file does not exists");
        //        if (!File.Exists(Select_BomTB.Text))
        //            MessageBox.Show("LOG file does not exists");
        //    }
        //}

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void SelTOB_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "Select Testorder File..";
            openFileDialog1.Filter = "Testorder|*testorder*";
            DialogResult res = openFileDialog1.ShowDialog();
            if (res == DialogResult.OK)
                Sel_TOTB.Text = openFileDialog1.FileName;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                TOtests.Clear();
                if (File.Exists(Sel_TOTB.Text))
                {
                    string[] TOfile = File.ReadAllLines(Sel_TOTB.Text);

                    foreach (string TOComp in TOfile)
                    {
                        //string tempcomp = "";
                        string testname = "";
                        // Extract Test Name
                        if (TOComp.Trim().ToLower().Contains("test ") || TOComp.Trim().ToLower().Contains("skip "))
                        {
                            //tempcomp = tempcomp + TOComp.Substring(TOComp.IndexOf("\"", 0) + 1, TOComp.IndexOf("\"", TOComp.IndexOf("\"", 0) + 1) - TOComp.IndexOf("\"", 0) - 1);
                            if (TOComp.Trim().ToLower().Substring(0,TOComp.IndexOf("\"") - 1).Contains("test "))
                            {
                                testname = testname + TOComp.Substring(0, TOComp.IndexOf("test ", 0) + 4).Trim() + "\t";
                                testname = testname + TOComp.Substring(TOComp.IndexOf("test ", 0) + 4, TOComp.IndexOf("\"", TOComp.IndexOf("test ", 0)) - TOComp.IndexOf("test ", 0) - 4).Trim() + "\t";
                                testname = testname + TOComp.Substring(TOComp.IndexOf("\"", 0) + 1, TOComp.IndexOf("\"", TOComp.IndexOf("\"", 0) + 1) - TOComp.IndexOf("\"", 0) - 1);
                            }
                            if (TOComp.Trim().ToLower().Substring(0, TOComp.IndexOf("\"") - 1).Contains("skip "))
                            {
                                testname = testname + TOComp.Substring(0, TOComp.IndexOf("skip ", 0) + 4).Trim() + "\t";
                                testname = testname + TOComp.Substring(TOComp.IndexOf("skip ", 0) + 4, TOComp.IndexOf("\"", TOComp.IndexOf("skip ", 0)) - TOComp.IndexOf("skip ", 0) - 4).Trim() + "\t";
                                testname = testname + TOComp.Substring(TOComp.IndexOf("\"", 0) + 1, TOComp.IndexOf("\"", TOComp.IndexOf("\"", 0) + 1) - TOComp.IndexOf("\"", 0) - 1);
                            }

                            //testname = TOComp.Substring(0, TOComp.IndexOf("\"", TOComp.IndexOf("\"", 0) + 1) + 1).Trim();
                            TOtests.Add(testname + "\t");
                            // Extract Version Name
                            //tempcomp = tempcomp + "\t";
                            testname = testname + "\t";
                            string versionname = "Base";
                            if (TOComp.ToLower().Contains("version"))
                                //tempcomp = tempcomp + TOComp.Substring(TOComp.IndexOf("\"", TOComp.ToLower().IndexOf("version", 0)) + 1, TOComp.IndexOf("\"", TOComp.IndexOf("\"", TOComp.ToLower().IndexOf("version", 0)) + 1) - TOComp.IndexOf("\"", TOComp.ToLower().IndexOf("version", 0)) - 1);
                                versionname = TOComp.Substring(TOComp.IndexOf("\"", TOComp.ToLower().IndexOf("version", 0)) + 1, TOComp.IndexOf("\"", TOComp.IndexOf("\"", TOComp.ToLower().IndexOf("version", 0)) + 1) - TOComp.IndexOf("\"", TOComp.ToLower().IndexOf("version", 0)) - 1).Trim();

                            // Extract Options
                            //tempcomp = tempcomp + "\t";
                            testname = TOtests[TOtests.FindIndex(item => item == testname)] = TOtests.Find(item => item == testname) + versionname + "\t";

                            string optionsname = "";
                            if (TOComp.Contains(";"))
                            {
                                if (TOComp.Contains("!"))
                                    //tempcomp = tempcomp + TOComp.Substring(TOComp.IndexOf(";", 0) + 1, TOComp.IndexOf("!", 0) - TOComp.IndexOf(";", 0) - 1);
                                    optionsname = TOComp.Substring(TOComp.IndexOf(";", 0) + 1, TOComp.IndexOf("!", 0) - TOComp.IndexOf(";", 0) - 1).Trim();
                                else
                                    //tempcomp = tempcomp + TOComp.Substring(TOComp.IndexOf(";", 0) + 1, TOComp.Length - TOComp.IndexOf(";", 0) - 1);
                                    optionsname = TOComp.Substring(TOComp.IndexOf(";", 0) + 1, TOComp.Length - TOComp.IndexOf(";", 0) - 1).Trim();
                            }

                            testname = TOtests[TOtests.FindIndex(item => item == testname)] = TOtests.Find(item => item == testname) + optionsname + "\t";

                            // Exrtact Comment
                            //tempcomp = tempcomp + "\t";
                            string comments = "";
                            if (TOComp.Contains("!"))
                                //tempcomp = tempcomp + TOComp.Substring(TOComp.IndexOf("!", 0) + 1, TOComp.Length - TOComp.IndexOf("!", 0) - 1);
                                comments = TOComp.Substring(TOComp.IndexOf("!", 0) + 1, TOComp.Length - TOComp.IndexOf("!", 0) - 1).Trim();
                            testname = TOtests[TOtests.FindIndex(item => item == testname)] = TOtests.Find(item => item == testname) + comments;
                        }

                    }
                    //WriteToExcel(@"C:\Users\uidw8018\Desktop\ICTCOMP.xlsx", "Testorder", TOtests);
                    
                    //FileStream myWriteStream = new FileStream(Path.GetFullPath(@"C:\Users\uidw8018\Desktop\Testorder" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".xls"), FileMode.Append, FileAccess.Write);
                    //byte[] newLine = Encoding.Default.GetBytes(Environment.NewLine);
                    //foreach (string writeline in tests)
                    //{
                    //    myWriteStream.Write(Encoding.ASCII.GetBytes(writeline.Trim()), 0, Encoding.ASCII.GetByteCount(writeline.Trim()));
                    //    myWriteStream.Write(newLine, 0, newLine.Length);
                    //}
                    //myWriteStream.Close();
                    StatusL.Text = "Testorder Extraction Complete";
                }
                else
                {
                    MessageBox.Show("Testorder file does not exists");
                }

                if (File.Exists(Sel_TPTB.Text))
                {
                    TPtests.Clear();
                    string[] TPfile = File.ReadAllLines(Sel_TPTB.Text);

                    foreach (string TPComp in TPfile)
                    {
                        //string tempcomp = "";
                        string testname = "";
                        // Extract Test Name
                        if (TPComp.Trim().ToLower().Contains("test \""))
                        {
                            bool valid = (TPComp.Trim().ToLower()[0] == 't');
                            testname = TPComp.Substring(TPComp.IndexOf("\"", 0) + 1, TPComp.IndexOf("\"", TPComp.IndexOf("\"", 0) + 1) - TPComp.IndexOf("\"", 0) - 1);
                            
                            if (testname.Contains("\\"))
                                testname = testname.Substring(testname.IndexOf("\\", 0) + 1);
                            if (testname.Contains("/"))
                                testname = testname.Substring(testname.IndexOf("/", 0) + 1);

                            if (testname.Trim().IndexOf("1%") == 0)
                                testname = testname.Substring(testname.IndexOf("1%", 0) + 2);
                            if (testname.Trim().IndexOf("2%") == 0)
                                testname = testname.Substring(testname.IndexOf("2%", 0) + 2);
                            if (testname != "")
                            {
                                // Extract Version Name
                                //tempcomp = tempcomp + "\t";

                                if (valid)
                                {
                                    TPtests.Add(testname + "\t" + "\t");
                                    testname = testname + "\t" + "\t";
                                }
                                else
                                {
                                    TPtests.Add(testname + "\t" + TPComp.Trim().Substring(0, TPComp.IndexOf("test \"", 0)).Trim() + "\t");
                                    testname = testname + "\t" + TPComp.Trim().Substring(0, TPComp.IndexOf("test \"", 0)).Trim() + "\t";
                                }

                                // Exrtact Comment
                                //tempcomp = tempcomp + "\t";
                                string comments = "";
                                if (TPComp.Substring(TPComp.IndexOf("test ", 0)).Contains("!"))
                                    //tempcomp = tempcomp + TOComp.Substring(TOComp.IndexOf("!", 0) + 1, TOComp.Length - TOComp.IndexOf("!", 0) - 1);
                                    comments = TPComp.Substring(TPComp.IndexOf("!", TPComp.IndexOf("test ", 0)), TPComp.Length - TPComp.IndexOf("!", TPComp.IndexOf("test ", 0))).Trim();
                                testname = TPtests[TPtests.FindIndex(item => item == testname)] = TPtests.Find(item => item == testname) + comments;
                            }
                           // else
                              //  TPtests.Add(TPComp.Trim());
                        }
                    }

                    //WriteToExcel(@"C:\Users\uidw8018\Desktop\ICTCOMP.xlsx", "Testplan", TPtests);
                    //FileStream myWriteStream = new FileStream(Path.GetFullPath(@"C:\Users\uidw8018\Desktop\Testplan" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".xls"), FileMode.Append, FileAccess.Write);
                    //byte[] newLine = Encoding.Default.GetBytes(Environment.NewLine);
                    //foreach (string writeline in tests)
                    //{
                    //    myWriteStream.Write(Encoding.ASCII.GetBytes(writeline.Trim()), 0, Encoding.ASCII.GetByteCount(writeline.Trim()));
                    //    myWriteStream.Write(newLine, 0, newLine.Length);
                    //}
                    //myWriteStream.Close();
                    StatusL.Text = "Testplan Extraction Complete";

                }
                else
                    MessageBox.Show("Testplan file does not exists");

                if (File.Exists(Sel_TOTB.Text))
                {
                    BOMtests.Clear();
                    string[] BOMfile = File.ReadAllLines(Select_BomTB.Text);

                    foreach (string BOMComp in BOMfile)
                    {
                        BOMtests.Add(BOMComp.Trim());
                    }
                    StatusL.Text = "BOM File Extraction Complete";
                }
                else
                    MessageBox.Show("BOM file does not exists");
                

                if (File.Exists(Select_LogTB.Text))
                {
                    LOGtests.Clear();
                    string[] Logfile = File.ReadAllLines(Select_LogTB.Text);

                    foreach (string LogComp in Logfile)
                    {
                        //string tempcomp = "";
                        string testname = "";
                        // Extract Test Name
                        if (LogComp.Trim().ToLower().Contains("@block|") || LogComp.Trim().ToLower().Contains("@d-t|"))
                        {
                            if (LogComp.Trim().ToLower().Contains("@block"))
                                testname = LogComp.Split('|').ElementAt(1);

                            if (LogComp.Trim().ToLower().Contains("@d-t"))
                                testname = LogComp.Split('|').ElementAt(5);

                            if (testname.Trim().IndexOf("1%") == 0)
                                testname = testname.Substring(testname.IndexOf("1%", 0) + 2);
                            if (testname.Trim().IndexOf("2%") == 0)
                                testname = testname.Substring(testname.IndexOf("2%", 0) + 2);

                            LOGtests.Add(testname.Trim());
                        }
                    }
                    //WriteToExcel(@"C:\Users\uidw8018\Desktop\ICTCOMP.xlsx");
                    
                    //FileStream myWriteStream = new FileStream(Path.GetFullPath(@"C:\Users\uidw8018\Desktop\Log" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".xls"), FileMode.Append, FileAccess.Write);
                    //byte[] newLine = Encoding.Default.GetBytes(Environment.NewLine);
                    //foreach (string writeline in tests)
                    //{
                    //    myWriteStream.Write(Encoding.ASCII.GetBytes(writeline.Trim()), 0, Encoding.ASCII.GetByteCount(writeline.Trim()));
                    //    myWriteStream.Write(newLine, 0, newLine.Length);
                    //}
                    //myWriteStream.Close();
                    StatusL.Text = "Log File Extraction Complete";
                }
                else
                    MessageBox.Show("Log file does not exists");

                WriteToExcel(@"C:\Users\uidw8018\Desktop\ICTCOMP"+DateTime.Now.ToString("YYMMddHHmmss")+".xlsx");

            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.ToString());
            }
        }

        private void Sel_TOTB_TextChanged(object sender, EventArgs e)
        {

        }

        private void SelTPB_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "Select Testplan File..";
            openFileDialog1.Filter = "Testplan|*testplan*";
            DialogResult res = openFileDialog1.ShowDialog();
            if (res == DialogResult.OK)
                Sel_TPTB.Text = openFileDialog1.FileName;
        }


        private void WriteToExcel(string expath)
        {
            try
                {
                    Cursor.Current = Cursors.AppStarting;
                    xlApp = new Microsoft.Office.Interop.Excel.Application();

                    object misValue = System.Reflection.Missing.Value;

                    if (xlApp == null)
                    {
                        MessageBox.Show("Excel is not properly installed!!");
                        return;
                    }


                    if (File.Exists(expath))
                        File.Delete(expath);

                    xlWorkBook = xlApp.Workbooks.Add();

                    string[] shtnames = new string[5] { "LOG", "Testplan", "Testorder", "BOM", "Comparison" };

                    //xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets.Add(misValue);
                    //xlWorkSheet.Name = "Sheet1";
                    //for (int s = 0; s < xlWorkBook.Sheets.Count; ++s)
                    //{
                    //    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets.get_Item(s + 1);
                    //    for (int h=0;h<shtnames.Length;++h)
                    //    if (xlWorkSheet.Name == shtnames[h])
                    //        xlWorkBook.Sheets[shtnames[h]].Delete();
                    //}

                    StatusL.Text = "Writing to Excel File";

                    for (int s = 0; s <= shtnames.Length - xlWorkBook.Sheets.Count; ++s)
                    {
                        xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets.Add(misValue);
                    }

                    for (int s = 0; s < shtnames.Length; ++s)
                    {
                        xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets.get_Item(s+1);
                        xlWorkSheet.Name = shtnames[s];
                    }

                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets["LOG"];
                    for (int i = 0; i < LOGtests.Count; i++)
                    {
                        StatusL.Text = "Extracting Log File to Excel";
                        for (int j = 0; j < LOGtests[i].Split('\t').Count(); ++j)
                            xlWorkSheet.Cells[i + 1, j + 1].Value2 = LOGtests[i].Split('\t').ElementAt(j).ToString();

                        // DONT CHANGE THE ORDER OF THE BELOW CONDITIONS ---------- Always Check for NOT OK First and then OK second as "OK" is also present in "NOT OK"
                        //if (CheckStatusArray[j].Contains("NOT OK"))
                        //    xlWorkSheet.get_Range(ExcelAddress[i], ExcelAddress[i]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                        //else if (CheckStatusArray[j].Contains("OK"))
                        //    xlWorkSheet.get_Range(ExcelAddress[i], ExcelAddress[i]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LimeGreen);
                    }

                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets["Testplan"];
                    for (int i = 0; i < TPtests.Count; i++)
                    {
                        StatusL.Text = "Extracting Testplan to Excel";
                        for (int j = 0; j < TPtests[i].Split('\t').Count(); ++j)
                            xlWorkSheet.Cells[i + 1, j + 1].Value2 = TPtests[i].Split('\t').ElementAt(j).ToString();

                        // DONT CHANGE THE ORDER OF THE BELOW CONDITIONS ---------- Always Check for NOT OK First and then OK second as "OK" is also present in "NOT OK"
                        //if (CheckStatusArray[j].Contains("NOT OK"))
                        //    xlWorkSheet.get_Range(ExcelAddress[i], ExcelAddress[i]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                        //else if (CheckStatusArray[j].Contains("OK"))
                        //    xlWorkSheet.get_Range(ExcelAddress[i], ExcelAddress[i]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LimeGreen);
                    }

                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets["Testorder"];
                    for (int i = 0; i < TOtests.Count; i++)
                    {
                        StatusL.Text = "Extracting Testorder to Excel";
                        for (int j = 0; j < TOtests[i].Split('\t').Count(); ++j)
                            xlWorkSheet.Cells[i + 1, j + 1].Value2 = TOtests[i].Split('\t').ElementAt(j).ToString();

                        // DONT CHANGE THE ORDER OF THE BELOW CONDITIONS ---------- Always Check for NOT OK First and then OK second as "OK" is also present in "NOT OK"
                        //if (CheckStatusArray[j].Contains("NOT OK"))
                        //    xlWorkSheet.get_Range(ExcelAddress[i], ExcelAddress[i]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                        //else if (CheckStatusArray[j].Contains("OK"))
                        //    xlWorkSheet.get_Range(ExcelAddress[i], ExcelAddress[i]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LimeGreen);
                    }

                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets["Comparison"];
                    int maxlength = Math.Max(TOtests.Count, Math.Max(TPtests.Count, Math.Max(BOMtests.Count, LOGtests.Count)));

                    List<string> comptests = new List<string>();    
                    
                        try
                        {
                            //bool cond1 = (i >= TOtests.Count) || (TOtests.Count == 0);
                            //bool cond2 = (i >= TPtests.Count) || (TPtests.Count == 0);
                            //bool cond3 = (i >= BOMtests.Count) || (BOMtests.Count == 0);
                            //bool cond4 = (i >= LOGtests.Count) || (LOGtests.Count == 0);

                            //if (cond1 && !cond2 && !cond3 && !cond4)
                            //    comptests.Add("\t" + "\t" + TPtests[i].ElementAt(0) + "\t" + BOMtests[i].ElementAt(0) + "\t" + LOGtests[i].ElementAt(0));
                            //else if (!cond1 && cond2 && !cond3 && !cond4)
                            //    comptests.Add(TOtests[i].ElementAt(2) + "\t" + "\t" + "\t" + BOMtests[i].ElementAt(0) + "\t" + LOGtests[i].ElementAt(0));
                            //else if (!cond1 && !cond2 && cond3 && !cond4)
                            //    comptests.Add(TOtests[i].ElementAt(2) + "\t" + TPtests[i].ElementAt(0) + "\t" + "\t" + "\t" + LOGtests[i].ElementAt(0));
                            //else if (!cond1 && !cond2 && !cond3 && cond4)
                            //    comptests.Add(TOtests[i].ElementAt(2) + "\t" + TPtests[i].ElementAt(0) + "\t" + BOMtests[i].ElementAt(0) + "\t" + "\t");
                            //else
                            //    comptests.Add(TOtests[i].ElementAt(2) + "\t" + TPtests[i].ElementAt(0) + "\t" + BOMtests[i].ElementAt(0) + "\t" + LOGtests[i].ElementAt(0));
                            List<string> BOMt = new List<string>();
                            List<string> TOt = new List<string>();
                            List<string> TPt = new List<string>();
                            List<string> LOGt = new List<string>();
                            List<string> TOCOM1 = new List<string>();
                            List<string> TOCOM2 = new List<string>();
                            List<string> TPCOM1 = new List<string>();
                            List<string> TPCOM2 = new List<string>();


                            StatusL.Text = "Preparing data to be compared";

                            for (int i = 0; i < TOtests.Count; i++)
                            {
                                //for (int j = 0; j < comptests[i].Split('\t').Count(); ++j)
                                TOt.Add(TOtests[i].Split('\t').ElementAt(2).ToString());
                                
                                    TOt.Sort();
                                    TOt = TOt.Distinct().ToList();

                                    TOCOM1.Add(TOtests[i].Split('\t').ElementAt(4).ToString());
                                    TOCOM1.Sort();
                                    TOCOM1 = TOCOM1.Distinct().ToList();

                                    TOCOM2.Add(TOtests[i].Split('\t').ElementAt(5).ToString());
                                    TOCOM2.Sort();
                                    TOCOM2 = TOCOM2.Distinct().ToList();
                                
                            }
                            
                            for (int i = 0; i < TPtests.Count; i++)
                            {
                                // for (int j = 0; j < comptests[i].Split('\t').Count(); ++j)
                                TPt.Add(TPtests[i].Split('\t').ElementAt(0).ToString());
                                
                                    TPt.Sort();
                                    TPt = TPt.Distinct().ToList();

                                    TPCOM1.Add(TPtests[i].Split('\t').ElementAt(1).ToString());
                                    TPCOM1.Sort();
                                    TPCOM1 = TPCOM1.Distinct().ToList();

                                    TPCOM2.Add(TPtests[i].Split('\t').ElementAt(2).ToString());
                                    TPCOM2.Sort();
                                    TPCOM2 = TPCOM2.Distinct().ToList();
                               
                            }
                            
                            for (int i = 0; i < BOMtests.Count; i++)
                            {
                                //for (int j = 0; j < comptests[i].Split('\t').Count(); ++j)
                                BOMt.Add(BOMtests[i]);
                                
                                    BOMt.Sort();
                                    BOMt = BOMt.Distinct().ToList();
                               
                            }
                            
                            for (int i = 0; i < LOGtests.Count; i++)
                            {
                                //for (int j = 0; j < comptests[i].Split('\t').Count(); ++j)
                                LOGt.Add(LOGtests[i].Split('\t').ElementAt(0).ToString());
                                    LOGt.Sort();
                                    LOGt = LOGt.Distinct().ToList();
                            }

                            //for (int i = 0; i < TOtests.Count; i++)
                            //{
                            //    // for (int j = 0; j < comptests[i].Split('\t').Count(); ++j)
                            //    TOCOM1.Add(TOtests[i].Split('\t').ElementAt(4).ToString());
                            //    TOCOM1.Sort();
                            //    TOCOM1 = TOCOM1.Distinct().ToList();
                            //}

                            //for (int i = 0; i < TOtests.Count; i++)
                            //{
                            //    // for (int j = 0; j < comptests[i].Split('\t').Count(); ++j)
                            //    TOCOM2.Add(TOtests[i].Split('\t').ElementAt(5).ToString());
                            //    TOCOM2.Sort();
                            //    TOCOM2 = TOCOM2.Distinct().ToList();
                            //}

                            //for (int i = 0; i < TPtests.Count; i++)
                            //{
                            //    //for (int j = 0; j < comptests[i].Split('\t').Count(); ++j)
                            //    TPCOM1.Add(TPt[i].Split('\t').ElementAt(1).ToString());
                            //    TPCOM1.Sort();
                            //    TPCOM1 = TPCOM1.Distinct().ToList();
                            //}

                            //for (int i = 0; i < TPtests.Count; i++)
                            //{
                            //    //for (int j = 0; j < comptests[i].Split('\t').Count(); ++j)
                            //    TPCOM2.Add(TPtests[i].Split('\t').ElementAt(2).ToString());
                            //    TPCOM2.Sort();
                            //    TPCOM2 = TPCOM2.Distinct().ToList();
                            //}


                            for (int f = 0; f < maxlength; ++f)
                            {
                                StatusL.Text = "Comparing data";
                                maxlength = Math.Max(TOt.Count, Math.Max(TPt.Count, Math.Max(BOMt.Count, LOGt.Count)));
                                string TO = "", TP = "", BOM = "", LOG = "", resstr = "";
                               
                                if (f >= TOt.Count)
                                    TO = "";
                                else
                                    TO = TOt[f];

                                if (f >= TPt.Count)
                                    TP = "";
                                else
                                    TP = TPt[f];

                                if (f >= BOMt.Count)
                                    BOM = "";
                                else
                                    BOM = BOMt[f];

                                if (f >= LOGt.Count)
                                    LOG = "";
                                else
                                    LOG = LOGt[f];


                                resstr = loweststring(new string[4] { TO, TP, BOM, LOG });

                                if (f >= TOt.Count)
                                {
                                    TOt.Insert(f, "");
                                    TOCOM1.Insert(f, "");
                                    TOCOM2.Insert(f, "");
                                }
                                else if (string.Compare(resstr, TOt[f], true) == -1)
                                {
                                    TOt.Insert(f, "");
                                    TOCOM1.Insert(f, "");
                                    TOCOM2.Insert(f, "");
                                }
                                else if (string.Compare(resstr, TOt[f], true) == 1)
                                {
                                    TOt.Insert(f, "");
                                    TPt.Insert(f, "");
                                    BOMt.Insert(f, "");
                                    LOGt.Insert(f, "");
                                    TOCOM1.Insert(f, "");
                                    TOCOM2.Insert(f, "");
                                    TPCOM1.Insert(f, "");
                                    TPCOM2.Insert(f, "");
                                }


                                if (f >= TPt.Count)
                                {
                                    TPt.Insert(f, "");
                                    TPCOM1.Insert(f, "");
                                    TPCOM2.Insert(f, "");
                                }
                                else if (string.Compare(resstr, TPt[f], true) == -1)
                                {
                                    TPt.Insert(f, "");
                                    TPCOM1.Insert(f, "");
                                    TPCOM2.Insert(f, "");
                                }
                                else if (string.Compare(resstr, TPt[f], true) == 1)
                                {
                                    TOt.Insert(f, "");
                                    TPt.Insert(f, "");
                                    BOMt.Insert(f, "");
                                    LOGt.Insert(f, "");
                                    TOCOM1.Insert(f, "");
                                    TOCOM2.Insert(f, "");
                                    TPCOM1.Insert(f, "");
                                    TPCOM2.Insert(f, "");
                                }


                                if (f >= BOMt.Count)
                                    BOMt.Insert(f, "");
                                else if (string.Compare(resstr, BOMt[f], true) == -1)
                                {
                                    BOMt.Insert(f, "");
                                }
                                else if (string.Compare(resstr, BOMt[f], true) == 1)
                                {
                                    TOt.Insert(f, "");
                                    TPt.Insert(f, "");
                                    BOMt.Insert(f, "");
                                    LOGt.Insert(f, "");
                                    TOCOM1.Insert(f, "");
                                    TOCOM2.Insert(f, "");
                                    TPCOM1.Insert(f, "");
                                    TPCOM2.Insert(f, "");
                                }


                                if (f >= LOGt.Count)
                                    LOGt.Insert(f, "");
                                else if (string.Compare(resstr, LOGt[f], true) == -1)
                                {
                                    LOGt.Insert(f, "");
                                }
                                else if (string.Compare(resstr, LOGt[f], true) == 1)
                                {
                                    TOt.Insert(f, "");
                                    TPt.Insert(f, "");
                                    BOMt.Insert(f, "");
                                    LOGt.Insert(f, "");
                                    TOCOM1.Insert(f, "");
                                    TOCOM2.Insert(f, "");
                                    TPCOM1.Insert(f, "");
                                    TPCOM2.Insert(f, "");
                                }
                            }

                            StatusL.Text = "Writing Compared data to excel";

                           xlWorkSheet.Cells[1, 1].Value2 = "Testorder";
                            for (int i = 0; i < TOt.Count; i++)
                            {
                                //for (int j = 0; j < comptests[i].Split('\t').Count(); ++j)
                                xlWorkSheet.Cells[i + 2, 1].Value2 = TOt[i];
                            }
                            xlWorkSheet.Cells[1, 2].Value2 = "Testplan";
                            for (int i = 0; i < TPt.Count; i++)
                            {
                                // for (int j = 0; j < comptests[i].Split('\t').Count(); ++j)
                                xlWorkSheet.Cells[i + 2, 2].Value2 = TPt[i];
                            }
                            xlWorkSheet.Cells[1, 3].Value2 = "BOM";
                            for (int i = 0; i < BOMt.Count; i++)
                            {
                                //for (int j = 0; j < comptests[i].Split('\t').Count(); ++j)
                                xlWorkSheet.Cells[i + 2, 3].Value2 = BOMt[i];
                            }
                            xlWorkSheet.Cells[1, 4].Value2 = "LOG";
                            for (int i = 0; i < LOGt.Count; i++)
                            {
                                //for (int j = 0; j < comptests[i].Split('\t').Count(); ++j)
                                xlWorkSheet.Cells[i + 2, 4].Value2 = LOGt[i];
                            }

                            xlWorkSheet.Cells[1, 5].Value2 = "Testorder Comments";
                            for (int i = 0; i < TOCOM1.Count; i++)
                            {
                                // for (int j = 0; j < comptests[i].Split('\t').Count(); ++j)
                                xlWorkSheet.Cells[i + 2, 5].Value2 = TOCOM1[i] + " <|> " + TOCOM2[i];
                            }
                            xlWorkSheet.Cells[1, 6].Value2 = "Testplan Comments";
                            for (int i = 0; i < TPCOM1.Count; i++)
                            {
                                //for (int j = 0; j < comptests[i].Split('\t').Count(); ++j)
                                xlWorkSheet.Cells[i + 2, 6].Value2 = TPCOM1[i] + " <|> " + TPCOM2[i];
                            }
                            StatusL.Text = "Comparison finished and results written to Excel";
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.ToString());
                        }



                        Excel.Style style = xlWorkBook.Styles.Add("myStyle");

                        style.Font.Name = "Arial";
                        style.Font.Bold = true;
                        style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        style.Font.Size = 12;
                        style.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

                        style.Interior.Pattern = Excel.XlPattern.xlPatternSolid;

                    ////xlWorkBook = xlApp.Workbooks.Add(misValue);
                    ////foreach(string addSTR in ExcelAddress)
                    //for (int i = 0, j = 0; i < ExcelAddress.Count; i++, j++)
                    //{
                    //    if (CheckStatusArray[j].Contains("NOT OK"))
                    //    {
                    //        style.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    //        xlWorkSheet.Range[ExcelAddress[i]].Style = "myStyle";
                    //    }
                    //    else if (CheckStatusArray[j] == "OK")
                    //    {
                    //        style.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LimeGreen);
                    //        xlWorkSheet.Range[ExcelAddress[i]].Style = "myStyle";
                    //    }
                    //    else
                    //    {
                    //        style.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    //        xlWorkSheet.Range[ExcelAddress[i]].Style = "myStyle";
                    //    }
                    //}
                    if (File.Exists(expath))
                        xlWorkBook.Save();
                    else
                    xlWorkBook.SaveAs(expath);
                    //StatusCB.Items.Add(DateTime.Now.ToLongTimeString() + " - Excel Report Updated Successfully!");
                    //StatusCB.BackColor = Color.LimeGreen;
                    Cursor.Current = Cursors.Default;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    //StatusCB.Items.Add(DateTime.Now.ToLongTimeString() +" - ERROR. Click \"Show Detail\" to view error.");
                    //StatusCB.BackColor = Color.Red;
                    //MessageBox.Show(ex.Message + Environment.NewLine + Environment.NewLine + ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    xlWorkBook.Close(false);
                    xlApp.Quit();

                    releaseObject(xlWorkSheet);
                    releaseObject(xlWorkBook);
                    releaseObject(xlApp);
                }

        }

        private string loweststring(string[] strarr)
        {
            List<string> temparr = new List<string>();
            foreach (string str in strarr)
            {
                if (str != "")
                    temparr.Add(str);
            }
            return temparr.Min();
        }
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void Select_LogTB_TextChanged(object sender, EventArgs e)
        {

        }
    }
}

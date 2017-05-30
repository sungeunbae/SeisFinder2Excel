using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using Excel= Microsoft.Office.Interop.Excel;


namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }



        private void btn_select_none_Click(object sender, EventArgs e)
        {

        }

        private void btn_convert_Click(object sender, EventArgs e)
        {

        }

        private void InitializeComponent()
        {
            this.btnOpen = new System.Windows.Forms.Button();
            this.btnSelectAll = new System.Windows.Forms.Button();
            this.btnSelectNone = new System.Windows.Forms.Button();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.checkedListBox2 = new System.Windows.Forms.CheckedListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.btnConvert = new System.Windows.Forms.Button();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.btnExit = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnOpen
            // 
            this.btnOpen.Location = new System.Drawing.Point(262, 26);
            this.btnOpen.Name = "btnOpen";
            this.btnOpen.Size = new System.Drawing.Size(75, 23);
            this.btnOpen.TabIndex = 0;
            this.btnOpen.Text = "Open";
            this.btnOpen.UseVisualStyleBackColor = true;
            // 
            // btnSelectAll
            // 
            this.btnSelectAll.Location = new System.Drawing.Point(262, 83);
            this.btnSelectAll.Name = "btnSelectAll";
            this.btnSelectAll.Size = new System.Drawing.Size(75, 23);
            this.btnSelectAll.TabIndex = 1;
            this.btnSelectAll.Text = "Select All";
            this.btnSelectAll.UseVisualStyleBackColor = true;
            // 
            // btnSelectNone
            // 
            this.btnSelectNone.Location = new System.Drawing.Point(262, 117);
            this.btnSelectNone.Name = "btnSelectNone";
            this.btnSelectNone.Size = new System.Drawing.Size(75, 23);
            this.btnSelectNone.TabIndex = 2;
            this.btnSelectNone.Text = "Select None";
            this.btnSelectNone.UseVisualStyleBackColor = true;
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(13, 29);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(243, 20);
            this.textBox2.TabIndex = 3;
            // 
            // checkedListBox2
            // 
            this.checkedListBox2.FormattingEnabled = true;
            this.checkedListBox2.Location = new System.Drawing.Point(13, 83);
            this.checkedListBox2.Name = "checkedListBox2";
            this.checkedListBox2.Size = new System.Drawing.Size(243, 259);
            this.checkedListBox2.TabIndex = 4;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(10, 67);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(94, 13);
            this.label1.TabIndex = 5;
            this.label1.Text = "Stations to select :";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 9);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(95, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "Input Data Folder :";
            // 
            // btnConvert
            // 
            this.btnConvert.Location = new System.Drawing.Point(94, 400);
            this.btnConvert.Name = "btnConvert";
            this.btnConvert.Size = new System.Drawing.Size(75, 23);
            this.btnConvert.TabIndex = 7;
            this.btnConvert.Text = "Convert";
            this.btnConvert.UseVisualStyleBackColor = true;
            this.btnConvert.Click += new System.EventHandler(this.btnConvert_Click);
            // 
            // textBox3
            // 
            this.textBox3.Location = new System.Drawing.Point(13, 371);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(243, 20);
            this.textBox3.TabIndex = 8;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 355);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 13);
            this.label3.TabIndex = 9;
            this.label3.Text = "Save As :";
            // 
            // btnBrowse
            // 
            this.btnBrowse.Location = new System.Drawing.Point(13, 400);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(75, 23);
            this.btnBrowse.TabIndex = 10;
            this.btnBrowse.Text = "Browse";
            this.btnBrowse.UseVisualStyleBackColor = true;
            // 
            // btnExit
            // 
            this.btnExit.Location = new System.Drawing.Point(262, 400);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(75, 23);
            this.btnExit.TabIndex = 11;
            this.btnExit.Text = "Exit";
            this.btnExit.UseVisualStyleBackColor = true;
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // Form1
            // 
            this.ClientSize = new System.Drawing.Size(349, 429);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnBrowse);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.textBox3);
            this.Controls.Add(this.btnConvert);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.checkedListBox2);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.btnSelectNone);
            this.Controls.Add(this.btnSelectAll);
            this.Controls.Add(this.btnOpen);
            this.Name = "Form1";
            this.Text = "SeisFinder2Excel";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private void btnConvert_Click(object sender, EventArgs e)
        {
            Object misValue = System.Reflection.Missing.Value;
            Excel.Application xls = new Excel.Application();
            Excel.Workbook xlsWorkBook = xls.Workbooks.Add(misValue);
            Excel.Worksheet xlsWorkSheet = (Excel.Worksheet)xlsWorkBook.Sheets[1];
            
            System.IO.StreamReader fileReader;

            String info1, info2;
            int nt;
            double dt;
            String filePathPrefix = "C:\\tmp\\AMBC\\AMBC";
            String excelExt = "xlsx";

            xlsWorkSheet.Cells[1, 1] = "Component";
            xlsWorkSheet.Cells[2, 1] = "No. of timesteps";
            xlsWorkSheet.Cells[3, 1] = "Size of timesteps";
            xlsWorkSheet.Cells[4, 1] = "Acceleration";

            String[] componentsCode = { "000", "090", "ver" };
            String[] componentsStr = { "X-axis (000)", "Y-axis (090)", "Z-axis (ver)" };

            for (int k=0;k<3;k++)
            {
                String ext = componentsCode[k];
                fileReader = new StreamReader(new FileStream(filePathPrefix + "." + ext, FileMode.Open));
                info1 = fileReader.ReadLine();
                info2 = fileReader.ReadLine();

                int LastNonEmpty = -1;
                String[] info2Array = info2.Split();
                for (int i = 0; i < info2Array.Length; i++)
                {
                    if (info2Array[i] != "") {
                        LastNonEmpty += 1;
                        info2Array[LastNonEmpty] = info2Array[i];
                    }
                }

                nt = Convert.ToInt32(info2Array[0]);
                dt = Convert.ToDouble(info2Array[1]);

                int [,] intRange = new int[nt,1];
                xlsWorkSheet.Cells[1, 2 + k] = componentsStr[k];
                xlsWorkSheet.Cells[2, 2 + k] = nt;
                xlsWorkSheet.Cells[3, 2 + k] = dt;

                LastNonEmpty = -1;

                String lines = fileReader.ReadToEnd();
                String[] strArray = lines.Split();
                Double[,] dblArray = new Double[nt,1]; //needs to be 2-d to be able to bulk write

                for (int i=0;i< strArray.Length;i++)
                {
                    if (strArray[i]!="")
                    {
                        LastNonEmpty += 1;
                        intRange[LastNonEmpty,0] = LastNonEmpty + 1;
                        dblArray[LastNonEmpty,0] = Convert.ToDouble(strArray[i].Replace("\n", ""));

                    }
                }
                int nt2 = LastNonEmpty + 1;
                if (k==0)
                {
                    xlsWorkSheet.Range["A5"].Resize[nt2, 1].Value = intRange;
                }
                Excel.Range cell;
                cell = xlsWorkSheet.Cells[5, 2 + k];
                xlsWorkSheet.Range[cell, cell].Resize[nt2, 1].Value = dblArray;
                
            }



            xlsWorkBook.SaveAs(filePathPrefix+"."+excelExt, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue,
        false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
        misValue, misValue, misValue, misValue, misValue);

            xlsWorkBook.Close();
            xls.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlsWorkSheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlsWorkBook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xls);
            MessageBox.Show("Done!");

        }

        private void btnExit_Click(object sender, EventArgs e)
        {

            // The user wants to exit the application. Close everything down.
            Application.Exit();
        }
    }
}

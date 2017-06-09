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
using System.Reflection;
//using Microsoft.WindowsAPICodePack.Dialogs;

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
        

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
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
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnOpen
            // 
            this.btnOpen.Location = new System.Drawing.Point(504, 27);
            this.btnOpen.Name = "btnOpen";
            this.btnOpen.Size = new System.Drawing.Size(75, 23);
            this.btnOpen.TabIndex = 0;
            this.btnOpen.Text = "Browse";
            this.btnOpen.UseVisualStyleBackColor = true;
            this.btnOpen.Click += new System.EventHandler(this.btnOpen_Click);
            // 
            // btnSelectAll
            // 
            this.btnSelectAll.Location = new System.Drawing.Point(12, 397);
            this.btnSelectAll.Name = "btnSelectAll";
            this.btnSelectAll.Size = new System.Drawing.Size(75, 23);
            this.btnSelectAll.TabIndex = 1;
            this.btnSelectAll.Text = "Select All";
            this.btnSelectAll.UseVisualStyleBackColor = true;
            this.btnSelectAll.Click += new System.EventHandler(this.btnSelectAll_Click);
            // 
            // btnSelectNone
            // 
            this.btnSelectNone.Location = new System.Drawing.Point(93, 397);
            this.btnSelectNone.Name = "btnSelectNone";
            this.btnSelectNone.Size = new System.Drawing.Size(75, 23);
            this.btnSelectNone.TabIndex = 2;
            this.btnSelectNone.Text = "Select None";
            this.btnSelectNone.UseVisualStyleBackColor = true;
            this.btnSelectNone.Click += new System.EventHandler(this.btnSelectNone_Click);
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(13, 29);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(468, 20);
            this.textBox2.TabIndex = 3;
            // 
            // checkedListBox2
            // 
            this.checkedListBox2.FormattingEnabled = true;
            this.checkedListBox2.Location = new System.Drawing.Point(12, 123);
            this.checkedListBox2.Name = "checkedListBox2";
            this.checkedListBox2.Size = new System.Drawing.Size(180, 259);
            this.checkedListBox2.TabIndex = 4;
            this.checkedListBox2.SelectedIndexChanged += new System.EventHandler(this.checkedListBox2_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(10, 104);
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
            this.btnConvert.Location = new System.Drawing.Point(352, 397);
            this.btnConvert.Name = "btnConvert";
            this.btnConvert.Size = new System.Drawing.Size(75, 23);
            this.btnConvert.TabIndex = 7;
            this.btnConvert.Text = "Convert";
            this.btnConvert.UseVisualStyleBackColor = true;
            this.btnConvert.Click += new System.EventHandler(this.btnConvert_Click);
            // 
            // textBox3
            // 
            this.textBox3.Location = new System.Drawing.Point(12, 70);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(469, 20);
            this.textBox3.TabIndex = 8;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 54);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(77, 13);
            this.label3.TabIndex = 9;
            this.label3.Text = "Output Folder :";
            // 
            // btnBrowse
            // 
            this.btnBrowse.Location = new System.Drawing.Point(504, 67);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(75, 23);
            this.btnBrowse.TabIndex = 10;
            this.btnBrowse.Text = "Browse";
            this.btnBrowse.UseVisualStyleBackColor = true;
            // 
            // btnExit
            // 
            this.btnExit.Location = new System.Drawing.Point(504, 397);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(75, 23);
            this.btnExit.TabIndex = 11;
            this.btnExit.Text = "Exit";
            this.btnExit.UseVisualStyleBackColor = true;
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // textBox4
            // 
            this.textBox4.Location = new System.Drawing.Point(211, 123);
            this.textBox4.Multiline = true;
            this.textBox4.Name = "textBox4";
            this.textBox4.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBox4.Size = new System.Drawing.Size(368, 259);
            this.textBox4.TabIndex = 12;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(211, 104);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(54, 13);
            this.label4.TabIndex = 13;
            this.label4.Text = "Progress :";
            // 
            // Form1
            // 
            this.ClientSize = new System.Drawing.Size(591, 432);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.textBox4);
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
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            string version = Assembly.GetExecutingAssembly().GetName().Version.ToString();
            this.Text = String.Format("SeisFinder2Excel  QuakeCoRE Soft - version {0}", version);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private void btnConvert_Click(object sender, EventArgs e)
        {
            string inputPath = this.textBox2.Text;
            string outputPath = this.textBox3.Text;
            Object misValue = System.Reflection.Missing.Value;
            Excel.Application xls = new Excel.Application();
            String excelExt = "xlsx";
            Directory.CreateDirectory(outputPath);

            this.textBox4.Clear();

            int numCheckedStations = checkedListBox2.CheckedItems.Count ;
            int numStationsProcessed=0;
            foreach (string stationCode in checkedListBox2.CheckedItems)
            {
                numStationsProcessed++;

                Excel.Workbook xlsWorkBook = xls.Workbooks.Add(misValue);
                Excel.Worksheet xlsWorkSheet = (Excel.Worksheet)xlsWorkBook.Sheets[1];

                System.IO.StreamReader fileReader;

                String info1, info2;
                int nt = 0;
                double dt;
                String filePathPrefix = inputPath+"\\"+ stationCode;
                String outputFile = outputPath + "\\" + stationCode + "." + excelExt;

                //this.textBox4.Text +="Processing "+stationCode + "("+numStationsProcessed+"/"+numCheckedStations+")  ";
                this.textBox4.AppendText("Processing " + stationCode + "(" + numStationsProcessed + "/" + numCheckedStations + ")  ");

                xlsWorkSheet.Cells[1, 1] = "Component";
                xlsWorkSheet.Cells[2, 1] = "No. of timesteps";
                xlsWorkSheet.Cells[3, 1] = "Size of timesteps (seconds)";
                xlsWorkSheet.Cells[4, 1] = "Time (seconds)";
                xlsWorkSheet.Cells[5, 1] = "Acceleration";

                String[] componentsCode = { "000", "090", "ver" };
                String[] componentsStr = { "X-axis (000)", "Y-axis (090)", "Z-axis (ver)" };

                for (int k = 0; k < 3; k++) //Iterates each component file
                {

                    this.textBox4.AppendText(Convert.ToString(k + 1)+"...");

                    String ext = componentsCode[k];
                    fileReader = new StreamReader(new FileStream(filePathPrefix + "." + ext, FileMode.Open));
                    info1 = fileReader.ReadLine();
                    info2 = fileReader.ReadLine();

                    int LastNonEmpty = -1;
                    String[] info2Array = info2.Split();
                    for (int i = 0; i < info2Array.Length; i++)
                    {
                        if (info2Array[i] != "")
                        {
                            LastNonEmpty += 1;
                            info2Array[LastNonEmpty] = info2Array[i];
                        }
                    }

                    nt = Convert.ToInt32(info2Array[0]); //number of time steps
                    dt = Convert.ToDouble(info2Array[1]); //time step size

                    Double[,] timeRange = new double[nt, 1];
                    xlsWorkSheet.Cells[1, 2 + k] = componentsStr[k];
                    xlsWorkSheet.Cells[2, 2 + k] = nt;
                    xlsWorkSheet.Cells[3, 2 + k] = dt;

                    LastNonEmpty = -1;

                    String lines = fileReader.ReadToEnd();
                    String[] strArray = lines.Split();
                    Double[,] dblArray = new Double[nt, 1]; //needs to be 2-d to be able to bulk write

                    for (int i = 0; i < strArray.Length; i++)
                    {
                        if (strArray[i] != "")
                        {
                            LastNonEmpty += 1;
                            timeRange[LastNonEmpty, 0] = LastNonEmpty * dt;
                            dblArray[LastNonEmpty, 0] = Convert.ToDouble(strArray[i].Replace("\n", ""));

                        }
                    }
                    int nt2 = LastNonEmpty + 1;
                    if (k == 0)
                    {
                        xlsWorkSheet.Range["A5"].Resize[nt2, 1].Value = timeRange; // first column filled with time steps
                    }
                    Excel.Range cell;
                    cell = xlsWorkSheet.Cells[5, 2 + k];
                    xlsWorkSheet.Range[cell, cell].Resize[nt2, 1].Value = dblArray; //bulk write at column "B, C, D"


                }
                Excel.WorksheetFunction wsf = xls.WorksheetFunction;
                Double maxAcc = wsf.Max(xlsWorkSheet.Range[xlsWorkSheet.Cells[5, 4], xlsWorkSheet.Cells[nt + 4, 4]]); //find the maximum and minimum
                Double minAcc = wsf.Min(xlsWorkSheet.Range[xlsWorkSheet.Cells[5, 4], xlsWorkSheet.Cells[nt + 4, 4]]);
                Double maxAmpAcc = Math.Max(maxAcc, -1.0*minAcc); //Select the max of absolute value of two

                for (int k = 0; k < 3; k++) // Plotting
                {
                    String ext = componentsCode[k];

                    Excel.Range chartRange;
                    Excel.ChartObjects xlCharts = (Excel.ChartObjects)xlsWorkSheet.ChartObjects(Type.Missing);
                    Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(200, 80 + 300 * k, 600, 250);
                    Excel.Chart chartPage = myChart.Chart;

                    chartRange = xlsWorkSheet.Range[xlsWorkSheet.Cells[5, 2 + k], xlsWorkSheet.Cells[nt + 4, 2 + k]];
                    
                    chartPage.SetSourceData(chartRange, misValue);
                    chartPage.ChartType = Excel.XlChartType.xlLine;
                    chartPage.HasTitle = true;
                    
                    Excel.Series series = (Excel.Series) chartPage.SeriesCollection(1);

                    Microsoft.Office.Interop.Excel.Axis xAxis = (Microsoft.Office.Interop.Excel.Axis)chartPage.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                    Microsoft.Office.Interop.Excel.Axis yAxis = (Microsoft.Office.Interop.Excel.Axis)chartPage.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                    xAxis.HasTitle = true;
                    xAxis.AxisTitle.Text = "Time (sec)";
                    xAxis.CategoryNames = (Excel.Range)xlsWorkSheet.Range["A5"].Resize[nt, 1];
                    xAxis.TickLabelPosition = Excel.XlTickLabelPosition.xlTickLabelPositionLow;
                    yAxis.HasTitle = true;
                    yAxis.AxisTitle.Text = "Acceleration (cm/s^2)";
                    yAxis.MinimumScale = (-1* maxAmpAcc * 3.0);
                    yAxis.MaximumScale = (maxAmpAcc * 3.0);

                    series.Name = ext;

                    
                    switch (k)
                    {
                        case 0:
                            chartPage.ChartTitle.Text = "["+ stationCode+"]: "+"Acceleration along X-axis (" +ext+")";
                            series.Border.Color = (int)Excel.XlRgbColor.rgbRed;
                            
                            break;
                        case 1:
                            chartPage.ChartTitle.Text = "["+ stationCode+"]: "+"Acceleration along Y-axis (" + ext + ")";
                            series.Border.Color = (int)Excel.XlRgbColor.rgbBlue;
                            break;
                        case 2:
                            chartPage.ChartTitle.Text = "[" + stationCode + "]: " + "Acceleration along Z-axis (" + ext + ")";
                            series.Border.Color = (int)Excel.XlRgbColor.rgbGreen;
                            break;
                    }
                    

                }



                this.textBox4.AppendText("Done..!" +Environment.NewLine);

                if (File.Exists(outputFile)) // delete file if it already exists
                {
                    File.Delete(outputFile);
                }

                xlsWorkBook.SaveAs(outputFile, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue,
            false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
            misValue, misValue, misValue, misValue, misValue);

                xlsWorkBook.Close();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlsWorkSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlsWorkBook);
            }
            
            xls.Quit();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(xls);
            //MessageBox.Show("Done!");
            this.textBox4.AppendText("Finished!!" + Environment.NewLine);

        }

        private void btnExit_Click(object sender, EventArgs e)
        {

            // The user wants to exit the application. Close everything down.
            Application.Exit();
        }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                fbd.SelectedPath = this.textBox2.Text;
                DialogResult result = fbd.ShowDialog();

                this.textBox3.Text = fbd.SelectedPath + "\\Output";

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    this.textBox2.Text = fbd.SelectedPath;
                    var files = Directory.EnumerateFiles(fbd.SelectedPath);

                    checkedListBox2.Items.Clear();

                    IDictionary<string, int> fileDict = new Dictionary<string, int>();

                    foreach (string currentFile in files)
                    {
                        if (currentFile.EndsWith(".000") || currentFile.EndsWith(".090")|| currentFile.EndsWith(".ver"))
                        {
                            string station = Path.GetFileNameWithoutExtension(@currentFile);
                            //string station = currentFile.Substring(0, currentFile.Length - ".000".Length);
                            if (currentFile.EndsWith(".000"))
                            {
                                if (fileDict.ContainsKey(station)) fileDict[station] = fileDict[station] | 1;
                                else
                                    fileDict[station] = 1;
                            }
                            else if (currentFile.EndsWith(".090"))
                            {
                                if (fileDict.ContainsKey(station)) fileDict[station] = fileDict[station] | 2;
                                else
                                    fileDict[station] = 2;
                            }
                            else if (currentFile.EndsWith(".ver"))
                            {
                                if (fileDict.ContainsKey(station)) fileDict[station] = fileDict[station] | 4;
                                else
                                    fileDict[station] = 4;
                            }
                        }
                        
                    }
                    for (int i = 0; i<fileDict.Count; i++)
                    {
                        var item = fileDict.ElementAt(i);
                        var itemKey = item.Key;
                        var itemValue = item.Value;
                        if (itemValue == 7) //All 000, 090, ver are present
                        {
                            checkedListBox2.Items.Add(itemKey, CheckState.Unchecked);
                        }
                    }
                }
            }
            //http://www.lyquidity.com/devblog/?p=136

            //Stream myStream = null;
            //OpenFileDialog openFileDialog1 = new OpenFileDialog();

            //openFileDialog1.InitialDirectory= System.Environment.SpecialFolder.Personal;
            //openFileDialog1.Filter = "Acc/Vel files|*.000;*.090;*.ver";
            //openFileDialog1.FilterIndex = 2;
            //openFileDialog1.RestoreDirectory = true;

            //if (openFileDialog1.ShowDialog() == DialogResult.OK)
            //{
            //    try
            //    {
            //        if ((myStream = openFileDialog1.OpenFile()) != null)
            //        {
            //            using (myStream)
            //            {
            //                // Insert code to read the stream here.
            //            }
            //        }
            //    }
            //    catch (Exception ex)
            //    {
            //        MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
            //    }
            //}
        }

        private void checkedListBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void btnSelectAll_Click(object sender, EventArgs e)
        {

            for (int i = 0; i < checkedListBox2.Items.Count; i++)
                checkedListBox2.SetItemChecked(i,true);

        }

        private void btnSelectNone_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < checkedListBox2.Items.Count; i++)
                checkedListBox2.SetItemChecked(i, false);
        }



        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }
    }
}

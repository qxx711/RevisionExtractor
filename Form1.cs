using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.IO;
using ICSharpCode.SharpZipLib.Zip;
using LibGit2Sharp;
using System.Threading;

namespace ReversionExtract
{
    public partial class Form1 : Form
    {
        FolderBrowserDialog folder = new FolderBrowserDialog();
        string OutputFolderPath = "";
        Microsoft.Office.Interop.Excel.Application app = null;
        int col;
        int row;
        private Thread t;
        private int current = 0;
        private int total = 1;
        Extractor extractor;
        private uint starttime;
        private uint endtime;
        private Boolean failToLink = false;
        public Form1()
        {
            InitializeComponent();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            folder.ShowDialog();
            textBox1.Text = folder.SelectedPath;
        }
     
        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                textBox2.Text = dialog.FileName;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            folder.ShowDialog();
            textBox3.Text = folder.SelectedPath;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            progressPresent.ForeColor = Color.Green;
            progressPresent.Text = "Running";

            String gitPath = textBox1.Text;
            String excelPath = textBox2.Text;
            OutputFolderPath = textBox3.Text;

            List<List<object>> InputMatrix;
            int RowNum;
            ReadExcel(excelPath, out InputMatrix, out RowNum);

            extractor = new GitExtractor(textBox1.Text, textBox3.Text, InputMatrix, RowNum);
            progressBar1.Value = 0;
            progressBar1.Visible = true;
            progressPresent.Visible = true;
            timer1.Start();
            button4.Text = "Extracting...";
            button4.Enabled = false;
            button4.Update();
            t = new Thread(new ThreadStart(extractor.doExtract));
            t.IsBackground = true;
            t.Start();
            progressPresent.Text = "success!";            

        }
        public void ReadExcel(string FilePath, out List<List<object>> InputMatrix, out int RowNum)
        {
            Microsoft.Office.Interop.Excel.Workbook workBook = null;
            app = new Microsoft.Office.Interop.Excel.Application();
            workBook = app.Workbooks.Open(FilePath);
            Worksheet worksheet = (Worksheet)workBook.Worksheets[5];
            col = worksheet.UsedRange.CurrentRegion.Columns.Count;
            row = worksheet.UsedRange.CurrentRegion.Rows.Count;
            InputMatrix = NewListMatrixOfObject(row, col);
            object[,] current;
            current = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[row, col]].Value2;

            int k = 1;
            for (int i = 0; i <= row - 1; i++)
            {
                if (current[i + 1, 7] != null)
                {
                    for (int j = 1; j <= col; j++)
                    {
                        InputMatrix[k - 1][j - 1] = current[i + 1, j];
                    }
                    k++;
                }
            }
            RowNum = k - 1;
            app.Quit();
            app = null;
            workBook = null;
        }
        public static List<List<object>> NewListMatrixOfObject(int iRowCount, int iColCount)
        {
            List<List<object>> matrix = new List<List<object>>(iRowCount);
            for (int i = 0; i < iRowCount; i++)
            {
                matrix.Add(new List<object>(iColCount));
                for (int j = 0; j < iColCount; j++)
                {
                    matrix[i].Add("");
                }
            }
            return matrix;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            current = extractor.current;
            total = extractor.total;
            if (total != 0)
            {
                int pos = (int)(((float)current) / total * 100);
                this.progressBar1.Value = (int)pos;
                this.progressPresent.Text = string.Format("{0}%", pos);
                if (current == total)
                {
                    starttime = extractor.starttime;
                    endtime = extractor.endtime;
                    setInformation();
                }
            }
        }
        public void setInformation()
        {
            this.button4.Enabled = true;
            this.button4.Text = "Extract";
            if (!failToLink)
            {
                this.info.Text = "耗时：" + Convert.ToString((endtime - starttime) / 1000) + "秒";
                this.info.ForeColor = System.Drawing.Color.Black;
            }
            else
            {
                this.info.Text = "Fail!";
                this.info.ForeColor = System.Drawing.Color.Red;
            }
            this.progressBar1.Visible = false;
            this.progressPresent.Visible = false;
        }


    }
}

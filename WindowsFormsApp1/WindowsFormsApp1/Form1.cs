using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        //comboBox1 is number of pathologists
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem != null)
            {
                GlobalVariables.Pathologists = int.Parse(comboBox1.SelectedItem.ToString());
                //MessageBox.Show(GlobalVariables.Pathologists.ToString()); //just for testing
            }
        }

        //comboBox2 is number of sheet
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.SelectedItem != null)
            {
                GlobalVariables.SheetNumber = int.Parse(comboBox2.SelectedItem.ToString());
                //MessageBox.Show(GlobalVariables.SheetNumber.ToString()); //just for testing

            }
        }

        //does magical things
        OpenFileDialog ofd = new OpenFileDialog();

        //button1 is browse
        private void button1_Click(object sender, EventArgs e)
        {

            ofd.Filter = "XLSX (*.xlsx)|*.xlsx|All files (*.*)|*.*";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = ofd.SafeFileName;
                GlobalVariables.FileName = ofd.SafeFileName;
                GlobalVariables.FileLocation = ofd.FileName;
                //MessageBox.Show(GlobalVariables.FileLocation); //just for testing
            }
        }

        //button2 is submit
        private void button2_Click(object sender, EventArgs e)
        {
            Excel excel = new Excel(GlobalVariables.FileLocation, GlobalVariables.SheetNumber);

            int sheet = GlobalVariables.SheetNumber;
            int c = excel.wb.Worksheets[sheet].UsedRange.Columns.Count;
            int r = excel.wb.Worksheets[sheet].UsedRange.Rows.Count;

            string[,] excelData = new string[r, c];

            for (int row = 1; row <= r; row++)
            {
                for (int col = 1; col <= c; col++)
                {
                    excelData[row - 1, col - 1] = Convert.ToString(excel.wb.Worksheets[sheet].Cells[row, col].Value2);
                    //MessageBox.Show(excelData[row - 1, col - 1]);
                }
            }

            // example
            previewImported(excelData);

            //MessageBox.Show(excel.ReadCell(0, 0));
            //MessageBox.Show("Columns: " + c + "\n" + "Rows: " + r);
            //sol 

            CloseFile(excel.wb);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //previewImported(excelData);
        }

        public void CloseFile(Workbook wb)
        {
            wb.Close(0);
        }

        public void previewImported(string[,] excelData)
        {
            int height = excelData.GetLength(0);
            int width = excelData.GetLength(1);

            this.dataGridView1.ColumnCount = width;

            for (int r = 0; r < height; r++)
            {
                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(this.dataGridView1);

                for (int c = 0; c < width; c++)
                {
                    row.Cells[c].Value = excelData[r, c];
                }

                this.dataGridView1.Rows.Add(row);
            }
        }

    }

    class Excel
    {
        string path = GlobalVariables.FileLocation;
        _Application excel = new _Excel.Application();
        public Workbook wb;
        public Worksheet ws;

        public Excel(string path, int sheet)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[sheet];
        }

        public string ReadCell(int i, int j)
        {
            i++;
            j++;

            if (ws.Cells[i, j].Value2 != null) return Convert.ToString(ws.Cells[i, j].Value2);
            else return "";
        }
    }
}

public static class GlobalVariables
{
    public static String FileName;
    public static String FileLocation;
    public static int Pathologists;
    public static int SheetNumber;
    public static int NumberInColumnOne = 1;
}
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
                MessageBox.Show(GlobalVariables.Pathologists.ToString()); //just for testing
            }
        }

        //comboBox2 is number of sheet
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.SelectedItem != null)
            {
                GlobalVariables.SheetNumber = int.Parse(comboBox2.SelectedItem.ToString());
                MessageBox.Show(GlobalVariables.SheetNumber.ToString()); //just for testing
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
                GlobalVariables.FileLocation = ofd.FileName;
                MessageBox.Show(GlobalVariables.FileLocation); //just for testing
            }
        }

        //button2 is submit
        private void button2_Click(object sender, EventArgs e)
        {
            
        }
    }

    class Excel
    {
        RowCount = m_ActiveWorkSheet.UsedRange.Rows.Count;
    }
}

public static class GlobalVariables
{
    public static String FileLocation;
    public static int Pathologists;
    public static int SheetNumber;
    public static int NumberInColumnOne = 1;
}
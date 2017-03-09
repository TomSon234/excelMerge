using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;       //microsoft Excel 14 object in references-> COM tab

namespace excelMerge
{
    public partial class Form1 : Form
    {
        public int rowCount = 0;
        public int colCount = 0;

        public Form1()
        {
            InitializeComponent();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox1.Text) || string.IsNullOrWhiteSpace(textBox2.Text) || string.IsNullOrWhiteSpace(textBox3.Text))
            {
                MessageBox.Show("Określ pliki źródłowe i wynikowe", "Wybierz pliki", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else {
                String plik1 = textBox1.Text;
                String plik2 = textBox2.Text;
                String wynik = textBox3.Text;            


                //Create COM Objects. Create a COM object for everything that is referenced
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@plik1);
                Excel.Workbook xlWorkbook2 = xlApp.Workbooks.Open(@plik2);
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excel._Worksheet xlWorksheet2 = xlWorkbook2.Sheets[1];

                Excel.Range xlRange = xlWorksheet.UsedRange;
                Excel.Range xlRange2 = xlWorksheet2.UsedRange;

                if (string.IsNullOrWhiteSpace(textBox6.Text))
                    rowCount = xlRange.Rows.Count;
                else
                    rowCount = Int32.Parse(textBox6.Text);

                if (string.IsNullOrWhiteSpace(textBox7.Text))
                    colCount = xlRange.Columns.Count;
                else
                    colCount = Int32.Parse(textBox7.Text);
            


                //iterate over the rows and columns and print to the console as it appears in the file
                //excel is not zero based!!
                for (int i = 1; i <= rowCount; i++)
                {
                    for (int j = 1; j <= colCount; j++)
                    {
                        //new line
                        if (j == 1)
                        {
                            //Console.Write("\r\n");
                        }

                        //write the value to the console
                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        {
                            //Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");
                        }
                        else
                        {
                            if (xlRange2.Cells[i, j] != null && xlRange2.Cells[i, j].Value2 != null)
                            {
                                xlRange.Cells[i, j] = xlRange2.Cells[i, j].Value2;
                                //Console.Write(xlRange2.Cells[i, j].Value2.ToString() + "\t");
                            }
                        }
                    }
                }

                xlWorksheet.SaveAs(@wynik);
                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //rule of thumb for releasing com objects:
                //  never use two dots, all COM objects must be referenced and released individually
                //  ex: [somthing].[something].[something] is bad

                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);
                Marshal.ReleaseComObject(xlRange2);
                Marshal.ReleaseComObject(xlWorksheet2);

                //close and release
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);
                xlWorkbook2.Close();
                Marshal.ReleaseComObject(xlWorkbook2);

                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);

            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                string file = openFileDialog1.FileName;
                textBox1.Text = file;

                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@file);
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];

                Excel.Range xlRange = xlWorksheet.UsedRange;
                textBox4.Text = xlRange.Rows.Count.ToString();
                textBox5.Text = xlRange.Columns.Count.ToString();

                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                string file = openFileDialog1.FileName;
                textBox2.Text = file;

                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@file);
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];

                Excel.Range xlRange = xlWorksheet.UsedRange;
                textBox4.Text = xlRange.Rows.Count.ToString();
                textBox5.Text = xlRange.Columns.Count.ToString();

                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult result = saveFileDialog1.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                string file = saveFileDialog1.FileName;
                textBox3.Text = file;
            }
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {

        }
    }
}

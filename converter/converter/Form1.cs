using System;
using System.Windows.Forms;
using System.IO;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System.Data;
using System.Reflection;
using System.Text;
using System.Diagnostics;

namespace converter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        public int numOfLastRow = 0;
        private void btnSelect_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.InitialDirectory = "D:\\";
            dialog.Filter = "Excel files (*.xls, *.xlsx, *.xlsm)|*.xls;*.xlsx;*.xlsm";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                label2.Text = "File is being processed...";
                try
                {
                    Microsoft.Office.Interop.Excel.Application myExcel;
                    Microsoft.Office.Interop.Excel.Workbook myWorkbook;
                    Microsoft.Office.Interop.Excel.Worksheet myWorksheet;

                    myExcel = new Microsoft.Office.Interop.Excel.Application();
                    myExcel.Workbooks.Open(@dialog.FileName);
                    myWorkbook = myExcel.ActiveWorkbook;
                    myWorksheet = (Excel.Worksheet)myWorkbook.Sheets[1];

                    StreamWriter sw = new StreamWriter(@"D:\outputFile.txt", false, Encoding.GetEncoding("Windows-1251"));

                    string str = "";
                    int count = 0;
                    var lastCell = myWorksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);

                    //записываем название, автора и вопросы
                    for (int i = 0; i < lastCell.Column; i++)
                    {
                        for (int j = 0; j < lastCell.Row; j++)
                        {
                            str = myWorksheet.Cells[j + 1, i + 1].Text.ToString();
                            if ((myWorksheet.Cells[j + 1, 1].Text.ToString() == "") || (myWorksheet.Cells[j + 1, 1].Text.ToString() == "    "))
                                count++;
                            sw.WriteLine(str, Encoding.GetEncoding("Windows-1251"));
                            if (count == 2)
                            {
                                numOfLastRow = j + 1;
                                break;
                            }
                        }
                        if (count == 2) break;
                    }
                    

                    //записываем вероятности
                    for (int j = numOfLastRow; j < lastCell.Row; j++)
                    {
                        for (int i = 0; i < lastCell.Column; i++)
                        {
                            str = myWorksheet.Cells[j + 1, i + 1].Text.ToString();
                            str = str.Replace(",", ".");
                            if ((myWorksheet.Cells[j + 1, i + 2].Text.ToString() == "") || (myWorksheet.Cells[j + 1, i + 2].Text.ToString() == "    "))
                                str += "\n";
                            else
                                str += ",";
                            sw.Write(str);
                        }
                    }
                    sw.Close();
                    myWorkbook.Close(false);
                    myExcel.Quit();

                }
                catch (Exception)
                {
                    MessageBox.Show("No file selected");
                }
            }
            string name = dialog.FileName;
            int position = name.LastIndexOf("\\");
            name = name.Substring(position + 1);
            label2.Text = "Selected file:";
            label3.Text = name;
            btnConvert.Enabled = true;
            btnSelect.Enabled = false;
            btnOpen.Enabled = false;
            MessageBox.Show("Processing completed");
        }

        private void btnConvert_Click(object sender, EventArgs e)
        {
            // здесь будет код по конвертированию файла :)
            File.Move(@"D:\outputFile.txt", @"D:\outputFile.mkb");
            //тут должна быть смена кодировки outputFile.mkb!!!!
            label2.Text = "File converted!";
            label3.Text = "";
            btnSelect.Enabled = true;
            btnOpen.Enabled = true;
            btnConvert.Enabled = false;
        }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            Process.Start("D:\\");
            label2.Text = "";
            btnSelect.Enabled = true;
            btnOpen.Enabled = false;
        }

        private void btnInfo_Click(object sender, EventArgs e)
        {
            MessageBox.Show("To change the format from *.xls or *.xlsx to *.mkb, you need to save the sheet as new Excel book and open the Excel file by clicking the \"Select file\" button.\n\nAfter that the \"Convert to *.mkb\" button will become active. To convert the file, click this button and wait for the file to be processed.\n\nAt the end of the process, you can open a file or select a new one for conversion.\n\nPath to the converted file: \"D:\\\"\n\nThe program is developed by Marina Kubrina(c). 2017");
        }
    }
}

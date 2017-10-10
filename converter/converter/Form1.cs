using System;
using System.Windows.Forms;
using System.IO;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace converter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        
        private void btnSelect_Click(object sender, EventArgs e)
        {
            Stream stream = null;
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.InitialDirectory = "D:\\";
            dialog.Filter = "Excel files (*.xls, *.xlsx, *.xlsm)|*.xls;*.xlsx;*.xlsm";
            if (dialog.ShowDialog() == DialogResult.OK)
                try
                {
                    /*excelappworkbook = excelapp.Workbooks.Open(dialog.FileName,
    Type.Missing, Type.Missing, 1, Type.Missing,
    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
    Type.Missing, Type.Missing);
                    excelsheets = excelappworkbook.Worksheets;
                    //Получаем ссылку на лист 1
                    excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);*/
                    // здесь будет код по открытию эксель-файла :)
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Файл не выбран ");
                }
            string name = dialog.FileName;
            int position = name.LastIndexOf("\\");
            name = name.Substring(position + 1);
            label2.Text = "Selected file:";
            label3.Text = name;
            btnConvert.Enabled = true;
            btnSelect.Enabled = false;
            btnOpen.Enabled = false;
        }

        private void btnConvert_Click(object sender, EventArgs e)
        {
            // здесь будет код по конвертированию файла :)
            MessageBox.Show("Типо готово");

            label2.Text = "File converted!";
            label3.Text = "";
            btnSelect.Enabled = true;
            btnOpen.Enabled = true;
            btnConvert.Enabled = false;
        }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            // здесь будет код по открытию файла
            MessageBox.Show("Типо файл открылся");

            label2.Text = "";
            btnSelect.Enabled = true;
            btnOpen.Enabled = false;
        }

        private void btnInfo_Click(object sender, EventArgs e)
        {
            MessageBox.Show("To change the format from *.xls or *.xlsx to *.mkb, you need to save the sheet called \"Expert\" and open the Excel file by clicking the \"Select file\" button.\n\nAfter that the \"Convert to *.mkb\" button will become active. To convert the file, click this button and wait for the file to be processed.\n\nAt the end of the process, you can open a file or select a new one for conversion.\n\nPath to the converted file: \"D:\\Knowledge base\"\n\nThe program is developed by Marina Kubrina(c). 2017");
        }
    }
}

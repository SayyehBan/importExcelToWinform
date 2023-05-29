using ExcelDataReader;
using System;
using System.Data;
using System.IO;
using System.Windows.Forms;

namespace importExcelToWinform
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Excel Files|*.xlsx;*.xls;*.xlsm";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog1.FileName;

                // Create an instance of FileStream to read from the Excel file
                using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    // Create an instance of IExcelDataReader to read the Excel file with 
                    using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        // Load the contents of the Excel file into a DataTable
                        DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                            {
                                UseHeaderRow = true
                            }
                        });

                        // Set the first table in the DataSet as the source of the DataGridView
                        dataGridView1.DataSource = result.Tables[0];
                    }
                }
            }
        }
    }
}
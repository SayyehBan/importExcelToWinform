دریافت فایل اکسل و انتقال آن به DataGridView در ویندوز فرم سی شارپ

سورس سی شارپ

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
        
        تصویر برنامه
        ![Snag_5f6fc3](https://github.com/SayyehBan/importExcelToWinform/assets/38620223/2daf85f8-ce29-4c40-be89-d9ad5800b0ea)

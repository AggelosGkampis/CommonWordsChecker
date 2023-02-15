using Excel;
using System;
using System.Data;
using System.IO;
using System.Windows.Forms;

namespace CommonWordsApplication
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        DataSet ds;

        private void ChooseFile_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog dialog = new OpenFileDialog()
            {
                Filter = "Excel Workbook |*.xlsx",
                ValidateNames = true
            })
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    FileStream fileStream = File.Open(dialog.FileName, FileMode.Open, FileAccess.Read);
                    IExcelDataReader reader = ExcelReaderFactory.CreateBinaryReader(fileStream);
                    reader.IsFirstRowAsColumnNames = true;
                    ds = reader.AsDataSet();
                    ResultGrid.DataSource = ds.Tables[0];
                }

        }

        private void ResultGrid_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            
            
            
        }

        private void ChooseFile_Click_1(object sender, EventArgs e)
        {

        }
    }
}

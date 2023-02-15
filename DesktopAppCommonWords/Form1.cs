using OfficeOpenXml;
using System.Data;
using System.Text;

namespace DesktopAppCommonWords
{
    public partial class Form1 : Form
    {
        private ExcelWorkbook workbook;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnchoose_Click(object sender, EventArgs e)
        {

            OpenFileDialog openFileDialog = new OpenFileDialog();

            openFileDialog.Filter = "Excel files (*.xls;*.xlsx)|*.xls;*.xlsx|All files (*.*)|*.*";
            openFileDialog.FilterIndex = 1;
            openFileDialog.RestoreDirectory = true;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                txtpath.Text = openFileDialog.FileName;

                // Import Excel file and store in variable using EPPlus
                using (var package = new ExcelPackage(new FileInfo(openFileDialog.FileName)))
                {
                    // Do something with the workbook, such as read or modify data
                    workbook = package.Workbook;
                    System.Data.DataTable table = Program.GetCommonWords(workbook);
                    dataGridView1.DataSource = table;
                    // When done, close the package and release the resources
                    package.Dispose();
                }
            }

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            // Create a new save file dialog to allow the user to choose a file name and location to save the exported Excel file
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel file (*.xlsx)|*.xlsx";
            saveFileDialog.Title = "Export Common Words to Excel";
            saveFileDialog.ShowDialog();

            // If the user selected a file name and location, export the data to the Excel file
            if (saveFileDialog.FileName != "")
            {
                try
                {
                    // Create a new Excel package and add a worksheet to it
                    using (ExcelPackage package = new ExcelPackage())
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Common Words");

                        // Write the column headers to the Excel worksheet
                        for (int i = 1; i <= dataGridView1.Columns.Count; i++)
                        {
                            worksheet.Cells[1, i].Value = dataGridView1.Columns[i - 1].HeaderText;
                        }

                        // Write each row of data to the Excel worksheet
                        for (int i = 1; i <= dataGridView1.Rows.Count; i++)
                        {
                            for (int j = 1; j <= dataGridView1.Columns.Count; j++)
                            {
                                worksheet.Cells[i + 1, j].Value = dataGridView1.Rows[i - 1].Cells[j - 1].Value;
                            }
                        }

                        // Save the Excel package to the selected file
                        package.SaveAs(new FileInfo(saveFileDialog.FileName));
                    }

                    MessageBox.Show("Common Words exported successfully to Excel file.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error exporting Common Words to Excel file: " + ex.Message);
                }
            }
        }

    }


}


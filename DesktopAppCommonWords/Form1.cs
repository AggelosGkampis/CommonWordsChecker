using OfficeOpenXml;

namespace DesktopAppCommonWords
{
    public partial class Form1 : Form
    {
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
                    ExcelWorkbook workbook = package.Workbook;
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
    }
}
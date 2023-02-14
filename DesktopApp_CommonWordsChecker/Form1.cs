using ExcelCommonWords;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DesktopApp_CommonWordsChecker
{
    public partial class Form : System.Windows.Forms.Form
    {       
        public Form()
        {
            InitializeComponent();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        private void MainForm_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void MainForm_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files.Length > 0)
            {
                var commonWords = new List<string>();
                var commonWordRowsInColumn1 = new List<string>();
                var commonWordRowsInColumn2 = new List<string>();

                var file = new FileInfo(files[0]);
                var arguments = new[] { file.FullName };
                CommonWordsCheck.Main(arguments);
                
                // retrieve the results from the console application
                // and populate the DataGridView control
                // ...
            }
        }
    }
}

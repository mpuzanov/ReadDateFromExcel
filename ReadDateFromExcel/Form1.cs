using ExcelDataReader;
using System;
using System.Data;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ReadDateFromExcel
{
    public partial class Form1 : Form
    {
        private string fileName = string.Empty;
        private DataTableCollection tableCollection = null;
        public Form1()
        {
            InitializeComponent();
        }

        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult res = openFileDialog1.ShowDialog();
                if (res == DialogResult.OK)
                {
                    fileName = openFileDialog1.FileName;
                    Text = fileName;
                    OpenExcelFile(fileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OpenExcelFile(string path)
        {
            FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read);
            IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);
            DataSet db = reader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = true
                }
            });
            tableCollection = db.Tables;
            toolStripComboBox1.Items.Clear();
            foreach (DataTable table in tableCollection)
            {
                toolStripComboBox1.Items.Add(table.TableName);
            }
            toolStripComboBox1.SelectedIndex = 0;
        }

        private void toolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable table = tableCollection[Convert.ToString(toolStripComboBox1.SelectedItem)];
            dataGridView1.DataSource = table;
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void сохранитьКакToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (tableCollection == null) return;
            try
            {
                DialogResult res = saveFileDialog1.ShowDialog();
                if (res == DialogResult.OK)
                {
                    fileName = saveFileDialog1.FileName;
                    Text = fileName;
                    SaveDataToExcelFile(FillData(dataGridView1), fileName);
                    MessageBox.Show("Ok");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private object[,] FillData(DataGridView dg)
        {
            object[,] data = new object[dg.RowCount, dg.ColumnCount];

            for (int j = 0; j < dg.Columns.Count; j++)
            {
                data[0, j] = dg.Columns[j].HeaderText;
            }

            for (int i = 0; i < dg.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dg.Columns.Count; j++)
                {
                    data[i + 1, j] = dg.Rows[i].Cells[j].Value;
                }
            }
            return data;
        }

        void SaveDataToExcelFile(object[,] data, string fileName)
        {
            int topRow = 1, leftCol = 1;
            int rows = data.GetUpperBound(0) + 1;
            int cols = data.GetUpperBound(1) + 1;

            Excel.Application XlApp = new Excel.Application();
            Excel.Workbook XlWorkBook = XlApp.Workbooks.Add(Type.Missing);
            Excel.Worksheet sheet = XlWorkBook.Worksheets.get_Item(1);

            object leftTop = sheet.Cells[topRow, leftCol];
            object rightBottom = sheet.Cells[rows, cols];

            Excel.Range range = sheet.get_Range(leftTop, rightBottom);

            range.Value2 = data;
            range.EntireColumn.AutoFit();
            range.EntireRow.AutoFit();

            XlApp.AlertBeforeOverwriting = false;
            XlWorkBook.SaveAs(fileName);
            //XlApp.Visible = true;
            XlApp.Quit();
        }
    }
}

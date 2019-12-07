using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace Excel_Differ
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private string _fileOne = "";
        private string _fileTwo = "";

        private string OpenFileToFileName()
        {
            string fileName = "";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fileName = openFileDialog1.FileName;
            }

            return fileName;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            _fileOne = OpenFileToFileName();
            textBox1.Text = _fileOne;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            _fileTwo = OpenFileToFileName();
            textBox2.Text = _fileTwo;
        }

        private List<string> ReadFirstLineFromExcelFile(string path)
        {
            var excelApp = new Excel.Application {Visible = true};
            var excelWorkbooks = excelApp.Workbooks;
            var excelWorkbook = excelWorkbooks.Open(path);
            var excelSheets = excelWorkbook.Sheets;
            var excelWorksheet = (Excel.Worksheet) excelSheets.Item[1];

            bool isRead = true;
            int index = 1;
            var list = new List<string>();

            while (isRead)
            {
                var p = (string) (excelWorksheet.Cells[1, index] as Excel.Range)?.Value;

                if (p == null)
                {
                    isRead = false;
                }
                else
                {
                    list.Add(p);
                    index++;
                }
            }

            excelWorkbook.Close();
            excelApp.Quit();

            return list;
        }

        private Tuple<List<string>, List<string>> DiffList(List<string> fileOne, List<string> fileTwo)
        {
            var include = new List<string>();
            var notInclude = new List<string>();
            var items = new SortedSet<string>();

            foreach (string itemOne in fileOne)
            {
                items.Add(itemOne);
            }

            foreach (string itemTwo in fileTwo)
            {
                items.Add(itemTwo);
            }

            foreach (string item in items)
            {
                if (fileOne.Contains(item) && fileTwo.Contains(item))
                {
                    include.Add(item);
                }
                else
                {
                    notInclude.Add(item);
                }
            }

            return new Tuple<List<string>, List<string>>(include, notInclude);
        }

        private void SaveReportToExcel(Tuple<List<string>, List<string>> report)
        {
            var excelApp = new Excel.Application {Visible = true};
            var excelWorkbooks = excelApp.Workbooks;
            var excelWorkbook = excelWorkbooks.Add();
            var excelSheets = excelWorkbook.Sheets;
            var excelWorksheet = (Excel.Worksheet) excelSheets.Item[1];

            var indexInclude = 2;
            excelWorksheet.Cells[1, 1] = @"Повторяются";
            foreach (var value in report.Item1)
            {
                excelWorksheet.Cells[indexInclude, 1] = value;
                indexInclude++;
            }

            var indexNotInclude = 2;
            excelWorksheet.Cells[1, 2] = @"Не повторяются";
            foreach (var value in report.Item2)
            {
                excelWorksheet.Cells[indexNotInclude, 2] = value;
                indexNotInclude++;
            }
        }

        private void Diff_Click(object sender, EventArgs e)
        {
            if (_fileOne != "" && _fileTwo != "")
            {
                var fileOne = ReadFirstLineFromExcelFile(_fileOne);
                var fileTwo = ReadFirstLineFromExcelFile(_fileTwo);

                var report = DiffList(fileOne, fileTwo);

                SaveReportToExcel(report);
            }
            else
            {
                MessageBox.Show(@"Не выбраны файлы для сравнения");
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            MessageBox.Show(Properties.Resources.Copyright);
        }
    }
}
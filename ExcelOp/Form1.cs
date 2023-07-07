using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelOp
{
    public partial class Form1 : Form
    {

        string[,] list = new string[1000, 50];
        public Form1()
        {
            InitializeComponent();
        }


        private void button1_Click(object sender, EventArgs e)
        {
            int n = ExportExcel();
            listBox1.Items.Clear();
            string s;
            for (int i = 0; i < n; i++) // по всем строкам
            {
                s = "";
                for (int j = 0; j < 50; j++) //по всем колонкам
                    s += " | " + list[i, j];
                listBox1.Items.Add(s);
            }
        }

        private int ExportExcel()
        {
            // Выбрать путь и имя файла в диалоговом окне
            OpenFileDialog ofd = new OpenFileDialog();
            // Задаем расширение имени файла по умолчанию (открывается папка с программой)
            ofd.DefaultExt = "*.xls;*.xlsx";
            // Задаем строку фильтра имен файлов, которая определяет варианты
            ofd.Filter = "файл Excel (Spisok.xlsx)|*.xlsx";
            // Задаем заголовок диалогового окна
            ofd.Title = "Выберите файл базы данных";
            if (!(ofd.ShowDialog() == DialogResult.OK))
            {
                MessageBox.Show(
                "Не удалось открыть файл",
                "Ошибка",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 0;
            }
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1]; //получить 1-й лист
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);//последнюю ячейку
                                                                                                // размеры базы
            int lastColumn = (int)lastCell.Column;
            int lastRow = (int)lastCell.Row;
            for (int j = 0; j < 50; j++) //по всем колонкам
                for (int i = 0; i < lastRow; i++) // по всем строкам
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text.ToString(); //считываем данные
            ObjWorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
            ObjWorkExcel.Quit(); // выйти из Excel
            GC.Collect(); // убрать за собой
            return lastRow;
        }
    }
}

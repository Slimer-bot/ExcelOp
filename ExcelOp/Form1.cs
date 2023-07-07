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
using Spire.Xls;
using System.IO;
using System.Drawing.Imaging;

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
            int n = ExportExcelText();
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

        private int ExportExcelText()
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "*.xls;*.xlsx";
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
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int lastColumn = (int)lastCell.Column;
            int lastRow = (int)lastCell.Row;
            for (int j = 0; j < lastColumn; j++) //по всем колонкам
                for (int i = 0; i < lastRow; i++) // по всем строкам
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text.ToString(); //считываем данные
            ObjWorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
            ObjWorkExcel.Quit(); // выйти из Excel
            GC.Collect(); // убрать за собой
            return lastRow;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "*.xls;*.xlsx";
            ofd.Filter = "файл Excel (Spisok.xlsx)|*.xlsx";
            // Задаем заголовок диалогового окна
            ofd.Title = "Выберите файл базы данных";
            if (!(ofd.ShowDialog() == DialogResult.OK))
            {
                MessageBox.Show(
                "Не удалось открыть файл",
                "Ошибка",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


        }

        private void button3_Click(object sender, EventArgs e)
        {
            string filePath = "";
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "*.xls;*.xlsx";
            ofd.Filter = "файл Excel (Spisok.xlsx)|*.xlsx";
            // Задаем заголовок диалогового окна
            ofd.Title = "Выберите файл базы данных";
            if (!(ofd.ShowDialog() == DialogResult.OK))
            {
                MessageBox.Show(
                "Не удалось открыть файл",
                "Ошибка",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else filePath = ofd.FileName;
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@filePath);
            Worksheet worksheet = workbook.Worksheets[0];
            using (MemoryStream ms = new MemoryStream())
            {
                worksheet.ToEMFStream(ms, 0, 3, 11, 11);
                Image image = Image.FromStream(ms);
                Bitmap images = ResetResolution(image as Metafile, 300);
                images.Save(@"C:\Users\Oleg_\OneDrive\Документы\Result.jpg", ImageFormat.Jpeg);
            }
            this.Close();
        }
        private static Bitmap ResetResolution(Metafile mf, float resolution)
        {
            int width = (int)(mf.Width * resolution / mf.HorizontalResolution);
            int height = (int)(mf.Height * resolution / mf.VerticalResolution);
            Bitmap bmp = new Bitmap(width, height);
            bmp.SetResolution(resolution, resolution);
            Graphics g = Graphics.FromImage(bmp);
            g.DrawImage(mf, 0, 0);
            g.Dispose();
            return bmp;
        }
    }
}

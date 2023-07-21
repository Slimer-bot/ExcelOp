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
using System.Threading;

namespace ExcelOp
{
    
    public partial class Form1 : Form
    {

        
        public Form1()
        {
            InitializeComponent();
            
        }

        int milsek = 10000;
        private void button1_Click(object sender, EventArgs e)
        {

        }



        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                milsek = Convert.ToInt32(textBox1.Text) * 1000;
            }
            catch
            {
                MessageBox.Show(
                "Периодичность не задана, будет использовано значение по умолчанию (10 сек)",
                "Уведомление",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            string filePath = "";
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "*.xls;*.xlsx";
            ofd.Filter = "файл Excel (Spisok.xlsx)|*.xlsx";
            ofd.Title = "Выберите файл базы данных";
            if (!(ofd.ShowDialog() == DialogResult.OK))
            {
                MessageBox.Show(
                "Не удалось открыть файл",
                "Ошибка",
                MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
            else
            {
                filePath = ofd.FileName;
                Workbook workbook = new Workbook();
                workbook.LoadFromFile(@filePath);
                Worksheet worksheet = workbook.Worksheets[1];
                int firstC = 1, firstR = 1, LastC = 9, LastR = 16;
                for (int y = 1; y <= 13; y++)
                {

                    for (int i = 1; i <= 7; i++)
                    {

                        using (MemoryStream ms = new MemoryStream())
                        {
                            worksheet.ToEMFStream(ms, firstR, firstC, LastR, LastC);
                            Image image = Image.FromStream(ms);
                            Bitmap images = ResetResolution(image as Metafile, 300);
                            images.Save(@"C:\Users\Oleg_\OneDrive\Документы\Отчеты\Result" + i + ".jpg", ImageFormat.Jpeg);
                            pictureBox1.Image = images;
                            GC.Collect();
                        }
                        switch(i)
                        {
                            case 1:
                                LastC += 1;
                                firstR += 16;
                                LastR += 16;
                                Sleep(milsek);
                                break;
                            
                            case 2:
                                firstR += 15;
                                LastR += 15;
                                Sleep(milsek);
                                break;
                            
                            case 3:
                                firstR += 14;
                                LastR += 14;
                                break;
                            
                            case 4:
                                firstR += 14;
                                LastR += 14;
                                break;
                            
                            case 5:
                                firstR += 7;
                                LastR += 7;
                                break;

                            case 6:
                                firstR += 14;
                                LastR += 14;
                                break;

                            //case 7:
                                //firstR += 5;
                                //LastR += 5;
                                //break;

                            default:
                                break;
                        }
                    }
                    GC.Collect(); // убрать за собой
                    firstR = 1;
                    firstC += 8;
                    LastC += 7;
                    LastR = 16;
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try 
            { 
                milsek = Convert.ToInt32(textBox1.Text) * 1000;
            }
            catch
            {
                MessageBox.Show(
                "Периодичность не задана, будет использовано значение по умолчанию (10 сек)",
                "Уведомление",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            string filePath = "";
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "*.xls;*.xlsx";
            ofd.Filter = "файл Excel (Spisok.xlsx)|*.xlsx";
            ofd.Title = "Выберите файл базы данных";
            if (!(ofd.ShowDialog() == DialogResult.OK))
            {
                MessageBox.Show(
                "Не удалось открыть файл",
                "Ошибка",
                MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
            else
            {
                filePath = ofd.FileName;
                Workbook workbook = new Workbook();
                workbook.LoadFromFile(@filePath);
                Worksheet worksheet = workbook.Worksheets[0];
                int firstC = 1, firstR = 1, LastC = 12, LastR = 17;
                for (int i = 1; i <= 6; i++)
                {
                    
                    using (MemoryStream ms = new MemoryStream())
                    {
                        worksheet.ToEMFStream(ms, firstR, firstC, LastR, LastC);
                        Image image = Image.FromStream(ms);
                        Bitmap images = ResetResolution(image as Metafile, 300);
                        Sleep(milsek);
                        //images.Save(@"C:\Users\Oleg_\OneDrive\Документы\Отчеты\Result" + i + ".jpg", ImageFormat.Jpeg);
                        pictureBox1.Image = images;
                        GC.Collect();
                    }
                    firstR += 17;
                    LastR += 17;
                    
                }
                GC.Collect(); // убрать за собой
            }
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

        public static void FindImg()
        {
            //
        }

        public static void Sleep(int milsek)
        {
            Thread.Sleep(milsek);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}

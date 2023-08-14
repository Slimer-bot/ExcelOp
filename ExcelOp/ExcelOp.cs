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
using System.Diagnostics;
using System.Threading;

namespace ExcelOp
{
    
    public partial class ExcelOp : Form
    {

        
        public ExcelOp()
        {
            InitializeComponent();

            timersec_toolStripTextBox1.Text = "Введите секунды"; // ДОБАВЬ НАСТРОЙКИ, В КОТОРЫХ МОЖНО ВЫБРАТЬ ВИДИМОСТЬ ТАЙМЕРА СЛЕВА СНИЗУ И ПРИДУМАЙ ЕЩЁ ПАРУ ШТУК (не таймеров, а фич)
            timersec_toolStripTextBox1.ForeColor = Color.Gray;

            label3.Hide();

#if DEBUG
            TopMost = false;
#endif
        }
        Stopwatch sw = new Stopwatch();
        
        //int milsek = 10000; 
        //int time = 0;
        int sec = 5000;

        // Ниже находится тестовая кнопка Олега

        /*private void button3_Click(object sender, EventArgs e)
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
        } */

        private static Bitmap ResetResolution(Metafile mf, float resolution) // Тут удалить надо создание потока или сделать чтоб статик был в классе
        {
            Thread drawimg = new Thread(() =>
            {
                Action draw = () =>
                {
                    int width = (int)(mf.Width * resolution / mf.HorizontalResolution);
                    int height = (int)(mf.Height * resolution / mf.VerticalResolution);
                    Bitmap bmp = new Bitmap(width, height);
                    bmp.SetResolution(resolution, resolution);
                    Graphics g = Graphics.FromImage(bmp);
                    g.DrawImage(mf, 0, 0);
                    g.Dispose();
                    return bmp;
                };
                if (InvokeRequired)
                    Invoke(draw);
                else
                    draw;
            };
            drawimg.Start();
        }

        public static void FindImg()
        {
            //
        }

        public static void Sleep(int sec) // Метод вместо таймера
        {
            Thread.Sleep(sec);
        }

        private void timer1_Tick(object sender, EventArgs e) // Таймер
        {
            //time++;
            //label3.Text = Convert.ToString(time);
        }

        private void открытьToolStripMenuItem_Click(object sender, EventArgs e) // Открывание файла первого листа
        {
            //CheckForIllegalCrossThreadCalls = false; // Это отключает перехват ошибки (лайфхак, но не очень хороший), потому что при тестировании кода он не будет ошибку показывать среди кучи кода и ты будешь сам тупить сидеть з:
            Thread first = new Thread(() => // Это и есть открывание нового потока
                // Делегат это ссылка на метод (delegate), но его можно заменить на лямбда оператор, так короче и мужик сказал, что так лучше будет :DDD
                {
                    Action firstlist = () =>
                    {
                        try
                        {
                            sw.Start();
                            try
                            {
                                sec = Convert.ToInt32(timersec_toolStripTextBox1.Text) * 1000;
                                MessageBox.Show(
                                    "Переодичность может быть некорректной на +-2 секунды, т.к. программа возобновляет работу только по истечению таймера, а для этого нужна пара секунд!\nОбратите внимание, первые пару слайдов программа запускает долго, т.к. генерирует изображения",
                                    "Уведомление",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            catch
                            {
                                MessageBox.Show(
                                    "Периодичность не задана, будет использовано значение по умолчанию (5 сек)\nПереодичность может быть некорректной на +-2 секунды, т.к. программа возобновляет работу только по истечению таймера, а для этого нужна пара секунд!\nОбратите внимание, первые пару слайдов программа запускает долго, т.к. генерирует изображения",
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
                                "Не удалось открыть файл/Файл не был выбран",
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
                                        Sleep(sec);
                                        //images.Save(@"C:\Users\Oleg_\OneDrive\Документы\Отчеты\Result" + i + ".jpg", ImageFormat.Jpeg);
                                        pictureBox1.Image = images;
                                        GC.Collect();
                                    }
                                    firstR += 17;
                                    LastR += 17;

                                }
                                GC.Collect(); // убрать за собой
                            }
                            sw.Stop();
                            label4.Text = sw.Elapsed.ToString();
                        }
                        catch
                        {
                            MessageBox.Show("Произошла критическая ошибка!", "Ошибка!");
                        }
                    };
                    // Можно конечно и без IF сделать Invoke((Action) (() => { код }));
                    // А ещё можно вовсе без делегата Action через Invoke((MethodInvoker) (() => { code })); это вообще феншуй феншуйский по феншую
                    if (InvokeRequired) // Это крч хрень какая-то которая проверяет, чтоб поток нормально запускался (хз, сам загугли я не знаю как объяснить :DDD)
                        Invoke(firstlist);
                    else
                        firstlist();
                });
            first.Start(); // Здесь начинается работа потока считывающая код выше
        }

        private void выйтиToolStripMenuItem_Click(object sender, EventArgs e) // Выход из приложения
        {
            this.Close();
        }

        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e) // О программе
        {
            MessageBox.Show("Приложение ExcelOp разработали студенты НГПУ им. Козьмы Минина группы ИСТ-20 Бижко Виталий и Задонский Олег", "Версия: 1.3.2");
        }

        private void timersec_toolStripTextBox1_Click(object sender, EventArgs e) // Заданные секунды таймера
        {
            if (timersec_toolStripTextBox1.Text == "Введите секунды")
            {
                timersec_toolStripTextBox1.Text = "";
                timersec_toolStripTextBox1.ForeColor = Color.Black;
            }
        }

        private void label4_Click(object sender, EventArgs e) // Это стопвотч
        {

        }

        private void таймерToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (timersec_toolStripTextBox1.Text == "")
            {
                timersec_toolStripTextBox1.Text = "Введите секунды";
                timersec_toolStripTextBox1.ForeColor = Color.Gray;
            }
        }

        private void спрятатьТаймерВУглуToolStripMenuItem_Click(object sender, EventArgs e) // Спрятать таймер
        {
            if (спрятатьТаймерВУглуToolStripMenuItem.Text == "Спрятать таймер в углу")
            {
                label4.Hide();
                спрятатьТаймерВУглуToolStripMenuItem.Text = "Показать таймер в углу";
            }
            else
            {
                label4.Show();
                спрятатьТаймерВУглуToolStripMenuItem.Text = "Спрятать таймер в углу";
            }
        }

        private void открытьПервыйЛистToolStripMenuItem_Click(object sender, EventArgs e) // Открывание файла второго листа
        {
            try
            {
                //timer1.Enabled = true;
                sw.Start();
                try
                {
                    //milsek = Convert.ToInt32(textBox1.Text) * 1000;
                    sec = Convert.ToInt32(timersec_toolStripTextBox1.Text) * 1000;
                    MessageBox.Show(
                    "Переодичность может быть некорректной на +-2 секунды, т.к. программа возобновляет работу только по истечению таймера, а для этого нужна пара секунд!\nОбратите внимание, первые пару слайдов программа запускает долго, т.к. генерирует изображения",
                    "Уведомление",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch
                {
                    MessageBox.Show(
                    "Периодичность не задана, будет использовано значение по умолчанию (5 сек)\nПереодичность может быть некорректной на +-2 секунды, т.к. программа возобновляет работу только по истечению таймера, а для этого нужна пара секунд!\nОбратите внимание, первые пару слайдов программа запускает долго, т.к. генерирует изображения",
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
                    "Не удалось открыть файл/Файл не был выбран",
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
                                images.Save(@"C:\Users\Public\Documents" + i + ".jpg", ImageFormat.Jpeg);
                                pictureBox1.Image = images;
                                GC.Collect();
                            }
                            switch (i)
                            {
                                case 1:
                                    //time = 0;
                                    LastC += 1;
                                    firstR += 16;
                                    LastR += 16;
                                    Sleep(sec);
                                    break;

                                case 2:
                                    //time = 0;
                                    firstR += 15;
                                    LastR += 15;
                                    Sleep(sec);
                                    break;

                                case 3:
                                    //time = 0;
                                    firstR += 14;
                                    LastR += 14;
                                    Sleep(sec);
                                    break;

                                case 4:
                                    //time = 0;
                                    firstR += 14;
                                    LastR += 14;
                                    Sleep(sec);
                                    break;

                                case 5:
                                    //time = 0;
                                    firstR += 7;
                                    LastR += 7;
                                    Sleep(sec);
                                    break;

                                case 6:
                                    //time = 0;
                                    firstR += 14;
                                    LastR += 14;
                                    Sleep(sec);
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
                sw.Stop();
                label4.Text = sw.Elapsed.ToString();
            }
            catch
            {
                MessageBox.Show("Произошла критическая ошибка!", "Ошибка!");
            }
        }
    }
}

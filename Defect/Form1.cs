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

namespace Pipeline
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            //Выбор файлов Excel
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            string filename1 = openFileDialog1.FileName;

            //Инициализация Excel
            Excel.Application firstApp;
            Excel.Workbook firstWorkbook;
            Excel.Worksheet firstWorksheet;
            Excel.Range firstRange;

            firstApp = new Excel.Application();
            firstWorkbook = firstApp.Workbooks.Open(filename1);
            firstWorksheet = firstWorkbook.Worksheets[1];
            firstRange = firstWorksheet.UsedRange;

            //Импорт данных
            var defects = new List<Defect>();
            try
            {
                for (int i = 2; i <= firstRange.Rows.Count; i++)
                {
                    var temp1 = new Defect();
                    temp1.Id = (int)firstRange.Cells[i, 1].Value;
                    temp1.Name = firstRange.Cells[i, 2].Value;
                    temp1.LengthCoordinate = firstRange.Cells[i, 3].Value;
                    temp1.HeightCoordinate = DateTime.Parse(firstRange.Cells[i, 4].Text);
                    temp1.Diameter = (int)firstRange.Cells[i, 5].Value;
                    temp1.SectionLength = firstRange.Cells[i, 6].Value;
                    defects.Add(temp1);
                }
            }
            catch
            {
                MessageBox.Show("Проверьте корректность данных в файле!\n" +
                    "Значения должны начинаться со второй строки.\n" +
                    "Порядок значений: id, Наименование, Расстояние от поперечного сварного шва, м, Местоположение дефекта в часовой системе (в виде чч:мм), Диаметр, мм, Длина секции, м.\n" +
                    "Ошибка считывания данных");
                Application.Exit();
            }
            CheckParameters(defects);
            PrintGrigView(defects);

            firstWorkbook.Close(false);
            firstApp.Quit();
            firstApp = null;
            firstWorkbook = null;
            firstWorksheet = null;
            firstRange = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();

        }



        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Graphics g = pictureBox1.CreateGraphics();
            g.Clear(Color.White);
            if (e.RowIndex>=0)
            {
                DataGridViewRow row = this.dataGridView1.Rows[e.RowIndex];

                var time = Convert.ToDateTime(row.Cells[3].Value);
                var alpha = time.Hour * 30 + time.Minute * 0.5;
                int h = Convert.ToInt32(row.Cells[4].Value);
                int s = Convert.ToInt32(row.Cells[5].Value)*1000;
                double heightDefect = (Math.PI * (h/2) * alpha) / 180;
                double lengthDefect = Convert.ToDouble(row.Cells[2].Value) * 1000;


                //Рисование развертки и местоположения дефекта
                g.DrawRectangle(new Pen(Brushes.Black, 2), 10, 10, s/20, (int)(Math.PI*h/29.73));
                g.DrawEllipse(new Pen(Brushes.Red, 2), 10 + (int)(lengthDefect/20), 10 + (int)(heightDefect / 29.73), 10, 10);
                
            }
        }

        //Вывод в DataGridView
        private void PrintGrigView(List<Defect> defects)
        {
            string id;
            string Name;
            string Length;
            string Height;
            string Diameter;
            string Section;

            foreach (Defect d in defects)
            {
                id = d.Id == 0 ? "" : d.Id.ToString();
                Name = d.Name;
                Length = d.LengthCoordinate.ToString();
                Height = d.HeightCoordinate.ToShortTimeString().ToString();
                Diameter = d.Diameter.ToString();
                Section = d.SectionLength.ToString();
                dataGridView1.Rows.Add(id, Name, Length, Height, Diameter, Section);
            }
        }

        //Проверка входящих значений
        private void CheckParameters(List<Defect> defects)
        {
            int l = 0;
            int diam = 0;
            int h = 0;
            foreach (Defect d in defects)
            {
                if (d.LengthCoordinate > d.SectionLength)
                {
                    l++;
                }

                if (d.Diameter < 520 || d.Diameter > 1420)
                {
                    diam++;
                }

                if (d.HeightCoordinate.Hour > 12)
                {
                    h++;
                }

                if(l>0 || diam>0 || h>0)
                {
                    MessageBox.Show("Проверьте указанные параметры:\n " +
                        "Расстояние от сварного шва до дефекта не может быть больше длины секции.\n" +
                        "Местоположение дефекта указывается в виде чч:мм и не может быть более 12.\n" +
                        "Диаметр трубопровода должен находиться в границах 520-1420 мм.");
                    Application.Exit();
                }
            }
        }
    }
}

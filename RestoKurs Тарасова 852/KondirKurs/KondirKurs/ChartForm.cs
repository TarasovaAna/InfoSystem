using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms.DataVisualization.Charting;
using System.Windows.Forms;

namespace KondirKurs
{
    public partial class ChartForm : Form
    {
        KondirDataSet.Report2DataTable data;
        string chartTitle;
        public ChartForm(KondirDataSet.Report2DataTable dt, string s)
        {
            InitializeComponent();
            data = dt;
            chartTitle = s;
        }
        private void ChartForm_Load(object sender, EventArgs e)
        {
            //Создаем DataView из таблицы данных
            DataView reportWiew = new DataView(data);
            //Устанавливаем ее в качестве источника данных
            chart1.Series[0].Points.DataBindXY(reportWiew, "SupplierName", reportWiew, "Total");
            //Вводим текст заголовка диаграммы
            chart1.Titles[0].Text = chartTitle;
        }
        public void CopyChartToClipboard()
        {
            //Копирование диграммы в буфер обмена
            //Объект Chart позволяет сохранение в файл в различных графических форматах
            //Чтобы не сохранять на диск воспользуемся MemoryStream
            using (MemoryStream ms = new MemoryStream())
            {
                chart1.SaveImage(ms, ChartImageFormat.Bmp);//сохраняем
                Bitmap bm = new Bitmap(ms);//считываем в переменную
                Clipboard.SetImage(bm);//копируем в буфер обмена
            }
        }
    }
}

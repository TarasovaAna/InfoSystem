using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Core;
using System.Windows.Forms;

namespace KondirKurs
{
    public partial class ReportForm : Form
    {
        public ReportForm()
        {
            InitializeComponent();
        }
        private void ReportForm_Load(object sender, EventArgs e)
        {
            DateTime d = DateTime.Today;
            dateTimePicker2.Value = d;
            dateTimePicker1.Value = new DateTime(d.Year, d.Month, 1);
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                report1TableAdapter1.Fill(kondirDataSet1.Report1, dateTimePicker1.Value, dateTimePicker2.Value);
                dataGridView1.DataSource = kondirDataSet1.Report1;
            }
            if (radioButton2.Checked)
            {
                report2TableAdapter1.Fill(kondirDataSet1.Report2, dateTimePicker1.Value, dateTimePicker2.Value);
                dataGridView1.DataSource = kondirDataSet1.Report2;
                ShowChart();
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (radioButton1.Checked)
                {
                    ReportExcel1();
                }
                if (radioButton2.Checked)
                {
                    ReportExcel2();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                if (radioButton1.Checked)
                {
                    ReportWord1();
                }
                if (radioButton2.Checked)
                {
                    ReportWord2();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void ReportExcel1()
        {
            Excel.Application app = new Excel.Application();
            app.Visible = true;
            Excel.Workbook wb = app.Workbooks.Add();
            Excel.Worksheet ws = wb.Worksheets[1];
            ws.Columns[1].ColumnWidth = 20;
            ws.Columns[2].ColumnWidth = 40;
            ws.Columns[3].ColumnWidth = 40;
            ws.Columns[4].ColumnWidth = 10;
            ws.Columns[5].ColumnWidth = 10;
            ws.Columns["A:E"].WrapText = true;
            Excel.Range rng = ws.Range["A1:E1"];
            ws.Cells[1, 1].value = "ОТЧЕТ О ПРОДАЖАХ ЗА " + dateTimePicker1.Value.ToShortDateString() +
            " - " + dateTimePicker2.Value.ToShortDateString();
            rng.Font.Bold = true;
            rng.Font.Size = 14;
            rng.MergeCells = true;
            ws.Range["A1:E2"].HorizontalAlignment =
           Excel.Constants.xlCenter;
            ws.Cells[2, 1].value = "Поставщик";
            ws.Cells[2, 2].value = "Товар";
            ws.Cells[2, 3].value = "Клиент";
            ws.Cells[2, 4].value = "Дата";
            ws.Cells[2, 5].value = "Сумма продаж";
            int i = 3;
            report1TableAdapter1.Fill(kondirDataSet1.Report1,
           dateTimePicker1.Value, dateTimePicker2.Value);
            foreach (KondirDataSet.Report1Row r in kondirDataSet1.Report1)
            {
                ws.Cells[i, 1].value = r.SupplierName;
                ws.Cells[i, 2].value = r.ProductName;
                ws.Cells[i, 3].value = r.ClientName;
                ws.Cells[i, 4].value = r.SaleDate;
                ws.Cells[i, 5].value = r.Total;
                i++;
            }
            ws.Cells[i, 1].value = "ИТОГО";
            ws.Range[ws.Cells[i, 1], ws.Cells[i, 3]].MergeCells =
           true;
            ws.Cells[i, 4].FormulaR1C1 = "=SUM(R[-" + (i - 3) +
           "]C:R[-1]C)";
            rng = ws.Range[ws.Cells[2, 1], ws.Cells[i, 4]];
            rng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
           Excel.XlLineStyle.xlContinuous;
            rng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle =
           Excel.XlLineStyle.xlContinuous;
            rng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
           Excel.XlLineStyle.xlContinuous;
            rng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
           Excel.XlLineStyle.xlContinuous;
            rng.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle =
            Excel.XlLineStyle.xlContinuous;
            rng.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
            Excel.XlLineStyle.xlContinuous;
        }
        private void ReportExcel2()
        {
            Excel.Application app = new Excel.Application();
            app.Visible = true;
            Excel.Workbook wb = app.Workbooks.Add();
            Excel.Worksheet ws = wb.Worksheets[1];
            ws.Columns[1].ColumnWidth = 40;
            ws.Columns[2].ColumnWidth = 20;
            ws.Columns["A:B"].WrapText = true;
            Excel.Range rng = ws.Range["A1:B1"];
            ws.Cells[1, 1].value = "ОТЧЕТ О ПРОДАЖАХ ЗА " +
           dateTimePicker1.Value.ToShortDateString() +
            " - " + dateTimePicker2.Value.ToShortDateString();
            rng.Font.Bold = true;
            rng.Font.Size = 14;
            rng.MergeCells = true;
            ws.Range["A1:B2"].HorizontalAlignment =
           Excel.Constants.xlCenter;
            ws.Cells[2, 1].value = "Поставщик";
            ws.Cells[2, 2].value = "Сумма продаж";
            int i = 3;
            report2TableAdapter1.Fill(kondirDataSet1.Report2,
           dateTimePicker1.Value, dateTimePicker2.Value);
            foreach (KondirDataSet.Report2Row r in kondirDataSet1.Report2)
            {
                ws.Cells[i, 1].value = r.SupplierName;
                ws.Cells[i, 2].value = r.Total;
                i++;
            }
            ws.Cells[i, 1].value = "ИТОГО";
            ws.Cells[i, 2].FormulaR1C1 = "=SUM(R[-" + (i - 3) + "]C:R[-1]C)";
            rng = ws.Range[ws.Cells[2, 1], ws.Cells[i, 2]];
            rng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
           Excel.XlLineStyle.xlContinuous;
            rng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle =
           Excel.XlLineStyle.xlContinuous;
            rng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
           Excel.XlLineStyle.xlContinuous;
            rng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
           Excel.XlLineStyle.xlContinuous;
            rng.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle =
            Excel.XlLineStyle.xlContinuous;
            rng.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
            Excel.XlLineStyle.xlContinuous;
            Excel.Shape sh =
           ws.Shapes.AddChart(Excel.XlChartType.xlPie);
            sh.Left = 0;
            sh.Top = (float)ws.Rows[i + 2].Top;
            float scale = (float)ws.Columns[3].Left / sh.Width;
            sh.ScaleWidth(scale, MsoTriState.msoFalse,
           MsoScaleFrom.msoScaleFromTopLeft);
            Excel.Chart ch = sh.Chart;
            ch.SetSourceData(ws.Range[ws.Cells[2, 1], ws.Cells[i - 1,
           2]]);
            ch.SeriesCollection(1).ApplyDataLabels();
        }
        private void ReportWord1()
        {
            Word.Application app = new Word.Application();
            app.Visible = true;
            Word.Document doc = app.Documents.Add();
            app.Selection.Font.Name = "Times New Roman";
            app.Selection.Font.Size = 14;
            app.Selection.Font.Bold = 1;
            app.Selection.TypeText("ОТЧЕТ О ПРОДАЖАХ ЗА " +
           dateTimePicker1.Value.ToShortDateString() +
            " - " + dateTimePicker2.Value.ToShortDateString());
            app.Selection.ParagraphFormat.Alignment =
           Word.WdParagraphAlignment.wdAlignParagraphCenter;
            app.Selection.TypeParagraph();
            app.Selection.ParagraphFormat.Alignment =
           Word.WdParagraphAlignment.wdAlignParagraphLeft;
            app.Selection.Font.Bold = 0;
            Word.Table t = doc.Tables.Add(app.Selection.Range, 1, 5);
            Object style = "Сетка таблицы";
            t.set_Style(style);
            Object unit = Word.WdUnits.wdCell;
            app.Selection.TypeText("Производитель");
            app.Selection.MoveRight(unit);
            app.Selection.TypeText("Товар");
            app.Selection.MoveRight(unit);
            app.Selection.TypeText("Клиент");
            app.Selection.MoveRight(unit);
            app.Selection.TypeText("Дата");
            app.Selection.MoveRight(unit);
            app.Selection.TypeText("Сумма продаж");
            report1TableAdapter1.Fill(kondirDataSet1.Report1,
           dateTimePicker1.Value, dateTimePicker2.Value);
            foreach (KondirDataSet.Report1Row r in kondirDataSet1.Report1)
            {
                app.Selection.MoveRight(unit);
                app.Selection.TypeText(r.SupplierName);
                app.Selection.MoveRight(unit);
                app.Selection.TypeText(r.ProductName);
                app.Selection.MoveRight(unit);
                app.Selection.TypeText(r.ClientName);
                app.Selection.MoveRight(unit);
                app.Selection.TypeText(r.SaleDate.ToShortDateString());
                app.Selection.MoveRight(unit);
                app.Selection.TypeText(r.Total.ToString());
            }
        }
        private void ReportWord2()
        {
            Word.Application app = new Word.Application();
            app.Visible = true;
            Word.Document doc = app.Documents.Add(Template:
           Application.StartupPath + "\\SalesReport.dotx.docx");
            app.Selection.GoTo(What: Word.WdGoToItem.wdGoToBookmark,
           Name: "date1");

            app.Selection.TypeText(dateTimePicker1.Value.ToShortDateString());
            app.Selection.GoTo(What: Word.WdGoToItem.wdGoToBookmark,
           Name: "date2");

            app.Selection.TypeText(dateTimePicker2.Value.ToShortDateString());
            app.Selection.GoTo(What: Word.WdGoToItem.wdGoToBookmark,
           Name: "sales");
            app.Selection.Range.ListFormat.ApplyListTemplateWithLevel(

           app.ListGalleries[Word.WdListGalleryType.wdNumberGallery].ListTemplates[1]);
            report2TableAdapter1.Fill(kondirDataSet1.Report2,
           dateTimePicker1.Value, dateTimePicker2.Value);
            foreach (KondirDataSet.Report2Row r in kondirDataSet1.Report2)
            {
                app.Selection.TypeText(r.SupplierName + ": " + r.Total + " руб.");
                app.Selection.TypeParagraph();
            }
            app.Selection.Range.ListFormat.RemoveNumbers();
            //Переходим к закладке с именем chart для вставки диаграммы
            app.Selection.GoTo(What: Word.WdGoToItem.wdGoToBookmark, Name: "chart");
            //Создаем диаграмму, копируем ее в буфер обмена, вставляем в Word и закрываем форму с диаграммой.
            ChartForm f = ShowChart();
            f.CopyChartToClipboard();
            f.Close();
            app.Selection.Paste();
            app.Selection.GoTo(What: Word.WdGoToItem.wdGoToBookmark, Name: "date3");
            app.Selection.TypeText(DateTime.Today.ToShortDateString());
        }
        private ChartForm ShowChart()
        {
            //В конструкторе передаем таблицу данных и заголовок диаграммы
            ChartForm f = new ChartForm(kondirDataSet1.Report2, "ОТЧЕТ О ПРОДАЖАХ ЗА " + dateTimePicker1.Value.ToShortDateString() + " - " + dateTimePicker2.Value.ToShortDateString());
            f.Show();
            //Располагаем левый верхний угол формы с диаграммой по центру данной формы
            f.Location = new Point(Location.X + Size.Width / 2, Location.Y + Size.Height / 2);
            return f;
        }


        

        
    }
}

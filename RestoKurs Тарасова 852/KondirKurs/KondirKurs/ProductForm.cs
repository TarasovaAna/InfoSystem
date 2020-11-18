using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text;
using System.Windows.Forms;

namespace KondirKurs
{
    public partial class ProductForm : Form
    {
        public ProductForm()
        {
            InitializeComponent();
        }

        private void supplierBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.supplierBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.kondirDataSet);

        }

        private void ProductForm_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "kondirDataSet.Product". При необходимости она может быть перемещена или удалена.
            this.productTableAdapter.Fill(this.kondirDataSet.Product);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "kondirDataSet.Supplier". При необходимости она может быть перемещена или удалена.
            this.supplierTableAdapter.Fill(this.kondirDataSet.Supplier);

        }

        private void supplierBindingSource_DataError(object sender, BindingManagerDataErrorEventArgs e)
        {
            MessageBox.Show("Ошибка в данных");
        }
        private void productBindingSource_DataError(object sender, BindingManagerDataErrorEventArgs e)
        {
            MessageBox.Show("Ошибка в данных");
        }
        private void supplierDataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            MessageBox.Show("Ошибка в данных");
        }
        private void productDataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            MessageBox.Show("Ошибка в данных");
        }
        private void productDataGridView_Enter(object sender, EventArgs e)
        {
            if (kondirDataSet.HasChanges(DataRowState.Added))
            {
                DataRowView dr =
               (DataRowView)supplierBindingSource.Current;
                if (dr != null)
                {
                    string st = dr["SupplierName"] as string;
                    if (st != null)
                    {
                        Validate();
                        supplierBindingSource.EndEdit();
                        tableAdapterManager.UpdateAll(this.kondirDataSet);
                        supplierTableAdapter.Fill(this.kondirDataSet.Supplier);
                        int pos = this.supplierBindingSource.Find("SupplierName", st);
                        supplierBindingSource.Position = pos;
                    }
                }
            }
        }
        private DataTable ReadExcelFile(string fileName)
        {
            Excel.Application app = new Excel.Application();
            app.Visible = true;
            Excel.Workbook wb = app.Workbooks.Open(fileName);
            Excel.Worksheet ws = wb.Worksheets[1];
            Excel.Range rng = ws.Range["A1"].CurrentRegion;
            int rowsCount = rng.Rows.Count;
            DataTable table = new DataTable();
            table.Columns.Add("SupplierName", typeof(string));
            table.Columns.Add("ProductName", typeof(string));
            table.Columns.Add("Price", typeof(decimal));
            for (int i = 2, j = 0; i <= rowsCount; i++, j++)
            {
                DataRow r = table.NewRow();
                r["SupplierName"] = rng.Cells[i, 1].Value;
                r["ProductName"] = rng.Cells[i, 2].Value;
                r["Price"] = rng.Cells[i, 3].Value;
                table.Rows.Add(r);
            }
            wb.Close(false);
            app.Quit();
            return table;
        }
        private void InsertData(DataTable table)
        {
            var manufacturers = table.AsEnumerable().
            Select(v => v.Field<string>("SupplierName")).
            //из добавляемых данных выбираем производителей
            Distinct(StringComparer.CurrentCultureIgnoreCase).
            //без повторений независимо от регистра 

            Except(kondirDataSet.Tables["Supplier"].AsEnumerable().
            Select(v => v.Field<string>("SupplierName")),
            //за исключением уже имеющихся в таблице производителей
            StringComparer.CurrentCultureIgnoreCase);
            //независимо от регистра
            //добавление новых производителей в таблицу производителей
            foreach (var m in manufacturers)
            {
                DataRow r =
               kondirDataSet.Tables["Supplier"].NewRow();
                r["SupplierName"] = m;
                kondirDataSet.Tables["Supplier"].Rows.Add(r);
            }
            tableAdapterManager.UpdateAll(kondirDataSet);//вставка производителей в БД
            supplierTableAdapter.Fill(this.kondirDataSet.Supplier);
            var products = from np in table.AsEnumerable()
                           join m in
                          kondirDataSet.Tables["Supplier"].AsEnumerable()
                           on
                          np.Field<string>("SupplierName").ToUpper()
                           equals
                          m.Field<string>("SupplierName").ToUpper()
                           join p in
                          kondirDataSet.Tables["Product"].AsEnumerable()
                           on new
                           {
                               f1 =
                          m.Field<int>("SupplierID"),
                               f2 =
                          np.Field<string>("ProductName").ToUpper(),
                               f3 = np.Field<decimal>("Price")
                           }
                          equals new
                          {
                              f1 =
                          p.Field<int>("SupplierID"),
                              f2 =
                          p.Field<string>("ProductName").ToUpper(),
                              f3 = p.Field<decimal>("Price")
                          }
                           into pg
                           where !pg.Any()
                           select
                           new
                           {
                               SupplierID =
                           m.Field<int>("SupplierID"),
                               ProductName =
                           np.Field<string>("ProductName"),
                               Price = np.Field<decimal>("Price")
                           }
            into np1
                           group np1 by new
                           {
                               f1 = np1.SupplierID,
                               f2 = np1.ProductName.ToUpper(),
                               np1.Price
                           }
            into np2
                           select
                          new
                          {
                              np2.First().SupplierID,
                              np2.First().ProductName,
                              np2.First().Price
                          };
            //вставка новых товаров в таблицу товаров
            foreach (var p in products)
            {
                DataRow r = kondirDataSet.Tables["Product"].NewRow();
                r["SupplierID"] = p.SupplierID;
                r["ProductName"] = p.ProductName;
                r["Price"] = p.Price;
                kondirDataSet.Tables["Product"].Rows.Add(r);
            }
            tableAdapterManager.UpdateAll(kondirDataSet);
        }
    }
}

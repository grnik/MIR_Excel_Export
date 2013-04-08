using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelToXML
{
    public partial class Main : Form
    {
        private Excel.Application excelapp;
        private Excel.Window excelWindow;

        private Excel.Workbooks excelappworkbooks;
        private Excel.Workbook excelappworkbook;

        private Excel.Sheets excelsheets;
        private Excel.Worksheet excelworksheet;
        private Excel.Range excelcells;


        public Main()
        {
            InitializeComponent();
        }

        private void btTransfer_Click(object sender, EventArgs e)
        {
            excelapp = new Excel.Application();
            excelapp.Visible = true;
            try
            {
                excelapp.Workbooks.Open(txtExcelFile.Text,
                      Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                      Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                      Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                      Type.Missing, Type.Missing);

                //Получаем набор ссылок на объекты Workbook (на созданные книги)
                excelappworkbooks = excelapp.Workbooks;
                //Получаем ссылку на книгу 1 - нумерация от 1
                excelappworkbook = excelappworkbooks[1];
                //Ссылку можно получить и так, но тогда надо знать имена книг,
                //причем, после сохранения - знать расширение файла
                //excelappworkbook=excelappworkbooks["Книга 1"];

                excelsheets = excelappworkbook.Worksheets;
                //Получаем ссылку на лист 1
                excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);
                //Выделение группы ячеек

                excelcells = excelworksheet.get_Range("A1", Type.Missing);
                string sStr = Convert.ToString(excelcells.Value2);

                sStr = excelworksheet.Cells[1, 1].Value2.ToString();
            }
            catch (Exception)
            {

                throw;
            }
            finally
            {
                excelapp.Quit();
            }
        }

        private void btExcelChoose_Click(object sender, EventArgs e)
        {
            if (ofdExcelFile.ShowDialog() == DialogResult.OK)
            {
                txtExcelFile.Text = ofdExcelFile.FileName;
            }
        }

        private void btXMLChoose_Click(object sender, EventArgs e)
        {
            if (sfdXMLFile.ShowDialog() == DialogResult.OK)
            {
                txtXMLFile.Text = sfdXMLFile.FileName;
            }
        }
    }
}

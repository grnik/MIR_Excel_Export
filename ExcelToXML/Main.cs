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
    using System.IO;

    public partial class Main : Form
    {
        const string channel = "MIR";

        private Excel.Application excelapp;
        private Excel.Window excelWindow;

        private Excel.Workbooks excelappworkbooks;
        private Excel.Workbook excelappworkbook;

        private Excel.Sheets excelsheets;
        private Excel.Worksheet excelworksheet;
        private Excel.Range excelcells;

        private StreamWriter file;

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

                this.CreateFile(txtXMLFile.Text);
                int row = 1;
                try
                {
                    DateTime date;
                    DateTime start = new DateTime(0);
                    DateTime stop = new DateTime(0);

                    while (excelworksheet.Cells[row, 1].Value2 != null)
                    {
                        string A = ReadCell("A", row);
                        string B = ReadCell("B", row);
                        string C = ReadCell("C", row);
                        date = Convert.ToDateTime(A);
                        //Если время пусто, то берем время предыдущей строки.
                        if (!String.IsNullOrEmpty(B))
                            start = date + DoubleToTimeSpan(Convert.ToDouble(B));
                        if (!String.IsNullOrEmpty(C))
                            stop = start + DoubleToTimeSpan(Convert.ToDouble(C));

                        string title = ReadCell(5, row);
                        string rating = "";
                        string desc = ReadCell(8, row);

                        AddProgramm(start, stop, channel, title, date, rating, desc);

                        row++;
                    }
                    MessageBox.Show("Было обработано " + (row - 1).ToString() + " строк");
                }
                catch (Exception exc)
                {
                    MessageBox.Show(
                        "При обработке файла возникла ошибка: " + exc.Message + " в строке " + row.ToString(),
                        "Конвертация файла",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
                finally
                {
                    this.CloseFile();
                }
            }
            catch (Exception exc)
            {
                MessageBox.Show(
                    "При открытии файла возникла ошибка: " + exc.Message,
                    "Конвертация файла",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                excelapp.Quit();
            }
        }

        private string ReadCell(int col, int row)
        {
            string res = "";
            excelcells = excelworksheet.Cells[row, col];
            if (excelcells.Value2 != null)
                res = excelworksheet.Cells[row, col].Value2.ToString();

            return res;
        }

        private string ReadCell(string col, int row)
        {
            string res = "";
            excelcells = excelworksheet.get_Range(col + row.ToString(), Type.Missing);
            if (excelcells.Value2 != null)
                res = excelcells.Value2.ToString();

            return res;
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

        void CreateFile(string fileName)
        {
            file = File.CreateText(fileName);
            file.WriteLine(@"<?xml version=""1.0"" encoding=""UTF-8"" ?>");
            file.WriteLine(@"<tv>");
            file.WriteLine(@"<channel id=""" + channel + @""">");
            file.WriteLine(@"<display-name lang=""ru"">МИР</display-name>");
            file.WriteLine(@"</channel>");

        }

        void CloseFile()
        {
            file.WriteLine(@"</tv>");
            file.Flush();
            file.Close();
        }

        void AddProgramm(DateTime start, DateTime stop, string channel, string title, DateTime date, string rating, string desc)
        {
            StringBuilder programm = new StringBuilder();
            programm.AppendFormat(@"<programme start=""{0} +0400"" stop=""{1} +0400"" channel=""{2}"">", start.ToString("yyyyMMddHHmmss"), stop.ToString("yyyyMMddHHmmss"), channel);
            programm.AppendFormat(@"<title lang=""ru"">{0}</title>", title);
            programm.AppendFormat(@"<date>{0}</date>", date.ToString("yyyyMMdd"));
            programm.AppendFormat(@"<rating system=""RU""><value>{0}</value></rating>", rating);
            programm.AppendFormat(@"<desc lang=""ru"">{0}</desc>", desc);
            programm.Append(@"</programme>");

            file.WriteLine(programm.ToString());
        }

        TimeSpan DoubleToTimeSpan(double tm)
        {
            tm = tm * 24;
            int hour = (int)Math.Truncate(tm);

            tm = (tm - hour) * 60;
            int minut = (int)Math.Truncate(tm);

            tm = (tm - minut) * 60;
            int second = (int)Math.Truncate(tm);

            return new TimeSpan(hour, minut, second);
        }
    }
}

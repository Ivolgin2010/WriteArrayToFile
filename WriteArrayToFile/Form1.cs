using System;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace WriteArrayToFile
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        // сохраняем данные в файл
        private void BtnSave_Click(object sender, EventArgs e)
        {

            string fileTest = @" C:\Users\i.geraskin\source\repos\CalcData.csv";



            //Объявляем приложение
            Excel.Application ex = new Microsoft.Office.Interop.Excel.Application
            {
                //Отобразить Excel
                Visible = true,

                //Количество листов в рабочей книге
                SheetsInNewWorkbook = 2
            };

            //Добавить рабочую книгу
            Excel.Workbook workBook = ex.Workbooks.Add(Type.Missing);

            //Отключить отображение окон с сообщениями
            ex.DisplayAlerts = false;

            //Получаем первый лист документа (счет начинается с 1)
            Excel.Worksheet sheet = (Excel.Worksheet)ex.Worksheets.get_Item(1);

            //Название листа (вкладки снизу)
            sheet.Name = "Данные";

            //Пример заполнения ячеек
            for (int i = 1; i <= 10; i++)
            {
                sheet.Cells[i, 1] = String.Format("Value {0}", i);
            }

            for (int i = 1; i <= 10; i++)
            {
                sheet.Cells[i, 2] = String.Format("Value {0}", i);
            }

            for (int i = 1; i <= 10; i++)
            {
                sheet.Cells[i, 3] = String.Format("Value {0}", i);
            }

            for (int i = 1; i <= 10; i++)
            {
                sheet.Cells[i, 4] = String.Format("Value {0}", i);
            }

            for (int i = 1; i <= 10; i++)
            {
                sheet.Cells[i, 5] = String.Format("Value {0}", i);
            }

            // добавляем пустую строку
            Excel.Range cellRange = (Excel.Range)sheet.Cells[1, 1];
            Excel.Range rowRange = cellRange.EntireRow;
            rowRange.Insert(Excel.XlInsertShiftDirection.xlShiftDown, false);

            // создаем заголовки
            sheet.Cells[1, 1] = String.Format("Depth");
            sheet.Cells[1, 2] = String.Format("Paraffins");
            sheet.Cells[1, 3] = String.Format("Nom. debit");
            sheet.Cells[1, 4] = String.Format("Temp_oil");
            sheet.Cells[1, 5] = String.Format("Temp_wire");

            MessageBox.Show("Данные сохранены", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);

            sheet.SaveAs(fileTest);


            //Excel.Application oApp;
            //Excel.Application oSheet;
            //Excel.Application oBook;

            //oApp = new Excel.Application();
            //oBook = oApp.Workbooks.Add();
            //oSheet = (Excel.Worksheet)oBook.Worksheets.get_Item(1);
            //oSheet.Cell[1, 1] = "some value";

            //using (StreamWriter streamWriter = new StreamWriter(fileTest))
            //{

            //}


            //sheet.SaveAs(fileTest);
            //oBook.Close();
            //oApp.Quit();

            //SaveFileDialog savefile = new SaveFileDialog
            //{
            //    // задаем имя файла по умолчанию
            //    FileName = "CalcData.csv",


            //    // фильтруем по типу
            //    Filter = "Excel файл CSV|*.csv|All files (*.*)|*.*"
            //};

            //// открываем диалоговое окно
            //if (savefile.ShowDialog() == DialogResult.OK)
            //{
            //    using (StreamWriter streamWriter = new StreamWriter(savefile.FileName))
            //    {                    

            //        using (var csvWriter = new CsvWriter(streamWriter))
            //        {
            //            // указываем разделитель (каждый заголовок запишется в свою ячейку)
            //            csvWriter.Configuration.Delimiter = ";";

            //            // записываем заголовки
            //            csvWriter.WriteField("Depth");
            //            csvWriter.WriteField("Paraffins");
            //            csvWriter.WriteField("Nom. debit");
            //            csvWriter.WriteField("Temp_oil");
            //            csvWriter.WriteField("Temp_wire");
            //            csvWriter.NextRecord();                          

            //        }
            //    }



            //    //    int.TryParse(textBox1.Text, out int TextBox1);
            //    //    for (int i = 0; i < TextBox1; i++)
            //    //    {
            //    //        // записываем данные в файл
            //    //        sw.WriteLine(i);                

            MessageBox.Show("Данные успешно сохранены", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        

        // закрываем окно
        private void BtnClose_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}

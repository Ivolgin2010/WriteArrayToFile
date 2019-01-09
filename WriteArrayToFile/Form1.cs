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
            //Объявляем приложение
            Excel.Application ex = new Microsoft.Office.Interop.Excel.Application
            {
                //Отобразить Excel
                Visible = false,

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

            int.TryParse(textBox1.Text, out int TextBox1);

            // заполняем первый столбец
            for (int i = 1; i <= TextBox1; i++)
            {
                //sheet.Cells[i, 1] = String.Format("Value {0}", i);
                sheet.Cells[i, 1] = string.Format("{0}",i);
            }

            // заполняем второй столбец
            for (int i = 1; i <= 10; i++)
            {
                sheet.Cells[i, 2] = string.Format("{0}", i);
            }

            // заполняем третий столбец
            for (int i = 1; i <= 10; i++)
            {
                sheet.Cells[i, 3] = string.Format("{0}", i);
            }

            // заполняем четвертый столбец
            for (int i = 1; i <= 10; i++)
            {
                sheet.Cells[i, 4] = string.Format("{0}", i);
            }

            // заполняем пятый столбец
            for (int i = 1; i <= 10; i++)
            {
                sheet.Cells[i, 5] = string.Format("{0}", i);
            }

            // добавляем пустую строку
            Excel.Range cellRange = (Excel.Range)sheet.Cells[1, 1];

            Excel.Range rowRange = cellRange.EntireRow;

            rowRange.Insert(Excel.XlInsertShiftDirection.xlShiftDown, false);

            // создаем заголовки
            sheet.Cells[1, 1] = string.Format("Depth");
            sheet.Cells[1, 2] = string.Format("Paraffins");
            sheet.Cells[1, 3] = string.Format("Nom_debit");
            sheet.Cells[1, 4] = string.Format("Temp_oil");
            sheet.Cells[1, 5] = string.Format("Temp_wire");           

            // сохраняем данные
            // задаем путь к файлу
            string fileTest = @"C:\Users\i.geraskin\source\repos\CalcData.csv";

            // сохраняем данные
            sheet.SaveAs(fileTest);

            // закрываем рабочую книгу
            workBook.Close();

            // закрываем приложение
            ex.Quit();



            //******************************************************************************************
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
            //*****************************************************************************************************************



            // выводим сообщение что все ок
            MessageBox.Show("Данные успешно сохранены", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }        

        // закрываем окно
        private void BtnClose_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}

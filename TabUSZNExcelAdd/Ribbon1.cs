using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;


namespace TabUSZNExcelAdd
{
    public partial class Ribbon1
    {
       
        //public string SelectedTemplatePath { get; private set; }
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }


        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            // Открытие формы выбора шаблона в блоке using для автоматического освобождения ресурсов
            using (Templates form = new Templates())
            {
                // Проверка, что форма была закрыта с результатом OK (пользователь выбрал шаблон)
                if (form.ShowDialog() == DialogResult.OK)
                {
                    // Получение пути к выбранному шаблону из формы
                    string templateFilePath = form.SelectedTemplatePath;
                    // Проверка существования файла шаблона по указанному пути
                    if (!File.Exists(templateFilePath))
                    {
                        // Если файл шаблона не найден, выводится сообщение об ошибке и выполнение прекращается
                        MessageBox.Show($"Файл шаблона не найден: {templateFilePath}");
                        return;
                    }
                    // Получение текущего экземпляра приложения Excel
                    Excel.Application excelApp = Globals.ThisAddIn.Application;
                    // Получение активного листа в текущей книге Excel
                    Excel.Worksheet activeSheet = (Excel.Worksheet)excelApp.ActiveSheet;

                    // Вызов метода для создания листов на основе шаблона в активной книге
                    CreateSheetsFromTemplateInActiveWorkbook2(templateFilePath);
                }
            }
        }
        
        private void CreateSheetsFromTemplateInActiveWorkbook2(string templateFilePath)
        {
            // Получаем текущее приложение Excel
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Workbook activeWorkbook = excelApp.ActiveWorkbook;

            // Открываем шаблон
            Excel.Workbook templateWorkbook = excelApp.Workbooks.Open(templateFilePath);

            try
            {
                // Копируем каждый лист шаблона в активную книгу
                foreach (Excel.Worksheet templateSheet in templateWorkbook.Sheets)
                {
                    // Если в активной книге нет листов, создаем новый
                    if (activeWorkbook.Sheets.Count == 0)
                    {
                        activeWorkbook.Sheets.Add();
                    }

                    // Получаем последний лист в активной книге
                    Excel.Worksheet targetSheet = (Excel.Worksheet)activeWorkbook.Sheets[activeWorkbook.Sheets.Count];

                    // Копируем данные из шаблона в целевой лист
                    //templateSheet.UsedRange.Copy(targetSheet.Range["A1"]);
                    
                    // Копируем данные, ширину столбцов и высоту строк из шаблона в целевой лист
                    templateSheet.UsedRange.EntireRow.Copy(targetSheet.Range["A1"]);
                    templateSheet.UsedRange.EntireColumn.Copy(targetSheet.Range["A1"]);

                    // Устанавливаем имя целевого листа равным имени листа шаблона
                    targetSheet.Name = templateSheet.Name;
                    // Освобождаем COM-объекты для освобождения памяти и избежания блокировок
                    Marshal.ReleaseComObject(templateSheet);
                    

                }
            
            }
            finally
            {
                // Закрываем шаблон без сохранения
                templateWorkbook.Close(false);
                // Освобождаем COM-объект шаблона
                Marshal.ReleaseComObject(templateWorkbook);
            }
        }

    }
}

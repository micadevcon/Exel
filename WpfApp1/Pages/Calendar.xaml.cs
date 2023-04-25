using System;
using System.Windows.Controls;
using System.Windows.Data;
using Excel = Microsoft.Office.Interop.Excel;
namespace WpfApp1.Pages
{
    /// <summary>
    /// Логика взаимодействия для Calendar.xaml
    /// </summary>
    public partial class Calendar : Page
    {
        public Calendar()
        {
            InitializeComponent();
            Excel.Application xlApp = new Excel.Application(); //Excel
            Excel.Workbook xlWB; //рабочая книга откуда будем копировать лист  
            //Excel.Workbook xlNewWB; //рабочая книга, в которую будем вставлять данные
            Excel.Worksheet xlSht; //лист Excel
            //Excel.Worksheet xlNewSht; //лист Excel

            xlWB = xlApp.Workbooks.Open(Environment.CurrentDirectory+"1/1.xmls"); //название файла Excel откуда будем копировать лист
            xlSht = xlWB.Worksheets["Лист"]; //название листа или 1-й лист в книге xlSht = xlWB.Worksheets[1];

            //xlNewWB = xlApp.Workbooks.Add(); //новая рабочая книга, куда будем вставлять данные
            //xlNewSht = xlNewWB.Sheets[1]; //первый лист по порядку - в него будем вставлять данные

            xlSht.Range["A1:A10"].UnMerge();
            xlSht.Range["A1:A10"].Copy(); //копируем диапазон ячеек


            xlSht.Range["B1:B10"].PasteSpecial(Excel.XlPasteType.xlPasteValues);

            //вставить только значения
            //xlSht.Range["D1"].PasteSpecial(Excel.XlPasteType.xlPasteAll); //вставить всё (формулы, форматы и т.д.)
            //xlSht.Range["D1"].PasteSpecial(Excel.XlPasteType.xlPasteFormulas); //вставить только формулы
            //xlSht.Range["D1"].PasteSpecial(Excel.XlPasteType.xlPasteFormats); //вставить только форматирование (заливка, граница, форматы ячеек и т.д.

            //xlApp.Visible = true; //отображаем Excel
            xlWB.Close(true); //true - сохранить изменения, false - не сохранять изменения в файле 
            xlApp.Quit(); //закрываем Excel
            GC.Collect();
            MessageBox.Show("Данные скопированы", "Excel", MessageBoxButtons.OK, MessageBoxIcon.Information);

        }
    }
}

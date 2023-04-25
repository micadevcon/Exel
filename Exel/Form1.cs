using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Exel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void sort(Object[,] mas, int pole)//сортировка по 
        {
            Object[,] s = new string[mas.GetLength(0)+1, mas.GetLength(1)+1];
            for (int i = 1; i < mas.GetLength(0); i++)
            {
                for (int j = i; j < mas.GetLength(0)+1; j++)
                    if (Convert.ToDouble(mas[i, pole]) > Convert.ToDouble(mas[j, pole]))
                        for (int k = 1; k < mas.GetLength(1)+1; k++)
                        {
                            s[j, k] = mas[j, k].ToString();
                            mas[j, k] = mas[i, k].ToString();
                            mas[i, k] = s[j, k].ToString();
                        }
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp = new Excel.Application(); //Excel
            Excel.Workbook Workbook; //рабочая книга откуда будем копировать лист  
           
            Excel.Worksheet Worksheet; //лист Excel
                                       //Properties.Settings.Default.

            Workbook = xlApp.Workbooks.Open(Properties.Settings.Default.PathFile1); //название файла Excel откуда будем копировать лист

            Worksheet = (Excel.Worksheet)Workbook.Worksheets[Properties.Settings.Default.NameList1];

            // xlSht = xlWB.Worksheets["Лист"]; //название листа или 1-й лист в книге xlSht = xlWB.Worksheets[1];
            //Workbook = xlApp.Workbooks.Open(Environment.CurrentDirectory + "\\1\\1.xlsx"); //название файла Excel откуда будем копировать лист
            //xlNewWB = xlApp.Workbooks.Add(); //новая рабочая книга, куда будем вставлять данные
            //xlNewSht = xlNewWB.Sheets[1]; //первый лист по порядку - в него будем вставлять данные
            /*xlSht.Range["A1:A10"].UnMerge();
            xlSht.Range["A1:A10"].Copy(); //копируем диапазон ячеек*/
            //xlSht.Range["B1:B10"].PasteSpecial(Excel.XlPasteType.xlPasteValues);
            //Worksheet.Cells[1, 1] = 123;
            /*Excel.Range r1 = Worksheet.Cells[1, 1];
            Excel.Range r2 = Worksheet.Cells[9, 9];
            Excel.Range range1 = Worksheet.get_Range(r1, r2);*/
            //richTextBox1.Text = Convert.ToString(Worksheet.Range[range1].Value)  ;
            //int iLastCol = Worksheet.Cells[1, Worksheet.Columns.Count].End[Excel.XlDirection.xlToLeft].Column; //последний заполненный столбец в 1-й строке
            //int iLastRow = Worksheet.Cells[Worksheet.Rows.Count, "A"].End[Excel.XlDirection.xlUp].Row; //последняя заполненная строка в столбце А
            /*string[,] saNames = new string[5, 2];

 saNames[0, 0] = "John";
 saNames[0, 1] = "Smith";*/


            //Convert.ToString(Worksheet.Range["A1:A10"].Copy());
            /*            dataGridView1.Rows.Add(Convert.ToString(Worksheet.Cells[2, 2].Value), Convert.ToString(Worksheet.Cells[2, 3].Value));
                       dataGridView1.Rows.Add(3, 9, "cv");
                       dataGridView1.Rows.Add(2, 5, "22s");*/
            //сортировка по дате, возрастание
            //////// dataGridView1.Sort(dataGridView1.Columns["Column2"], System.ComponentModel.ListSortDirection.Ascending);
            //вставить только значения

            //xlApp.Visible = true; //отображаем Excel
            // Workbook.Close(true); //true - сохранить изменения, false - не сохранять изменения в файле 
            //Workbook.Close(false);

            int iLastCol = Worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
            int iLastRow = Worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
           
            Object[,] saRet;
            saRet = (System.Object[,])Worksheet.get_Range("B2", "G420").get_Value();  


           // sort(saRet, 5);
            for (int i = 1; i < saRet.GetLength(0)+1; i++)
            {
                dataGridView1.Rows.Add(saRet[i, 1], saRet[i, 2], saRet[i, 3], saRet[i, 4], saRet[i, 5], saRet[i, 6]);
            }

            xlApp.Quit(); //закрываем Excel
            GC.Collect();
            MessageBox.Show("Данные скопированы", "Excel", MessageBoxButtons.OK, MessageBoxIcon.Information);
            
        }

        private void contextMenuStrip1_Opening(object sender, System.ComponentModel.CancelEventArgs e)
        {
            
        }

        private void настройкиToolStripMenuItem_Click(object sender, EventArgs e)//настройки
        {
            Setting form2 = new Setting();
            form2.Show();
        }
    }
}

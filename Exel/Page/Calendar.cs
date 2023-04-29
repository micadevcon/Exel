using Exel.Page;
using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Exel
{
    public partial class Calendar : Form
    {
        public Calendar()
        {
            
            InitializeComponent();
            Main();
            
        }
        private void checkFalse() 
        {
            radioButton1.Checked = false;
            radioButton2.Checked = false;
            radioButton3.Checked = false;
            radioButton4.Checked = false;
        }
        private object checkNull(object mas)// 
        {
            if (mas == null)
                return "-";
            else
                return mas.ToString();
        }
        private int Main()
        {
            loading form3 = new loading();
            form3.Show();
            form3.progressBar1.Maximum = 100;
            form3.progressBar1.Value = 5;
            Excel.Application xlApp1 = new Excel.Application(); //Excel
            
            Excel.Worksheet Worksheet1; //лист Excel
            Excel.Worksheet Worksheet2;

            Excel.Application xlApp2 = new Excel.Application();
            Excel.Workbook Workbook2 ;
            Excel.Worksheet Worksheet3 ;

            Excel.Workbook Workbook1 ; //рабочая книга откуда будем копировать лист  

           
            try
            {
                if (string.IsNullOrWhiteSpace(Properties.Settings.Default.PathFile1) ||
               string.IsNullOrWhiteSpace(Properties.Settings.Default.PathFile2) ||
               string.IsNullOrWhiteSpace(Properties.Settings.Default.NameList1) ||
               string.IsNullOrWhiteSpace(Properties.Settings.Default.NameList2) ||
               string.IsNullOrWhiteSpace(Properties.Settings.Default.NameList3))
                    throw new Exception("Изначальные данные не заполнены! Пожалуйста, перейдите в меню настройки и заполните пустые поля");
               if (File.Exists(Properties.Settings.Default.PathFile1))
                    Workbook1 = (Excel.Workbook)(xlApp1.Workbooks.Add(Properties.Settings.Default.PathFile1)); //название файла Excel откуда будем копировать лист
                else 
                    throw new Exception("первый файл не существует");
                if (File.Exists(Properties.Settings.Default.PathFile2))
                    Workbook2 = (Excel.Workbook)(xlApp2.Workbooks.Add(Properties.Settings.Default.PathFile2)); //название файла Excel откуда будем копировать лист
                else 
                    throw new Exception("второй файл не существует");
                form3.progressBar1.Value = 15;
                //название листа в файле
                Worksheet1 = (Excel.Worksheet)Workbook1.Worksheets[Properties.Settings.Default.NameList1];
                Worksheet2 = (Excel.Worksheet)Workbook1.Worksheets[Properties.Settings.Default.NameList2];
                Worksheet3 = (Excel.Worksheet)Workbook2.Worksheets[Properties.Settings.Default.NameList3];

                //последний занятый столбец и строка с данными
                int iFirstCol1 = Worksheet1.Cells.SpecialCells(Excel.XlCellType.xlCellTypeConstants).Column;
                int iFirstRow1 = Worksheet1.Cells.SpecialCells(Excel.XlCellType.xlCellTypeConstants).Row;
                int iLastCol1 = Worksheet1.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                int iLastRow1 = Worksheet1.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

                int iFirstCol2 = Worksheet2.Cells.SpecialCells(Excel.XlCellType.xlCellTypeConstants).Column;
                int iFirstRow2 = Worksheet2.Cells.SpecialCells(Excel.XlCellType.xlCellTypeConstants).Row;
                int iLastCol2 = Worksheet2.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                int iLastRow2 = Worksheet2.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

                int iFirstCol3 = Worksheet3.Cells.SpecialCells(Excel.XlCellType.xlCellTypeConstants).Column;
                int iFirstRow3 = Worksheet3.Cells.SpecialCells(Excel.XlCellType.xlCellTypeConstants).Row;
                int iLastCol3 = Worksheet3.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                int iLastRow3 = Worksheet3.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

                
                //копирование данных с листа екселя
                Excel.Range rangeFirst;
                Excel.Range rangeEnd;

                Object[,] masExel1;
                rangeFirst = Worksheet1.Cells[iFirstRow1, iFirstCol1];
                rangeEnd = Worksheet1.Cells[iLastRow1, iLastCol1];
                masExel1 = (System.Object[,])Worksheet1.get_Range(rangeFirst, rangeEnd).get_Value();

                Object[,] masExel2;
                rangeFirst = Worksheet2.Cells[iFirstRow2, iFirstCol2];
                rangeEnd = Worksheet2.Cells[iLastRow2, iLastCol2];
                masExel2 = (System.Object[,])Worksheet2.get_Range(rangeFirst, rangeEnd).get_Value();

                Object[,] masExel3;
                rangeFirst = Worksheet3.Cells[iFirstRow3, iFirstCol3];
                rangeEnd = Worksheet3.Cells[iLastRow3, iLastCol3];
                masExel3 = (System.Object[,])Worksheet3.get_Range(rangeFirst, rangeEnd).get_Value();

                dataGridView1.Columns[1].ValueType = typeof(DateTime);
                //отображение в таблице данных
                form3.progressBar1.Value = 20;
                for (int i = 2; i < masExel1.GetLength(0) + 1; i++)
                {
                    object nameList = Properties.Settings.Default.NameList1;
                    object data = Convert.ToDateTime(masExel1[i, 8]).ToShortDateString();
                    object organiz = checkNull(masExel1[i, 1]);
                    object dogovor = checkNull(masExel1[i, 2]);
                    object atp = checkNull(masExel1[i, 3]);
                    object programm = checkNull(masExel1[i, 4]);
                    object numPeople = checkNull(masExel1[i, 5]);
                    object numTime = checkNull(masExel1[i, 6]);
                    object dataC = checkNull(masExel1[i, 7]);
                    object sum = checkNull(masExel1[i, 9]);
                    object numNotifStart = "-";
                    object numNotifservice = "-";
                    dataGridView1.Rows.Add(
                        nameList, Convert.ToDateTime(data), organiz, dogovor,
                        atp, programm, numPeople, numTime,
                        dataC, sum, numNotifStart, numNotifservice
                        );
                }
                form3.progressBar1.Value = 40;

                for (int i = 2; i < masExel2.GetLength(0) + 1; i++)
                {
                    object nameList = Properties.Settings.Default.NameList2;
                    object data = Convert.ToDateTime(masExel2[i, 8]).ToShortDateString();
                    object organiz = checkNull(masExel2[i, 1]);
                    object dogovor = checkNull(masExel2[i, 2]);
                    object atp = checkNull(masExel2[i, 3]);
                    object programm = checkNull(masExel2[i, 4]);
                    object numPeople = checkNull(masExel2[i, 5]);
                    object numTime = checkNull(masExel2[i, 6]);
                    object dataC = checkNull(masExel2[i, 7]);
                    object sum = checkNull(masExel2[i, 9]);
                    object numNotifStart = "-";
                    object numNotifservice = "-";

                    dataGridView1.Rows.Add(
                        nameList, Convert.ToDateTime(data), organiz, dogovor,
                        atp, programm, numPeople, numTime,
                        dataC, sum, numNotifStart, numNotifservice
                        );
                }
                form3.progressBar1.Value = 60;
                for (int i = 2; i < masExel3.GetLength(0) + 1; i++)
                {
                    object nameList = Properties.Settings.Default.NameList3;
                    object data = Convert.ToDateTime(masExel3[i, 3]);//
                    object organiz = checkNull(masExel3[i, 2]);//
                    object dogovor = checkNull(masExel3[i, 1]);//
                    object atp = "-";
                    object programm = "-";
                    object numPeople = "-";
                    object numTime = "-";
                    object dataC = "-";
                    object sum = "-";
                    object numNotifStart = checkNull(masExel3[i, 4]);//
                    object numNotifservice = checkNull(masExel3[i, 5]);//
                    dataGridView1.Rows.Add(
                        nameList, Convert.ToDateTime(data), organiz, dogovor,
                        atp, programm, numPeople, numTime,
                        dataC, sum, numNotifStart, numNotifservice
                        );
                }
                form3.progressBar1.Value = 80;
                //сортировка 2 столбца по возрастанию
                dataGridView1.Sort(dataGridView1.Columns["Column2"], System.ComponentModel.ListSortDirection.Ascending);
                form3.progressBar1.Value = 100;
                Workbook1.Close(false);
                Workbook2.Close(false);
                
                xlApp1.Quit(); //закрываем Excel
                xlApp2.Quit();
                Marshal.ReleaseComObject(xlApp1);
                Marshal.ReleaseComObject(xlApp2);

                GC.Collect();
                form3.Close();
                return 1;
                

            }
            catch (Exception e)
            {
                form3.Close();
                xlApp1.Quit();
                xlApp2.Quit();
                Marshal.ReleaseComObject(xlApp1);
                Marshal.ReleaseComObject(xlApp2);
                MessageBox.Show($"Ошибка: {e.Message}", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return 0;
                throw;
            }


            
        }


        private void настройкиToolStripMenuItem_Click(object sender, EventArgs e)//настройки
        {
            Setting form2 = new Setting();
            form2.Show();
        }

        private void dataGridView1_Scroll(object sender, ScrollEventArgs e)
        {
            dataGridView1.Update();
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void обновитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            Main();
        }

        private void radioButton4_Click(object sender, EventArgs e)
        {
            checkFalse();
            radioButton4.Checked = true;
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Cells[1].Value != null && Convert.ToDateTime(row.Cells[1].Value) < DateTime.Today)
                    row.Visible = false;
                else
                    row.Visible = true;
            }
        }

        private void radioButton3_Click(object sender, EventArgs e)
        {
            checkFalse();
            radioButton3.Checked = true;
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Cells[1].Value != null && Convert.ToDateTime(row.Cells[1].Value).AddDays(7) < DateTime.Today)
                    row.Visible = false;
                else
                    row.Visible = true;
            }
        }

        private void radioButton2_Click(object sender, EventArgs e)
        {
            checkFalse();
            radioButton2.Checked = true;
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Cells[1].Value != null && Convert.ToDateTime(row.Cells[1].Value).AddDays(30) < DateTime.Today)
                    row.Visible = false;
                else
                    row.Visible = true;
            }
        }

        private void radioButton1_Click(object sender, EventArgs e)
        {
            checkFalse();
            radioButton1.Checked = true;
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Cells[1].Value != null)
                    row.Visible = true;
            }
        }

        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            info form3 = new info();
            form3.Show();
        }
    }
}

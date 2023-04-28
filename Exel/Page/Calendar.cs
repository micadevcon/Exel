using System;
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
            if (string.IsNullOrWhiteSpace(Properties.Settings.Default.PathFile1) ||
                string.IsNullOrWhiteSpace(Properties.Settings.Default.PathFile2) ||
                string.IsNullOrWhiteSpace(Properties.Settings.Default.NameList1) ||
                string.IsNullOrWhiteSpace(Properties.Settings.Default.NameList2) ||
                string.IsNullOrWhiteSpace(Properties.Settings.Default.NameList3))

                MessageBox.Show("Изначальные данные не заполнены! Пожалуйста, перейдите в меню настройки и заполните пустые поля", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Information);
            else
                main();
        }
        private object checkNull(object mas)// 
        {
            if (mas == null)
                return "-";
            else
                return mas.ToString();
        }
        private void sort(Object[,] mas, int pole)//сортировка по дате 
        {
            Object[,] s = new string[mas.GetLength(0) + 1, mas.GetLength(1) + 1];
            for (int i = 1; i < mas.GetLength(0); i++)
            {
                for (int j = i; j < mas.GetLength(0) + 1; j++)
                    if (Convert.ToDouble(mas[i, pole]) > Convert.ToDouble(mas[j, pole]))
                        for (int k = 1; k < mas.GetLength(1) + 1; k++)
                        {
                            s[j, k] = mas[j, k].ToString();
                            mas[j, k] = mas[i, k].ToString();
                            mas[i, k] = s[j, k].ToString();
                        }
            }
        }
        private void main()
        {
            Excel.Application xlApp1 = new Excel.Application(); //Excel
            Excel.Workbook Workbook1; //рабочая книга откуда будем копировать лист  
            Excel.Worksheet Worksheet1; //лист Excel
            Excel.Worksheet Worksheet2;

            Excel.Application xlApp2 = new Excel.Application();
            Excel.Workbook Workbook2 ;
            Excel.Worksheet Worksheet3 ;
            var sheets = xlApp1.Workbooks;

            try
            {
              
                Workbook1 = sheets.Open(Properties.Settings.Default.PathFile1); //название файла Excel откуда будем копировать лист
            }
            catch (Exception)
            {
                Marshal.ReleaseComObject(xlApp1);
                Marshal.ReleaseComObject(Workbook1);
                
                /*GC.Collect();
                GC.WaitForPendingFinalizers();*/
                MessageBox.Show("Неверный путь файлов!", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                
                throw;
            }
            Workbook2 = xlApp2.Workbooks.Open(Properties.Settings.Default.PathFile2); //название файла Excel откуда будем копировать лист

            Worksheet1 = (Excel.Worksheet)Workbook1.Worksheets[Properties.Settings.Default.NameList1];//название листа в файле
            Worksheet2 = (Excel.Worksheet)Workbook1.Worksheets[Properties.Settings.Default.NameList2];//название листа в файле

            
            Worksheet3 = (Excel.Worksheet)Workbook2.Worksheets[Properties.Settings.Default.NameList3];//название листа в файле

            //последний занятый столбец и строка с данными
            //richTextBox1.Text = Worksheet1.Cells.SpecialCells(Excel.XlCellType.xlCellTypeConstants).Row + " "+Worksheet1.Cells.SpecialCells(Excel.XlCellType.xlCellTypeConstants).Column;
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

            // sort(masExel1, 5); хзхзхз
            dataGridView1.Columns[1].ValueType = typeof(DateTime);
            //отображение в таблице данных
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
            //сортировка 2 столбца по возрастанию
            dataGridView1.Sort(dataGridView1.Columns["Column2"], System.ComponentModel.ListSortDirection.Ascending);
            Workbook1.Close(false);
            Workbook2.Close(false);
            xlApp1.Quit(); //закрываем Excel
            xlApp2.Quit();
            GC.Collect();

        }

        private void contextMenuStrip1_Opening(object sender, System.ComponentModel.CancelEventArgs e)
        {

        }

        private void настройкиToolStripMenuItem_Click(object sender, EventArgs e)//настройки
        {
            Setting form2 = new Setting();
            form2.Show();
        }



        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Cells[1].Value != null && Convert.ToDateTime(row.Cells[1].Value).AddDays(30) < DateTime.Today)
                    row.Visible = false;
                else
                    row.Visible = true;
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            dataGridView1.Rows[2].Visible = true;
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                row.Visible = true;
            }
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Cells[1].Value != null && Convert.ToDateTime(row.Cells[1].Value).AddDays(7) < DateTime.Today)
                    row.Visible = false;
                else
                    row.Visible = true;
            }
        }

        private void dataGridView1_Scroll(object sender, ScrollEventArgs e)
        {
            dataGridView1.Update();
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Cells[1].Value != null && Convert.ToDateTime(row.Cells[1].Value) < DateTime.Today)
                    row.Visible = false;
                else
                    row.Visible = true;
            }
        }
    }
}

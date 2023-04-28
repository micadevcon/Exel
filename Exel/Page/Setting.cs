using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Exel
{
    public partial class Setting : Form
    {
        public Setting()
        {
            InitializeComponent();
            //переменные для хранения настроек
            textBox1.Text = Properties.Settings.Default.PathFile1.ToString();
            textBox2.Text = Properties.Settings.Default.PathFile2.ToString();
            textBox3.Text = Properties.Settings.Default.NameList1.ToString();
            textBox4.Text = Properties.Settings.Default.NameList2.ToString();
            textBox5.Text = Properties.Settings.Default.NameList3.ToString();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (            string.IsNullOrWhiteSpace(textBox1.Text) ||
                            string.IsNullOrWhiteSpace(textBox2.Text) ||
                            string.IsNullOrWhiteSpace(textBox3.Text) ||
                            string.IsNullOrWhiteSpace(textBox4.Text) ||
                            string.IsNullOrWhiteSpace(textBox5.Text))
            {
                MessageBox.Show("Не все поля заполнены!!", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else 
            {
                //сохранение данных
                Properties.Settings.Default.PathFile1 = textBox1.Text;
                Properties.Settings.Default.PathFile2 = textBox2.Text;
                Properties.Settings.Default.NameList1 = textBox3.Text;
                Properties.Settings.Default.NameList2 = textBox4.Text;
                Properties.Settings.Default.NameList3 = textBox5.Text;
                //сохранение данных
                Properties.Settings.Default.Save();
                MessageBox.Show("Данные сохранены! Для вступления изменений требуется перезагрузка приложения.", "Оповещение", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string filePath = string.Empty;
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    filePath = openFileDialog.FileName;
                    textBox1.Text = filePath;
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string filePath = string.Empty;
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    filePath = openFileDialog.FileName;
                    textBox2.Text = filePath;
                }
            }
        }
    }
}

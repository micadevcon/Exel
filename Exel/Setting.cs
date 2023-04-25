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
            textBox6.Text = Properties.Settings.Default.ColomnList1.ToString();
            textBox7.Text = Properties.Settings.Default.ColomnList2.ToString();
            textBox8.Text = Properties.Settings.Default.ColomnList3.ToString();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //сохранение данных
            Properties.Settings.Default.PathFile1 = textBox1.Text;
            Properties.Settings.Default.PathFile2 = textBox2.Text;
            Properties.Settings.Default.NameList1 = textBox3.Text;
            Properties.Settings.Default.NameList2 = textBox4.Text;
            Properties.Settings.Default.NameList3 = textBox5.Text;
            Properties.Settings.Default.ColomnList1 = Int32.Parse(textBox6.Text);
            Properties.Settings.Default.ColomnList2 = Int32.Parse(textBox7.Text);
            Properties.Settings.Default.ColomnList3 = Int32.Parse(textBox8.Text);
            //сохранение данных
            Properties.Settings.Default.Save();
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

using System;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Constitution_des_classes
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = @"C:\",
                Title = @"Browse Text Files",

                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = "xlsx",
                Filter = @"xlsx files (*.xlsx)|*.xlsx",
                FilterIndex = 2,
                RestoreDirectory = true,

                ReadOnlyChecked = true,
                ShowReadOnly = true
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                label2.Text = openFileDialog1.FileName;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int compteurBulletins = 0;
            char Classe = 'A';
            var excelApplication = new Microsoft.Office.Interop.Excel.Application();

            var fichierEcolesXlsx = excelApplication.Workbooks.Open(label2.Text);
            var feuilleEcoles = (Worksheet) fichierEcolesXlsx.ActiveSheet;
            int dernierRang = feuilleEcoles.Cells.Find("*", System.Reflection.Missing.Value,
                System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                XlSearchOrder.xlByRows, XlSearchDirection.xlPrevious,
                false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
            var range = feuilleEcoles.Range["B5:B" + dernierRang];
            TabControl tabControl = new TabControl();
            tabControl.Dock = DockStyle.Fill;
            //if (range[5, 2].Text.Contains("6"))
            {
                for (int i = 1; i <= Int16.Parse(textBox1.Text); i++)
                {
                    TabPage tp = new TabPage("6" + Classe);
                    tabControl.TabPages.Add(tp);

                    
                    this.Controls.Add(tabControl);
                    Classe++;
                }
            }
        }
    }
}
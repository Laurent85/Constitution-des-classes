using Microsoft.Office.Interop.Excel;
using System;
using System.Windows.Forms;

namespace Constitution_des_classes
{
    public partial class Principal : Form
    {
        public Principal()
        {
            InitializeComponent();
        }

        public int Filles = 0;
        public int Garcons = 0;
        public int Eleves = 0;

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void btn_Parcourir(object sender, EventArgs e)
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
                lblCheminFichierExcel.Text = openFileDialog1.FileName;
            }
        }

        private void btn_Valider_Config(object sender, EventArgs e)
        {
            char classe = 'A';
            var excelApplication = new Microsoft.Office.Interop.Excel.Application();

            var fichierEcolesXlsx = excelApplication.Workbooks.Open(lblCheminFichierExcel.Text);
            var feuilleEcoles = (Worksheet)fichierEcolesXlsx.ActiveSheet;
            int dernierRang = feuilleEcoles.Cells.Find("*", System.Reflection.Missing.Value,
                System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                XlSearchOrder.xlByRows, XlSearchDirection.xlPrevious,
                false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
            var range = feuilleEcoles.Range["B5:B" + dernierRang];
            tabControl.Dock = DockStyle.Fill;
            TabPage tp1 = new TabPage("Tous les élèves");
            tabControl.TabPages.Add(tp1);
            Controls.Add(tabControl);
            string division = range[5, 1].Text.Substring(0, 1);

            for (int i = 1; i <= Int16.Parse(txbNombreClasses.Text); i++)
            {
                TabPage tp = new TabPage(division + classe);
                tabControl.TabPages.Add(tp);

                classe++;
            }

            TabPage t = tabControl.TabPages[1];
            //tabControl.SelectedTab = t; //go to tab
            ListView liste = new ListView();
            liste.Dock = DockStyle.Fill;
            liste.LabelEdit = true;
            // Allow the user to rearrange columns.
            liste.AllowColumnReorder = true;
            // Display check boxes.
            liste.CheckBoxes = false;
            // Select the item and subitems when selection is made.
            liste.FullRowSelect = true;
            // Display grid lines.
            liste.GridLines = true;
            // Sort the items in the list in ascending order.
            //liste.Sorting = SortOrder.Ascending;
            liste.Scrollable = true;
            // Set to details view.
            liste.View = View.Details;

            t.Controls.Add(liste);

            var range1 = feuilleEcoles.Range["A5:G" + dernierRang];
            int rowCount = range1.Rows.Count;
            int colCount = range1.Columns.Count;
            ListViewItem itm = new ListViewItem();
            string[] arr = new string[7];

            liste.Columns.Add(range1[0, 1].Text, -2, HorizontalAlignment.Left);
            liste.Columns.Add(range1[0, 2].Text, -2, HorizontalAlignment.Left);
            liste.Columns.Add(range1[0, 3].Text, -2, HorizontalAlignment.Left);
            liste.Columns.Add(range1[0, 4].Text, -2, HorizontalAlignment.Left);
            liste.Columns.Add(range1[0, 5].Text, -2, HorizontalAlignment.Left);
            liste.Columns.Add(range1[0, 6].Text, -2, HorizontalAlignment.Left);
            liste.Columns.Add(range1[0, 7].Text, -2, HorizontalAlignment.Left);

            for (int i = 0; i <= rowCount; i++)
            {

                

                for (int j = 1; j <= 7; j++)
                {
                    try
                    {
                        ListViewItem lvitem = new ListViewItem();
                        lvitem.Text = range1.Cells[i, j].Value2.ToString();
                        arr[j - 1] = lvitem.Text;
                    }
                    catch (Exception ex)
                    {
                        //MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
                    }
                }
                itm = new ListViewItem(arr);
                liste.Items.Add(itm);

                // Traitement garçons/filles
                if (range1.Cells[i, 5].Value2.ToString().Contains("M")) Garcons++;
                if (range1.Cells[i, 5].Value2.ToString().Contains("F")) Filles++;


            }
            label1.Text = Garcons.ToString();
            label2.Text = Filles.ToString();
        }
    }
}
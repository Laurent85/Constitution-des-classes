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
            char classe = 'A';
            var excelApplication = new Microsoft.Office.Interop.Excel.Application();

            var fichierEcolesXlsx = excelApplication.Workbooks.Open(label2.Text);
            var feuilleEcoles = (Worksheet) fichierEcolesXlsx.ActiveSheet;
            int dernierRang = feuilleEcoles.Cells.Find("*", System.Reflection.Missing.Value,
                System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                XlSearchOrder.xlByRows, XlSearchDirection.xlPrevious,
                false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
            var range = feuilleEcoles.Range["B5:B" + dernierRang];
            //TabControl tabControl = new TabControl();
            tabControl.Dock = DockStyle.Fill;
            TabPage tp1 = new TabPage("Tous les élèves");
            tabControl.TabPages.Add(tp1);
            Controls.Add(tabControl);
            string division = range[5, 1].Text.Substring(0, 1);
            //if (range[5, 1].Text.Contains("6"))
            {
                for (int i = 1; i <= Int16.Parse(textBox1.Text); i++)
                {
                    TabPage tp = new TabPage(division + classe);
                    tabControl.TabPages.Add(tp);

                    //tp.Controls.Add(new System.Windows.Forms.Button());
                    //this.Controls.Add(tabControl);
                    classe++;
                }
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
            liste.Sorting = SortOrder.Ascending;
            liste.Scrollable = true;
            // Set to details view.
            liste.View = View.Details;
            


            t.Controls.Add(liste);

            var range1 = feuilleEcoles.Range["A5:G" + dernierRang];
            int rowCount = range1.Rows.Count;
            int colCount = range1.Columns.Count;
            ListViewItem itm = new ListViewItem();
            string[] arr = new string[7];

            // Add a column with width 20 and left alignment.
            liste.Columns.Add(range1[0,1].Text, -2, HorizontalAlignment.Left);
            liste.Columns.Add(range1[0, 2].Text, -2, HorizontalAlignment.Left);
            liste.Columns.Add(range1[0, 3].Text, -2, HorizontalAlignment.Left);
            liste.Columns.Add(range1[0, 4].Text, -2, HorizontalAlignment.Left);
            liste.Columns.Add(range1[0, 5].Text, -2, HorizontalAlignment.Left);
            liste.Columns.Add(range1[0, 6].Text, -2, HorizontalAlignment.Left);
            liste.Columns.Add(range1[0, 7].Text, -2, HorizontalAlignment.Left);


            for (int i = 4; i <= rowCount; i++)
            {
                
                for (int j = 1; j <= 7; j++)
                {


                    try
                    {

                        
                        
                        ListViewItem lvitem = new ListViewItem();
                        lvitem.Text = range1.Cells[i, j].Value2.ToString();
                        arr[j-1] = lvitem.Text;
                        //lvitem.SubItems.Add(range1.Cells[i, j].Value2.ToString());
                        //lvitem.SubItems.Add(range1.Cells[i, 3].Value2.ToString());
                        
                        

                    }
                    catch (Exception ex)
                    {
                        //MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
                    }

                    
                }
                itm = new ListViewItem(arr);
                liste.Items.Add(itm);
            }
        }

        private void ImportExcel(string filename, ListView liste)
        {
            Microsoft.Office.Interop.Excel.Application xla = new Microsoft.Office.Interop.Excel.Application();
            xla.Visible = true;

            // Load the workbook
            Microsoft.Office.Interop.Excel.Workbook wb = xla.Workbooks.Open(filename);

            // Get the worksheet
            Microsoft.Office.Interop.Excel.Worksheet ws = wb.Worksheets[1];

            // Retrieve number of columns, number of rows
            int numberOfColumns = GetNumberOfColumns(ws);
            int numberOfRows = GetNumberOfRows(ws);

            for (int r = 1; r <= numberOfRows; r++)
            {
                // Put your own code here to add the listviewitem
                ListViewItem item = liste.Items.Add(ws.Cells[r, 1].Value);

                for (int c = 1; c <= numberOfColumns; c++)
                {
                    // Put your own code here to add the subitems
                    item.SubItems.Add(ws.Cells[r, c].Value.ToString());
                }
            }

            // Quit Excel
            xla.Quit();
            xla = null;
        }

        /// <summary>
        /// Gets the number of rows by searching the first null row
        /// </summary>
        /// <param name="ws">worksheet to search</param>
        /// <returns>number of rows</returns>
        /// <remarks>there probably is a better way to do this</remarks>
        private int GetNumberOfRows(Microsoft.Office.Interop.Excel.Worksheet ws)
        {
            int numberOfRows = 5;
            while (ws.Cells[1, numberOfRows].Value != null)
            {
                numberOfRows += 1;
            }
            numberOfRows -= 1; // substract 1 to get the last filled column

            return numberOfRows;
        }

        /// <summary>
        /// Gets the number of columns by searching the first null row
        /// </summary>
        /// <param name="ws">worksheet to search</param>
        /// <returns>number of columns</returns>
        /// <remarks>there probably is a better way to do this</remarks>
        private int GetNumberOfColumns(Microsoft.Office.Interop.Excel.Worksheet ws)
        {
            int numberOfColumns = 1;
            while (ws.Cells[1, numberOfColumns].Value != null)
            {
                numberOfColumns += 1;
            }
            numberOfColumns -= 1; // substract 1 to get the last filled column

            return numberOfColumns;
        }
    }
}
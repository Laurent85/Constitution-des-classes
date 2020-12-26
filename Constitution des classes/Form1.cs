using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Windows.Forms;

namespace Constitution_des_classes
{
    public partial class Principal : Form
    {
        public Principal()
        {
            InitializeComponent();
        }

        public int NbFilles = 0;
        public int NbGarcons = 0;
        public int NbElevesTotal;
        public int MoyenneElevesClasse;
        public int NbDivisions;
        public int NbEcoles = 0;
        public int NbLignesEcoles = 0;
        public int NbLignesMariagesOptions = 0;
        public int NumLigneExiste;
        public int VerifieLigneExiste = 0;
        public int NbOptions = 0;
        public int NbMariagesOptions = 0;
        public DataGridView listeEleves = new DataGridView();
        public DataGridView listeEcoles = new DataGridView();
        public DataGridView listeOptions = new DataGridView();
        public DataGridView listeMariagesOptions = new DataGridView();
        public Range range;

        private void Form1_Load(object sender, EventArgs e)
        {
            TuerProcessus("Excel");
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

        private void paramètresListe(DataGridView liste)
        {
            liste.Dock = DockStyle.Fill;
            liste.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            liste.DoubleBuffered(true);
            liste.EditMode = DataGridViewEditMode.EditOnEnter;
        }

        private void CréationOnglet(TabPage NomOnglet, string TitreOnglet, DataGridView liste)
        {
            NomOnglet = new TabPage(TitreOnglet);
            tabPrincipal.TabPages.Add(NomOnglet);
            NomOnglet.Controls.Add(liste);
            paramètresListe(liste);
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
            range = feuilleEcoles.Range["A5:J" + dernierRang];
            tabPrincipal.Dock = DockStyle.Fill;

            CréationOnglet(new TabPage("OngletEleves"), "Tous les élèves", listeEleves);
            CréationOnglet(new TabPage("OngletEcoles"), "Ecoles primaires", listeEcoles);
            CréationOnglet(new TabPage("OngletOptions"), "Options", listeOptions);
            CréationOnglet(new TabPage("OngletMariagesOptions"), "Mariages d'options", listeMariagesOptions);

            string division = range[5, 2].Text.Substring(0, 1);
            NbDivisions = Int16.Parse(txbNombreClasses.Text);

            for (int i = 1; i <= NbDivisions; i++)
            {
                TabPage OngletsClasses = new TabPage(division + classe);
                tabPrincipal.TabPages.Add(OngletsClasses);
                var tableau_classe = new DataGridView();
                tableau_classe.Name = "liste" + division + classe;
                //var v = ("effectif" + division + classe).ToString();
                //System.Windows.Forms.Label vi = new System.Windows.Forms.Label();
                //vi.Name = v;
                //vi.Text = "toto";
                //groupBox1.Controls.Add(vi);
                paramètresListe(tableau_classe);
                OngletsClasses.Controls.Add(tableau_classe);
                tableau_classe.Columns.Add(range[0, 3].Text, range[0, 3].Text);
                tableau_classe.Columns.Add(range[0, 4].Text, range[0, 4].Text);
                tableau_classe.Columns.Add(range[0, 7].Text, range[0, 7].Text);
                tableau_classe.Columns.Add(range[0, 8].Text, range[0, 8].Text);
                tableau_classe.Columns.Add(range[0, 9].Text, range[0, 9].Text);
                tableau_classe.Columns.Add(range[0, 10].Text, range[0, 10].Text);
                classe++;
            }

            int rowCount = range.Rows.Count;

            listeEleves.Columns.Add(range[0, 1].Text, range[0, 1].Text);
            listeEleves.Columns.Add(range[0, 2].Text, range[0, 2].Text);
            listeEleves.Columns.Add(range[0, 3].Text, range[0, 3].Text);
            listeEleves.Columns.Add(range[0, 4].Text, range[0, 4].Text);
            listeEleves.Columns.Add(range[0, 5].Text, range[0, 5].Text);
            listeEleves.Columns.Add(range[0, 6].Text, range[0, 6].Text);
            listeEleves.Columns.Add(range[0, 7].Text, range[0, 7].Text);
            listeEleves.Columns.Add(range[0, 8].Text, range[0, 8].Text);
            listeEleves.Columns.Add(range[0, 9].Text, range[0, 9].Text);
            listeEleves.Columns.Add(range[0, 10].Text, range[0, 10].Text);

            listeEcoles.Columns.Add("Nom", "Nom");
            listeEcoles.Columns.Add("Elèves", "Elèves");

            listeOptions.Columns.Add("Nom", "Nom");
            listeOptions.Columns.Add("Elèves", "Elèves");

            listeMariagesOptions.Columns.Add("Nom", "Nom");
            listeMariagesOptions.Columns.Add("Elèves", "Elèves");

            DataGridViewComboBoxColumn ColonneComboListeEcoles = new DataGridViewComboBoxColumn();
            listeEcoles.Columns.Add(ColonneComboListeEcoles);
            ColonneComboListeEcoles.HeaderText = "Liste des Elèves           ";

            DataGridViewComboBoxColumn ColonneComboListeMariagesOptions = new DataGridViewComboBoxColumn();
            listeMariagesOptions.Columns.Add(ColonneComboListeMariagesOptions);
            ColonneComboListeMariagesOptions.HeaderText = "Liste des Elèves           ";

            for (int i = 1; i <= rowCount; i++)
            {
                listeEleves.Rows.Add();

                for (int j = 1; j <= 10; j++)
                {
                    ListViewItem cellule = new ListViewItem();

                    if ((range.Cells[i, j].Value2) != null)
                    {
                        cellule.Text = range.Cells[i, j].Value2.ToString();
                        listeEleves.Rows[i - 1].Cells[j - 1].Value = cellule.Text;

                        if (j == 5)
                        {
                            RechercherTexte(listeEcoles, cellule.Text, 0);

                            if (VerifieLigneExiste == 1)
                            {
                                NbEcoles = Int16.Parse(listeEcoles.Rows[NumLigneExiste].Cells[1].Value.ToString()) + 1;
                                listeEcoles.Rows[NumLigneExiste].Cells[1].Value = NbEcoles.ToString();
                                (listeEcoles.Rows[NumLigneExiste].Cells[2] as DataGridViewComboBoxCell).Items.Add(listeEleves.Rows[i - 1].Cells[j - 3].Value);
                            }

                            if (VerifieLigneExiste == 0)
                            {
                                NbEcoles = 1;
                                listeEcoles.Rows.Add(cellule.Text, NbEcoles.ToString());

                                DataGridViewComboBoxCell CelluleComboBoxListeEcoles = new DataGridViewComboBoxCell();
                                CelluleComboBoxListeEcoles.DropDownWidth = 200;
                                CelluleComboBoxListeEcoles.Items.Add(listeEleves.Rows[i - 1].Cells[j - 3].Value);
                                listeEcoles.Rows[NbLignesEcoles].Cells[2] = CelluleComboBoxListeEcoles;
                                NbEcoles = 0;
                                NbLignesEcoles++;
                            }
                        }

                        if (j == 6 || j == 7 || j == 8 || j == 9 || j == 10)
                        {
                            RechercherTexte(listeOptions, cellule.Text, 0);

                            if (VerifieLigneExiste == 1)
                            {
                                NbOptions = Int16.Parse(listeOptions.Rows[NumLigneExiste].Cells[1].Value.ToString()) + 1;
                                listeOptions.Rows[NumLigneExiste].Cells[1].Value = NbOptions.ToString();
                            }

                            if (VerifieLigneExiste == 0)
                            {
                                NbOptions = 1;
                                listeOptions.Rows.Add(cellule.Text, NbOptions.ToString());
                                NbOptions = 0;
                            }
                        }

                        if (j == 6)
                        {
                            cellule.Text = "";
                            for (int c = 7; c <= 10; c++)
                            {
                                if ((range.Cells[i, c].Value2) != null)

                                {
                                    if (c == 7)
                                    {
                                        cellule.Text = range.Cells[i, c].Value2.ToString();
                                    }
                                    else
                                        cellule.Text = cellule.Text + "/" + range.Cells[i, c].Value2.ToString();
                                }
                            }

                            RechercherTexte(listeMariagesOptions, cellule.Text, 0);

                            if (VerifieLigneExiste == 1)
                            {
                                NbMariagesOptions = Int16.Parse(listeMariagesOptions.Rows[NumLigneExiste].Cells[1].Value.ToString()) + 1;
                                listeMariagesOptions.Rows[NumLigneExiste].Cells[1].Value = NbMariagesOptions.ToString();
                                (listeMariagesOptions.Rows[NumLigneExiste].Cells[2] as DataGridViewComboBoxCell).Items.Add(listeEleves.Rows[i - 1].Cells[j - 4].Value);
                                //if (cellule.Text == "/ANGLAIS LV1/ESPAGNOL LV2")
                                //{
                                //    DataGridView liste = (DataGridView)this.Controls.Find("liste4A", true)[0];
                                //    liste.Rows.Add(range.Cells[i, 3].Value2.ToString(), range.Cells[i, 4].Value2.ToString(), range.Cells[i, 8].Value2.ToString());
                                //}
                            }

                            if (VerifieLigneExiste == 0)
                            {
                                NbMariagesOptions = 1;
                                listeMariagesOptions.Rows.Add(cellule.Text, NbMariagesOptions.ToString());
                                DataGridViewComboBoxCell CelluleComboBoxListeMariagesOptions = new DataGridViewComboBoxCell();
                                CelluleComboBoxListeMariagesOptions.DropDownWidth = 200;
                                CelluleComboBoxListeMariagesOptions.Items.Add(listeEleves.Rows[i - 1].Cells[j - 4].Value);
                                listeMariagesOptions.Rows[NbLignesMariagesOptions].Cells[2] = CelluleComboBoxListeMariagesOptions;
                                NbMariagesOptions = 0;
                                NbLignesMariagesOptions++;
                                //if (cellule.Text == "/ANGLAIS LV1/ESPAGNOL LV2")
                                //{
                                //    DataGridView liste = (DataGridView)this.Controls.Find("liste4A", true)[0];
                                //    liste.Rows.Add(range.Cells[i, 3].Value2.ToString(), range.Cells[i, 4].Value2.ToString(), range.Cells[i, 8].Value2.ToString());
                                //}
                            }
                        }
                    }
                }

                if (range.Cells[i, 4].Value2.ToString().Contains("M")) NbGarcons++;
                if (range.Cells[i, 4].Value2.ToString().Contains("F")) NbFilles++;
            }

            lblGarcons.Text = NbGarcons.ToString() + " garçons";
            lblFilles.Text = NbFilles.ToString() + " filles";
            NbElevesTotal = NbGarcons + NbFilles;
            lblTotalEleves.Text = NbElevesTotal.ToString() + " élèves au total";
            int moyenne = (NbElevesTotal / NbDivisions);
            int reste = (NbElevesTotal % NbDivisions);
            listeMariagesOptions.Columns[0].Width = -1;
            listeEcoles.Columns[0].Width = -1;
            listeOptions.Columns[0].Width = -1;
        }

        private void TuerProcessus(string processus)
        {
            var process = System.Diagnostics.Process.GetProcessesByName(processus);
            foreach (var p in process)
            {
                if (!string.IsNullOrEmpty(p.ProcessName))
                {
                    try
                    {
                        p.Kill();
                    }
                    catch
                    {
                        // ignored
                    }
                }
            }
        }

        private void RechercherTexte(DataGridView liste, string recherche, int colonne)
        {
            VerifieLigneExiste = 0;
            foreach (DataGridViewRow row in liste.Rows)
            {
                if (row.Cells[colonne].Value != null)
                {
                    if (row.Cells[colonne].Value.ToString().Equals(recherche))
                    {
                        VerifieLigneExiste = 1;
                        NumLigneExiste = row.Index;
                    }
                }
            }
            //return false;
        }

        private void btn_Constituer_Classes_Click(object sender, EventArgs e)
        {
            if (listeMariagesOptions.Rows[0].Cells[0].Value.Equals("ESPAGNOL LV2"))
            {
                //for (int i = 0; i < (listeMariagesOptions.Rows[0].Cells[2] as DataGridViewComboBoxCell).Items.Count; i++)
                for (int i = 0; i < 29; i++)
                {
                    string nomEleve = ((listeMariagesOptions.Rows[0].Cells[2] as DataGridViewComboBoxCell).Items[i].ToString());
                    string option1 = (listeMariagesOptions.Rows[0].Cells[0].Value.ToString());
                    string sexe = "";
                    foreach (DataGridViewRow row in listeEleves.Rows)
                    {
                        if (row.Cells[2].Value.ToString().Equals(nomEleve))
                        {
                            sexe = row.Cells[3].Value.ToString();
                            break;
                        }
                    }
                    DataGridView liste = (DataGridView)this.Controls.Find("liste4E", true)[0];
                    liste.Rows.Add(nomEleve,sexe, option1);
                    (listeMariagesOptions.Rows[0].Cells[2] as DataGridViewComboBoxCell).Items.Remove((listeMariagesOptions.Rows[0].Cells[2] as DataGridViewComboBoxCell).Items[i]);
                    NbMariagesOptions = Int16.Parse(listeMariagesOptions.Rows[0].Cells[1].Value.ToString()) - 1;
                    listeMariagesOptions.Rows[0].Cells[1].Value = NbMariagesOptions.ToString();
                }
            }
        }
    }

    public static class ExtensionMethods
    {
        public static void DoubleBuffered(this DataGridView dgv, bool setting)
        {
            Type dgvType = dgv.GetType();
            PropertyInfo pi = dgvType.GetProperty("DoubleBuffered",
                BindingFlags.Instance | BindingFlags.NonPublic);
            pi.SetValue(dgv, setting, null);
        }
    }
}
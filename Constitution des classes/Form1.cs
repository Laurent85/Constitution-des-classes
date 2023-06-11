using Microsoft.Office.Interop.Excel;
using System;
using System.Text;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Xceed.Document.NET;
using Xceed.Words.NET;
using Border = Xceed.Document.NET.Border;
using BorderStyle = Xceed.Document.NET.BorderStyle;
using Label = System.Windows.Forms.Label;
using Range = Microsoft.Office.Interop.Excel.Range;
using System.IO;

namespace Constitution_des_classes
{
    public partial class Principal : Form
    {
        public Principal()
        {
            InitializeComponent();
        }

        public int NbFilles;
        public int NbGarcons;
        public int NbElevesTotal;
        public int MoyenneElevesClasse;
        public int MoyenneElevesClasseEtReste;
        public int Reste;
        public static int NbDivisions;
        public int NbEcoles;
        public int NbLignesEcoles;
        public int NbLignesMariagesOptions;
        public int NbLignesOptions;
        public int NumLigneExiste;
        public int VerifieLigneExiste;
        public int NbOptions;
        public int NbMariagesOptions;
        public int IndexRang;
        public int DernierRang;
        public int NbDoublons;
        public static string Division;
        public DataGridView ListeEleves = new DataGridView();
        public DataGridView ListeEcoles = new DataGridView();
        public DataGridView ListeOptions = new DataGridView();
        public DataGridView ListeMariagesOptions = new DataGridView();
        public DataGridView ListeBilan = new DataGridView();
        public Range Range;
        private readonly List<Label> _lblEffectifs = new List<Label>();
        private readonly List<System.Windows.Forms.TextBox> _txbEffectifs = new List<System.Windows.Forms.TextBox>();
        private readonly List<Label> _lblRemplirClasse = new List<Label>();
        private readonly List<Label> _lblOptions = new List<Label>();
        private readonly List<Label> _lblNbOptions = new List<Label>();
        private readonly List<Label> _lblMariagesOptions = new List<Label>();
        private readonly List<Label> _lblNbMariagesOptions = new List<Label>();
        private readonly List<Label> _lblClassesMariagesOptions = new List<Label>();
        public readonly List<Label> NomDuPp = new List<Label>();
        private readonly List<System.Windows.Forms.CheckBox> _cbxMariagesOptions = new List<System.Windows.Forms.CheckBox>();
        public static string[] ListePp = new string[6];

        private void Form1_Load(object sender, EventArgs e)
        {
            TuerProcessus("Excel");
            cbxNbAjoutEleves.SelectedIndex = 0;
            cbxAnnée.SelectedIndex = 0;
            btnValiderConfig.Enabled = false;
            btnWord.Enabled = false;
            btnPP.Enabled = false;
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

                var excelApplication = new Microsoft.Office.Interop.Excel.Application();

                var fichierEcolesXlsx = excelApplication.Workbooks.Open(lblCheminFichierExcel.Text);
                var feuilleEcoles = (Worksheet)fichierEcolesXlsx.ActiveSheet;
                int dernierRang = feuilleEcoles.Cells.Find("*", Missing.Value,
                    Missing.Value, Missing.Value,
                    XlSearchOrder.xlByRows, XlSearchDirection.xlPrevious,
                    false, Missing.Value, Missing.Value).Row;

                Range = feuilleEcoles.Range["A5:J" + dernierRang];
                Division = Range[5, 2].Text.Substring(0, 1);
                if (Division == "3")
                {
                    lblNiveauInit.Text = @"Niveau 3ème";
                    lblNiveau.Text = @"Niveau 3ème";
                }
                if (Division == "4")
                {
                    lblNiveauInit.Text = @"Niveau 4ème";
                    lblNiveau.Text = @"Niveau 4ème";
                }
                if (Division == "5")
                {
                    lblNiveauInit.Text = @"Niveau 5ème";
                    lblNiveau.Text = @"Niveau 5ème";
                }
                if (Division == "6")
                {
                    lblNiveauInit.Text = @"Niveau 6ème";
                    lblNiveau.Text = @"Niveau 6ème";
                }
                excelApplication.ActiveWorkbook.Close(false);
                excelApplication.Quit();
            }
        }

        private void paramètresListe(DataGridView liste)
        {
            liste.Dock = DockStyle.Fill;
            liste.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            liste.DoubleBuffered(true);
            liste.EditMode = DataGridViewEditMode.EditOnEnter;
        }

        private void CréationOnglet(TabPage nomOnglet, string titreOnglet, DataGridView liste)
        {
            if (nomOnglet == null) throw new ArgumentNullException(nameof(nomOnglet));
            nomOnglet = new TabPage(titreOnglet);
            tabPrincipal.TabPages.Add(nomOnglet);
            nomOnglet.Controls.Add(liste);
            paramètresListe(liste);
        }

        private void btn_Valider_Config(object sender, EventArgs e)
        {
            Attente attente = new Attente();
            attente.Show();
            char classe = 'A';
            var excelApplication = new Microsoft.Office.Interop.Excel.Application();

            var fichierEcolesXlsx = excelApplication.Workbooks.Open(lblCheminFichierExcel.Text);
            var feuilleEcoles = (Worksheet)fichierEcolesXlsx.ActiveSheet;
            int dernierRang = feuilleEcoles.Cells.Find("*", Missing.Value,
                Missing.Value, Missing.Value,
                XlSearchOrder.xlByRows, XlSearchDirection.xlPrevious,
                false, Missing.Value, Missing.Value).Row;
            Range = feuilleEcoles.Range["A5:M" + dernierRang];
            tabPrincipal.Dock = DockStyle.Fill;

            CréationOnglet(new TabPage("OngletEleves"), "Tous les élèves", ListeEleves);
            CréationOnglet(new TabPage("OngletEcoles"), "Ecoles primaires", ListeEcoles);
            CréationOnglet(new TabPage("OngletOptions"), "Options", ListeOptions);
            CréationOnglet(new TabPage("OngletMariagesOptions"), "Mariages d'options", ListeMariagesOptions);

            foreach (Range cell in Range)
            {
                if (cell.Text.Contains("Accompagnement"))
                {
                    cell.Value = @"ADP";
                }

                if (cell.Text.Contains("EUROPEEN"))
                {
                    cell.Value = @"ANGLAIS EURO";
                }

                if (cell.Text.Contains("LCA"))
                {
                    cell.Value = @"LATIN";
                }

                if (cell.Text.Contains("ESPAGNOL"))
                {
                    cell.Value = @"ESPAGNOL";
                }

                if (cell.Text.Contains("ALLEMAND"))
                {
                    cell.Value = @"ALLEMAND";
                }
            }

            Division = Range[5, 2].Text.Substring(0, 1);
            NbDivisions = Int16.Parse(cbxNombreClasses.Text);
            int y = 40;
            int y1 = 40;
            int y2 = 80;
            int x1 = 400;

            for (int i = 0; i < NbDivisions; i++)
            {
                TabPage ongletsClasses = new TabPage(Division + classe);
                tabPrincipal.TabPages.Add(ongletsClasses);
                var tableauClasse = new DataGridView { Name = "liste" + Division + classe };

                paramètresListe(tableauClasse);
                ongletsClasses.Controls.Add(tableauClasse);
                tableauClasse.Columns.Add(Range[0, 3].Text, Range[0, 3].Text);
                tableauClasse.Columns.Add(Range[0, 4].Text, Range[0, 4].Text);
                tableauClasse.Columns.Add(Range[0, 7].Text, Range[0, 7].Text);
                tableauClasse.Columns.Add(Range[0, 8].Text, Range[0, 8].Text);
                tableauClasse.Columns.Add(Range[0, 9].Text, Range[0, 9].Text);
                tableauClasse.Columns.Add(Range[0, 10].Text, Range[0, 10].Text);
                ListeBilan.Columns.Add(Division + classe, Division + classe);
                ListeBilan.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Raised;
                ListeBilan.Columns[i].Width = 1500 / NbDivisions;

                ListeBilan.ColumnHeadersDefaultCellStyle.BackColor = Color.LightGray;
                ListeBilan.EnableHeadersVisualStyles = false;
                ListeBilan.Columns[i].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                ListeBilan.Columns[i].HeaderCell.Style.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, FontStyle.Bold, GraphicsUnit.Pixel);
                classe++;
            }

            grpBilan.Controls.Add(ListeBilan);
            ListeBilan.Dock = DockStyle.Fill;
            for (int i = 1; i <= 15; i++)
            {
                ListeBilan.Rows.Add();
            }

            ListeBilan.Rows[0].Cells[0].Selected = false;

            int rowCount = Range.Rows.Count;

            ListeEleves.Columns.Add(Range[0, 1].Text, Range[0, 1].Text);
            ListeEleves.Columns.Add(Range[0, 2].Text, Range[0, 2].Text);
            ListeEleves.Columns.Add(Range[0, 3].Text, Range[0, 3].Text);
            ListeEleves.Columns.Add(Range[0, 4].Text, Range[0, 4].Text);
            ListeEleves.Columns.Add(Range[0, 5].Text, Range[0, 5].Text);
            ListeEleves.Columns.Add(Range[0, 6].Text, Range[0, 6].Text);
            ListeEleves.Columns.Add(Range[0, 7].Text, Range[0, 7].Text);
            ListeEleves.Columns.Add(Range[0, 8].Text, Range[0, 8].Text);
            ListeEleves.Columns.Add(Range[0, 9].Text, Range[0, 9].Text);
            ListeEleves.Columns.Add(Range[0, 10].Text, Range[0, 10].Text);
            ListeEleves.Columns.Add(Range[0, 11].Text, Range[0, 11].Text);
            ListeEleves.Columns.Add(Range[0, 12].Text, Range[0, 12].Text);
            ListeEleves.Columns.Add(Range[0, 13].Text, Range[0, 13].Text);

            ListeEcoles.Columns.Add("Nom", "Nom");
            ListeEcoles.Columns.Add("Elèves", "Elèves");

            ListeOptions.Columns.Add("Nom", "Nom");
            ListeOptions.Columns.Add("Elèves", "Elèves");

            ListeMariagesOptions.Columns.Add("Nom", "Nom");
            ListeMariagesOptions.Columns.Add("Elèves", "Elèves");

            DataGridViewComboBoxColumn colonneComboListeEcoles = new DataGridViewComboBoxColumn();
            ListeEcoles.Columns.Add(colonneComboListeEcoles);
            colonneComboListeEcoles.HeaderText = @"Liste des Elèves           ";

            DataGridViewComboBoxColumn colonneComboListeMariagesOptions = new DataGridViewComboBoxColumn();
            ListeMariagesOptions.Columns.Add(colonneComboListeMariagesOptions);
            colonneComboListeMariagesOptions.HeaderText = @"Liste des Elèves           ";

            DataGridViewComboBoxColumn colonneComboListeOptions = new DataGridViewComboBoxColumn();
            ListeOptions.Columns.Add(colonneComboListeOptions);
            colonneComboListeOptions.HeaderText = @"Liste des Elèves           ";

            for (int i = 1; i <= rowCount; i++)
            {
                ListeEleves.Rows.Add();

                for (int j = 1; j <= 13; j++)
                {
                    ListViewItem cellule = new ListViewItem();

                    if ((Range.Cells[i, j].Value2) != null)
                    {
                        cellule.Text = Range.Cells[i, j].Value2.ToString();
                        ListeEleves.Rows[i - 1].Cells[j - 1].Value = cellule.Text;

                        if (j == 5)
                        {
                            RechercherTexte(ListeEcoles, cellule.Text, 0);

                            if (VerifieLigneExiste == 1)
                            {
                                NbEcoles = Int16.Parse(ListeEcoles.Rows[NumLigneExiste].Cells[1].Value.ToString()) + 1;
                                ListeEcoles.Rows[NumLigneExiste].Cells[1].Value = NbEcoles.ToString();
                                (ListeEcoles.Rows[NumLigneExiste].Cells[2] as DataGridViewComboBoxCell)?.Items.Add(
                                    ListeEleves.Rows[i - 1].Cells[2].Value);
                            }

                            if (VerifieLigneExiste == 0)
                            {
                                NbEcoles = 1;
                                ListeEcoles.Rows.Add(cellule.Text, NbEcoles.ToString());

                                DataGridViewComboBoxCell celluleComboBoxListeEcoles = new DataGridViewComboBoxCell();
                                celluleComboBoxListeEcoles.DropDownWidth = 200;
                                celluleComboBoxListeEcoles.Items.Add(ListeEleves.Rows[i - 1].Cells[2].Value);
                                ListeEcoles.Rows[NbLignesEcoles].Cells[2] = celluleComboBoxListeEcoles;
                                NbEcoles = 0;
                                NbLignesEcoles++;
                            }
                        }

                        if (j == 7 || j == 8 || j == 9 || j == 10 || j == 11 || j == 12 || j == 13)
                        {
                            RechercherTexte(ListeOptions, cellule.Text, 0);

                            if (VerifieLigneExiste == 1)
                            {
                                NbOptions = Int16.Parse(ListeOptions.Rows[NumLigneExiste].Cells[1].Value.ToString()) +
                                            1;
                                ListeOptions.Rows[NumLigneExiste].Cells[1].Value = NbOptions.ToString();
                                (ListeOptions.Rows[NumLigneExiste].Cells[2] as DataGridViewComboBoxCell)?.Items.Add(
                                    ListeEleves.Rows[i - 1].Cells[2].Value);
                            }

                            if (VerifieLigneExiste == 0)
                            {
                                NbOptions = 1;
                                ListeOptions.Rows.Add(cellule.Text, NbOptions.ToString());
                                DataGridViewComboBoxCell celluleComboBoxListeOptions = new DataGridViewComboBoxCell();
                                celluleComboBoxListeOptions.DropDownWidth = 200;
                                celluleComboBoxListeOptions.Items.Add(ListeEleves.Rows[i - 1].Cells[2].Value);
                                ListeOptions.Rows[NbLignesOptions].Cells[2] = celluleComboBoxListeOptions;
                                NbOptions = 0;
                                NbLignesOptions++;
                            }
                        }

                        if ((j == 6) && (Range.Cells[i, j].Value2 != null))
                        {
                            cellule.Text = "";
                            for (int c = 7; c <= 13; c++)
                            {
                                if ((Range.Cells[i, c].Value2) != null)

                                {
                                    if ((c == 7) && ((Range.Cells[i, c].Value2) != null))
                                    {
                                        cellule.Text = Range.Cells[i, c].Value2.ToString();
                                    }
                                    else if (!(cellule.Text.Contains(Range.Cells[i, c].Value2.ToString())))
                                    {
                                        cellule.Text = cellule.Text + @"/" + Range.Cells[i, c].Value2.ToString();
                                    }
                                }
                            }

                            RechercherTexte(ListeMariagesOptions, cellule.Text, 0);

                            if (VerifieLigneExiste == 1)
                            {
                                NbMariagesOptions =
                                    Int16.Parse(ListeMariagesOptions.Rows[NumLigneExiste].Cells[1].Value.ToString()) +
                                    1;
                                ListeMariagesOptions.Rows[NumLigneExiste].Cells[1].Value = NbMariagesOptions.ToString();
                                (ListeMariagesOptions.Rows[NumLigneExiste].Cells[2] as DataGridViewComboBoxCell)?.Items
                                    .Add(ListeEleves.Rows[i - 1].Cells[2].Value);
                            }

                            if (VerifieLigneExiste == 0)
                            {
                                NbMariagesOptions = 1;
                                ListeMariagesOptions.Rows.Add(cellule.Text, NbMariagesOptions.ToString());
                                DataGridViewComboBoxCell celluleComboBoxListeMariagesOptions =
                                    new DataGridViewComboBoxCell();
                                celluleComboBoxListeMariagesOptions.DropDownWidth = 200;
                                celluleComboBoxListeMariagesOptions.Items.Add(ListeEleves.Rows[i - 1].Cells[2].Value);
                                ListeMariagesOptions.Rows[NbLignesMariagesOptions].Cells[2] =
                                    celluleComboBoxListeMariagesOptions;
                                NbMariagesOptions = 0;
                                NbLignesMariagesOptions++;
                            }
                        }
                    }
                }

                if (Range.Cells[i, 4].Value2.ToString().Contains("M")) NbGarcons++;
                if (Range.Cells[i, 4].Value2.ToString().Contains("F")) NbFilles++;
            }

            lblNiveau.Text = @"Niveau " + Division + @"ème";
            lblGarcons.Text = NbGarcons.ToString() + @" garçons";
            lblFilles.Text = NbFilles.ToString() + @" filles";
            lblNbClasses.Text = NbDivisions.ToString() + @" classes";
            lblNbOptions.Text = ListeOptions.Rows.Count - 1 + @" options";
            lblNbGroupesOptions.Text = ListeMariagesOptions.Rows.Count - 1 + @" groupes d'options";
            NbElevesTotal = NbGarcons + NbFilles;
            lblTotalEleves.Text = NbElevesTotal.ToString() + @" élèves au total";
            MoyenneElevesClasse = (NbElevesTotal / NbDivisions);
            Reste = (NbElevesTotal % NbDivisions);
            ListeMariagesOptions.Columns[0].Width = -1;
            ListeEcoles.Columns[0].Width = -1;
            ListeOptions.Columns[0].Width = -1;

            classe = 'A';
            for (int i = 0; i < NbDivisions; i++)
            {
                _lblEffectifs.Add(new Label());
                _lblEffectifs[i].Name = "lblEffectif_" + i;
                _lblEffectifs[i].Text = Division + classe;
                _lblEffectifs[i].Location = new System.Drawing.Point(20, y);
                grpEffectifs.Controls.Add(_lblEffectifs[i]);

                _txbEffectifs.Add(new System.Windows.Forms.TextBox());
                _txbEffectifs[i].Name = "txbEffectif_" + Division + classe;
                _txbEffectifs[i].Width = 50;
                _txbEffectifs[i].Location = new System.Drawing.Point(120, y);
                grpEffectifs.Controls.Add(_txbEffectifs[i]);

                _lblRemplirClasse.Add(new Label());
                _lblRemplirClasse[i].Name = "lblRemplir_" + Division + classe;
                _lblRemplirClasse[i].Text = @"0";
                _lblRemplirClasse[i].Location = new System.Drawing.Point(180, y);
                grpEffectifs.Controls.Add(_lblRemplirClasse[i]);
                y = y + 30;

                if (Reste > 0)
                {
                    MoyenneElevesClasseEtReste = MoyenneElevesClasse + 1;
                    _txbEffectifs[i].Text = MoyenneElevesClasseEtReste.ToString();
                    Reste = Reste - 1;
                }
                else
                {
                    _txbEffectifs[i].Text = MoyenneElevesClasse.ToString();
                }

                classe++;
            }

            int nbOptions = ListeOptions.Rows.Count - 1;

            for (int i = 0; i < nbOptions; i++)
            {
                _lblOptions.Add(new Label());
                _lblOptions[i].Name = "lbl" + ListeOptions.Rows[i].Cells[0].Value;
                _lblOptions[i].Text = ListeOptions.Rows[i].Cells[0].Value.ToString();
                _lblOptions[i].Location = new System.Drawing.Point(7, y1);
                grpOptions.Controls.Add(_lblOptions[i]);

                _lblNbOptions.Add(new Label());
                _lblNbOptions[i].Name = "lblNb" + ListeOptions.Rows[i].Cells[0].Value;
                _lblNbOptions[i].Text = ListeOptions.Rows[i].Cells[1].Value.ToString();
                _lblNbOptions[i].Location = new System.Drawing.Point(150, y1);
                grpOptions.Controls.Add(_lblNbOptions[i]);
                y1 = y1 + 30;
            }

            int nbMariagesOptions = ListeMariagesOptions.Rows.Count - 1;

            for (int i = 0; i < nbMariagesOptions; i++)
            {
                _lblMariagesOptions.Add(new Label());
                _lblMariagesOptions[i].Name = "lblOption" + ListeMariagesOptions.Rows[i].Cells[0].Value;
                _lblMariagesOptions[i].AutoSize = true;
                _lblMariagesOptions[i].Text = ListeMariagesOptions.Rows[i].Cells[0].Value.ToString();
                _lblMariagesOptions[i].Location = new System.Drawing.Point(7, y2);
                grpMariagesOptions.Controls.Add(_lblMariagesOptions[i]);

                _lblNbMariagesOptions.Add(new Label());
                _lblNbMariagesOptions[i].Name = "lblNb" + ListeMariagesOptions.Rows[i].Cells[0].Value;
                _lblNbMariagesOptions[i].Text = ListeMariagesOptions.Rows[i].Cells[1].Value.ToString();
                _lblNbMariagesOptions[i].Location = new System.Drawing.Point(300, y2);
                grpMariagesOptions.Controls.Add(_lblNbMariagesOptions[i]);
                y2 = y2 + 30;
            }

            classe = 'A';
            for (int i = 0; i < NbDivisions; i++)
            {
                _lblClassesMariagesOptions.Add(new Label());
                _lblClassesMariagesOptions[i].Name = "lblClasse" + Division + classe;
                _lblClassesMariagesOptions[i].AutoSize = true;
                _lblClassesMariagesOptions[i].Text = Division + classe;
                _lblClassesMariagesOptions[i].Location = new System.Drawing.Point(x1, 40);
                grpMariagesOptions.Controls.Add(_lblClassesMariagesOptions[i]);
                classe++;
                x1 = x1 + 50;
            }

            int p = 0;
            int y5 = 80;
            foreach (Control labelOption in grpMariagesOptions.Controls)
            {
                if (labelOption is Label)
                {
                    if (labelOption.Name.Contains("lblOption"))
                    {
                        int x5 = 403;
                        foreach (Control labelClasse in grpMariagesOptions.Controls)
                        {
                            if (labelClasse is Label)
                            {
                                if (labelClasse.Name.Contains("lblClasse"))
                                {
                                    _cbxMariagesOptions.Add(new System.Windows.Forms.CheckBox());
                                    _cbxMariagesOptions[p].Name = labelOption.Text + "_" + labelClasse.Text;
                                    _cbxMariagesOptions[p].AutoSize = true;
                                    _cbxMariagesOptions[p].Text = null;
                                    _cbxMariagesOptions[p].Location = new System.Drawing.Point(x5, y5);
                                    _cbxMariagesOptions[p].CheckStateChanged += Classe_Cochee;
                                    grpMariagesOptions.Controls.Add(_cbxMariagesOptions[p]);
                                    x5 = x5 + 50;
                                    p++;
                                }
                            }
                        }
                    }
                }

                y5 = y5 + 15;
            }

            VerifierCasesCochables();
            btnWord.Enabled = true;
            btnPP.Enabled = true;
            attente.Close();
            //fichierEcolesXlsx.Close();
            excelApplication.ActiveWorkbook.Close(false);
            excelApplication.Quit();
        }

        private void Classe_Cochee(object sender, EventArgs e)
        {
            System.Windows.Forms.CheckBox chbxOptionClasse = (System.Windows.Forms.CheckBox)sender;
            string option = chbxOptionClasse.Name.Before("_");
            string classe = chbxOptionClasse.Name.After("_");
            string nbOptionGrpOptions;
            int effectifADistribuer = 0;

            if (chbxOptionClasse.Checked)
            {
                #region Effectifs groupe mariages options

                foreach (Control lblMariagesOption in grpMariagesOptions.Controls)
                {
                    if ((lblMariagesOption is Label) && (lblMariagesOption.Name.Contains("lblNb")))
                    {
                        // ReSharper disable once PossibleNullReferenceException
                        int index = ListeBilan.Columns[classe].Index;
                        if (chbxOptionClasse.Name.Before("_") == lblMariagesOption.Name.After("lblNb"))
                        {
                            if (Int16.Parse(lblMariagesOption.Text) > MoyenneElevesClasse)
                            {
                                chbxOptionClasse.Tag = MoyenneElevesClasse;

                                if (cbxNbAjoutEleves.SelectedIndex != 0)
                                {
                                    if ((Int16.Parse(cbxNbAjoutEleves.Text) <
                                         Int16.Parse(lblMariagesOption.Text)))
                                    {
                                        chbxOptionClasse.Tag = Int16.Parse(cbxNbAjoutEleves.Text);
                                    }
                                }
                            }
                            else
                            {
                                if (cbxNbAjoutEleves.SelectedIndex != 0)
                                {
                                    if ((Int16.Parse(cbxNbAjoutEleves.Text) <
                                         Int16.Parse(lblMariagesOption.Text)))
                                    {
                                        chbxOptionClasse.Tag = Int16.Parse(cbxNbAjoutEleves.Text);
                                    }
                                    else
                                    {
                                        chbxOptionClasse.Tag = Int16.Parse(lblMariagesOption.Text);
                                    }
                                }
                                else
                                {
                                    chbxOptionClasse.Tag = Int16.Parse(lblMariagesOption.Text);
                                }
                            }

                            lblMariagesOption.Text = (Int16.Parse(lblMariagesOption.Text) - Int16.Parse(chbxOptionClasse.Tag.ToString())).ToString();
                            if (Int16.Parse(lblMariagesOption.Text) > 0)
                            {
                                lblMariagesOption.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
                                lblMariagesOption.ForeColor = Color.Black;
                            }
                            else
                            {
                                lblMariagesOption.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold);
                                lblMariagesOption.ForeColor = Color.Red;
                            }

                            #endregion Effectifs groupe mariages options

                            #region Ajout tableau bilan

                            foreach (DataGridViewRow ligne in ListeBilan.Rows)
                            {
                                if (ligne.Cells[index].Value == null)
                                {
                                    ligne.Cells[index].Value =
                                        chbxOptionClasse.Tag + " " + chbxOptionClasse.Name.Before("_");
                                    break;
                                }
                            }

                            #endregion Ajout tableau bilan

                            cbxNbAjoutEleves.SelectedIndex = 0;
                        }
                    }
                }

                #region Effectifs groupe options

                foreach (Control lblOption in grpOptions.Controls)
                {
                    if ((lblOption is Label) && (lblOption.Name.Contains("lblNb")))
                    {
                        nbOptionGrpOptions = lblOption.Name.After("lblNb");
                        if (chbxOptionClasse.Name.Contains(nbOptionGrpOptions))
                        {
                            lblOption.Text = (Int16.Parse(lblOption.Text) - Int16.Parse(chbxOptionClasse.Tag.ToString())).ToString();
                            if (Int16.Parse(lblOption.Text) > 0)
                            {
                                lblOption.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
                                lblOption.ForeColor = Color.Black;
                            }
                            else
                            {
                                lblOption.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold);
                                lblOption.ForeColor = Color.Red;
                            }
                        }
                    }
                }

                #endregion Effectifs groupe options

                #region Effectifs groupe effectifs

                foreach (Control lblEffectif in grpEffectifs.Controls)
                {
                    int effectif = 0;
                    string nomClasse = lblEffectif.Name.After("_");
                    if ((lblEffectif is Label) && (lblEffectif.Name.Contains(nomClasse)) && (lblEffectif.Name.Contains("lblRemplir")))
                    {
                        effectif = Int16.Parse(lblEffectif.Text);
                    }

                    if ((lblEffectif is Label) && (chbxOptionClasse.Name.After("_") == lblEffectif.Name.After("_")))
                    {
                        {
                            lblEffectif.Text = (Int16.Parse(lblEffectif.Text) + Int16.Parse(chbxOptionClasse.Tag.ToString())).ToString();
                            effectifADistribuer = Int16.Parse(chbxOptionClasse.Tag.ToString());
                            effectif = Int16.Parse(lblEffectif.Text);
                        }
                    }

                    foreach (Control txbEffectif in grpEffectifs.Controls)
                    {
                        if ((txbEffectif is System.Windows.Forms.TextBox) && (txbEffectif.Name.Contains(nomClasse)))
                        {
                            int effectifMax;
                            {
                                effectifMax = Int16.Parse(txbEffectif.Text);
                            }
                            if ((effectif == effectifMax) && (lblEffectif is Label) && (lblEffectif.Name.Contains(nomClasse)))
                            {
                                lblEffectif.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold);
                                lblEffectif.ForeColor = Color.Red;
                            }
                            if ((effectif < effectifMax) && (lblEffectif is Label) && (lblEffectif.Name.Contains(nomClasse)))
                            {
                                lblEffectif.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
                                lblEffectif.ForeColor = Color.Black;
                            }
                        }
                    }
                }

                #endregion Effectifs groupe effectifs

                ConstituerClassesAjouter(option, classe, effectifADistribuer);
            }

            if (!chbxOptionClasse.Checked)
            {
                #region Effectifs groupe mariages options

                foreach (Control lblMariagesOption in grpMariagesOptions.Controls)
                {
                    if ((lblMariagesOption is Label) && (lblMariagesOption.Name.Contains("lblNb")))
                    {
                        if (chbxOptionClasse.Name.Before("_") == lblMariagesOption.Name.After("lblNb"))
                        {
                            lblMariagesOption.Text = (Int16.Parse(lblMariagesOption.Text) + Int16.Parse(chbxOptionClasse.Tag.ToString())).ToString();
                            if (Int16.Parse(lblMariagesOption.Text) > 0)
                            {
                                lblMariagesOption.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
                                lblMariagesOption.ForeColor = Color.Black;
                            }
                            else
                            {
                                lblMariagesOption.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold);
                                lblMariagesOption.ForeColor = Color.Red;
                            }
                        }
                    }
                }

                #endregion Effectifs groupe mariages options

                #region Effectifs groupe options

                foreach (Control lblOption in grpOptions.Controls)
                {
                    if ((lblOption is Label) && (lblOption.Name.Contains("lblNb")))
                    {
                        nbOptionGrpOptions = lblOption.Name.After("lblNb");
                        if (chbxOptionClasse.Name.Contains(nbOptionGrpOptions))
                        {
                            lblOption.Text = (Int16.Parse(lblOption.Text) + Int16.Parse(chbxOptionClasse.Tag.ToString())).ToString();
                            if (Int16.Parse(lblOption.Text) > 0)
                            {
                                lblOption.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
                                lblOption.ForeColor = Color.Black;
                            }
                            else
                            {
                                lblOption.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold);
                                lblOption.ForeColor = Color.Red;
                            }
                        }
                    }
                }

                #endregion Effectifs groupe options

                #region Effectifs groupe effectifs

                foreach (Control lblEffectif in grpEffectifs.Controls)
                {
                    int effectif = 0;
                    string nomClasse = lblEffectif.Name.After("_");
                    if ((lblEffectif is Label) && (lblEffectif.Name.Contains(nomClasse)) && (lblEffectif.Name.Contains("lblRemplir")))
                    {
                        effectif = Int16.Parse(lblEffectif.Text);
                    }

                    if ((lblEffectif is Label) && (chbxOptionClasse.Name.After("_") == lblEffectif.Name.After("_")))
                    {
                        {
                            lblEffectif.Text = (Int16.Parse(lblEffectif.Text) - Int16.Parse(chbxOptionClasse.Tag.ToString())).ToString();
                            effectif = Int16.Parse(lblEffectif.Text);
                        }
                    }

                    foreach (Control txbEffectif in grpEffectifs.Controls)
                    {
                        if ((txbEffectif is System.Windows.Forms.TextBox) && (txbEffectif.Name.Contains(nomClasse)))
                        {
                            int effectifMax;
                            {
                                effectifMax = Int16.Parse(txbEffectif.Text);
                            }
                            if ((effectif == effectifMax) && (lblEffectif is Label) && (lblEffectif.Name.Contains(nomClasse)))
                            {
                                lblEffectif.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold);
                                lblEffectif.ForeColor = Color.Red;
                            }
                            if ((effectif < effectifMax) && (lblEffectif is Label) && (lblEffectif.Name.Contains(nomClasse)))
                            {
                                lblEffectif.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
                                lblEffectif.ForeColor = Color.Black;
                            }
                        }
                    }
                }

                #endregion Effectifs groupe effectifs

                #region Ajout tableau bilan

                int index = ListeBilan.Columns[classe].Index;
                foreach (DataGridViewRow ligne in ListeBilan.Rows)
                {
                    if (ligne.Cells[index].Value != null)
                    {
                        if (ligne.Cells[index].Value.ToString() ==
                            chbxOptionClasse.Tag + " " + chbxOptionClasse.Name.Before("_"))
                        {
                            ligne.Cells[index].Value = null;
                            break;
                        }
                    }
                }

                for (int i = 0; i <= ListeBilan.Rows.Count - 2; i++)
                {
                    if (ListeBilan.Rows[i].Cells[index].Value == null)
                    {
                        {
                            ListeBilan.Rows[i].Cells[index].Value = ListeBilan.Rows[i + 1].Cells[index].Value;
                            ListeBilan.Rows[i + 1].Cells[index].Value = null;
                        }
                    }
                }

                #endregion Ajout tableau bilan

                ConstituerClassesRetirer(option, classe);
            }

            VerifierCasesCochables();
            VérifierOptionsClasses();
        }

        private void TuerProcessus(string processus)
        {
            var process = Process.GetProcessesByName(processus);
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
        }

        private void VerifierCasesCochables()
        {
            int effectifEnCours = 0;
            int effectifMaxi = 0;
            int effectifRestant = 0;

            foreach (Control chbxMariagesOption in grpMariagesOptions.Controls)
            {
                if ((chbxMariagesOption is System.Windows.Forms.CheckBox))
                {
                    string nomOption = chbxMariagesOption.Name.Before("_");
                    string nomClasse = chbxMariagesOption.Name.After("_");
                    string nomLabelEffectifRestant = "lblNb" + nomOption;

                    foreach (Control lblRestant in grpMariagesOptions.Controls)
                    {
                        if ((lblRestant is Label) && (lblRestant.Name == nomLabelEffectifRestant))
                        {
                            effectifRestant = Int16.Parse(lblRestant.Text);
                        }
                    }

                    foreach (Control lblEffectifs in grpEffectifs.Controls)
                    {
                        if ((lblEffectifs is Label) && (lblEffectifs.Name == "lblRemplir_" + nomClasse))
                        {
                            effectifEnCours = Int16.Parse(lblEffectifs.Text);
                        }
                        if ((lblEffectifs is System.Windows.Forms.TextBox) && (lblEffectifs.Name.Contains(nomClasse)))
                        {
                            effectifMaxi = Int16.Parse(lblEffectifs.Text);
                        }

                        if (effectifEnCours == effectifMaxi)
                        {
                            if (((System.Windows.Forms.CheckBox)chbxMariagesOption)
                                .Checked ==
                                false)
                            {
                                chbxMariagesOption.Enabled = false;
                            }
                            else
                            {
                                chbxMariagesOption.Enabled = true;
                            }
                        }
                    }

                    if (cbxNbAjoutEleves.Text == @"Maxi")
                    {
                        if ((effectifEnCours + effectifRestant > effectifMaxi) && (((System.Windows.Forms.CheckBox)chbxMariagesOption)
                            .Checked ==
                            false))
                        {
                            chbxMariagesOption.Enabled = false;
                        }
                        else
                        {
                            chbxMariagesOption.Enabled = true;
                        }
                    }
                    if (cbxNbAjoutEleves.Text != @"Maxi")
                    {
                        if ((effectifEnCours + Int16.Parse(cbxNbAjoutEleves.Text) > effectifMaxi) && (((System.Windows.Forms.CheckBox)chbxMariagesOption)
                            .Checked ==
                            false))
                        {
                            chbxMariagesOption.Enabled = false;
                        }
                        else
                        {
                            chbxMariagesOption.Enabled = true;
                        }
                    }

                    if (effectifRestant == 0)
                    {
                        if (((System.Windows.Forms.CheckBox)chbxMariagesOption)
                            .Checked ==
                            false)
                        {
                            chbxMariagesOption.Enabled = false;
                        }
                        else
                        {
                            chbxMariagesOption.Enabled = true;
                        }
                    }
                }
            }
        }

        private void ConstituerClassesAjouter(string option, string classe, int nombreEleves)
        {
            foreach (DataGridViewRow ligne in ListeMariagesOptions.Rows)
            {
                if (ligne.Cells[0].Value != null)
                {
                    if (ligne.Cells[0].Value.ToString() == option)
                    {
                        DataGridView liste = (DataGridView)Controls.Find("liste" + classe, true)[0]; // ex : "liste4E"

                        {
                            for (int i = 0; i < nombreEleves; i++)
                            {
                                string nomEleveMariageOptions =
                                    ((ligne.Cells[2] as DataGridViewComboBoxCell)?.Items[i]
                                        .ToString());
                                string mariageOptions = (ligne.Cells[0].Value.ToString());
                                string sexe = "";
                                foreach (DataGridViewRow row in ListeEleves.Rows)
                                {
                                    if (row.Cells[2].Value.ToString().Equals(nomEleveMariageOptions))
                                    {
                                        sexe = row.Cells[3].Value.ToString();
                                        break;
                                    }
                                }

                                liste.Rows.Add(nomEleveMariageOptions, sexe, mariageOptions);
                            }
                            for (int i = nombreEleves - 1; i >= 0; i--)
                            {
                                string nomEleveMariageOptions =
                                    ((ligne.Cells[2] as DataGridViewComboBoxCell)?.Items[i]
                                        .ToString());
                                foreach (DataGridViewRow row in ListeEleves.Rows)
                                {
                                    if (row.Cells[2].Value.ToString().Equals(nomEleveMariageOptions))
                                    {
                                        break;
                                    }
                                }

                                (ligne.Cells[2] as DataGridViewComboBoxCell)?.Items.Remove(
                                    (ligne.Cells[2] as DataGridViewComboBoxCell)?.Items[i] ??
                                    throw new InvalidOperationException());

                                ligne.Cells[1].Value = (ligne.Cells[2] as DataGridViewComboBoxCell)?.Items.Count;
                            }
                        }
                    }
                }
            }
        }

        private void ConstituerClassesRetirer(string option, string classe)
        {
            DataGridView liste = (DataGridView)Controls.Find("liste" + classe, true)[0]; // ex : "liste4E"

            {
                int nombreLignes = liste.Rows.Count;
                for (int i = nombreLignes - 2; i >= 0; i--)
                {
                    {
                        if (liste.Rows[i].Cells[2].Value.ToString() == option)
                        {
                            foreach (DataGridViewRow ligne in ListeMariagesOptions.Rows)
                            {
                                if (ligne.Cells[0].Value != null)
                                {
                                    if (ligne.Cells[0].Value.ToString() == option)
                                    {
                                        (ligne.Cells[2] as DataGridViewComboBoxCell)?.Items.Add(liste.Rows[i].Cells[0]
                                            .Value.ToString());
                                        ligne.Cells[1].Value = (ligne.Cells[2] as DataGridViewComboBoxCell)?.Items
                                            .Count;
                                    }
                                }
                            }

                            foreach (DataGridViewRow ligne in ListeOptions.Rows)
                            {
                                if (ligne.Cells[0].Value != null)
                                {
                                    if (ligne.Cells[0].Value.ToString() == option)
                                    {
                                        (ligne.Cells[2] as DataGridViewComboBoxCell)?.Items.Add(liste.Rows[i].Cells[0]
                                            .Value.ToString());
                                    }
                                }
                            }

                            liste.Rows.Remove(liste.Rows[i]);
                        }
                    }
                }
            }
        }

        private void cbxNbAjoutEleves_SelectedIndexChanged(object sender, EventArgs e)
        {
            VerifierCasesCochables();
        }

        //private void RestartProgram()
        //{
        //    // Get file path of current process
        //    var filePath = Assembly.GetExecutingAssembly().Location;
        //    //var filePath = Application.ExecutablePath;  // for WinForms

        //    // Start program
        //    Process.Start(filePath);

        //    // For Windows Forms app
        //    System.Windows.Forms.Application.Exit();

        //    // For all Windows application but typically for Console app.
        //    //Environment.Exit(0);
        //}

        private void btnWord_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                var dossier = folderBrowserDialog1.SelectedPath;

                DocX doc = DocX.Create(dossier + @"\Structure_niveau_" + Division + "ème.docx");
                doc.PageLayout.Orientation = Xceed.Document.NET.Orientation.Landscape;
                doc.MarginBottom = 30;
                doc.MarginTop = 30;
                doc.MarginLeft = 30;
                doc.MarginRight = 30;

                var tableau = doc.AddTable(1, NbDivisions);
                tableau.Alignment = Alignment.center;
                tableau.Design = TableDesign.Custom;
                tableau.AutoFit = AutoFit.Contents;

                var title = doc.InsertParagraph("Structure " + cbxAnnée.Text + "   -   Niveau " + Division + "ème");
                title.FontSize(18).Font(new Xceed.Document.NET.Font("Calibri"));
                title.Color(Color.DarkRed);
                title.Bold();
                title.UnderlineColor(Color.DarkRed);
                title.Alignment = Alignment.center;
                doc.InsertParagraph();

                var résumé1 = doc.InsertParagraph(NbElevesTotal + " élèves = " + NbDivisions + " divisions");
                résumé1.FontSize(12).Font(new Xceed.Document.NET.Font("Calibri"));
                résumé1.Color(Color.BlueViolet);
                résumé1.Bold();
                résumé1.Alignment = Alignment.center;
                doc.InsertParagraph();

                string résuméOptions = "";
                for (int i = 0; i < ListeOptions.Rows.Count - 2; i++)
                {
                    résuméOptions = résuméOptions + ListeOptions.Rows[i].Cells[1].Value + " " + ListeOptions.Rows[i].Cells[0].Value + " - ";
                }
                résuméOptions = résuméOptions + ListeOptions.Rows[ListeOptions.Rows.Count - 2].Cells[1].Value + " " + ListeOptions.Rows[ListeOptions.Rows.Count - 2].Cells[0].Value;
                var résumé2 = doc.InsertParagraph(résuméOptions);
                résumé2.FontSize(10).Font(new Xceed.Document.NET.Font("Calibri"));
                résumé2.Color(Color.BlueViolet);
                résumé2.Alignment = Alignment.center;
                doc.InsertParagraph();

                char classe = 'A';
                tableau.InsertRow();
                tableau.InsertRow();
                int effectif = 0;
                for (int i = 0; i < NbDivisions; i++)
                {
                    foreach (Control lblEffectif in grpEffectifs.Controls)
                    {
                        if ((lblEffectif is Label) && (lblEffectif.Name.Contains(Division + classe)) &&
                            (lblEffectif.Name.Contains("lblRemplir")))
                        {
                            effectif = Int16.Parse(lblEffectif.Text);
                        }
                    }

                    tableau.Rows[0].Cells[i].Paragraphs.First().Append(Division + classe + " (" + effectif + ")");
                    tableau.Rows[0].Cells[i].Paragraphs.First().Color(Color.Black);
                    tableau.Rows[0].Cells[i].Paragraphs.First().Bold();
                    tableau.Rows[0].Cells[i].Paragraphs.First().FontSize(14).Font(new Xceed.Document.NET.Font("Calibri"));
                    tableau.Rows[0].Cells[i].FillColor = (Color.LightPink);
                    tableau.Rows[0].Cells[i].Paragraphs.First().Alignment = Alignment.center;
                    Border b = new Border(BorderStyle.Tcbs_single, BorderSize.one, 0, Color.Gray);
                    Border b1 = new Border(BorderStyle.Tcbs_single, BorderSize.two, 0, Color.Gray);
                    tableau.Rows[0].Cells[i].SetBorder(TableCellBorderType.Left, b);
                    tableau.Rows[0].Cells[i].SetBorder(TableCellBorderType.Right, b);
                    tableau.Rows[0].Cells[i].SetBorder(TableCellBorderType.Bottom, b1);
                    tableau.Rows[0].Cells[i].SetBorder(TableCellBorderType.Top, b1);

                    tableau.Rows[1].Cells[i].Paragraphs.First().Append("PP : " + ListePp[i]);
                    tableau.Rows[1].Cells[i].Paragraphs.First().Color(Color.DarkBlue);
                    tableau.Rows[1].Cells[i].Paragraphs.First().FontSize(9).Font(new Xceed.Document.NET.Font("Calibri"));
                    tableau.Rows[1].Cells[i].FillColor = (Color.LightPink);
                    tableau.Rows[1].Cells[i].Paragraphs.First().Alignment = Alignment.center;
                    b = new Border(BorderStyle.Tcbs_single, BorderSize.one, 0, Color.Gray);
                    tableau.Rows[1].Cells[i].SetBorder(TableCellBorderType.Left, b);
                    tableau.Rows[1].Cells[i].SetBorder(TableCellBorderType.Right, b);
                    tableau.Rows[1].Cells[i].SetBorder(TableCellBorderType.Bottom, b1);

                    DataGridView liste = (DataGridView)Controls.Find("liste" + Division + classe, true)[0]; // ex : "liste4E"
                    string options = "";
                    foreach (DataGridViewRow row in liste.Rows)
                    {
                        if (row.Cells[2].Value != null)
                        {
                            string groupeOptions = "/" + row.Cells[2].Value;
                            string[] authorInfo = groupeOptions.Split('/');

                            foreach (string info in authorInfo)
                            {
                                if (!options.Contains(info))
                                {
                                    options = options + info + "  ";
                                }
                            }
                        }
                    }
                    tableau.Rows[2].Cells[i].Paragraphs.First().Append(options.ToLower());
                    tableau.Rows[2].Cells[i].Paragraphs.First().Color(Color.Black);
                    tableau.Rows[2].Cells[i].Paragraphs.First().Bold();
                    tableau.Rows[2].Cells[i].Paragraphs.First().FontSize(11).Font(new Xceed.Document.NET.Font("Calibri"));
                    tableau.Rows[2].Cells[i].FillColor = (Color.LightBlue);
                    tableau.Rows[2].Cells[i].Paragraphs.First().Alignment = Alignment.center;
                    tableau.Rows[2].Cells[i].SetBorder(TableCellBorderType.Left, b);
                    tableau.Rows[2].Cells[i].SetBorder(TableCellBorderType.Right, b);
                    tableau.Rows[2].Cells[i].SetBorder(TableCellBorderType.Bottom, b1);
                    classe++;
                }

                {
                    foreach (DataGridViewColumn colonneBilan in ListeBilan.Columns)
                    {
                        string classeBilan = colonneBilan.HeaderText;
                        int nbLignes = 2;

                        foreach (DataGridViewRow ligneBilan in ListeBilan.Rows)
                        {
                            if (ligneBilan.Cells[colonneBilan.Index].Value != null)
                            {
                                tableau.InsertRow();
                                nbLignes++;

                                tableau.Rows[nbLignes].Cells[colonneBilan.Index].Paragraphs.First().Append(ligneBilan.Cells[colonneBilan.Index].Value.ToString().ToLower());
                                tableau.Rows[nbLignes].Cells[colonneBilan.Index].FillColor = (Color.LightYellow);
                                tableau.Rows[nbLignes].Cells[colonneBilan.Index].Paragraphs.First().Bold();
                                tableau.Rows[nbLignes].Cells[colonneBilan.Index].Paragraphs.First().FontSize(10).Font(new Xceed.Document.NET.Font("Calibri"));
                                tableau.Rows[nbLignes].Cells[colonneBilan.Index].Paragraphs.First().Color(Color.Red);
                                Border b = new Border(BorderStyle.Tcbs_single, BorderSize.one, 0, Color.Gray);
                                Border b1 = new Border(BorderStyle.Tcbs_single, BorderSize.one, 0, Color.LightGray);
                                tableau.Rows[nbLignes].Cells[colonneBilan.Index].SetBorder(TableCellBorderType.Left, b);
                                tableau.Rows[nbLignes].Cells[colonneBilan.Index].SetBorder(TableCellBorderType.Right, b);
                                tableau.Rows[nbLignes].Cells[colonneBilan.Index].SetBorder(TableCellBorderType.Top, b1);
                                tableau.Rows[nbLignes].Cells[colonneBilan.Index].Paragraphs.First().Alignment = Alignment.center;

                                DataGridView liste = (DataGridView)Controls.Find("liste" + classeBilan, true)[0]; // ex : "liste4E"

                                foreach (DataGridViewRow ligne in liste.Rows)
                                {
                                    if (ligne.Cells[2].Value != null)
                                    {
                                        int supr = ligneBilan.Cells[colonneBilan.Index].Value.ToString().IndexOf(" ", StringComparison.Ordinal);
                                        string option = ligneBilan.Cells[colonneBilan.Index].Value.ToString()
                                            .Remove(0, supr + 1);
                                        if (ligne.Cells[2].Value.ToString() == option)
                                        {
                                            tableau.InsertRow();
                                            nbLignes++;

                                            if (!chkAffecterEleves.Checked)
                                            {
                                                if (ligneBilan.Cells[colonneBilan.Index].Style.ForeColor == Color.Red)
                                                {
                                                    tableau.Rows[nbLignes].Cells[colonneBilan.Index].Paragraphs
                                                        .First().Append(" -");
                                                }
                                                else
                                                {
                                                    tableau.Rows[nbLignes].Cells[colonneBilan.Index].Paragraphs
                                                        .First().Append(ligne.Cells[0].Value.ToString());
                                                }
                                            }
                                            else
                                            {
                                                tableau.Rows[nbLignes].Cells[colonneBilan.Index].Paragraphs
                                                                                                        .First().Append(ligne.Cells[0].Value.ToString());
                                            }

                                            tableau.Rows[nbLignes].Cells[colonneBilan.Index].SetBorder(TableCellBorderType.Top, new Border(BorderStyle.Tcbs_none, BorderSize.one, 1, Color.Red));
                                            tableau.Rows[nbLignes].Cells[colonneBilan.Index].Paragraphs.First().Italic();
                                            tableau.Rows[nbLignes].Cells[colonneBilan.Index].Paragraphs.First().FontSize(8).Font(new Xceed.Document.NET.Font("Calibri"));
                                            tableau.Rows[nbLignes].Cells[colonneBilan.Index].SetBorder(TableCellBorderType.Left, b);
                                            tableau.Rows[nbLignes].Cells[colonneBilan.Index].SetBorder(TableCellBorderType.Right, b);
                                        }
                                    }
                                }
                            }
                        }
                    }

                    for (int i = tableau.Rows.Count - 1; i >= 0; i--)
                    {
                        int compteur = 0;
                        for (int j = 0; j < NbDivisions; j++)
                        {
                            if (tableau.Rows[i].Cells[j].Paragraphs[0].Text == "")
                            {
                                {
                                    compteur++;

                                    if (compteur == NbDivisions)
                                    {
                                        tableau.Rows[i].Remove();
                                        compteur = 0;
                                    }
                                }
                            }
                        }
                    }
                }

                doc.InsertTable(tableau);
                doc.Save();
                Process.Start("WINWORD.EXE", dossier + @"\Structure_niveau_" + Division + "ème.docx");
            }
        }

        private void ChangementLblChemin(object sender, EventArgs e)
        {
            if ((lblCheminFichierExcel.Text.Contains("xls")) && (Regex.IsMatch(cbxNombreClasses.Text, @"^\d+$")) &&
                (cbxAnnée.Text != null))
            {
                btnValiderConfig.Enabled = true;
            }
            else
            {
                btnValiderConfig.Enabled = false;
            }
        }

        private void cbxNombreClasses_SelectedIndexChanged(object sender, EventArgs e)
        {
            if ((lblCheminFichierExcel.Text.Contains("xls")) && (Regex.IsMatch(cbxNombreClasses.Text, @"^\d+$")) &&
                (cbxAnnée.Text != null))
            {
                btnValiderConfig.Enabled = true;
            }
            else
            {
                btnValiderConfig.Enabled = false;
            }
        }

        private void cbxAnnée_SelectedIndexChanged(object sender, EventArgs e)
        {
            if ((lblCheminFichierExcel.Text.Contains("xls")) && (Regex.IsMatch(cbxNombreClasses.Text, @"^\d+$")) &&
                (cbxAnnée.Text != null))
            {
                btnValiderConfig.Enabled = true;
            }
            else
            {
                btnValiderConfig.Enabled = false;
            }
        }

        public void btnPP_Click(object sender, EventArgs e)
        {
            char classe = 'A';
            for (int nbPp = 0; nbPp < NbDivisions; nbPp++)
            {
                NomDuPp.Add(new Label());
                NomDuPp[nbPp].Name = "PP_" + Division + classe;
                classe++;
            }

            Form2 form2 = new Form2();
            form2.Show();
        }

        private void VérifierOptionsClasses()
        {
            string[] optionsTrouvees = new string[10];
            int i = 0;
            int j = 0;

            foreach (DataGridViewColumn classes in ListeBilan.Columns)
            {
                foreach (DataGridViewRow optionClasse in ListeBilan.Rows)
                {
                    optionClasse.Cells[classes.Index].Style.ForeColor = Color.Black;
                }
            }

            foreach (DataGridViewRow optionsMariage in ListeMariagesOptions.Rows)
            {
                foreach (DataGridViewColumn classes in ListeBilan.Columns)
                {
                    foreach (DataGridViewRow optionClasse in ListeBilan.Rows)
                    {
                        if ((optionClasse.Cells[classes.Index].Value != null) && (optionsMariage.Cells[0].Value != null))
                        {
                            int supr = optionClasse.Cells[classes.Index].Value.ToString().IndexOf(" ", StringComparison.Ordinal);
                            string option = optionClasse.Cells[classes.Index].Value.ToString()
                                .Remove(0, supr + 1);
                            if (option == optionsMariage.Cells[0].Value.ToString())
                            {
                                i++;
                            }
                        }
                    }
                }

                if (i > 1)
                {
                    optionsTrouvees[j] = optionsMariage.Cells[0].Value?.ToString();
                    j++;
                }

                i = 0;
            }

            foreach (string item in optionsTrouvees)
            {
                foreach (DataGridViewColumn classes in ListeBilan.Columns)
                {
                    foreach (DataGridViewRow optionClasse in ListeBilan.Rows)
                    {
                        if ((optionClasse.Cells[classes.Index].Value != null) && (item != null))
                        {
                            int supr = optionClasse.Cells[classes.Index].Value.ToString().IndexOf(" ", StringComparison.Ordinal);
                            string option = optionClasse.Cells[classes.Index].Value.ToString()
                                .Remove(0, supr + 1);
                            if (option == item)
                            {
                                optionClasse.Cells[classes.Index].Style.ForeColor = Color.Red;
                            }
                        }
                    }
                }
            }
        }

        private void btnNettoyageFichierExcel_Click(object sender, EventArgs e)
        {
            progressBar1.Visible = true;
            ThreadNettoyage.RunWorkerAsync();
        }

        private void ThreadNettoyageMéthode(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            var excelApplication = new Microsoft.Office.Interop.Excel.Application();

            var fichierEcolesXlsx = excelApplication.Workbooks.Open(lblCheminFichierExcel.Text);
            var feuilleEcoles = (Worksheet)fichierEcolesXlsx.ActiveSheet;
            DernierRang = feuilleEcoles.Cells.Find("*", Missing.Value,
                Missing.Value, Missing.Value,
                XlSearchOrder.xlByRows, XlSearchDirection.xlPrevious,
                false, Missing.Value, Missing.Value).Row;

            Range = feuilleEcoles.Range["A5:M" + DernierRang];
            NbDoublons = 0;

            for (IndexRang = 1; IndexRang < DernierRang - 3; IndexRang++)
            {
                for (int i = 1; i < 14; i++)
                {
                    if ((i == 7) && (Range[IndexRang, 7].Value == null))
                    {
                        Range[IndexRang, 7].Value = "PAS DE LV2";
                    }

                    if (Range[IndexRang, i].Value == "Aucune option")
                    {
                        Range[IndexRang, i].Value = "";
                    }

                    for (int j = i + 1; j < 15; j++)
                    {
                        if (Range[IndexRang, j].Value != null)
                        {
                            if (Range[IndexRang, i].Value == Range[IndexRang, j].Value)
                            {
                                Range[IndexRang, j].Value = "";
                                NbDoublons++;
                            }
                        }
                    }

                    
                }
                ThreadNettoyage.ReportProgress(IndexRang);
            }

            fichierEcolesXlsx.Save();
            fichierEcolesXlsx.Close();
            //excelApplication.ActiveWorkbook.Close(false);
            //excelApplication.Quit();
        }

        private void ThreadNettoyageProgression(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            progressBar1.Maximum = DernierRang;
            // Change the value of the ProgressBar to the BackgroundWorker progress.
            progressBar1.Value = e.ProgressPercentage;
            // Set the text.
            
        }

        private void ThreadNettoyageTerminé(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
             progressBar1.Value = 0;
             lblNbDoublons.Text = NbDoublons + @" doublon(s) corrigé(s)";
        }
    }

    public static class ExtensionMethods
    {
        public static void DoubleBuffered(this DataGridView dgv, bool setting)
        {
            Type dgvType = dgv.GetType();
            PropertyInfo pi = dgvType.GetProperty("DoubleBuffered",
                BindingFlags.Instance | BindingFlags.NonPublic);
            if (pi != null) pi.SetValue(dgv, setting, null);
        }
    }

    internal static class SubstringExtensions
    {
        /// <summary>
        /// Get string value between [first] a and [last] b.
        /// </summary>
        public static string Between(this string value, string a, string b)
        {
            int posA = value.IndexOf(a, StringComparison.Ordinal);
            int posB = value.LastIndexOf(b, StringComparison.Ordinal);
            if (posA == -1)
            {
                return "";
            }
            if (posB == -1)
            {
                return "";
            }
            int adjustedPosA = posA + a.Length;
            if (adjustedPosA >= posB)
            {
                return "";
            }
            return value.Substring(adjustedPosA, posB - adjustedPosA);
        }

        /// <summary>
        /// Get string value after [first] a.
        /// </summary>
        public static string Before(this string value, string a)
        {
            int posA = value.IndexOf(a, StringComparison.Ordinal);
            if (posA == -1)
            {
                return "";
            }
            return value.Substring(0, posA);
        }

        /// <summary>
        /// Get string value after [last] a.
        /// </summary>
        public static string After(this string value, string a)
        {
            int posA = value.LastIndexOf(a, StringComparison.Ordinal);
            if (posA == -1)
            {
                return "";
            }
            int adjustedPosA = posA + a.Length;
            if (adjustedPosA >= value.Length)
            {
                return "";
            }
            return value.Substring(adjustedPosA);
        }
    }
}
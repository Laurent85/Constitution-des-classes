﻿namespace Constitution_des_classes
{
    partial class Principal
    {
        /// <summary>
        /// Variable nécessaire au concepteur.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Nettoyage des ressources utilisées.
        /// </summary>
        /// <param name="disposing">true si les ressources managées doivent être supprimées ; sinon, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Code généré par le Concepteur Windows Form

        /// <summary>
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Principal));
            this.btnParcourir = new System.Windows.Forms.Button();
            this.lblClasses = new System.Windows.Forms.Label();
            this.lblCheminFichierExcel = new System.Windows.Forms.Label();
            this.btnValiderConfig = new System.Windows.Forms.Button();
            this.tabPrincipal = new System.Windows.Forms.TabControl();
            this.Configuration = new System.Windows.Forms.TabPage();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.label3 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.btnNettoyageFichier = new System.Windows.Forms.Button();
            this.lblNiveauInit = new System.Windows.Forms.Label();
            this.cbxNombreClasses = new System.Windows.Forms.ComboBox();
            this.chkAffecterEleves = new System.Windows.Forms.CheckBox();
            this.btnPP = new System.Windows.Forms.Button();
            this.lblAnnée = new System.Windows.Forms.Label();
            this.cbxAnnée = new System.Windows.Forms.ComboBox();
            this.btnWord = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.grpBilan = new System.Windows.Forms.GroupBox();
            this.cbxNbAjoutEleves = new System.Windows.Forms.ComboBox();
            this.grpResume = new System.Windows.Forms.GroupBox();
            this.lblNiveau = new System.Windows.Forms.Label();
            this.lblNbGroupesOptions = new System.Windows.Forms.Label();
            this.lblNbOptions = new System.Windows.Forms.Label();
            this.lblNbClasses = new System.Windows.Forms.Label();
            this.lblGarcons = new System.Windows.Forms.Label();
            this.lblFilles = new System.Windows.Forms.Label();
            this.lblTotalEleves = new System.Windows.Forms.Label();
            this.grpMariagesOptions = new System.Windows.Forms.GroupBox();
            this.grpOptions = new System.Windows.Forms.GroupBox();
            this.grpEffectifs = new System.Windows.Forms.GroupBox();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.ThreadNettoyage = new System.ComponentModel.BackgroundWorker();
            this.lblNbDoublons = new System.Windows.Forms.Label();
            this.tabPrincipal.SuspendLayout();
            this.Configuration.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.panel1.SuspendLayout();
            this.grpResume.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnParcourir
            // 
            this.btnParcourir.Location = new System.Drawing.Point(21, 64);
            this.btnParcourir.Name = "btnParcourir";
            this.btnParcourir.Size = new System.Drawing.Size(137, 23);
            this.btnParcourir.TabIndex = 0;
            this.btnParcourir.Text = "Fichier des élèves...";
            this.btnParcourir.UseVisualStyleBackColor = true;
            this.btnParcourir.Click += new System.EventHandler(this.btn_Parcourir);
            // 
            // lblClasses
            // 
            this.lblClasses.AutoSize = true;
            this.lblClasses.Location = new System.Drawing.Point(19, 148);
            this.lblClasses.Name = "lblClasses";
            this.lblClasses.Size = new System.Drawing.Size(110, 13);
            this.lblClasses.TabIndex = 1;
            this.lblClasses.Text = "Combien de classes ?";
            // 
            // lblCheminFichierExcel
            // 
            this.lblCheminFichierExcel.AutoSize = true;
            this.lblCheminFichierExcel.Location = new System.Drawing.Point(179, 69);
            this.lblCheminFichierExcel.Name = "lblCheminFichierExcel";
            this.lblCheminFichierExcel.Size = new System.Drawing.Size(142, 13);
            this.lblCheminFichierExcel.TabIndex = 3;
            this.lblCheminFichierExcel.Text = "Chemin du fichier des élèves";
            this.lblCheminFichierExcel.TextChanged += new System.EventHandler(this.ChangementLblChemin);
            // 
            // btnValiderConfig
            // 
            this.btnValiderConfig.Location = new System.Drawing.Point(21, 191);
            this.btnValiderConfig.Name = "btnValiderConfig";
            this.btnValiderConfig.Size = new System.Drawing.Size(137, 23);
            this.btnValiderConfig.TabIndex = 4;
            this.btnValiderConfig.Text = "Valider";
            this.btnValiderConfig.UseVisualStyleBackColor = true;
            this.btnValiderConfig.Click += new System.EventHandler(this.btn_Valider_Config);
            // 
            // tabPrincipal
            // 
            this.tabPrincipal.Controls.Add(this.Configuration);
            this.tabPrincipal.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabPrincipal.Location = new System.Drawing.Point(0, 0);
            this.tabPrincipal.Name = "tabPrincipal";
            this.tabPrincipal.SelectedIndex = 0;
            this.tabPrincipal.Size = new System.Drawing.Size(1652, 1023);
            this.tabPrincipal.TabIndex = 5;
            // 
            // Configuration
            // 
            this.Configuration.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.Configuration.Controls.Add(this.pictureBox1);
            this.Configuration.Controls.Add(this.label3);
            this.Configuration.Controls.Add(this.panel1);
            this.Configuration.Controls.Add(this.label1);
            this.Configuration.Controls.Add(this.grpBilan);
            this.Configuration.Controls.Add(this.cbxNbAjoutEleves);
            this.Configuration.Controls.Add(this.grpResume);
            this.Configuration.Controls.Add(this.grpMariagesOptions);
            this.Configuration.Controls.Add(this.grpOptions);
            this.Configuration.Controls.Add(this.grpEffectifs);
            this.Configuration.Location = new System.Drawing.Point(4, 22);
            this.Configuration.Name = "Configuration";
            this.Configuration.Padding = new System.Windows.Forms.Padding(3);
            this.Configuration.Size = new System.Drawing.Size(1644, 997);
            this.Configuration.TabIndex = 0;
            this.Configuration.Text = "Tableau de bord";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.InitialImage = global::Constitution_des_classes.Properties.Resources.LOGO1;
            this.pictureBox1.Location = new System.Drawing.Point(31, 18);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(158, 69);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 18;
            this.pictureBox1.TabStop = false;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(1416, 51);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(99, 13);
            this.label3.TabIndex = 17;
            this.label3.Text = "Combien d\'élèves ?";
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.LightGray;
            this.panel1.Controls.Add(this.lblNbDoublons);
            this.panel1.Controls.Add(this.progressBar1);
            this.panel1.Controls.Add(this.btnNettoyageFichier);
            this.panel1.Controls.Add(this.lblNiveauInit);
            this.panel1.Controls.Add(this.cbxNombreClasses);
            this.panel1.Controls.Add(this.chkAffecterEleves);
            this.panel1.Controls.Add(this.btnPP);
            this.panel1.Controls.Add(this.lblAnnée);
            this.panel1.Controls.Add(this.cbxAnnée);
            this.panel1.Controls.Add(this.btnWord);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.btnParcourir);
            this.panel1.Controls.Add(this.lblClasses);
            this.panel1.Controls.Add(this.lblCheminFichierExcel);
            this.panel1.Controls.Add(this.btnValiderConfig);
            this.panel1.Location = new System.Drawing.Point(31, 106);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(813, 230);
            this.panel1.TabIndex = 16;
            // 
            // btnNettoyageFichier
            // 
            this.btnNettoyageFichier.Location = new System.Drawing.Point(21, 93);
            this.btnNettoyageFichier.Name = "btnNettoyageFichier";
            this.btnNettoyageFichier.Size = new System.Drawing.Size(137, 23);
            this.btnNettoyageFichier.TabIndex = 19;
            this.btnNettoyageFichier.Text = "Nettoyer le fichier excel";
            this.btnNettoyageFichier.UseVisualStyleBackColor = true;
            this.btnNettoyageFichier.Click += new System.EventHandler(this.btnNettoyageFichierExcel_Click);
            // 
            // lblNiveauInit
            // 
            this.lblNiveauInit.AutoSize = true;
            this.lblNiveauInit.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblNiveauInit.ForeColor = System.Drawing.Color.CornflowerBlue;
            this.lblNiveauInit.Location = new System.Drawing.Point(311, 12);
            this.lblNiveauInit.Name = "lblNiveauInit";
            this.lblNiveauInit.Size = new System.Drawing.Size(0, 26);
            this.lblNiveauInit.TabIndex = 12;
            // 
            // cbxNombreClasses
            // 
            this.cbxNombreClasses.FormattingEnabled = true;
            this.cbxNombreClasses.Items.AddRange(new object[] {
            "1",
            "2",
            "3",
            "4",
            "5",
            "6"});
            this.cbxNombreClasses.Location = new System.Drawing.Point(146, 145);
            this.cbxNombreClasses.Name = "cbxNombreClasses";
            this.cbxNombreClasses.Size = new System.Drawing.Size(48, 21);
            this.cbxNombreClasses.TabIndex = 11;
            this.cbxNombreClasses.SelectedIndexChanged += new System.EventHandler(this.cbxNombreClasses_SelectedIndexChanged);
            // 
            // chkAffecterEleves
            // 
            this.chkAffecterEleves.AutoSize = true;
            this.chkAffecterEleves.Location = new System.Drawing.Point(654, 110);
            this.chkAffecterEleves.Name = "chkAffecterEleves";
            this.chkAffecterEleves.Size = new System.Drawing.Size(131, 30);
            this.chkAffecterEleves.TabIndex = 10;
            this.chkAffecterEleves.Text = "  Affecter des élèves\r\n  aux options en rouge";
            this.chkAffecterEleves.UseVisualStyleBackColor = true;
            // 
            // btnPP
            // 
            this.btnPP.Location = new System.Drawing.Point(649, 64);
            this.btnPP.Name = "btnPP";
            this.btnPP.Size = new System.Drawing.Size(128, 23);
            this.btnPP.TabIndex = 9;
            this.btnPP.Text = "Affecter les PP";
            this.btnPP.UseVisualStyleBackColor = true;
            this.btnPP.Click += new System.EventHandler(this.btnPP_Click);
            // 
            // lblAnnée
            // 
            this.lblAnnée.AutoSize = true;
            this.lblAnnée.Location = new System.Drawing.Point(225, 148);
            this.lblAnnée.Name = "lblAnnée";
            this.lblAnnée.Size = new System.Drawing.Size(77, 13);
            this.lblAnnée.TabIndex = 8;
            this.lblAnnée.Text = "Année scolaire";
            // 
            // cbxAnnée
            // 
            this.cbxAnnée.FormattingEnabled = true;
            this.cbxAnnée.Items.AddRange(new object[] {
            "2021-2022",
            "2022-2023",
            "2023-2024",
            "2024-2025",
            "2025-2026",
            "2026-2027",
            "2027-2028",
            "2028-2029",
            "2029-2030"});
            this.cbxAnnée.Location = new System.Drawing.Point(316, 145);
            this.cbxAnnée.Name = "cbxAnnée";
            this.cbxAnnée.Size = new System.Drawing.Size(78, 21);
            this.cbxAnnée.TabIndex = 7;
            this.cbxAnnée.SelectedIndexChanged += new System.EventHandler(this.cbxAnnée_SelectedIndexChanged);
            // 
            // btnWord
            // 
            this.btnWord.Location = new System.Drawing.Point(649, 161);
            this.btnWord.Name = "btnWord";
            this.btnWord.Size = new System.Drawing.Size(128, 23);
            this.btnWord.TabIndex = 6;
            this.btnWord.Text = "Enregistrer sous Word";
            this.btnWord.UseMnemonic = false;
            this.btnWord.UseVisualStyleBackColor = true;
            this.btnWord.Click += new System.EventHandler(this.btnWord_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.CornflowerBlue;
            this.label2.Location = new System.Drawing.Point(27, 12);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(140, 26);
            this.label2.TabIndex = 5;
            this.label2.Text = "Initialisation";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Forte", 30F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.CornflowerBlue;
            this.label1.Location = new System.Drawing.Point(472, 27);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(460, 44);
            this.label1.TabIndex = 15;
            this.label1.Text = "Constitution des classes";
            // 
            // grpBilan
            // 
            this.grpBilan.BackColor = System.Drawing.Color.LightGray;
            this.grpBilan.Location = new System.Drawing.Point(31, 776);
            this.grpBilan.Name = "grpBilan";
            this.grpBilan.Size = new System.Drawing.Size(1567, 195);
            this.grpBilan.TabIndex = 14;
            this.grpBilan.TabStop = false;
            this.grpBilan.Text = "Bilan";
            // 
            // cbxNbAjoutEleves
            // 
            this.cbxNbAjoutEleves.BackColor = System.Drawing.Color.LightGray;
            this.cbxNbAjoutEleves.FormattingEnabled = true;
            this.cbxNbAjoutEleves.ItemHeight = 13;
            this.cbxNbAjoutEleves.Items.AddRange(new object[] {
            "Maxi",
            "1",
            "2",
            "3",
            "4",
            "5",
            "6",
            "7",
            "8",
            "9",
            "10",
            "11",
            "12",
            "13",
            "14",
            "15",
            "16",
            "17",
            "18",
            "19",
            "20",
            "21",
            "22",
            "23",
            "24",
            "25",
            "26",
            "27",
            "28",
            "29",
            "30",
            "31"});
            this.cbxNbAjoutEleves.Location = new System.Drawing.Point(1531, 43);
            this.cbxNbAjoutEleves.MaxDropDownItems = 40;
            this.cbxNbAjoutEleves.Name = "cbxNbAjoutEleves";
            this.cbxNbAjoutEleves.Size = new System.Drawing.Size(67, 21);
            this.cbxNbAjoutEleves.TabIndex = 13;
            this.cbxNbAjoutEleves.SelectedIndexChanged += new System.EventHandler(this.cbxNbAjoutEleves_SelectedIndexChanged);
            // 
            // grpResume
            // 
            this.grpResume.BackColor = System.Drawing.Color.LightGray;
            this.grpResume.Controls.Add(this.lblNiveau);
            this.grpResume.Controls.Add(this.lblNbGroupesOptions);
            this.grpResume.Controls.Add(this.lblNbOptions);
            this.grpResume.Controls.Add(this.lblNbClasses);
            this.grpResume.Controls.Add(this.lblGarcons);
            this.grpResume.Controls.Add(this.lblFilles);
            this.grpResume.Controls.Add(this.lblTotalEleves);
            this.grpResume.Location = new System.Drawing.Point(31, 366);
            this.grpResume.Name = "grpResume";
            this.grpResume.Size = new System.Drawing.Size(207, 390);
            this.grpResume.TabIndex = 12;
            this.grpResume.TabStop = false;
            this.grpResume.Text = "Résumé";
            // 
            // lblNiveau
            // 
            this.lblNiveau.AutoSize = true;
            this.lblNiveau.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblNiveau.ForeColor = System.Drawing.Color.Blue;
            this.lblNiveau.Location = new System.Drawing.Point(30, 48);
            this.lblNiveau.Name = "lblNiveau";
            this.lblNiveau.Size = new System.Drawing.Size(79, 25);
            this.lblNiveau.TabIndex = 11;
            this.lblNiveau.Text = "Niveau";
            // 
            // lblNbGroupesOptions
            // 
            this.lblNbGroupesOptions.AutoSize = true;
            this.lblNbGroupesOptions.Location = new System.Drawing.Point(32, 274);
            this.lblNbGroupesOptions.Name = "lblNbGroupesOptions";
            this.lblNbGroupesOptions.Size = new System.Drawing.Size(145, 13);
            this.lblNbGroupesOptions.TabIndex = 10;
            this.lblNbGroupesOptions.Text = "Nombre de groupes d\'options";
            // 
            // lblNbOptions
            // 
            this.lblNbOptions.AutoSize = true;
            this.lblNbOptions.Location = new System.Drawing.Point(32, 243);
            this.lblNbOptions.Name = "lblNbOptions";
            this.lblNbOptions.Size = new System.Drawing.Size(89, 13);
            this.lblNbOptions.TabIndex = 9;
            this.lblNbOptions.Text = "Nombre d\'options";
            // 
            // lblNbClasses
            // 
            this.lblNbClasses.AutoSize = true;
            this.lblNbClasses.Location = new System.Drawing.Point(32, 211);
            this.lblNbClasses.Name = "lblNbClasses";
            this.lblNbClasses.Size = new System.Drawing.Size(97, 13);
            this.lblNbClasses.TabIndex = 8;
            this.lblNbClasses.Text = "Nombre de classes";
            // 
            // lblGarcons
            // 
            this.lblGarcons.AutoSize = true;
            this.lblGarcons.Location = new System.Drawing.Point(29, 144);
            this.lblGarcons.Name = "lblGarcons";
            this.lblGarcons.Size = new System.Drawing.Size(72, 13);
            this.lblGarcons.TabIndex = 5;
            this.lblGarcons.Text = "Total garçons";
            // 
            // lblFilles
            // 
            this.lblFilles.AutoSize = true;
            this.lblFilles.Location = new System.Drawing.Point(29, 170);
            this.lblFilles.Name = "lblFilles";
            this.lblFilles.Size = new System.Drawing.Size(54, 13);
            this.lblFilles.TabIndex = 6;
            this.lblFilles.Text = "Total filles";
            // 
            // lblTotalEleves
            // 
            this.lblTotalEleves.AutoSize = true;
            this.lblTotalEleves.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTotalEleves.Location = new System.Drawing.Point(29, 109);
            this.lblTotalEleves.Name = "lblTotalEleves";
            this.lblTotalEleves.Size = new System.Drawing.Size(77, 13);
            this.lblTotalEleves.TabIndex = 7;
            this.lblTotalEleves.Text = "Total élèves";
            // 
            // grpMariagesOptions
            // 
            this.grpMariagesOptions.BackColor = System.Drawing.Color.LightGray;
            this.grpMariagesOptions.Location = new System.Drawing.Point(889, 104);
            this.grpMariagesOptions.Name = "grpMariagesOptions";
            this.grpMariagesOptions.Size = new System.Drawing.Size(709, 652);
            this.grpMariagesOptions.TabIndex = 11;
            this.grpMariagesOptions.TabStop = false;
            this.grpMariagesOptions.Text = "Mariages d\'options";
            // 
            // grpOptions
            // 
            this.grpOptions.BackColor = System.Drawing.Color.LightGray;
            this.grpOptions.Location = new System.Drawing.Point(569, 366);
            this.grpOptions.Name = "grpOptions";
            this.grpOptions.Size = new System.Drawing.Size(275, 390);
            this.grpOptions.TabIndex = 10;
            this.grpOptions.TabStop = false;
            this.grpOptions.Text = "Options";
            // 
            // grpEffectifs
            // 
            this.grpEffectifs.BackColor = System.Drawing.Color.LightGray;
            this.grpEffectifs.Location = new System.Drawing.Point(272, 366);
            this.grpEffectifs.Name = "grpEffectifs";
            this.grpEffectifs.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.grpEffectifs.Size = new System.Drawing.Size(264, 390);
            this.grpEffectifs.TabIndex = 9;
            this.grpEffectifs.TabStop = false;
            this.grpEffectifs.Text = "Effectifs classes";
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(182, 92);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(175, 23);
            this.progressBar1.Step = 1;
            this.progressBar1.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.progressBar1.TabIndex = 20;
            // 
            // ThreadNettoyage
            // 
            this.ThreadNettoyage.WorkerReportsProgress = true;
            this.ThreadNettoyage.DoWork += new System.ComponentModel.DoWorkEventHandler(this.ThreadNettoyageMéthode);
            this.ThreadNettoyage.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.ThreadNettoyageProgression);
            this.ThreadNettoyage.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.ThreadNettoyageTerminé);
            // 
            // lblNbDoublons
            // 
            this.lblNbDoublons.AutoSize = true;
            this.lblNbDoublons.Location = new System.Drawing.Point(373, 98);
            this.lblNbDoublons.Name = "lblNbDoublons";
            this.lblNbDoublons.Size = new System.Drawing.Size(0, 13);
            this.lblNbDoublons.TabIndex = 21;
            // 
            // Principal
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1652, 1023);
            this.Controls.Add(this.tabPrincipal);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Principal";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Constitution des classes";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.tabPrincipal.ResumeLayout(false);
            this.Configuration.ResumeLayout(false);
            this.Configuration.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.grpResume.ResumeLayout(false);
            this.grpResume.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnParcourir;
        private System.Windows.Forms.Label lblClasses;
        private System.Windows.Forms.Label lblCheminFichierExcel;
        private System.Windows.Forms.Button btnValiderConfig;
        private System.Windows.Forms.TabControl tabPrincipal;
        private System.Windows.Forms.TabPage Configuration;
        private System.Windows.Forms.Label lblFilles;
        private System.Windows.Forms.Label lblGarcons;
        private System.Windows.Forms.Label lblTotalEleves;
        private System.Windows.Forms.GroupBox grpEffectifs;
        private System.Windows.Forms.GroupBox grpOptions;
        private System.Windows.Forms.GroupBox grpResume;
        private System.Windows.Forms.GroupBox grpMariagesOptions;
        private System.Windows.Forms.ComboBox cbxNbAjoutEleves;
        private System.Windows.Forms.GroupBox grpBilan;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lblNbGroupesOptions;
        private System.Windows.Forms.Label lblNbOptions;
        private System.Windows.Forms.Label lblNbClasses;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnWord;
        private System.Windows.Forms.Label lblAnnée;
        private System.Windows.Forms.ComboBox cbxAnnée;
        private System.Windows.Forms.Button btnPP;
        private System.Windows.Forms.CheckBox chkAffecterEleves;
        private System.Windows.Forms.Label lblNiveau;
        private System.Windows.Forms.ComboBox cbxNombreClasses;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label lblNiveauInit;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Button btnNettoyageFichier;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.ComponentModel.BackgroundWorker ThreadNettoyage;
        private System.Windows.Forms.Label lblNbDoublons;
    }
}


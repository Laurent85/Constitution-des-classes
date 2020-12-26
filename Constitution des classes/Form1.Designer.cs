namespace Constitution_des_classes
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
            this.btnParcourir = new System.Windows.Forms.Button();
            this.lblClasses = new System.Windows.Forms.Label();
            this.txbNombreClasses = new System.Windows.Forms.TextBox();
            this.lblCheminFichierExcel = new System.Windows.Forms.Label();
            this.btnValiderConfig = new System.Windows.Forms.Button();
            this.tabPrincipal = new System.Windows.Forms.TabControl();
            this.Configuration = new System.Windows.Forms.TabPage();
            this.lblFilles = new System.Windows.Forms.Label();
            this.lblGarcons = new System.Windows.Forms.Label();
            this.lblTotalEleves = new System.Windows.Forms.Label();
            this.btn_Constituer_Classes = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.tabPrincipal.SuspendLayout();
            this.Configuration.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnParcourir
            // 
            this.btnParcourir.Location = new System.Drawing.Point(6, 20);
            this.btnParcourir.Name = "btnParcourir";
            this.btnParcourir.Size = new System.Drawing.Size(116, 23);
            this.btnParcourir.TabIndex = 0;
            this.btnParcourir.Text = "Fichier des élèves...";
            this.btnParcourir.UseVisualStyleBackColor = true;
            this.btnParcourir.Click += new System.EventHandler(this.btn_Parcourir);
            // 
            // lblClasses
            // 
            this.lblClasses.AutoSize = true;
            this.lblClasses.Location = new System.Drawing.Point(12, 66);
            this.lblClasses.Name = "lblClasses";
            this.lblClasses.Size = new System.Drawing.Size(110, 13);
            this.lblClasses.TabIndex = 1;
            this.lblClasses.Text = "Combien de classes ?";
            // 
            // txbNombreClasses
            // 
            this.txbNombreClasses.Location = new System.Drawing.Point(146, 63);
            this.txbNombreClasses.Name = "txbNombreClasses";
            this.txbNombreClasses.Size = new System.Drawing.Size(57, 20);
            this.txbNombreClasses.TabIndex = 2;
            // 
            // lblCheminFichierExcel
            // 
            this.lblCheminFichierExcel.AutoSize = true;
            this.lblCheminFichierExcel.Location = new System.Drawing.Point(143, 30);
            this.lblCheminFichierExcel.Name = "lblCheminFichierExcel";
            this.lblCheminFichierExcel.Size = new System.Drawing.Size(142, 13);
            this.lblCheminFichierExcel.TabIndex = 3;
            this.lblCheminFichierExcel.Text = "Chemin du fichier des élèves";
            // 
            // btnValiderConfig
            // 
            this.btnValiderConfig.Location = new System.Drawing.Point(15, 111);
            this.btnValiderConfig.Name = "btnValiderConfig";
            this.btnValiderConfig.Size = new System.Drawing.Size(75, 23);
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
            this.tabPrincipal.Size = new System.Drawing.Size(1380, 591);
            this.tabPrincipal.TabIndex = 5;
            // 
            // Configuration
            // 
            this.Configuration.Controls.Add(this.groupBox1);
            this.Configuration.Controls.Add(this.btn_Constituer_Classes);
            this.Configuration.Controls.Add(this.lblTotalEleves);
            this.Configuration.Controls.Add(this.lblFilles);
            this.Configuration.Controls.Add(this.lblGarcons);
            this.Configuration.Controls.Add(this.btnParcourir);
            this.Configuration.Controls.Add(this.btnValiderConfig);
            this.Configuration.Controls.Add(this.lblCheminFichierExcel);
            this.Configuration.Controls.Add(this.txbNombreClasses);
            this.Configuration.Controls.Add(this.lblClasses);
            this.Configuration.Location = new System.Drawing.Point(4, 22);
            this.Configuration.Name = "Configuration";
            this.Configuration.Padding = new System.Windows.Forms.Padding(3);
            this.Configuration.Size = new System.Drawing.Size(1372, 565);
            this.Configuration.TabIndex = 0;
            this.Configuration.Text = "Configuration";
            this.Configuration.UseVisualStyleBackColor = true;
            // 
            // lblFilles
            // 
            this.lblFilles.AutoSize = true;
            this.lblFilles.Location = new System.Drawing.Point(173, 260);
            this.lblFilles.Name = "lblFilles";
            this.lblFilles.Size = new System.Drawing.Size(35, 13);
            this.lblFilles.TabIndex = 6;
            this.lblFilles.Text = "label2";
            // 
            // lblGarcons
            // 
            this.lblGarcons.AutoSize = true;
            this.lblGarcons.Location = new System.Drawing.Point(173, 223);
            this.lblGarcons.Name = "lblGarcons";
            this.lblGarcons.Size = new System.Drawing.Size(35, 13);
            this.lblGarcons.TabIndex = 5;
            this.lblGarcons.Text = "label1";
            // 
            // lblTotalEleves
            // 
            this.lblTotalEleves.AutoSize = true;
            this.lblTotalEleves.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTotalEleves.Location = new System.Drawing.Point(173, 290);
            this.lblTotalEleves.Name = "lblTotalEleves";
            this.lblTotalEleves.Size = new System.Drawing.Size(41, 13);
            this.lblTotalEleves.TabIndex = 7;
            this.lblTotalEleves.Text = "label3";
            // 
            // btn_Constituer_Classes
            // 
            this.btn_Constituer_Classes.Location = new System.Drawing.Point(176, 369);
            this.btn_Constituer_Classes.Name = "btn_Constituer_Classes";
            this.btn_Constituer_Classes.Size = new System.Drawing.Size(122, 23);
            this.btn_Constituer_Classes.TabIndex = 8;
            this.btn_Constituer_Classes.Text = "Constituer les classes";
            this.btn_Constituer_Classes.UseVisualStyleBackColor = true;
            this.btn_Constituer_Classes.Click += new System.EventHandler(this.btn_Constituer_Classes_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Location = new System.Drawing.Point(692, 192);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(328, 263);
            this.groupBox1.TabIndex = 9;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "groupBox1";
            // 
            // Principal
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1380, 591);
            this.Controls.Add(this.tabPrincipal);
            this.Name = "Principal";
            this.Text = "Constitution des classes";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.tabPrincipal.ResumeLayout(false);
            this.Configuration.ResumeLayout(false);
            this.Configuration.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnParcourir;
        private System.Windows.Forms.Label lblClasses;
        private System.Windows.Forms.TextBox txbNombreClasses;
        private System.Windows.Forms.Label lblCheminFichierExcel;
        private System.Windows.Forms.Button btnValiderConfig;
        private System.Windows.Forms.TabControl tabPrincipal;
        private System.Windows.Forms.TabPage Configuration;
        private System.Windows.Forms.Label lblFilles;
        private System.Windows.Forms.Label lblGarcons;
        private System.Windows.Forms.Label lblTotalEleves;
        private System.Windows.Forms.Button btn_Constituer_Classes;
        private System.Windows.Forms.GroupBox groupBox1;
    }
}


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
            this.tabControl = new System.Windows.Forms.TabControl();
            this.Configuration = new System.Windows.Forms.TabPage();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.tabControl.SuspendLayout();
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
            // tabControl
            // 
            this.tabControl.Controls.Add(this.Configuration);
            this.tabControl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl.Location = new System.Drawing.Point(0, 0);
            this.tabControl.Name = "tabControl";
            this.tabControl.SelectedIndex = 0;
            this.tabControl.Size = new System.Drawing.Size(1380, 591);
            this.tabControl.TabIndex = 5;
            // 
            // Configuration
            // 
            this.Configuration.Controls.Add(this.label2);
            this.Configuration.Controls.Add(this.label1);
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
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(173, 223);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(35, 13);
            this.label1.TabIndex = 5;
            this.label1.Text = "label1";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(173, 260);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(35, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "label2";
            // 
            // Principal
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1380, 591);
            this.Controls.Add(this.tabControl);
            this.Name = "Principal";
            this.Text = "Constitution des classes";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.tabControl.ResumeLayout(false);
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
        private System.Windows.Forms.TabControl tabControl;
        private System.Windows.Forms.TabPage Configuration;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
    }
}


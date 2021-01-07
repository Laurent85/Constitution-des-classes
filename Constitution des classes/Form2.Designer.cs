
namespace Constitution_des_classes
{
    partial class Form2
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.panelPP = new System.Windows.Forms.Panel();
            this.btnValiderPp = new System.Windows.Forms.Button();
            this.panelPP.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.RoyalBlue;
            this.label1.Location = new System.Drawing.Point(57, 38);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(413, 26);
            this.label1.TabIndex = 0;
            this.label1.Text = "Attribution des professeurs principaux";
            // 
            // panelPP
            // 
            this.panelPP.Controls.Add(this.btnValiderPp);
            this.panelPP.Location = new System.Drawing.Point(62, 112);
            this.panelPP.Name = "panelPP";
            this.panelPP.Size = new System.Drawing.Size(411, 401);
            this.panelPP.TabIndex = 1;
            // 
            // btnValiderPp
            // 
            this.btnValiderPp.Location = new System.Drawing.Point(141, 364);
            this.btnValiderPp.Name = "btnValiderPp";
            this.btnValiderPp.Size = new System.Drawing.Size(128, 23);
            this.btnValiderPp.TabIndex = 0;
            this.btnValiderPp.Text = "Valider les PP";
            this.btnValiderPp.UseVisualStyleBackColor = true;
            this.btnValiderPp.Click += new System.EventHandler(this.btnValiderPp_Click);
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(535, 559);
            this.Controls.Add(this.panelPP);
            this.Controls.Add(this.label1);
            this.Name = "Form2";
            this.Text = "Form2";
            this.Load += new System.EventHandler(this.Form2_Load);
            this.Shown += new System.EventHandler(this.Form2_Load);
            this.panelPP.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Panel panelPP;
        private System.Windows.Forms.Button btnValiderPp;
    }
}
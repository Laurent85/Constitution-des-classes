using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace Constitution_des_classes
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        public Principal Principal = new Principal();

        public List<System.Windows.Forms.TextBox> TxbPp = new List<System.Windows.Forms.TextBox>();
        public List<Label> LblClassePp = new List<Label>();
        public List<TextBox> TxbClassePp = new List<TextBox>();

        public void Form2_Load(object sender, EventArgs e)
        {
            char classe = 'A';
            int y = 40;

            for (int i = 0; i < Principal.NbDivisions; i++)
            {
                LblClassePp.Add(new Label());
                LblClassePp[i].Name = "_lblClassePp_" + i;
                LblClassePp[i].Text = Principal.Division + classe;
                LblClassePp[i].Location = new System.Drawing.Point(100, y);
                panelPP.Controls.Add(LblClassePp[i]);

                TxbClassePp.Add(new TextBox());
                TxbClassePp[i].Name = "_txbClassePp_" + i;
                TxbClassePp[i].Text = Principal.ListePp[i];
                TxbClassePp[i].Location = new System.Drawing.Point(200, y);
                TxbClassePp[i].Width = 150;
                panelPP.Controls.Add(TxbClassePp[i]);
                y = y + 50;
                classe++;
            }
        }

        private void btnValiderPp_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < Principal.NbDivisions; i++)
            {
                Principal.ListePp[i] = TxbClassePp[i].Text;
            }

            this.Close();
        }
    }
}
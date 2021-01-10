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

        public List<Label> LblClassePp = new List<Label>();
        public List<ComboBox> CbxClassePp = new List<ComboBox>();

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

                CbxClassePp.Add(new ComboBox());
                CbxClassePp[i].Name = "_txbClassePp_" + i;
                CbxClassePp[i].Location = new System.Drawing.Point(200, y);
                CbxClassePp[i].Width = 150;
                panelPP.Controls.Add(CbxClassePp[i]);
                string resourceData = Properties.Resources.Profs;
                string[] words = resourceData.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string lignes in words)
                {
                    CbxClassePp[i].Items.Add(lignes);
                }

                CbxClassePp[i].Text = Principal.ListePp[i];
                y = y + 50;
                classe++;
            }
        }

        private void btnValiderPp_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < Principal.NbDivisions; i++)
            {
                Principal.ListePp[i] = CbxClassePp[i].Text;
            }

            this.Close();
        }
    }
}
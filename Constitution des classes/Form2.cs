using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Constitution_des_classes
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private readonly Principal _principal = new Principal();
        
        private readonly List<System.Windows.Forms.TextBox> _txbPp = new List<System.Windows.Forms.TextBox>();
        private readonly List<Label> _lblClassePp = new List<Label>();

        private void Form2_Load(object sender, EventArgs e)
        {
            _principal.NomDuPp[0].Text = "";
            char classe = 'A';
            int y = 40;

            for (int i = 0; i < _principal.NbDivisions; i++)
            {
                _lblClassePp.Add(new Label());
                _lblClassePp[i].Name = "_lblClassePp_" + i;
                _lblClassePp[i].Text = _principal.Division + classe;
                _lblClassePp[i].Location = new System.Drawing.Point(20, y);
                panelPP.Controls.Add(_lblClassePp[i]);
                classe++;
            }
        }
    }
}

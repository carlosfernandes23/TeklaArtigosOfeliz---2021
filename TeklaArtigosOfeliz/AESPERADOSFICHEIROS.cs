using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TeklaArtigosOfeliz
{
    
    public partial class AESPERADOSFICHEIROS : Form
    {
        public static bool cancel = false;
        public AESPERADOSFICHEIROS(string Peça)
        {
            InitializeComponent();
            timer1.Enabled = true;
            label2.Text = Peça;
        }
        int X = 10;
        private void timer1_Tick(object sender, EventArgs e)
        {
            X = X - 1;
            label3.Text = "A tentar novamente dentro de " + X + " S";
            if (X==0)
            {
                this.Close();                         
            }
         
        }

        private void button1_Click(object sender, EventArgs e)
        {
            cancel = true;
            this.Close();
        }

        private void AESPERADOSFICHEIROS_Load(object sender, EventArgs e)
        {

        }
    }
}

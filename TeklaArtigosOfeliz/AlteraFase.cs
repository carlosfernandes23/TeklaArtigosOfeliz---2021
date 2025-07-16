using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tekla.Structures.Model;

namespace TeklaArtigosOfeliz
{
    public partial class AlteraFase : Form
    {
        public AlteraFase()
        {
            InitializeComponent();
            Model model = new Model();
           PhaseCollection fases = model.GetPhases();
            foreach (Phase item in fases)
            {
                comboBox1.Items.Add(item.PhaseNumber);
                comboBox1.SelectedIndex=0;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

         ArrayList r =  ComunicaTekla.ListadePecasSelec();
            foreach  (Tekla.Structures.Model.Part item in r)
            {
                item.SetPhase(new Phase(int.Parse(comboBox1.Text.ToString())));
            }
            MessageBox.Show(this, "FASE ALTERADA A " +r.Count+"PEÇAS.");
        }

        private void AlteraFase_Load(object sender, EventArgs e)
        {

        }
    }
}

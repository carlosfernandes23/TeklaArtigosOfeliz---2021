using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tekla.Structures.Model;
using TSM = Tekla.Structures.Model;

namespace TeklaArtigosOfeliz
{
    public partial class Frm_ExportarNC1: Form
    {
        public Frm_ExportarNC1()
        {
            InitializeComponent();
        }

        private void Frm_ExportarNC1_Load(object sender, EventArgs e)
        {

        }

       

        private void button9_Click(object sender, EventArgs e)
        {
            LBLestado.Text = "A criar CNC ";
            TSM.Operations.Operation.CreateNCFilesFromSelected("OFELIZ_chapas", PASTAEXPORTACAO.Text + @"\");
            TSM.Operations.Operation.CreateNCFilesFromSelected("OFELIZ_perfis", PASTAEXPORTACAO.Text + @"\");
            TSM.Operations.Operation.CreateNCFilesFromSelected("OFELIZ_madres", PASTAEXPORTACAO.Text + @"\");
            LBLestado.Text = "CNC Criados com sucesso";
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            if (Directory.Exists(PASTAEXPORTACAO.Text))
            {
                string[] NCfiles = Directory.GetFiles(PASTAEXPORTACAO.Text, "*.nc1", SearchOption.TopDirectoryOnly);
                if (NCfiles.Length != 0)
                {
                    List<string> myfiles = new List<string>();
                    foreach (var item in NCfiles)
                    {
                        myfiles.Add(item);
                    }
                    dstv_dxf.CRIAR(myfiles);
                    LBLestado.Text = "Ficheiros convertidos";
                }
                else
                {
                    LBLestado.Text = "Não existe ficheiros nc1";
                }
            }
            else
            {
                LBLestado.Text = "Não existe o Caminho";
            }
        }
    }
}

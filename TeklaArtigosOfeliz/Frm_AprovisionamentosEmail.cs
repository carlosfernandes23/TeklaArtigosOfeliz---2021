using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection.Emit;
using System.Windows.Forms;
using Tekla.Structures.Model;
using Image = System.Drawing.Image;
using Outlook = Microsoft.Office.Interop.Outlook;
using Point = Tekla.Structures.Geometry3d.Point;
using TSM = Tekla.Structures.Model;

namespace TeklaArtigosOfeliz
{
    public partial class Frm_AprovisionamentosEmail : Form
    {
        public Frm_AprovisionamentosEmail()
        {
            InitializeComponent();
            Chb_alw_top.CheckedChanged += Chb_alw_top_CheckedChanged;
            TopMost = Chb_alw_top.Checked;
        }

        private void AprovisionamentosEmail_Load(object sender, EventArgs e)
        {

        }                         

        public void EnviarAprovisionamentosOpenEmailPreviewAndCreateEmail()
        {
            string SubjectEnviarAprovisionamentos = string.Empty;

            try
            {                                                                             
                    
                Model modelo = new Model();
                string PastaModelo = modelo.GetInfo().ModelPath;
                DirectoryInfo up = new DirectoryInfo(PastaModelo);
                string ultimaPasta = up.Name;
                string nomeProjeto = modelo.GetProjectInfo().Name;
                string numeroprojeto = modelo.GetProjectInfo().ProjectNumber;
                string ano10 = string.Empty;

                string imagemOfelizFilePath = @"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\4 Produção\Desenvolvimentos\Ficheiros Temp tekla artigos (Nao Apagar)\ofeliz_logo.png";

                string nomeUsuario = Environment.UserName;
                nomeUsuario = nomeUsuario.Replace('.', ' ');
                nomeUsuario = string.Join(" ", nomeUsuario.Split(' ').Select(p => char.ToUpper(p[0]) + p.Substring(1).ToLower()));

                if (numeroprojeto.Contains("PT"))
                {
                    ano10 = "20" + numeroprojeto.Substring(2, 2);
                }
                else
                {
                    ano10 = "20" + numeroprojeto.Substring(0, 2);
                }
                string caminho = @"\\Marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\11 Partilhada\" + ano10 + @"\DAP\" + numeroprojeto + "\\" + textBox2.Text;

                ultimaPasta = ultimaPasta.Replace("_", "-");

                SubjectEnviarAprovisionamentos = ultimaPasta + " -- APROVISIONAMENTO";

                string saudacao = GetSaudacao();

                string corpoEmail = "<html><body contenteditable=\"false\">";
                corpoEmail += "<font face = 'Calibri ' size = '3' > <p>" + saudacao + "</font></p>";

                corpoEmail += "<font face='Calibri ' size='3'>Venho por este meio informar que já solicitei o aprovisionamento "
                        + "<span style='color:red;'>" + textBox1.Text + "</span>&nbsp"
                        + "da obra em assunto. </font><br>";

                corpoEmail += "<font face = 'Calibri ' size = '3' ><p><b> FASE " + textBox2.Text + ": &nbsp" + "</b>"
                              + "<font face = 'Calibri ' size = '3' ;style='color:#5B9BD5;'>"
                              + "<u><a href='file:///" + caminho + "';style='color:#5B9BD5;'>" + caminho + "</a></u>" + "</font>";


               corpoEmail += "<font face = 'Calibri ' size = '3' > <p> Melhores Cumprimentos,</p> </font> <br>";
               corpoEmail += "<font face = 'Calibri' size = '3' > <b>" + nomeUsuario + "</b> </Font> <br>";
               corpoEmail += "<font face = 'Calibri' size = '3' > Construção Metálica | Preparador </Font> <br>";
               corpoEmail += "<font face = 'Calibri' size = '3' > T + 351 253 080 609 * </font> <br>";
               corpoEmail += "<font color='red' font face = 'Calibri ' size = '3'> ofeliz.com </font> <br>";
               corpoEmail += "<p><a href='https://www.ofeliz.com'><img src='file:///" + imagemOfelizFilePath.Replace("\\", "/") + "' width='127' height='34'></a></p>";

              corpoEmail += "<i><font color='Light grey' font face = 'Calibri ' size = '1.5'> Alvará Nº 10553 – Pub. *Chamada para a rede fixa nacional. </font> </i><br>";
              corpoEmail += "<i><font color='green' font face = 'Calibri ' size = '1.5'> Antes de imprimir este e-mail tenha em consideração o meio ambiente. </font> </i><br>";
              corpoEmail += "</body></html>";

              string textbox1 = textBox1.Text;
              string textbox2 = textBox2.Text;

             this.Visible = false;
             Frm_Corpo_de_Texto_Email_Enviar_Aprovisionamentos previewForm = new Frm_Corpo_de_Texto_Email_Enviar_Aprovisionamentos("Enviar Email de Aprovisionamentos", corpoEmail, SubjectEnviarAprovisionamentos, textbox1, textbox2, caminho);
             previewForm.ShowDialog(this);               
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Erro ao Conectar com o Tekla , tente novamente " + ex.Message);
            }
        }


        private string GetSaudacao()
        {
            DateTime horaAtual = DateTime.Now;
            if (horaAtual.Hour < 12 || (horaAtual.Hour == 12 && horaAtual.Minute < 30))
            {
                return "Bom Dia, ";
            }
            else
            {
                return "Boa Tarde, ";
            }
        }

        private void Chb_alw_top_CheckedChanged(object sender, EventArgs e)
        {
            if (Chb_alw_top.Checked)
            {
                TopMost = true;
            }
            else
            {
                TopMost = false;
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox1.Text) || string.IsNullOrEmpty(textBox2.Text))
            {
                MessageBox.Show(this, "Por favor, preencha os campos de texto antes de enviar o e-mail.", "Alerta", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                EnviarAprovisionamentosOpenEmailPreviewAndCreateEmail();
            }
        }
    }
}

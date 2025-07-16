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
    public partial class Frm_EnviarVariosEmail: Form
    {

        public Frm_EnviarVariosEmail()
        {
            InitializeComponent();
            Chb_alw_top.CheckedChanged += Chb_alw_top_CheckedChanged;
            TopMost = Chb_alw_top.Checked;
        }

        public string richTextBox1Value
        {
            get { return richTextBox1.Text; }
            set { richTextBox1.Text = value; }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(richTextBox1.Text))
            {
                MessageBox.Show(this, "Por favor, preencha os campos de texto antes de enviar o e-mail.", "Alerta", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (comboBox1.SelectedIndex == 0)
            {
                EnviarEmailRevisaoPecasEConjuntos();
            }
            else if (comboBox1.SelectedIndex == 1)
            {
                EnviarEmailRevisãoProjetodeExecução(); 
            }
            else if (comboBox1.SelectedIndex == 2)
            {
                EnviarEmailRevisãoEstudoPrevio();
            }
            else if (comboBox1.SelectedIndex == 3)
            {
                EnviarEmailRevisaoDesenhosdeMontagem();
            }
            else if (comboBox1.SelectedIndex == 4)
            {
                EnviarEmailAprovaçãoEstudoPrevio();
            }
            else if (comboBox1.SelectedIndex == 5)
            {
                EnviarEmailAprovaçãoProjetodeExecução();
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

        public void EnviarEmailRevisaoPecasEConjuntos()
        {
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

                ultimaPasta = ultimaPasta.Replace("_", "-");
                string SubjectEnviar = ultimaPasta + " -- Revisão Peças/Conjuntos";
                string saudacao = GetSaudacao();

                string corpoEmail = "<html><body contenteditable=\"false\">";
                corpoEmail += "<p style=\"font-family: Calibri; font-size: 14px;\">" + saudacao + "</p>";
                corpoEmail += "<p style=\"font-family: Calibri; font-size: 14px;\">Venho por este meio informar que os seguintes peças/conjuntos em anexo foram sujeitos a revisões da obra em assunto. &nbsp;</p>";
                corpoEmail += "<p style=\"font-family: Calibri; font-size: 14px; color: red;\"><b>" + richTextBox1.Text.Replace("/", "<br>") + "</b></p>";
                corpoEmail += "<font face = 'Calibri ' size = '3' > <p> Melhores Cumprimentos,</p> </font> <br>";
                corpoEmail += "<font face = 'Calibri' size = '3' > <b>" + nomeUsuario + "</b> </Font> <br>";
                corpoEmail += "<font face = 'Calibri' size = '3' > Construção Metálica | Preparador </Font> <br>";
                corpoEmail += "<font face = 'Calibri' size = '3' > T + 351 253 080 609 * </font> <br>";
                corpoEmail += "<font color='red' font face = 'Calibri ' size = '3'> ofeliz.com </font> <br>";
                corpoEmail += "<p><a href='https://www.ofeliz.com'><img src='file:///" + imagemOfelizFilePath.Replace("\\", "/") + "' width='127' height='34'></a></p>";
                corpoEmail += "<i><font color='Light grey' font face = 'Calibri ' size = '1.5'> Alvará Nº 10553 – Pub. *Chamada para a rede fixa nacional. </font> </i><br>";
                corpoEmail += "<i><font color='green' font face = 'Calibri ' size = '1.5'> Antes de imprimir este e-mail tenha em consideração o meio ambiente. </font> </i><br>";
                corpoEmail += "</body></html>";
                string richText = richTextBox1.Text;

                this.Visible = false;
                Frm_Corpo_de_Texto_Email_RevisaoPecaseConjuntos previewForm = new Frm_Corpo_de_Texto_Email_RevisaoPecaseConjuntos("Email de Revisões ( Peças / Conjuntos )", corpoEmail, SubjectEnviar, richText);
                previewForm.ShowDialog(this);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Erro ao Conectar com o Tekla , tente novamente " + ex.Message);
            }
        }

        public void EnviarEmailRevisãoProjetodeExecução()
        {
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
                ultimaPasta = ultimaPasta.Replace("_", "-");
                string SubjectEnviar = ultimaPasta + " -- Revisão ao projeto execução";
                string saudacao = GetSaudacao();

                string corpoEmail = "<html><body contenteditable=\"false\">";
                corpoEmail += "<p style=\"font-family: Calibri; font-size: 14px;\">" + saudacao + "</p>";
                corpoEmail += "<span style=\"font-family: Calibri; font-size: 14px;\">Em anexo envio nova revisão ao desenho &nbsp;</span>";
                corpoEmail += "<span style=\"font-family: Calibri; font-size: 14px; color: black;\"><b>" + richTextBox1.Text.Replace("/", "<br>") + "</b></span>";
                corpoEmail += "<span style=\"font-family: Calibri; font-size: 14px;\">&nbsp para aprovação da obra em assunto. &nbsp;</span>";
                corpoEmail += "<font face = 'Calibri ' size = '3' > <p> Melhores Cumprimentos,</p> </font> <br>";
                corpoEmail += "<font face = 'Calibri' size = '3' > <b>" + nomeUsuario + "</b> </Font> <br>";
                corpoEmail += "<font face = 'Calibri' size = '3' > Construção Metálica | Preparador </Font> <br>";
                corpoEmail += "<font face = 'Calibri' size = '3' > T + 351 253 080 609 * </font> <br>";
                corpoEmail += "<font color='red' font face = 'Calibri ' size = '3'> ofeliz.com </font> <br>";
                corpoEmail += "<p><a href='https://www.ofeliz.com'><img src='file:///" + imagemOfelizFilePath.Replace("\\", "/") + "' width='127' height='34'></a></p>";
                corpoEmail += "<i><font color='Light grey' font face = 'Calibri ' size = '1.5'> Alvará Nº 10553 – Pub. *Chamada para a rede fixa nacional. </font> </i><br>";
                corpoEmail += "<i><font color='green' font face = 'Calibri ' size = '1.5'> Antes de imprimir este e-mail tenha em consideração o meio ambiente. </font> </i><br>";
                corpoEmail += "</body></html>";
                string richText = richTextBox1.Text;

                this.Visible = false;
                Frm_Corpo_de_Texto_Email_RevisaoPecaseConjuntos previewForm = new Frm_Corpo_de_Texto_Email_RevisaoPecaseConjuntos("Email de Revisões ( Proj Execução )", corpoEmail, SubjectEnviar, richText);
                previewForm.ShowDialog(this);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Erro ao Conectar com o Tekla , tente novamente " + ex.Message);
            }
        }

        public void EnviarEmailRevisãoEstudoPrevio()
        {
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
                ultimaPasta = ultimaPasta.Replace("_", "-");
                string SubjectEnviar = ultimaPasta + " -- Revisão ao Estudo Prévio";
                string saudacao = GetSaudacao();

                string corpoEmail = "<html><body contenteditable=\"false\">";
                corpoEmail += "<p style=\"font-family: Calibri; font-size: 14px;\">" + saudacao + "</p>";
                corpoEmail += "<span style=\"font-family: Calibri; font-size: 14px;\">Em anexo envio nova revisão ao estudo &nbsp;</span>";
                corpoEmail += "<span style=\"font-family: Calibri; font-size: 14px; color: black;\"><b>" + richTextBox1.Text.Replace("/", "<br>") + "</b></span>";
                corpoEmail += "<span style=\"font-family: Calibri; font-size: 14px;\">&nbsp para aprovação da obra em assunto. &nbsp;</span>";
                corpoEmail += "<font face = 'Calibri ' size = '3' > <p> Melhores Cumprimentos,</p> </font> <br>";
                corpoEmail += "<font face = 'Calibri' size = '3' > <b>" + nomeUsuario + "</b> </Font> <br>";
                corpoEmail += "<font face = 'Calibri' size = '3' > Construção Metálica | Preparador </Font> <br>";
                corpoEmail += "<font face = 'Calibri' size = '3' > T + 351 253 080 609 * </font> <br>";
                corpoEmail += "<font color='red' font face = 'Calibri ' size = '3'> ofeliz.com </font> <br>";
                corpoEmail += "<p><a href='https://www.ofeliz.com'><img src='file:///" + imagemOfelizFilePath.Replace("\\", "/") + "' width='127' height='34'></a></p>";
                corpoEmail += "<i><font color='Light grey' font face = 'Calibri ' size = '1.5'> Alvará Nº 10553 – Pub. *Chamada para a rede fixa nacional. </font> </i><br>";
                corpoEmail += "<i><font color='green' font face = 'Calibri ' size = '1.5'> Antes de imprimir este e-mail tenha em consideração o meio ambiente. </font> </i><br>";
                corpoEmail += "</body></html>";
                string richText = richTextBox1.Text;

                this.Visible = false;
                Frm_Corpo_de_Texto_Email_RevisaoPecaseConjuntos previewForm = new Frm_Corpo_de_Texto_Email_RevisaoPecaseConjuntos("Email de Revisões ( Estudo Prévio )", corpoEmail, SubjectEnviar, richText);
                previewForm.ShowDialog(this);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Erro ao Conectar com o Tekla , tente novamente " + ex.Message);
            }
        }

        public void EnviarEmailRevisaoDesenhosdeMontagem()
        {
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
                ultimaPasta = ultimaPasta.Replace("_", "-");
                string SubjectEnviar = ultimaPasta + " -- Revisão ao desenho de montagem ";
                string saudacao = GetSaudacao();

                string corpoEmail = "<html><body contenteditable=\"false\">";
                corpoEmail += "<p style=\"font-family: Calibri; font-size: 14px;\">" + saudacao + "</p>";
                corpoEmail += "<span style=\"font-family: Calibri; font-size: 14px;\">Em anexo envio nova revisão ao desenhos de Montagem &nbsp;</span>";
                corpoEmail += "<span style=\"font-family: Calibri; font-size: 14px; color: black;\"><b>" + richTextBox1.Text.Replace("/", "<br>") + "</b></span>";
                corpoEmail += "<span style=\"font-family: Calibri; font-size: 14px;\">&nbsp da obra em assunto. &nbsp;</span>";
                corpoEmail += "<font face = 'Calibri ' size = '3' > <p> Melhores Cumprimentos,</p> </font> <br>";
                corpoEmail += "<font face = 'Calibri' size = '3' > <b>" + nomeUsuario + "</b> </Font> <br>";
                corpoEmail += "<font face = 'Calibri' size = '3' > Construção Metálica | Preparador </Font> <br>";
                corpoEmail += "<font face = 'Calibri' size = '3' > T + 351 253 080 609 * </font> <br>";
                corpoEmail += "<font color='red' font face = 'Calibri ' size = '3'> ofeliz.com </font> <br>";
                corpoEmail += "<p><a href='https://www.ofeliz.com'><img src='file:///" + imagemOfelizFilePath.Replace("\\", "/") + "' width='127' height='34'></a></p>";
                corpoEmail += "<i><font color='Light grey' font face = 'Calibri ' size = '1.5'> Alvará Nº 10553 – Pub. *Chamada para a rede fixa nacional. </font> </i><br>";
                corpoEmail += "<i><font color='green' font face = 'Calibri ' size = '1.5'> Antes de imprimir este e-mail tenha em consideração o meio ambiente. </font> </i><br>";
                corpoEmail += "</body></html>";
                string richText = richTextBox1.Text;

                this.Visible = false;
                Frm_Corpo_de_Texto_Email_RevisaoPecaseConjuntos previewForm = new Frm_Corpo_de_Texto_Email_RevisaoPecaseConjuntos("Email de Revisões ( Desenhos de Montagem )", corpoEmail, SubjectEnviar, richText);
                previewForm.ShowDialog(this);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Erro ao Conectar com o Tekla , tente novamente " + ex.Message);
            }
        }

        public void EnviarEmailAprovaçãoEstudoPrevio()
        {
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
                ultimaPasta = ultimaPasta.Replace("_", "-");
                string SubjectEnviar = ultimaPasta + " -- Aprovação";
                string saudacao = GetSaudacao();

                string corpoEmail = "<html><body contenteditable=\"false\">";
                corpoEmail += "<p style=\"font-family: Calibri; font-size: 14px;\">" + saudacao + "</p>";
                corpoEmail += "<span style=\"font-family: Calibri; font-size: 14px;\">Em anexo envio o desenho do estudo &nbsp;</span>";
                corpoEmail += "<span style=\"font-family: Calibri; font-size: 14px; color: black;\"><b>" + richTextBox1.Text.Replace("/", "<br>") + "</b></span>";
                corpoEmail += "<span style=\"font-family: Calibri; font-size: 14px;\">&nbsp para aprovação da obra em assunto. &nbsp;</span>";
                corpoEmail += "<font face = 'Calibri ' size = '3' > <p> Melhores Cumprimentos,</p> </font> <br>";
                corpoEmail += "<font face = 'Calibri' size = '3' > <b>" + nomeUsuario + "</b> </Font> <br>";
                corpoEmail += "<font face = 'Calibri' size = '3' > Construção Metálica | Preparador </Font> <br>";
                corpoEmail += "<font face = 'Calibri' size = '3' > T + 351 253 080 609 * </font> <br>";
                corpoEmail += "<font color='red' font face = 'Calibri ' size = '3'> ofeliz.com </font> <br>";
                corpoEmail += "<p><a href='https://www.ofeliz.com'><img src='file:///" + imagemOfelizFilePath.Replace("\\", "/") + "' width='127' height='34'></a></p>";
                corpoEmail += "<i><font color='Light grey' font face = 'Calibri ' size = '1.5'> Alvará Nº 10553 – Pub. *Chamada para a rede fixa nacional. </font> </i><br>";
                corpoEmail += "<i><font color='green' font face = 'Calibri ' size = '1.5'> Antes de imprimir este e-mail tenha em consideração o meio ambiente. </font> </i><br>";
                corpoEmail += "</body></html>";
                string richText = richTextBox1.Text;

                this.Visible = false;
                Frm_Corpo_de_Texto_Email_RevisaoPecaseConjuntos previewForm = new Frm_Corpo_de_Texto_Email_RevisaoPecaseConjuntos("Email de Aprovação ( Estudo prévio )", corpoEmail, SubjectEnviar, richText);
                previewForm.ShowDialog(this);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Erro ao Conectar com o Tekla , tente novamente " + ex.Message);
            }
        }

        public void EnviarEmailAprovaçãoProjetodeExecução()
        {
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
                ultimaPasta = ultimaPasta.Replace("_", "-");
                string SubjectEnviar = ultimaPasta + " -- Aprovação";
                string saudacao = GetSaudacao();

                string corpoEmail = "<html><body contenteditable=\"false\">";
                corpoEmail += "<p style=\"font-family: Calibri; font-size: 14px;\">" + saudacao + "</p>";
                corpoEmail += "<span style=\"font-family: Calibri; font-size: 14px;\">Em anexo envio o desenho &nbsp;</span>";
                corpoEmail += "<span style=\"font-family: Calibri; font-size: 14px; color: black;\"><b>" + richTextBox1.Text.Replace("/", "<br>") + "</b></span>";
                corpoEmail += "<span style=\"font-family: Calibri; font-size: 14px;\">&nbsp para aprovação da obra em assunto.  &nbsp;</span>";
                corpoEmail += "<font face = 'Calibri ' size = '3' > <p> Melhores Cumprimentos,</p> </font> <br>";
                corpoEmail += "<font face = 'Calibri' size = '3' > <b>" + nomeUsuario + "</b> </Font> <br>";
                corpoEmail += "<font face = 'Calibri' size = '3' > Construção Metálica | Preparador </Font> <br>";
                corpoEmail += "<font face = 'Calibri' size = '3' > T + 351 253 080 609 * </font> <br>";
                corpoEmail += "<font color='red' font face = 'Calibri ' size = '3'> ofeliz.com </font> <br>";
                corpoEmail += "<p><a href='https://www.ofeliz.com'><img src='file:///" + imagemOfelizFilePath.Replace("\\", "/") + "' width='127' height='34'></a></p>";
                corpoEmail += "<i><font color='Light grey' font face = 'Calibri ' size = '1.5'> Alvará Nº 10553 – Pub. *Chamada para a rede fixa nacional. </font> </i><br>";
                corpoEmail += "<i><font color='green' font face = 'Calibri ' size = '1.5'> Antes de imprimir este e-mail tenha em consideração o meio ambiente. </font> </i><br>";
                corpoEmail += "</body></html>";
                string richText = richTextBox1.Text;

                this.Visible = false;
                Frm_Corpo_de_Texto_Email_RevisaoPecaseConjuntos previewForm = new Frm_Corpo_de_Texto_Email_RevisaoPecaseConjuntos("Email de Aprovação ( Proj Execução )", corpoEmail, SubjectEnviar, richText);
                previewForm.ShowDialog(this);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Erro ao Conectar com o Tekla , tente novamente " + ex.Message);
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

        private void Frm_EnviarEmailRevisaoPeçaseConjuntos_Load(object sender, EventArgs e)
        {
           
        }

        private void richTextBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (this.Text.Contains("Email de Revisões ( Peças / Conjuntos )"))
            {
                if (e.KeyCode == Keys.Enter)
                {
                    e.SuppressKeyPress = true;
                    richTextBox1.AppendText("/");
                    richTextBox1.AppendText("-");
                    richTextBox1.AppendText("*");
                    richTextBox1.AppendText("+");
                    richTextBox1.AppendText(Environment.NewLine);
                }
            }
        }        

        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == 0)
            {
                label1.Text = "Declare o Nome das Peças ou Conjuntos a baixo";
            }
            else if (comboBox1.SelectedIndex == 1)
            {
                label1.Text = "Declare o nome do desenho do projeto a enviar";
            }
            else if (comboBox1.SelectedIndex == 2)
            {
                label1.Text = "Declare o nome do desenho do estudo a enviar ";
            }
            else if (comboBox1.SelectedIndex == 3)
            {
                label1.Text = "Declare o nome do desenho de Montangem a enviar";
            }
            else if (comboBox1.SelectedIndex == 4)
            {
                label1.Text = "Declare o nome do desenho do estudo a enviar";
            }
            else if (comboBox1.SelectedIndex == 5)
            {
                label1.Text = "Declare o nome do desenho do projeto a enviar";
            }
        }

    }
}

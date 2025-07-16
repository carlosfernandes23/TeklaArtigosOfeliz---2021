using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection.Emit;
using System.Windows.Forms;
using Tekla.Structures.Model;
using Image = System.Drawing.Image;
using Outlook = Microsoft.Office.Interop.Outlook;
using Point = Tekla.Structures.Geometry3d.Point;
using TSM = Tekla.Structures.Model;
using Tekla.Structures.Filtering;
using System.Linq;
using System.Threading;


namespace TeklaArtigosOfeliz
{
    public partial class Frm_EnviarParaLotearEmail : Form
    {
        public Frm_EnviarParaLotearEmail()
        {
            InitializeComponent();
            Chb_alw_top.CheckedChanged += Chb_alw_top_CheckedChanged;
            TopMost = Chb_alw_top.Checked;
        }

        public string TextBox1Value
        {
            get { return textBox1.Text; }
            set { textBox1.Text = value; }
        }

        public string TextBox2Value
        {
            get { return textBox2.Text; }
            set { textBox2.Text = value; }
        }

        private void EnviarParaFabricoEmail_Load(object sender, EventArgs e)
        {

        }


        public void ClickButton9()
        {
            button9.PerformClick();
        }
              
        public void EnviarFabricoOpenEmailPreviewAndCreateEmail()
        {            
            string SubjectEnviarfabrico = string.Empty;
                       
            try
            {
                System.Threading.Tasks.Task.Delay(1500).Wait();
                Process msscreenclip = Process.Start("ms-screenclip:");

                if (msscreenclip != null)
                {

                    Thread.Sleep(30000);

                }
                System.Threading.Tasks.Task.Delay(2500).Wait();
                //MessageBox.Show(this, "Email Pronto", "Aguardando Captura", MessageBoxButtons.OK, MessageBoxIcon.Information);

                if (Clipboard.ContainsImage())
                {
                    Image screenshot = Clipboard.GetImage();

                    screenshot = ResizeImageToWidthFabrico(screenshot, 20);

                    string serverPath = @"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\4 Produção\Desenvolvimentos\Ficheiros Temp tekla artigos (Nao Apagar)\Fabric";

                    string baseFileName = "screenshot_" + DateTime.Now.ToString("yyyyMMdd_HH");

                    string fileName = baseFileName;
                    int counter = 1;

                    while (File.Exists(Path.Combine(serverPath, fileName + ".png")))
                    {
                        fileName = baseFileName + "_" + counter;
                        counter++;
                    }

                    string filePath = Path.Combine(serverPath, fileName + ".png");
                    screenshot.Save(filePath, System.Drawing.Imaging.ImageFormat.Png);

                    string tempImagePath = Path.Combine(Path.GetTempPath(), filePath);
                    screenshot.Save(tempImagePath, System.Drawing.Imaging.ImageFormat.Png);

                    string imagemOfelizFilePath = @"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\4 Produção\Desenvolvimentos\Ficheiros Temp tekla artigos (Nao Apagar)\ofeliz_logo.png";

                    Model modelo = new Model();
                    string nomeProjeto = modelo.GetProjectInfo().Name;
                    string PastaModelo = modelo.GetInfo().ModelPath;
                    DirectoryInfo up = new DirectoryInfo(PastaModelo);
                    string ultimaPasta = up.Name;
                    string nomeDaObra = string.Empty;

                    string nomeUsuario = Environment.UserName;

                    nomeUsuario = nomeUsuario.Replace('.', ' ');
                    nomeUsuario = string.Join(" ", nomeUsuario.Split(' ').Select(p => char.ToUpper(p[0]) + p.Substring(1).ToLower()));

                    ultimaPasta = ultimaPasta.Replace("_", "-");
                    SubjectEnviarfabrico = ultimaPasta + " -- APROVAÇÃO";

                    string saudacao = GetSaudacao();

                    string corpoEmail = "<html><body contenteditable=\"false\">";
                    corpoEmail += "<font face = 'Calibri ' size = '3' > <p>" + saudacao + "</font></p>";

                    corpoEmail += "<font face='Calibri ' size='3'><p>"
                             + "Venho por este meio informar que a modelação&nbsp "
                             + "<span style='color:red;'>" + textBox1.Text + ",&nbsp</span> "
                             + "está terminada, "
                             + "agradeço que analise e proceda com envio para fabrico da obra em assunto. </font></p>";

                    corpoEmail += "<font face = 'Calibri ' size = '3' ><b><u> NOTA IMPORTANTE: </u></b>";

                    corpoEmail += "<br><font face='Calibri ' size='3'>- O material a fabricar encontra-se modelado como&nbsp"
                            + "<span style='color:red;'><u>" + "Fase " + textBox2.Text + "</u></span>"
                            + "<span style='color:red;'></span> do gestor de fases. </font><br>";

                    corpoEmail += "<font face = 'Calibri ' size = '3' ><p><b><u> Modelo: </u></b>"
                           + "<font face = 'Calibri ' size = '3' style='color:#5B9BD5;'>"
                           + "<u><a href='file:///" + PastaModelo + "' style='color:#5B9BD5;'>" + PastaModelo + "</a></u>" + "</font></p>";

                    corpoEmail += "<font face = 'Calibri ' size = '3' ><p><b><u> PERSPETIVA DO MATERIAL A FABRICAR: </u></b></p></font><br>";

                    corpoEmail += "<img src='file:///" + tempImagePath.Replace("\\", "/") + "' width='755' />";

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
                    Frm_Corpo_de_Texto_Email_Enviar_Para_Lotear previewForm = new Frm_Corpo_de_Texto_Email_Enviar_Para_Lotear("Email para Enviar para Lotear", corpoEmail, SubjectEnviarfabrico, tempImagePath , PastaModelo, textbox1 , textbox2);
                    previewForm.ShowDialog(this);
                }
                else
                {
                    MessageBox.Show(this, "Não foi possível capturar a imagem. Certifique-se de que a captura foi feita corretamente.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Erro ao Conectar com o Tekla , tente novamente " + ex.Message);
            }
        }

        public static Image ResizeImageToWidthFabrico(Image image, double widthInCm)
        {
            double widthInPixels = widthInCm * 37.795;

            double aspectRatio = (double)image.Height / image.Width;
            int newHeight = (int)(widthInPixels * aspectRatio);

            Bitmap resizedImage = new Bitmap(image, new System.Drawing.Size((int)widthInPixels, newHeight));
            return resizedImage;
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
                EnviarFabricoOpenEmailPreviewAndCreateEmail();
            }
        }
    }
}
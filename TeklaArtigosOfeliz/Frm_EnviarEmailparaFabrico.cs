using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using static Tekla.Structures.Filtering.Categories.PartFilterExpressions;
using Image = System.Drawing.Image;
using TSM = Tekla.Structures.Model;
using System.Threading.Tasks;
using CefSharp.DevTools.Page;
using System.Drawing.Imaging; 

namespace TeklaArtigosOfeliz
{
    public partial class Frm_EnviarEmailparaFabrico : Form
    {
        public Frm_EnviarEmailparaFabrico()
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

        private void FabricoEmail_Load(object sender, EventArgs e)
        {

        }

        public void ClickButton9()
        {
            button9.PerformClick();
        }
        public void FabricoOpenEmailPreviewAndCreateEmail()
        {
            string dataObra = string.Empty;
            string lote = string.Empty;
            string Subjectfabrico = string.Empty;

            try
            {
                ArrayList peças = new ArrayList(ComunicaTekla.ListadePecasdoConjSelec());
                ArrayList conjuntos = new ArrayList(ComunicaTekla.ListadeConjuntosSelec());
                ArrayList objectos = new ArrayList(peças);

                bool encontrouPrimeiraPeca = false;
                int index = 0;

                List<string> tudo = new List<string>();

                IEnumerable dis = tudo.Distinct();

                while (index < peças.Count && !encontrouPrimeiraPeca)
                {
                    TSM.Part peca = (TSM.Part)peças[index];

                    if (peca != null)
                    {
                        bool loteSuccess = peca.GetReportProperty("USERDEFINED.lote_number", ref lote);
                        bool dataSuccess = peca.GetReportProperty("USERDEFINED.lote_data", ref dataObra);

                        if (loteSuccess && dataSuccess)
                        {
                            encontrouPrimeiraPeca = true;

                            MessageBox.Show(this, $"Peça com Lote: {lote}, \nData da Obra: {dataObra}");
                        }

                    }

                    index++;
                }

                if (!encontrouPrimeiraPeca)
                {
                    MessageBox.Show(this, "Selecione em Modo Conjunto no Tekla as peças");
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Erro: " + ex.Message);
            }
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

                    TSM.Model modelo = new TSM.Model();
                    string nomeProjeto = modelo.GetProjectInfo().Name;
                    string PastaModelo = modelo.GetInfo().ModelPath;
                    DirectoryInfo up = new DirectoryInfo(PastaModelo);
                    string ultimaPasta = up.Name;
                    string nomeDaObra = string.Empty;

                    string nomeUsuario = Environment.UserName;

                    nomeUsuario = nomeUsuario.Replace('.', ' ');
                    nomeUsuario = string.Join(" ", nomeUsuario.Split(' ').Select(p => char.ToUpper(p[0]) + p.Substring(1).ToLower()));

                    ultimaPasta = ultimaPasta.Replace("_", "-");
                    Subjectfabrico = ultimaPasta + " -- FABRICO";

                    string saudacao = GetSaudacao();

                    string corpoEmail = "<html><body contenteditable=\"false\">";
                    corpoEmail += "<font face = 'Calibri ' size = '3' > <p>" + saudacao + "</font></p>";

                    corpoEmail += "<font face='Calibri ' size='3'><p>Material pronto para fabrico (" + "<span style='color:red;'><u></span></u>"
                                 + "<span style='color:red;'><u>" + textBox1.Text + "</u></span>"
                                 + "<span style='color:red;'><u></span></u>).</font></p>";

                    corpoEmail += "<font face = 'Calibri' size = '3' ><span style='color:red;'><u><p> Lote " + lote + "&nbsp: " + dataObra + "</u></span></p> </font>";

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

                    this.Visible = false;
                    Frm_Corpo_de_Texto_Email_Fabrico previewForm = new Frm_Corpo_de_Texto_Email_Fabrico("Enviar Email para Fabrico", corpoEmail, Subjectfabrico, tempImagePath, lote, dataObra, textbox1);
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
            if (string.IsNullOrEmpty(textBox1.Text))
            {
                MessageBox.Show(this, "Por favor, preencha o campo de texto antes de enviar o e-mail.", "Alerta", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                FabricoOpenEmailPreviewAndCreateEmail();
            }
        }
    }
}

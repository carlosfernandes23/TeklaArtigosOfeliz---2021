using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Tekla.Structures.Model;
using Image = System.Drawing.Image;
using Outlook = Microsoft.Office.Interop.Outlook;
using Point = Tekla.Structures.Geometry3d.Point;
using TSM = Tekla.Structures.Model;
using static SautinSoft.HtmlToRtf;
using SautinSoft;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;



namespace TeklaArtigosOfeliz
{
    public partial class Frm_Quantificacao : Form
    {
        //private string filePath = @"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\4 Produção\Desenvolvimentos\Ficheiros Temp tekla artigos (Nao Apagar)\Diretor de Obra Base de dados\Nao APAGAR\Textbox1.txt";
       
        public Frm_Inico _formpai;

        public string TextBox1Valuequantificacao
        {
            get { return textBox1.Text; }
            set { textBox1.Text = value; }
        }
        
        public int? NumeroQuantificacaoManual { get; set; }
        public bool ReduzirNumeroQuantificacao { get; set; } = false;


        public Frm_Quantificacao()
        {
            InitializeComponent();
            button3 = this.button3;
            VerificarEExecutar();

            //if (File.Exists(filePath))  
            //{
            //    string savedText = File.ReadAllText(filePath);  
            //    textBox1.Text = savedText;  
            //}

            Chb_alw_top.CheckedChanged += Chb_alw_top_CheckedChanged;
            TopMost = Chb_alw_top.Checked;
            Model modelo = new Model();
            label11.Text = modelo.GetProjectInfo().ProjectNumber;
        }

        private void FrmQuantificacao_Load(object sender, EventArgs e)
        {
            DateTime dataFutura = DateTime.Now.AddDays(7);
            dateTimePicker1.Value = dataFutura;

            string obra = label11.Text.Trim();
            string ano10 = null;
            int i = 0;

            if (NumeroQuantificacaoManual.HasValue)
            {
                int valor = NumeroQuantificacaoManual.Value;
                numericUpDown1.Value = valor;
            }
            else
            {
                if (obra.Contains("PT"))
                {
                    ano10 = "20" + obra.Substring(2, 2);
                }
                else
                {
                    ano10 = "20" + obra.Substring(0, 2);
                }

                if (System.IO.Directory.Exists(@"\\Marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\1 Obras\" + ano10 + @"\" + obra + @"\1.8 Projeto\1.8.3 Quantificação de material"))
                {
                    do
                    {
                        i++;
                    } while (System.IO.Directory.Exists(@"\\Marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\1 Obras\" + ano10 + @"\" + obra + @"\1.8 Projeto\1.8.3 Quantificação de material\" + i.ToString("000")));
                }

                numericUpDown1.Value = i;
            }

            List<string> subempreiteiro = new List<string>();
            ComunicaBDprimavera a = new ComunicaBDprimavera();
            a.ConectarBD();
            subempreiteiro = a.Procurarbd("SELECT [CDU_MTSubempreiteiro] FROM [PRIOFELIZ].[dbo].[MT_View_MTLocalDescarga]");
            a.DesonectarBD();

            foreach (var item in subempreiteiro)
            {
                localdedescagacb.Items.Add(item);
            }

            localdedescagacb.Text = "Uni. Negócios CM";
        }


        public void ClickButton9()
        {
            button9.PerformClick();
        }

        // CODIGO ANTES DE REMOVER O GLASS E O C25/30

        //private void button53_Click(object sender, EventArgs e)
        //{
        //    if (string.IsNullOrEmpty(textBox1.Text.Trim()))
        //    {
        //        MessageBox.Show("Por favor, preencha o campo de texto antes de enviar o e-mail.", "Alerta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //    }
        //    else
        //    {
        //        string subempreiteiro = null;
        //        ComunicaBDprimavera a = new ComunicaBDprimavera();
        //        a.ConectarBD();
        //        subempreiteiro = a.Procurarbd("SELECT [CDU_MTLocal] FROM [PRIOFELIZ].[dbo].[MT_View_MTLocalDescarga] where [CDU_MTSubempreiteiro]='" + localdedescagacb.Text + "'")[0];
        //        a.DesonectarBD();


        //        Model model = new Model();
        //        string obra = label11.Text.Trim();
        //        string ano10 = null;
        //        if (model.GetProjectInfo().ProjectNumber == obra)
        //        {
        //            if (obra.Contains("PT"))
        //            {
        //                ano10 = "20" + obra.Substring(2, 2);
        //            }
        //            else
        //            {
        //                ano10 = "20" + obra.Substring(0, 2);
        //            }
        //            string CaminhoQuantifica = @"\\Marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\1 Obras\" + ano10 + @"\" + obra + @"\1.8 Projeto\1.8.3 Quantificação de material\" + Convert.ToInt32(numericUpDown1.Text).ToString("000");
        //            if (!System.IO.Directory.Exists(CaminhoQuantifica))
        //            {
        //                ArrayList lista = new ArrayList(ComunicaTekla.ListadePecasSelec());
        //                if (lista.Count != 0)
        //                {
        //                    string teste = CaminhoQuantifica + "\\" + obra + "Q" + int.Parse(numericUpDown1.Text).ToString("000") + "SG.86.xls";
        //                    ComunicaTekla A = new ComunicaTekla();

        //                    ArrayList PECASAQUANTIFICAR = new ArrayList();
        //                    foreach (TSM.Part part in lista)
        //                    {
        //                        string quant = null;
        //                        part.GetUserProperty("QUANTIFICACAO", ref quant);

        //                        if (string.IsNullOrEmpty(quant))
        //                        {
        //                            if (part.Name.Contains("BR") || part.Name.Contains("BQ") || part.Profile.ProfileString.Contains("CF") || part.Profile.ProfileString.Contains("HF") || part.Profile.ProfileString.Contains("TGCHS"))
        //                            {
        //                                ComunicaTekla.EnviaproPriedadePeca(part, "QUANTIFICACAO", int.Parse(numericUpDown1.Text).ToString("000") + "-123" + dateTimePicker1.Value.ToShortDateString() + "-" + subempreiteiro);
        //                                PECASAQUANTIFICAR.Add(part);
        //                                MessageBox.Show("Foram identificados, nesta quantificação, perfis CFHS/HFCHS. Caso a finalidade destes perfis seja para guardas, por favor, não se esqueça de aprovisionar as curvas necessárias.");
        //                            }
        //                            else if (part.Profile.ProfileString.Contains("CH") || part.Profile.ProfileString.Contains("CG") || part.Profile.ProfileString.Contains("VRSM") || part.Profile.ProfileString.Contains("NM") || part.Profile.ProfileString.Contains("WM") || part.Profile.ProfileString.Contains("C1") || part.Profile.ProfileString.Contains("C2") || part.Profile.ProfileString.Contains("C3") || part.Profile.ProfileString.Contains("Z") || part.Profile.ProfileString.Contains("PERNO") || part.Profile.ProfileString.Contains("H60") || part.Profile.ProfileString.Contains("PL") || part.Profile.ProfileString.Contains("MAX") || part.Profile.ProfileString.Contains("SUPEROMEGA") || part.Profile.ProfileString.Contains("CONE"))
        //                            {
        //                            }
        //                            else if (part.Material.MaterialString.Contains("C45E"))
        //                            {
        //                                ComunicaTekla.EnviaproPriedadePeca(part, "QUANTIFICACAO", int.Parse(numericUpDown1.Text).ToString("000") + "-" + dateTimePicker1.Value.ToShortDateString() + "-" + subempreiteiro);
        //                                PECASAQUANTIFICAR.Add(part);
        //                            }
        //                            else if (part.Material.MaterialString.Contains("C") || part.Material.MaterialString.Contains("NEO") || part.Material.MaterialString.Contains("TEF"))
        //                            {

        //                            }
        //                            else
        //                            {
        //                                ComunicaTekla.EnviaproPriedadePeca(part, "QUANTIFICACAO", int.Parse(numericUpDown1.Text).ToString("000") + "-" + dateTimePicker1.Value.ToShortDateString() + "-" + subempreiteiro);
        //                                PECASAQUANTIFICAR.Add(part);
        //                            }
        //                        }
        //                    }
        //                    ComunicaTekla.selectinmodel(PECASAQUANTIFICAR);
        //                    if (PECASAQUANTIFICAR.Count != 0)
        //                    {
        //                        Directory.CreateDirectory(CaminhoQuantifica);
        //                        Directory.CreateDirectory(CaminhoQuantifica + @"\Q");
        //                        Directory.CreateDirectory(CaminhoQuantifica + @"\N");

        //                        new TeklaMacroBuilder.MacroBuilder()
        //                        .Callback("acmd_display_report_dialog", "", "main_frame")
        //                        .ValueChange("xs_report_dialog", "report_display_type", "1")
        //                        .ValueChange("xs_report_dialog", "display_created_report", "0")
        //                        .ListSelect("xs_report_dialog", "xs_report_list", "0-6SG.086.01.xls")
        //                        .ValueChange("xs_report_dialog", "user_title1", dateTimePicker1.Value.ToShortDateString())
        //                        .ValueChange("xs_report_dialog", "user_title2", int.Parse(numericUpDown1.Text).ToString("000"))
        //                        .ValueChange("xs_report_dialog", "user_title3", subempreiteiro)
        //                        .ValueChange("xs_report_dialog", "xs_report_file", teste.Replace("\\", "\\\\"))
        //                        .PushButton("xs_report_selected", "xs_report_dialog")
        //                        .PushButton("xs_report_cancel", "xs_report_dialog").CommandEnd().Run();
        //                        new TeklaMacroBuilder.MacroBuilder()
        //                        .Callback("acmd_display_report_dialog", "", "main_frame")
        //                        .ValueChange("xs_report_dialog", "report_display_type", "1")
        //                        .ValueChange("xs_report_dialog", "display_created_report", "1")
        //                        .PushButton("xs_report_selected", "xs_report_dialog")
        //                        .ValueChange("xs_report_dialog", "user_title1", "Data")
        //                        .ValueChange("xs_report_dialog", "user_title2", "001")
        //                        .ValueChange("xs_report_dialog", "user_title3", "")
        //                        .ValueChange("xs_report_dialog", "xs_report_file", "")
        //                        .PushButton("xs_report_selected", "xs_report_dialog")
        //                        .PushButton("xs_report_cancel", "xs_report_dialog").CommandEnd().Run();

        //                        MessageBox.Show("Quantificação de material criada com sucesso da " + obra + " da fase " + int.Parse(numericUpDown1.Text).ToString("000"), "Quantificação de material", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //                    }
        //                    else
        //                    {
        //                        MessageBox.Show("Não existem peças a quantificar", "Quantificação de material", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //                    }

        //                }
        //            }
        //            else
        //            {
        //                DialogResult A = MessageBox.Show("Atenção que o número da quantificação já existe deseja apagar a pasta?" + Environment.NewLine + "Se responder  “Sim” irá apagar a pasta e poderá retirar depois novamente a quantificação." + Environment.NewLine + "Se responder “Não” não vai fazer nada." + Environment.NewLine + "Se responder Cancelar vai abrir a pasta no explorador do Windows.", "EXPORTAÇÃO", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation);

        //                if (A == DialogResult.Yes)
        //                {
        //                    if (System.IO.Directory.Exists(CaminhoQuantifica))
        //                    {
        //                        Directory.Delete(CaminhoQuantifica, true);
        //                        MessageBox.Show("Pasta Apagada", "EXPORTAÇÃO", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //                    }
        //                    else
        //                    {
        //                        MessageBox.Show("Pasta ja não existia", "EXPORTAÇÃO", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //                    }
        //                }
        //                else if (A == DialogResult.No)
        //                {

        //                }
        //                else if (A == DialogResult.Cancel)
        //                {
        //                    Process.Start(CaminhoQuantifica);
        //                }

        //            }
        //        }
        //        else
        //        {
        //            MessageBox.Show("O projeto atual deste programa não é o projeto atual do Tekla.", "ERRO", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //        }

        //        DialogResult resultado2 = MessageBox.Show("Quer importar no Primavera?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

        //        if (resultado2 == DialogResult.Yes)
        //        {
        //            Guardartextbox1();

        //            MessageBox.Show("Não esquecer de importar no Primavera!", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //            AppAbrirPrimavera primaveraHandler = new AppAbrirPrimavera();
        //            primaveraHandler.AbrirPrimaveira();

        //            int quantificacaoMaterial = int.Parse(numericUpDown1.Text);
        //            ExportIFCPlugin exportPlugin = new ExportIFCPlugin();
        //            exportPlugin.Run(obra, quantificacaoMaterial);

        //            OpenEmailPreviewAndCreateEmail();
        //        }
        //        else
        //        { }

        //    }
        //}

        private void label22_Click(object sender, EventArgs e)
        {
            string rui = null;
            ArrayList lista = ComunicaTekla.ListadePecasSelec();
            foreach (TSM.Part part in lista)
            {
                part.GetUserProperty("Requisitos", ref rui);
                if (!string.IsNullOrEmpty(rui))
                {
                    MessageBox.Show(this, rui.ToString());
                }
            }
        }            

        public void OpenEmailPreviewAndCreateEmail()
        {
            Model model = new Model();
            string obra = label11.Text.Trim();
            string nomeDaObra = string.Empty;
            int i = 0;
            int quantificacaoMaterial = int.Parse(numericUpDown1.Text);
            string Subject = string.Empty;

            string ano10 = null;
            if (model.GetProjectInfo().ProjectNumber == obra)
            {
                if (obra.Contains("PT"))
                {
                    ano10 = "20" + obra.Substring(2, 2);
                }
                else
                {
                    ano10 = "20" + obra.Substring(0, 2);
                }

                Chb_alw_top.Checked = false;

                try
                {
                    System.Threading.Tasks.Task.Delay(1500).Wait();
                    Process msscreenclip = Process.Start("ms-screenclip:");

                    if (msscreenclip != null)
                    {

                        Thread.Sleep(30000);

                    }

                    System.Threading.Tasks.Task.Delay(2000).Wait();
                    Chb_alw_top.Checked = true;
                    //MessageBox.Show(this, "Email Pronto", "Aguardando Captura", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    if (Clipboard.ContainsImage())
                    {
                        Image screenshot = Clipboard.GetImage();

                        screenshot = ResizeImageToWidth(screenshot, 20);

                        string serverPath = @"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\4 Produção\Desenvolvimentos\Ficheiros Temp tekla artigos (Nao Apagar)\Quantif";

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

                        string nomeUsuario = Environment.UserName;

                        nomeUsuario = nomeUsuario.Replace('.', ' ');
                        nomeUsuario = string.Join(" ", nomeUsuario.Split(' ').Select(p => char.ToUpper(p[0]) + p.Substring(1).ToLower()));

                        ultimaPasta = ultimaPasta.Replace("_", "-");
                        string numericupdown = int.Parse(numericUpDown1.Text).ToString("000");
                        Subject = ultimaPasta + " -- QUANTIFICAÇÃO DE MATERIAL";

                        string linkTexto = @"Y:\" + ano10 + @"\" + obra + @"\1.8 Projeto\1.8.3 Quantificação de material\" + quantificacaoMaterial.ToString("000");
                        string caminho = @"\\Marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\1 Obras\" + ano10 + "\\" + obra + "\\1.8 Projeto\\1.8.3 Quantificação de material\\" + quantificacaoMaterial.ToString("000");

                        string saudacao = GetSaudacao();

                        string corpoEmail = "<html><body contenteditable=\"false\">";
                        corpoEmail += "<font face = 'Calibri ' size = '3' > <p>" + saudacao + "</font></p>";
                        corpoEmail += "<font face='Calibri ' size='3'><p>Conforme solicitado, informo que já foi emitida no Primavera a quantificação de material&nbsp;"
                                    + "<span style='color:#00B0F0; display:inline-block; margin-right:10px;'><u>"
                                    + numericupdown
                                    + "</u></span>"
                                    + "<span style='color:red;'>\" " + textBox1.Text + " </span>"
                                    + "<span style='color:red;'>\"</span> da obra em assunto. </font></p>";

                        corpoEmail += "<font face = 'Calibri ' size = '3' ><p><b><u> QUANTIFICAÇÃO DE MATERIAL: </u></b>" + "<b><u>" + "</u></b> </font>" + " <font face = 'Calibri ' size = '3' style='color:Blue ><b><u> ";

                        corpoEmail += "<font face = 'Calibri ' size = '3' style='color:#5B9BD5;'>"
                                    + "<u><a href='file:///" + caminho + "' style='color:#5B9BD5;'> " + linkTexto + "</a></u>" + "</font></p>";

                        corpoEmail += "<font face = 'Calibri ' size = '3' ><p><b><u> PERSPETIVA DO MATERIAL A ENCOMENDAR: </u></b></p> </font>";

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
                        Frm_Corpo_de_Texto_Email_Quantificacao previewForm = new Frm_Corpo_de_Texto_Email_Quantificacao("Enviar Email da quantificação de Material", corpoEmail, Subject, tempImagePath, caminho, linkTexto, numericupdown, textbox1);
                        previewForm.ShowDialog(this);

                    }
                    else
                    {
                        MessageBox.Show(this, "Não foi possível capturar a imagem. Certifique-se de que a captura foi feita corretamente.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "Erro ao abrir a Ferramenta de Recorte ou enviar o e-mail: " + ex.Message);
                }
            }
        }


        public static Image ResizeImageToWidth(Image image, double widthInCm)
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

        private string UploadImageToServer(Image image)
        {
            try
            {
                string serverPath = @"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\4 Produção\Desenvolvimentos\Ficheiros Temp tekla artigos (Nao Apagar)\Quantif";

                string baseFileName = "screenshot_" + DateTime.Now.ToString("yyyyMMdd_HH");

                string fileName = baseFileName;
                int counter = 1;

                while (File.Exists(Path.Combine(serverPath, fileName + ".png")))
                {
                    fileName = baseFileName + "_" + counter;
                    counter++;
                }

                string filePath = Path.Combine(serverPath, fileName + ".png");

                image.Save(filePath, System.Drawing.Imaging.ImageFormat.Png);

                return filePath;
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Erro ao enviar imagem para o servidor: " + ex.Message);
                return null;
            }
        }

        public class ExportIFCPlugin
        {
            public void Run(string obra, int quantificacaoMaterial)
            {
                try
                {
                    string ano10 = string.Empty;

                    if (obra.Contains("PT"))
                    {
                        ano10 = "20" + obra.Substring(2, 2);
                    }
                    else
                    {
                        ano10 = "20" + obra.Substring(0, 2);
                    }

                    if (string.IsNullOrEmpty(ano10) || string.IsNullOrEmpty(obra))
                    {
                        MessageBox.Show("Ano ou Obra não definidos corretamente.");
                        return;
                    }

                    int i = quantificacaoMaterial;


                    string basePath = @"\\Marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\1 Obras\";

                    if (obra.Contains("PT"))
                    {
                        ano10 = "20" + obra.Substring(2, 2);
                    }
                    else
                    {
                        ano10 = "20" + obra.Substring(0, 2);
                    }

                    string caminhoBase = Path.Combine(basePath, ano10, obra, "1.8 Projeto", "1.8.3 Quantificação de material");

                    string pastaQuantificacao = Path.Combine(caminhoBase, i.ToString("000"));
                    if (!Directory.Exists(pastaQuantificacao))
                    {
                        Directory.CreateDirectory(pastaQuantificacao);
                    }

                    string outputFileName = Path.Combine(pastaQuantificacao, obra + "_Q" + i.ToString("000") + ".ifc");

                    var componentInput = new ComponentInput();
                    componentInput.AddOneInputPosition(new Point(0, 0, 0));

                    var comp = new TSM.Component(componentInput)
                    {
                        Name = "ExportIFC",
                        Number = BaseComponent.PLUGIN_OBJECT_NUMBER
                    };

                    comp.LoadAttributesFromFile("standard");
                    comp.SetAttribute("OutputFile", outputFileName);
                    comp.Insert();

                    // comp.SetAttribute("CreateAll", 0);  

                    FileCleaner cleaner = new FileCleaner();
                    cleaner.DeleteLogFiles(ano10, obra, i);

                    //MessageBox.Show("Caminho do arquivo de saída: " + outputFileName);
                    //MessageBox.Show("Exportação IFC concluída.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erro: " + ex.Message);
                }
            }
        }

        public class FileCleaner
        {
            public void DeleteLogFiles(string ano10, string obra, int i)
            {
                try
                {
                    string caminhoQuantifica = @"\\Marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\1 Obras\" + ano10 + @"\" + obra + @"\1.8 Projeto\1.8.3 Quantificação de material\" + i.ToString("000");

                    if (Directory.Exists(caminhoQuantifica))
                    {
                        string[] logFiles = Directory.GetFiles(caminhoQuantifica, "*.log");

                        foreach (var file in logFiles)
                        {
                            File.Delete(file);

                        }

                        //MessageBox.Show("Arquivos .log deletados com sucesso.", "Limpeza de arquivos", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("O diretório não foi encontrado.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erro ao tentar excluir arquivos: " + ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void Guardartextbox1()
        {
            //File.WriteAllText(filePath, textBox1.Text);  
        }

                  

        private System.Threading.Timer timer;
        private static DateTime dataReferencia = new DateTime(2024, 1, 1);

        private static void VerificarEExecutar()
        {
            if (VerificarIntervaloDeDoisMeses())
            {
                EliminarimagensPasta();
            }
            
        }

       public static void EliminarimagensPasta()
        {
            string directoryPath = @"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\4 Produção\Desenvolvimentos\Ficheiros Temp tekla artigos (Nao Apagar)\Quantif";

            if (Directory.Exists(directoryPath))
            {
                string[] files = Directory.GetFiles(directoryPath);

                DateTime twoYearsAgo = DateTime.Now.AddYears(-2);

                foreach (string filePath in files)
                {
                    try
                    {
                        FileInfo fileInfo = new FileInfo(filePath);

                        DateTime creationTime = fileInfo.CreationTime;

                        if (creationTime < twoYearsAgo)
                        {
                            File.Delete(filePath);
                        }
                        else
                        {
                            
                        }
                    }
                    catch (Exception ex)
                    {
                        // Caso haja algum erro (ex.: permissão de acesso), exibe a mensagem de erro
                        MessageBox.Show($"Erro ao processar o arquivo {filePath}: {ex.Message}");
                    }
                }
            }
            else
            {
                MessageBox.Show("O diretório especificado não existe.");
            }
        } 

        private static bool VerificarIntervaloDeDoisMeses()
        {
            int mesesDeDiferenca = (DateTime.Now.Year - dataReferencia.Year) * 12 + DateTime.Now.Month - dataReferencia.Month;

            return mesesDeDiferenca % 2 == 0;


        }
        public class AppAbrirPrimavera
        {
            [DllImport("user32.dll")]
            public static extern bool SetForegroundWindow(IntPtr hWnd);

            [DllImport("user32.dll")]
            public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

            [DllImport("user32.dll")]
            public static extern void SwitchToThisWindow(IntPtr hWnd, bool fAltTab);

            const int SW_RESTORE = 9;

            public void AbrirPrimaveira()
            {
                try
                {
                    string nomeProcesso = "Erp900LE";
                    string appPath = @"C:\Program Files (x86)\PRIMAVERA\SG900\Apl\Erp900LE.exe";

                    var processos = Process.GetProcessesByName(nomeProcesso);

                    if (processos.Length > 0)
                    {
                        var processoExistente = processos[0];
                        IntPtr hWnd = processoExistente.MainWindowHandle;

                        if (hWnd != IntPtr.Zero)
                        {
                            ShowWindow(hWnd, SW_RESTORE);
                            SetForegroundWindow(hWnd);
                        }
                        else
                        {
                            // Tenta usar o método alternativo
                            SwitchToThisWindow(processoExistente.Handle, true);
                        }
                    }
                    else
                    {
                        if (File.Exists(appPath))
                        {
                            var processoNovo = Process.Start(appPath);
                            processoNovo.WaitForInputIdle(); // Espera a janela

                            IntPtr hWnd = processoNovo.MainWindowHandle;

                            if (hWnd != IntPtr.Zero)
                            {
                                ShowWindow(hWnd, SW_RESTORE);
                                SetForegroundWindow(hWnd);
                            }
                            else
                            {
                                MessageBox.Show("Primavera iniciado, mas não foi possível aceder à janela principal.");
                            }
                        }
                        else
                        {
                            MessageBox.Show("O Primavera não foi encontrado no PC.");
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erro ao tentar abrir o Primavera: " + ex.Message);
                }
            }
        }


        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            
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

        private void button7_Click(object sender, EventArgs e)
        {
            Model modelo = new Model();
            label11.Text = modelo.GetProjectInfo().ProjectNumber;
            string obra = label11.Text.Trim();
            string ano10 = null;
            int i = 0;
            if (obra.Contains("PT"))
            {
                ano10 = "20" + obra.Substring(2, 2);
            }
            else
            {
                ano10 = "20" + obra.Substring(0, 2);
            }

            if (System.IO.Directory.Exists(@"\\Marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\1 Obras\" + ano10 + @"\" + obra + @"\1.8 Projeto\1.8.3 Quantificação de material"))
            {
                do
                {
                    i++;
                } while (System.IO.Directory.Exists(@"\\Marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\1 Obras\" + ano10 + @"\" + obra + @"\1.8 Projeto\1.8.3 Quantificação de material\" + i.ToString("000")));
            }

            if (System.IO.File.Exists(@"\\Marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\1 Obras\" + ano10 + @"\" + obra + @"\1.8 Projeto\1.8.3 Quantificação de material\" + (i - 1).ToString("000") + "\\" + label11.Text + "Q" + (i - 1).ToString("000") + "SG.86.xls"))
            {
                Process.Start(@"\\Marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\1 Obras\" + ano10 + @"\" + obra + @"\1.8 Projeto\1.8.3 Quantificação de material\" + (i - 1).ToString("000") + "\\" + label11.Text + "Q" + (i - 1).ToString("000") + "SG.86.xls");
            }
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            Model modelo = new Model();
            label11.Text = modelo.GetProjectInfo().ProjectNumber;
            string obra = label11.Text.Trim();
            string ano10 = null;
            int i = 0;
            if (obra.Contains("PT"))
            {
                ano10 = "20" + obra.Substring(2, 2);
            }
            else
            {
                ano10 = "20" + obra.Substring(0, 2);
            }

            if (System.IO.Directory.Exists(@"\\Marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\1 Obras\" + ano10 + @"\" + obra + @"\1.8 Projeto\1.8.3 Quantificação de material"))
            {
                do
                {
                    i++;
                } while (System.IO.Directory.Exists(@"\\Marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\1 Obras\" + ano10 + @"\" + obra + @"\1.8 Projeto\1.8.3 Quantificação de material\" + i.ToString("000")));
            }

            if (System.IO.Directory.Exists(@"\\Marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\1 Obras\" + ano10 + @"\" + obra + @"\1.8 Projeto\1.8.3 Quantificação de material\" + (i - 1).ToString("000")))
            {
                Process.Start(@"\\Marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\1 Obras\" + ano10 + @"\" + obra + @"\1.8 Projeto\1.8.3 Quantificação de material\" + (i - 1).ToString("000"));
            }
        }

        private void guna2Button2_Click(object sender, EventArgs e)
        {
            Guardartextbox1();
            Chb_alw_top.Checked = false;
            OpenEmailPreviewAndCreateEmail();

        }

        private void guna2Button3_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox1.Text.Trim()) || textBox1.Text.Trim().Equals("Estrutura", StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show(this, "Por favor, preencha o campo da Descrição do Material antes de enviar o e-mail.", "Alerta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                string subempreiteiro = null;
                ComunicaBDprimavera a = new ComunicaBDprimavera();
                a.ConectarBD();
                subempreiteiro = a.Procurarbd("SELECT [CDU_MTLocal] FROM [PRIOFELIZ].[dbo].[MT_View_MTLocalDescarga] where [CDU_MTSubempreiteiro]='" + localdedescagacb.Text + "'")[0];
                a.DesonectarBD();

                Model model = new Model();
                string obra = label11.Text.Trim();
                string ano10 = null;
                if (model.GetProjectInfo().ProjectNumber == obra)
                {
                    if (obra.Contains("PT"))
                    {
                        ano10 = "20" + obra.Substring(2, 2);
                    }
                    else
                    {
                        ano10 = "20" + obra.Substring(0, 2);
                    }
                    string CaminhoQuantifica = @"\\Marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\1 Obras\" + ano10 + @"\" + obra + @"\1.8 Projeto\1.8.3 Quantificação de material\" + Convert.ToInt32(numericUpDown1.Text).ToString("000");

                    if (!System.IO.Directory.Exists(CaminhoQuantifica))
                    {
                        ArrayList lista = new ArrayList(ComunicaTekla.ListadePecasSelec());
                        if (lista.Count != 0)
                        {
                            for (int i = lista.Count - 1; i >= 0; i--)
                            {
                                TSM.Part part = (TSM.Part)lista[i];

                                if (part.Material.MaterialString.Contains("GLASS") || part.Material.MaterialString.Contains("C25/30"))
                                {
                                    lista.RemoveAt(i);
                                }
                            }

                            string teste = CaminhoQuantifica + "\\" + obra + "Q" + int.Parse(numericUpDown1.Text).ToString("000") + "SG.86.xls";
                            ComunicaTekla A = new ComunicaTekla();

                            ArrayList PECASAQUANTIFICAR = new ArrayList();
                            bool mensagemExibida = false;

                            foreach (TSM.Part part in lista)
                            {
                                string quant = null;
                                part.GetUserProperty("QUANTIFICACAO", ref quant);

                                if ((part.Profile.ProfileString.Contains("CF") || part.Profile.ProfileString.Contains("HF")) && !mensagemExibida)
                                {
                                    MessageBox.Show(this, "Foram identificados, nesta quantificação, perfis CFHS/HFCHS. Caso a finalidade destes perfis seja para guardas, por favor, não se esqueça de aprovisionar as curvas necessárias.", "ATENÇÃO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    mensagemExibida = true;
                                }

                                if (string.IsNullOrEmpty(quant))
                                {
                                    if (part.Name.Contains("BR") || part.Name.Contains("BQ") || part.Profile.ProfileString.Contains("CF") || part.Profile.ProfileString.Contains("HF") || part.Profile.ProfileString.Contains("TGCHS"))
                                    {
                                        ComunicaTekla.EnviaproPriedadePeca(part, "QUANTIFICACAO", int.Parse(numericUpDown1.Text).ToString("000") + "-123" + dateTimePicker1.Value.ToShortDateString() + "-" + subempreiteiro);
                                        PECASAQUANTIFICAR.Add(part);
                                    }
                                    else if (part.Profile.ProfileString.Contains("CH") || part.Profile.ProfileString.Contains("CG") || part.Profile.ProfileString.Contains("VRSM") || part.Profile.ProfileString.Contains("NM") || part.Profile.ProfileString.Contains("WM") || part.Profile.ProfileString.Contains("C1") || part.Profile.ProfileString.Contains("C2") || part.Profile.ProfileString.Contains("C3") || part.Profile.ProfileString.Contains("Z") || part.Profile.ProfileString.Contains("PERNO") || part.Profile.ProfileString.Contains("H60") || part.Profile.ProfileString.Contains("PL") || part.Profile.ProfileString.Contains("MAX") || part.Profile.ProfileString.Contains("SUPEROMEGA") || part.Profile.ProfileString.Contains("CONE"))
                                    {
                                    }
                                    else if (part.Material.MaterialString.Contains("C45E"))
                                    {
                                        ComunicaTekla.EnviaproPriedadePeca(part, "QUANTIFICACAO", int.Parse(numericUpDown1.Text).ToString("000") + "-" + dateTimePicker1.Value.ToShortDateString() + "-" + subempreiteiro);
                                        PECASAQUANTIFICAR.Add(part);
                                    }
                                    else if (part.Material.MaterialString.Contains("C") || part.Material.MaterialString.Contains("NEO") || part.Material.MaterialString.Contains("TEF"))
                                    {

                                    }
                                    else
                                    {
                                        ComunicaTekla.EnviaproPriedadePeca(part, "QUANTIFICACAO", int.Parse(numericUpDown1.Text).ToString("000") + "-" + dateTimePicker1.Value.ToShortDateString() + "-" + subempreiteiro);
                                        PECASAQUANTIFICAR.Add(part);
                                    }
                                }
                            }
                            ComunicaTekla.selectinmodel(PECASAQUANTIFICAR);
                            if (PECASAQUANTIFICAR.Count != 0)
                            {
                                Directory.CreateDirectory(CaminhoQuantifica);
                                Directory.CreateDirectory(CaminhoQuantifica + @"\Q");
                                Directory.CreateDirectory(CaminhoQuantifica + @"\N");

                                new TeklaMacroBuilder.MacroBuilder()
                                .Callback("acmd_display_report_dialog", "", "main_frame")
                                .ValueChange("xs_report_dialog", "report_display_type", "1")
                                .ValueChange("xs_report_dialog", "display_created_report", "0")
                                .ListSelect("xs_report_dialog", "xs_report_list", "0-6SG.086.01.xls")
                                .ValueChange("xs_report_dialog", "user_title1", dateTimePicker1.Value.ToShortDateString())
                                .ValueChange("xs_report_dialog", "user_title2", int.Parse(numericUpDown1.Text).ToString("000"))
                                .ValueChange("xs_report_dialog", "user_title3", subempreiteiro)
                                .ValueChange("xs_report_dialog", "xs_report_file", teste.Replace("\\", "\\\\"))
                                .PushButton("xs_report_selected", "xs_report_dialog")
                                .PushButton("xs_report_cancel", "xs_report_dialog").CommandEnd().Run();
                                new TeklaMacroBuilder.MacroBuilder()
                                .Callback("acmd_display_report_dialog", "", "main_frame")
                                .ValueChange("xs_report_dialog", "report_display_type", "1")
                                .ValueChange("xs_report_dialog", "display_created_report", "1")
                                .PushButton("xs_report_selected", "xs_report_dialog")
                                .ValueChange("xs_report_dialog", "user_title1", "Data")
                                .ValueChange("xs_report_dialog", "user_title2", "001")
                                .ValueChange("xs_report_dialog", "user_title3", "")
                                .ValueChange("xs_report_dialog", "xs_report_file", "")
                                .PushButton("xs_report_selected", "xs_report_dialog")
                                .PushButton("xs_report_cancel", "xs_report_dialog").CommandEnd().Run();

                                MessageBox.Show(this, "Quantificação de material criada com sucesso da " + obra + " da fase " + int.Parse(numericUpDown1.Text).ToString("000"), "Quantificação de material", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                MessageBox.Show(this, "Não existem peças a quantificar", "Quantificação de material", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }

                        }
                    }
                    else
                    {
                        DialogResult A = MessageBox.Show(this, "Atenção que o número da quantificação já existe deseja apagar a pasta?" + Environment.NewLine + "Se responder  “Sim” irá apagar a pasta e poderá retirar depois novamente a quantificação." + Environment.NewLine + "Se responder “Não” não vai fazer nada." + Environment.NewLine + "Se responder Cancelar vai abrir a pasta no explorador do Windows.", "EXPORTAÇÃO", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation);

                        if (A == DialogResult.Yes)
                        {
                            if (System.IO.Directory.Exists(CaminhoQuantifica))
                            {
                                Directory.Delete(CaminhoQuantifica, true);
                                MessageBox.Show(this, "Pasta Apagada", "EXPORTAÇÃO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                MessageBox.Show(this, "Pasta ja não existia", "EXPORTAÇÃO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                        else if (A == DialogResult.No)
                        {

                        }
                        else if (A == DialogResult.Cancel)
                        {
                            Process.Start(CaminhoQuantifica);
                        }

                    }
                }
                else
                {
                    MessageBox.Show(this, "O projeto atual deste programa não é o projeto atual do Tekla.", "ERRO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                System.Threading.Tasks.Task.Delay(2000).Wait();

                DialogResult resultadoPrimavera = MessageBox.Show(this, "Quer importar no Primavera?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (resultadoPrimavera == DialogResult.Yes)
                {
                    Guardartextbox1();
                    Chb_alw_top.Checked = false;
                    AppAbrirPrimavera primaveraHandler = new AppAbrirPrimavera();
                    primaveraHandler.AbrirPrimaveira();

                    int quantificacaoMaterial = int.Parse(numericUpDown1.Text);
                    ExportIFCPlugin exportPlugin = new ExportIFCPlugin();
                    exportPlugin.Run(obra, quantificacaoMaterial);

                    System.Threading.Tasks.Task.Delay(5000).Wait();
                    Chb_alw_top.Checked = true;
                    DialogResult resultadoEmail = MessageBox.Show(this, "Pretende Criar Email?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (resultadoEmail == DialogResult.Yes)
                    {
                        OpenEmailPreviewAndCreateEmail();

                    }
                    else
                    { }
                }
                else
                { }

               



            }
        }

        private void guna2NumericUpDown1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void guna2Button6_Click(object sender, EventArgs e)
        {
            ArrayList PECAS = new ArrayList();
            PECAS = ComunicaTekla.ListadePecasSelec();
            ComunicaTekla.EnviaproPriedadePeca(PECAS, "Requisitos", LblRequesitos.Text);

            MessageBox.Show(this, "Foram adicionados os requisitos a " + PECAS.Count + " peças." + Environment.NewLine + LblRequesitos.Text, "INFORMAÇÃO", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void guna2Button4_Click(object sender, EventArgs e)
        {
            button123456.PerformClick();
        }

        private void guna2Button5_Click(object sender, EventArgs e)
        {
            LblRequesitos.Text = "";

        }

        private void button33_Click(object sender, EventArgs e)
        {
            string req = null;

            Button b = (Button)sender;
            if (b.Name == "button123456")
            {
                req = CbRequesitos.Text;
            }

            if (!string.IsNullOrEmpty(req.Trim()))
            {
                if (LblRequesitos.Text == "")
                {
                    LblRequesitos.Text = req;
                }
                else
                {
                    LblRequesitos.Text = LblRequesitos.Text + " | " + req;
                }
                CbRequesitos.Text = "";
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            Guardartextbox1();

            OpenEmailPreviewAndCreateEmail();
        }
    }
}




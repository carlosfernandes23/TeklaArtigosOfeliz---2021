using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Collections;
using System.Globalization;
using System.Diagnostics;
using System.Printing;
using System.Text.RegularExpressions;
using TSM = Tekla.Structures.Model;
using Tekla.Structures.Model;
using TSD = Tekla.Structures.Drawing;
using Tekla.Structures.Filtering;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Filtering.Categories;
using Tekla.Structures.Drawing;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Reflection.Emit;
using Image = System.Drawing.Image;
using Outlook = Microsoft.Office.Interop.Outlook;
using Point = Tekla.Structures.Geometry3d.Point;
using Excel = Microsoft.Office.Interop.Excel;
using Tekla.Structures.ModelInternal;


namespace TeklaArtigosOfeliz
{

    public partial class Frm_Inico : Form
    {

        public static string CaminhoModelo = null;
        public string versaotekla = "2021.0";
        public static string PastaPartilhada = null;

        public static string PastaReservatorioFicheiros = null;
        public static string ano = null;
        public static string _fase = null;

        private string filePath;
        private Timer checkTimer;

        public string fase
        {
            get { return _fase; }
            set { _fase = value; }
        }

        public static string _fase500 = null;

        public string fase500
        {
            get { return _fase500; }
            set { _fase500 = value; }
        }

        public static string _fase1000 = null;

        public string fase1000
        {
            get { return _fase1000; }
            set { _fase1000 = value; }
        }

        public static List<string> str = new List<string>();

        public Frm_Inico()
        {
            InitializeComponent();
            Pdfsoldadura();
            //////////////////////////////
            try
            {
                carregafase();
                carregafase500();
                carregafase1000();

                ///////////////////////////////
                StreamReader leitor = new StreamReader("config.txt");
                PastaPartilhada = leitor.ReadLine();
                //////////////////////////////////////////
                string[] caminho = CaminhoModelo.Split('\\');
                foreach (var item in caminho)
                {
                    if (item.StartsWith("20") && item.Count(char.IsNumber) == 4)
                    {
                        ano = item;
                    }
                }
            }
            catch (Exception EX)
            {
                label3.Visible = true;
                label1.Visible = false;
                label11.Visible = false;
                //MessageBox.Show(this, EX.Message, "ERRO ABRIR APP");
                MessageBox.Show(this, $"O Tekla Structures não está aberto ou a versão instalada não é compatível.\n\nDetalhes do erro: {EX.Message}", "Erro ao abrir a aplicação", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        //void carregafase()
        //{
        //    Model modelo = new Model();
        //    CaminhoModelo = modelo.GetInfo().ModelPath;
        //    label11.Text = modelo.GetProjectInfo().ProjectNumber;
        //    string[] _caminho = CaminhoModelo.Split('\\');
        //    bool fab = true;
        //    string caminho = null;
        //    foreach (string item in _caminho)
        //    {
        //        if (item!= "1.8 Projeto")
        //        {
        //            if (fab)
        //            {
        //                caminho += item + "\\";
        //            }
        //        }else
        //        {
        //            fab = false;
        //        }
        //    }

        //    int f = 0;
        //    if (System.IO.Directory.Exists(caminho))
        //    {
        //        do
        //        {
        //            f++;
        //            fase = f.ToString("000");
        //        } while (System.IO.Directory.Exists(caminho + @"\1.9 Gestão de fabrico\" + fase));
        //        PastaReservatorioFicheiros = caminho + @"\1.9 Gestão de fabrico\";

        //    }
        //}


        //void carregafase500()
        //{
        //    Model modelo = new Model();
        //    string[] _caminho = CaminhoModelo.Split('\\');
        //    bool fab = true;
        //    string caminho = null;
        //    foreach (string item in _caminho)
        //    {
        //        if (item != "1.8 Projeto")
        //        {
        //            if (fab)
        //            {
        //                caminho += item + "\\";
        //            }
        //        }
        //        else
        //        {
        //            fab = false;
        //        }
        //    }
        //    int f = 499; // Começando em 499
        //    fase500 = null; // Inicializa a variável
        //    if (System.IO.Directory.Exists(caminho))
        //    {
        //        do
        //        {
        //            f++;
        //            fase500 = f.ToString("000");
        //        } while (System.IO.Directory.Exists(caminho + @"\1.9 Gestão de fabrico\" + fase500));

        //        // Verifique se fase500 foi definida
        //        if (fase500 == null)
        //        {
        //            MessageBox.Show(this, "Fase 500 não encontrada.");
        //}
        //    }
        //}


        //void carregafase1000()
        //{
        //    Model modelo = new Model();
        //    string[] _caminho = CaminhoModelo.Split('\\');
        //    bool fab = true;
        //    string caminho = null;
        //    foreach (string item in _caminho)
        //    {
        //        if (item != "1.8 Projeto")
        //        {
        //            if (fab)
        //            {
        //                caminho += item + "\\";
        //            }
        //        }
        //        else
        //        {
        //            fab = false;
        //        }
        //    }
        //    int f = 999;
        //    fase1000 = null;
        //    if (System.IO.Directory.Exists(caminho))
        //    {
        //        do
        //        {
        //            f++;
        //            fase1000 = f.ToString("000");
        //        } while (System.IO.Directory.Exists(caminho + @"\1.9 Gestão de fabrico\" + fase1000));

        //    }
        //}


        public void carregafase()
        {
            Model modelo = new Model();
            CaminhoModelo = modelo.GetInfo().ModelPath;
            label11.Text = modelo.GetProjectInfo().ProjectNumber;

            string[] _caminho = CaminhoModelo.Split('\\');
            bool fab = true;
            string caminho = null;

            foreach (string item in _caminho)
            {
                if (item != "1.8 Projeto")
                {
                    if (fab)
                    {
                        caminho += item + "\\";
                    }
                }
                else
                {
                    fab = false;
                }
            }

            if (string.IsNullOrEmpty(caminho) || !Directory.Exists(caminho))
            {
                return; 
            }
            string pastaGestaoFabrico = Path.Combine(caminho, "1.9 Gestão de fabrico");
            PastaReservatorioFicheiros = pastaGestaoFabrico + "\\";
            Directory.CreateDirectory(pastaGestaoFabrico);

            var todosOsNomesDePasta = Directory.GetDirectories(pastaGestaoFabrico)
                                               .Select(Path.GetFileName)
                                               .ToList();

            int proximaVaga = 0;
            string fasePotencial;

            do
            {
                proximaVaga++;
                fasePotencial = proximaVaga.ToString("000");

            } while (todosOsNomesDePasta.Any(nomePasta => nomePasta.StartsWith(fasePotencial)));


            const int LIMITE_FASE = 499;

            int numeroFinal = Math.Min(proximaVaga, LIMITE_FASE);

            fase = numeroFinal.ToString("000");

            if (fase == null)
            {
                MessageBox.Show(this, "Fase de 000 a 499 não encontrada.");
            }
        }

        public void carregafase500()
        {
            Model modelo = new Model();

            string[] _caminho = CaminhoModelo.Split('\\');
            bool fab = true;
            string caminho = null;

            foreach (string item in _caminho)
            {
                if (item != "1.8 Projeto")
                {
                    if (fab)
                    {
                        caminho += item + "\\";
                    }
                }
                else
                {
                    fab = false;
                }
            }

            if (string.IsNullOrEmpty(caminho) || !Directory.Exists(caminho))
            {
                return; 
            }

            string pastaGestaoFabrico = Path.Combine(caminho, "1.9 Gestão de fabrico");
            Directory.CreateDirectory(pastaGestaoFabrico);

            var todosOsNomesDePasta = Directory.GetDirectories(pastaGestaoFabrico)
                                              .Select(Path.GetFileName)
                                              .ToList();

            int proximaVaga = 499;
            string fasePotencial;

            do
            {
                proximaVaga++;
                fasePotencial = proximaVaga.ToString("000");

            } while (todosOsNomesDePasta.Any(nomePasta => nomePasta.StartsWith(fasePotencial)));


            fase500 = proximaVaga.ToString("000");
            if (fase500 == null)
            {
                MessageBox.Show(this, "Fase 500 não encontrada.");
            }        
        }

        public void carregafase1000()
        {
            Model modelo = new Model();           

            string[] _caminho = CaminhoModelo.Split('\\');
            bool fab = true;
            string caminho = null;

            foreach (string item in _caminho)
            {
                if (item != "1.8 Projeto")
                {
                    if (fab)
                    {
                        caminho += item + "\\";
                    }
                }
                else
                {
                    fab = false;
                }
            }

            if (string.IsNullOrEmpty(caminho) || !Directory.Exists(caminho))
            {
                return; 
            }

            string pastaGestaoFabrico = Path.Combine(caminho, "1.9 Gestão de fabrico");
            Directory.CreateDirectory(pastaGestaoFabrico);

            var todosOsNomesDePasta = Directory.GetDirectories(pastaGestaoFabrico)
                                              .Select(Path.GetFileName)
                                              .ToList();

            int proximaVaga = 999;
            string fasePotencial;

            do
            {
                proximaVaga++;
                fasePotencial = proximaVaga.ToString("000");

            } while (todosOsNomesDePasta.Any(nomePasta => nomePasta.StartsWith(fasePotencial)));


            fase1000 = proximaVaga.ToString("000");

            if (fase1000 == null)
            {
                MessageBox.Show(this, "Fase 1000 não encontrada.");
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

        private void timer1_Tick(object sender, EventArgs e)
        {
            label5.Text = "Semana " + getweek() + " Dia " + DateTime.Now.ToString();
        }

        public int getweek()
        {
            CultureInfo ciCurr = CultureInfo.CurrentCulture;
            int weekNum = ciCurr.Calendar.GetWeekOfYear(DateTime.Now, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
            return weekNum;
        }


        private object _changedObjectHandlerLock = new object();
    
        private void Form1_Load(object sender, EventArgs e)
        {
            //_events = new Tekla.Structures.Model.Events();
            //_events.ModelObjectChanged += Events_ModelObjectChangedEvent;
            //_events.Register();
            Iniciar();            
            InitWatcher();
            this.FormClosed += Frm_Inico_FormClosed;
        }

        //void Events_ModelObjectChangedEvent(List<ChangeData> changes)
        //{
        //    /* Make sure that the inner code block is running synchronously */
        //    ArrayList b = new ArrayList();
        //    lock (_changedObjectHandlerLock)
        //    {
        //        var i = 1;
                
        //        foreach (var change in changes)
        //        {
        //            if (change.Object is Beam &&i==1)
        //            {
        //                b.Add(change.Object);
        //            }
        //            i++;
        //        }
        //    }

        //    ComunicaTekla.selectinmodel(b);


        //}

    
        private void label11_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(label11.Text);
        }       
        private void label5_MouseClick(object sender, MouseEventArgs e)
        {
            Model m = new Model();

            if (e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                if (Form.ModifierKeys == Keys.Control)
                {
                    Microsoft.Office.Interop.Outlook._Application oApp;
                    Microsoft.Office.Interop.Outlook._MailItem oMsg;
                    oApp = new Microsoft.Office.Interop.Outlook.Application();
                    oMsg = oApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
                    oMsg.To = "";
                    oMsg.CC = "";
                    string strBody3 = "<html><body>" + "<br>" + "<br>" + "</body></html>";

                    oMsg.BodyFormat = Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatHTML;
                    oMsg.Subject = m.GetProjectInfo().ProjectNumber + "-" + m.GetProjectInfo().Builder;
                    oMsg.HTMLBody = strBody3 + ReadSignature();
                    oMsg.Display();
                }
                else
                {
                    string path = @"C:\R\1.jpg";
                    if (File.Exists(path))
                    {
                        File.Delete(path);
                    }
                    try
                    {
                        Clipboard.GetImage().Save(path, System.Drawing.Imaging.ImageFormat.Png);
                    }
                    catch (System.Exception)
                    {   }

                    Microsoft.Office.Interop.Outlook._Application oApp;
                    Microsoft.Office.Interop.Outlook._MailItem oMsg;
                    oApp = new Microsoft.Office.Interop.Outlook.Application();
                    oMsg = oApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
                    oMsg.To = "";
                    oMsg.CC = "";
                    oMsg.BodyFormat = Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatHTML;

                    string strBody3 = "<html><body>" +
                        "<p style=margin:0cm;>Boa tarde,</p>" +
                        "<p style=margin:0cm;>Venho por este meio informar que a modelação está terminada, agradeço que analise e que proceda com o envio para fabrico da obra em assunto.</p> " +
                        "<br>" +
                        "<p style=margin:0cm><b><u>NOTAS IMPORTANTES:</u></b></p> " +
                        "<p style=margin:0cm>- O material a fabricar encontra-se modelado como Fase do gestor de Fases da obra em assunto</p> " +
                        "<br>" +
                        "<br>" +
                        "<a href=" + m.GetInfo().ModelPath.Replace(" ", "%20") + "> Caminho do modelo </a>" +
                        "<br>" +
                        "<br>" +
                        "<p style=margin:1cm>" + @"<img src=""" + path + @""">" + "</p>" +
                        "<br>" +
                        "<br>" +

                        "</body></html>";

                    oMsg.Subject = m.GetProjectInfo().ProjectNumber + "-" + m.GetProjectInfo().Builder;
                    oMsg.HTMLBody = strBody3;
                    oMsg.HTMLBody = oMsg.HTMLBody + ReadSignature();
                    oMsg.Display();
                    if (File.Exists(path))
                    {
                        File.Delete(path);
                    }
                }
            }
        }

        private string ReadSignature()
        {
            string appDataDir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Microsoft\\Signatures";
            string signature = string.Empty;
            DirectoryInfo diInfo = new DirectoryInfo(appDataDir);

            if (diInfo.Exists)
            {
                FileInfo[] fiSignature = diInfo.GetFiles("*.htm");
                FileInfo[] fiSignatureimg = diInfo.GetFiles("*.png", SearchOption.AllDirectories);

                //2016_ficheiros/image001.png
                if (fiSignature.Length > 0)
                {
                    StreamReader sr = new StreamReader(fiSignature[0].FullName, Encoding.Default);
                    signature = sr.ReadToEnd();

                    if (!string.IsNullOrEmpty(signature))
                    {
                        string fileName = fiSignature[0].Name.Replace(fiSignature[0].Extension, string.Empty);
                        signature = signature.Replace(fileName + "_files/", appDataDir + "/" + fileName + "_files/");
                    }
                }
                string TESTE = fiSignatureimg[0].DirectoryName.Split('\\').Last() + "/" + fiSignatureimg[0].Name;
                signature = signature.Replace(TESTE, fiSignatureimg[0].FullName);
            }


            return signature;
        }

        private void exportarNC1ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            Frm_ExportarNC1 F = new Frm_ExportarNC1();
            F.ShowDialog();
            this.Visible = true;
        }
       
        

        private void button65_Click(object sender, EventArgs e)
        {
            new TeklaMacroBuilder.MacroBuilder().PushButton("dia_draw_lock_off", "Drawing_selection")
                                   .PushButton("dia_draw_freeze_off", "Drawing_selection")
                                   .PushButton("dia_draw_ready_for_issue_off", "Drawing_selection")
                                   .PushButton("dia_draw_issue_off", "Drawing_selection").Run();
        }

        private void parametrosToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            this.Visible=false;
            Frm_Parametros F = new Frm_Parametros();
            F.ShowDialog();
            this.Visible = true;
        }

        private void quantificaçãoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            Frm_Quantificacao F = new Frm_Quantificacao();
            F.ShowDialog();
            this.Visible = true;
        }

        private void criarFasesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            Frm_CriarFase F = new Frm_CriarFase(this);
            F.ShowDialog();
            this.Visible = true;
        }

        private void soldaduraToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            Frm_Soldadura F = new Frm_Soldadura();
            F.ShowDialog();
            this.Visible = true;
        }

        private void ferramentasToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            Frm_DesenhosFerramentas F = new Frm_DesenhosFerramentas(this);
            F.ShowDialog();
            this.Visible = true;
        }

        private void macrosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process OPEN = new Process();
            OPEN.StartInfo.FileName = @"\\Marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\4 Produção\Desenvolvimentos\36.Ligaçõestekla\2021\LigacoesTekla.exe";
            OPEN.Start();
        }

        private void abrirObraToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Chb_alw_top.Checked = false;
            Process OPEN = new Process();
            OPEN.StartInfo.FileName = @"\\Marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\4 Produção\Desenvolvimentos\38-Abrir_Obra\" + "ABRIROBRA_PRETO.exe";
            OPEN.StartInfo.Arguments = label11.Text;
            OPEN.Start();
        }

        private void alteraFaseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            AlteraFase F = new AlteraFase();
            F.ShowDialog();
            this.Visible = true;
        }

        private void testeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LBLestado.Visible = true;
            ComunicaTekla.imprimepdf(ComunicaTekla.ListadeConjuntosSelec(),ComunicaTekla.ListadePecasSelec(),LBLestado);
        }      
        
        private void pDFSoldaduraToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            Frm_PDFsoldaduraescolha F = new Frm_PDFsoldaduraescolha();
            F.ShowDialog();
            this.Visible = true;

        }
        private void alterarBaseDeDadosCPEDapToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            Frm_ViewBD F = new Frm_ViewBD();
            F.ShowDialog();
            this.Visible = true;
        }
        
        private void enviarParaFabricoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            Frm_EnviarEmailparaFabrico F = new Frm_EnviarEmailparaFabrico();
            F.ShowDialog();
            this.Visible = true;
        }

        private void enviarEmailParaLotearToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            Frm_EnviarParaLotearEmail F = new Frm_EnviarParaLotearEmail();
            F.ShowDialog();
            this.Visible = true;
        }

        private void enviarEmialAprovToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            Frm_AprovisionamentosEmail F = new Frm_AprovisionamentosEmail();
            F.ShowDialog();
            this.Visible = true;
        }
        private void revisãoPeçasEConjuntosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            Frm_EnviarVariosEmail F = new Frm_EnviarVariosEmail();
            F.ShowDialog();
            this.Visible = true;
        }
        private void limparTodasAsUDADaPeçaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LimparUDAS();
        }

        private void desenhoToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            this.Size = new System.Drawing.Size(425, 350);
            webBrowser1.Visible = true;
        }

        public void Pdfsoldadura()
        {            
            string filePath = @"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\5 Soldadura\5.2 Soldadores\Lista de Soldadores\131 Lista de soldadores - CM.xlsm";

            string debugFolderPath = AppDomain.CurrentDomain.BaseDirectory;

            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(filePath);

            string outputPdfPath = Path.Combine(debugFolderPath, $"{fileNameWithoutExtension}.pdf");

            if (ShouldGeneratePdf(filePath, outputPdfPath))
            {
                ConvertExcelToPdf(filePath, outputPdfPath);

                MessageBox.Show(this, "A Lista de soldadores.pdf foi atualizada com Sucesso!.");
            }
            else
            {
               // MessageBox.Show("O arquivo PDF já está atualizado. Nenhuma ação necessária.");
            }

        }

        // Função para verificar se o arquivo PDF deve ser gerado
        static bool ShouldGeneratePdf(string excelFilePath, string pdfFilePath)
        {
            if (File.Exists(pdfFilePath))
            {
                DateTime excelLastWriteTime = File.GetLastWriteTime(excelFilePath);
                DateTime pdfLastWriteTime = File.GetLastWriteTime(pdfFilePath);

                if (excelLastWriteTime > pdfLastWriteTime)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return true;
            }
        }

        static void ConvertExcelToPdf(string inputFilePath, string outputFilePath)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = null;

            try
            {
                workbook = excelApp.Workbooks.Open(inputFilePath);

                Excel.Worksheet worksheet = workbook.Sheets[1];

                int lastRow = worksheet.Cells[worksheet.Rows.Count, 1].End(Excel.XlDirection.xlUp).Row;

                // Define o intervalo de células a ser exportado até a última linha com dados (até a coluna N)
                Excel.Range range = worksheet.Range["A1:N" + lastRow];

                worksheet.PageSetup.PrintArea = range.Address;

                worksheet.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, outputFilePath);

                
            }
            catch (Exception ex)
            {   }
            finally
            {
                if (workbook != null)
                {
                    workbook.Close(false);
                }
                excelApp.Quit();
            }
        }

        private void lotesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            string caminhoExe = @"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\4 Produção\Desenvolvimentos\37.Lotes\2021\Lotes\LOTES.exe";

            if (System.IO.File.Exists(caminhoExe))
            {
                System.Diagnostics.Process.Start(caminhoExe);
            }
            else
            {
                MessageBox.Show(this, "O arquivo não foi encontrado: " + caminhoExe, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            this.Visible = true;
        }

        private void PASTAEXPORTACAO_TextChanged(object sender, EventArgs e)
        {    }

        private void nESTDESENVOLVIMENTOToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            //Nest n = new Nest();
            //n.fazernesting(PASTAEXPORTACAO.Text);
            this.Visible = false;
            Frm_Nest n = new Frm_Nest();
            n.ShowDialog();
            this.Visible = true;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            carregafase();
            carregafase500();
            carregafase1000();
            ///////////////////////////////
            StreamReader leitor = new StreamReader("config.txt");
            PastaPartilhada = leitor.ReadLine();
            //////////////////////////////////////////
            string[] caminho = CaminhoModelo.Split('\\');
            foreach (var item in caminho)
            {
                if (item.StartsWith("20") && item.Count(char.IsNumber) == 4)
                {
                    ano = item;
                }
            }
        }

        private void LimparUDAS()
        {
            try
            {
                ArrayList pecas = ComunicaTekla.ListadePecasdoConjSelec();
                ArrayList conj = ComunicaTekla.ListadeConjuntosSelec();

                if (pecas.Count == 0 && conj.Count == 0)
                {
                    MessageBox.Show(this, "Por favor selecione pelo menos um conjunto.", "SEM CONJUNTOS SELECIONADOS", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                foreach (TSM.Part peca in pecas)
                {
                    if (pecas.Count == 1)
                    {
                        AlteraPrefixoumaPeça(peca);
                        peca.SetUserProperty("Artigo", "");
                        peca.SetUserProperty("Destinata_ext", "");
                        peca.SetUserProperty("forcar_destino", "");
                        peca.SetUserProperty("Operacoes", "");
                        peca.SetUserProperty("Artigo_interno", "");
                        peca.SetUserProperty("Fase", "");
                        peca.SetUserProperty("lote_number", "");
                        peca.SetUserProperty("lote_data", "");
                        peca.SetUserProperty("comment", "");
                        peca.SetUserProperty("ComprAdicional", "");
                        peca.SetUserProperty("Requisitos", "");
                        peca.SetUserProperty("QUANTIFICACAO", "");
                        peca.SetUserProperty("CHAPA_LACADA", "");
                        peca.SetUserProperty("Esp_chapa", "");
                        peca.SetUserProperty("Ralespcor", "");
                        peca.SetUserProperty("MaterialRevest", "");
                        peca.SetUserProperty("Local_descarga", "");
                        peca.SetUserProperty("Comentarioprep", "");
                        peca.SetUserProperty("pecaanterior", "");
                        peca.SetUserProperty("conjuntoanterior", "");
                        peca.SetUserProperty("parametroaxiliar", "");

                    }
                    else
                    {
                        AlteraPrefixo(pecas);
                        peca.SetUserProperty("Artigo", "");
                        peca.SetUserProperty("Destinata_ext", "");
                        peca.SetUserProperty("forcar_destino", "");
                        peca.SetUserProperty("Operacoes", "");
                        peca.SetUserProperty("Artigo_interno", "");
                        peca.SetUserProperty("Fase", "");
                        peca.SetUserProperty("lote_number", "");
                        peca.SetUserProperty("lote_data", "");
                        peca.SetUserProperty("comment", "");
                        peca.SetUserProperty("ComprAdicional", "");
                        peca.SetUserProperty("Requisitos", "");
                        peca.SetUserProperty("QUANTIFICACAO", "");
                        peca.SetUserProperty("CHAPA_LACADA", "");
                        peca.SetUserProperty("Esp_chapa", "");
                        peca.SetUserProperty("Ralespcor", "");
                        peca.SetUserProperty("MaterialRevest", "");
                        peca.SetUserProperty("Local_descarga", "");
                        peca.SetUserProperty("Comentarioprep", "");
                        peca.SetUserProperty("pecaanterior", "");
                        peca.SetUserProperty("conjuntoanterior", "");
                        peca.SetUserProperty("parametroaxiliar", "");

                    }
                }

                MessageBox.Show(this, "Todas as UDA's de " + pecas.Count + " Peças e " + conj.Count + " Conjuntos, foram limpas com sucesso.", "Êxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Erro: " + ex.Message);
                this.Close();
            }
        }

        public static void AlteraPrefixo(ArrayList peças)
        {
            foreach (TSM.Part part in peças)
            {
                if (part != null)
                {
                    if (part.Profile.ProfileString.Contains("CHA") || part.Profile.ProfileString.Contains("PL"))
                    {
                        part.PartNumber.StartNumber = 1;
                        part.PartNumber.Prefix = "C";
                        part.AssemblyNumber.StartNumber = 1;
                        part.AssemblyNumber.Prefix = "CJ";
                        part.Modify();
                    }
                    else
                    {

                        if (part.Material.MaterialString.Contains("8,8") || part.Material.MaterialString.Contains("5,8") || part.Material.MaterialString.Contains("10,9") || part.Profile.ProfileString.Contains("NUT_M"))
                        {
                            part.PartNumber.StartNumber = 1;
                            part.PartNumber.Prefix = "H";
                            part.AssemblyNumber.StartNumber = 1;
                            part.AssemblyNumber.Prefix = "H";
                            part.Modify();
                        }
                        else
                        {
                            part.PartNumber.StartNumber = 1;
                            part.PartNumber.Prefix = "P";
                            part.AssemblyNumber.StartNumber = 1;
                            part.AssemblyNumber.Prefix = "CJ";
                            part.Modify();
                        }
                    }
                }
            }
        }

        public static void AlteraPrefixoumaPeça(TSM.Part part)
        {
            if (part != null)
            {
                if (part.Profile.ProfileString.Contains("CHA") || part.Profile.ProfileString.Contains("PL"))
                {
                    part.PartNumber.StartNumber = 1;
                    part.PartNumber.Prefix = "C";
                    part.AssemblyNumber.StartNumber = 1;
                    part.AssemblyNumber.Prefix = "CJ";
                    part.Modify();
                }
                else
                {

                    if (part.Material.MaterialString.Contains("8,8") || part.Material.MaterialString.Contains("5,8") || part.Material.MaterialString.Contains("10,9") || part.Profile.ProfileString.Contains("NUT_M"))
                    {
                        part.PartNumber.StartNumber = 1;
                        part.PartNumber.Prefix = "H";
                        part.AssemblyNumber.StartNumber = 1;
                        part.AssemblyNumber.Prefix = "H";
                        part.Modify();
                    }
                    else
                    {
                        part.PartNumber.StartNumber = 1;
                        part.PartNumber.Prefix = "P";
                        part.AssemblyNumber.StartNumber = 1;
                        part.AssemblyNumber.Prefix = "CJ";
                        part.Modify();
                    }
                }
            }
        }

        private string caminhoLog;
        private string usuario;
        private string linhaEntrada;

        public void Iniciar()
        {
            string numeromodelo = label11.Text;

            if (string.IsNullOrWhiteSpace(numeromodelo) || numeromodelo.Equals("Sem Obra", StringComparison.OrdinalIgnoreCase))
            {
                return; 
            }

            string caminhoBase = @"\\marconi\OFELIZ\OFELIZ\OFM\2.AN\2.CM\DP\4 Produção\Desenvolvimentos\Ficheiros Temp tekla artigos (Nao Apagar)\Obras";
            caminhoLog = Path.Combine(caminhoBase, numeromodelo + "_usuarios_ativos.txt");
            usuario = Environment.UserName;
            linhaEntrada = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} | {usuario}";

            try
            {
                if (!File.Exists(caminhoLog))
                {
                    File.Create(caminhoLog).Close(); 
                }

                File.AppendAllLines(caminhoLog, new[] { linhaEntrada });

                var usuariosAtivos = File.ReadAllLines(caminhoLog)
                                    .Select(l => l.Split('|').Length > 1 ? l.Split('|')[1].Trim() : null)
                                    .Where(nome => !string.IsNullOrWhiteSpace(nome) &&
                                                   !nome.Equals(usuario, StringComparison.OrdinalIgnoreCase))
                                    .Distinct()
                                    .ToList();

                if (usuariosAtivos.Any())
                {
                    string mensagem = $"Atenção: Existe {usuariosAtivos.Count} usuário(s) a trabalhar nesta obra:\n" +
                                      string.Join("\n", usuariosAtivos);
                    MessageBox.Show(mensagem, "Alerta de User Ativos na Obra", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao gravar log de usuário:\n{ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            Application.ApplicationExit += new EventHandler(OnApplicationExit);
        }

        private void OnApplicationExit(object sender, EventArgs e)
        {
            try
            {
                if (!File.Exists(caminhoLog)) return;

                var linhas = File.ReadAllLines(caminhoLog).ToList();
                linhas.RemoveAll(l => l.Contains(usuario));
                File.WriteAllLines(caminhoLog, linhas);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao limpar log de usuário:\n{ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
                      

        private void InitWatcher()
        {
            Model modelo = new Model();
            string caminhoModelo = modelo.GetInfo().ModelPath;

            string nomeUsuario = Environment.UserName;
            filePath = Path.Combine(caminhoModelo, nomeUsuario + ".txt");

            checkTimer = new Timer();
            checkTimer.Interval = 5000;
            checkTimer.Tick += CheckDrawingAndUpdateFile;
            checkTimer.Start();

            CheckDrawingAndUpdateFile(null, null);
        }

        private void CheckDrawingAndUpdateFile(object sender, EventArgs e)
        {
            DrawingHandler drawingHandler = new DrawingHandler();
            if (!drawingHandler.GetConnectionStatus())
                return;

            Drawing currentDrawing = drawingHandler.GetActiveDrawing();
            if (currentDrawing == null)
                return;

            string drawingName = currentDrawing.Name;
            string drawingMark = currentDrawing.Mark;

            string nomeUsuario = Environment.UserName;
            nomeUsuario = nomeUsuario.Replace('.', ' ');
            nomeUsuario = string.Join(" ", nomeUsuario.Split(' ').Select(p => char.ToUpper(p[0]) + p.Substring(1).ToLower()));

            string novaLinha = $"{nomeUsuario} -- {drawingMark} de/o {drawingName}";

            if (!File.Exists(filePath) || File.ReadAllText(filePath).Trim() != novaLinha.Trim())
            {
                File.WriteAllText(filePath, novaLinha);
            }

            MostrarTxtsDasSiglasNoWebBrowser();
        }

        private void Frm_Inico_FormClosed(object sender, FormClosedEventArgs e)
        {
            checkTimer?.Stop();

            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
        }

        private void MostrarTxtsDasSiglasNoWebBrowser()
        {
            List<string> siglas = new List<string>();
            Model modelo = new Model();
            string caminhoModelo = modelo.GetInfo().ModelPath;

            ComunicaBaseDados comunicaBD = new ComunicaBaseDados();

            try
            {
                comunicaBD.ConectarBDArtigo();

                string query = "SELECT [nome.sigla] FROM dbo.nPreparadores1";
                DataTable resultado = comunicaBD.ProcurarbdArtigo(query);

                foreach (DataRow row in resultado.Rows)
                {
                    string sigla = row[0].ToString().Trim();
                    if (!string.IsNullOrEmpty(sigla))
                    {
                        siglas.Add(sigla);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao acessar banco de dados: " + ex.Message);
                return;
            }
            finally
            {
                comunicaBD.DesonectarBDArtigo();
            }

            //string html = "<html><head><style>body { font-family: Consolas; } pre { background: #eee; padding: 10px; border: 1px solid #ccc; }</style></head><body>";

            string html = @"
                            <html>
                            <head>
                            <style>
                                body {
                                    font-family: Consolas;
                                    background-color: rgb(227,227,227); /* Cor de fundo da página */
                                    margin: 20px;
                                }
                                pre {
                                    background-color: #ffffff; /* Cor de fundo das caixas de texto */
                                    padding: 10px;
                                    border: 1px solid #ccc;
                                }
                                h4 {
                                    color: #333;
                                }
                            </style>
                            </head>
                            <body>";

            bool encontrouArquivo = false;

            foreach (string sigla in siglas)
            {
                string caminhoTxt = Path.Combine(caminhoModelo, sigla + ".txt");

                if (File.Exists(caminhoTxt))
                {
                    string conteudo = File.ReadAllText(caminhoTxt)
                                      .Replace("[", "")
                                      .Replace("]", "");

                    string nomeUsuario = Environment.UserName;
                    nomeUsuario = nomeUsuario.Replace('.', ' ');
                    nomeUsuario = string.Join(" ", nomeUsuario.Split(' ').Select(p => char.ToUpper(p[0]) + p.Substring(1).ToLower()));

                    html += $"<pre><b>{System.Net.WebUtility.HtmlEncode(conteudo)}</b></pre>";
                    encontrouArquivo = true;
                }
            }

            html += "</body></html>";

            if (!encontrouArquivo)
            {
                html = "<html><body><p></p></body></html>";
            }

            webBrowser1.DocumentText = html;
        }


    }
}
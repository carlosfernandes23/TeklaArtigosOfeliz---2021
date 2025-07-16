using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Printing;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using Tekla.Structures.Drawing;
using Tekla.Structures.Filtering;
using Tekla.Structures.Filtering.Categories;
using Tekla.Structures.Model;
using static TeklaArtigosOfeliz.Frm_Quantificacao;
using TSM = Tekla.Structures.Model;
using Outlook = Microsoft.Office.Interop.Outlook;
using Image = System.Drawing.Image;
using Tekla.Structures.Model.Operations;
using Tekla.Structures.Plugins;
using Tekla.Structures.Model.UI;
using Point = Tekla.Structures.Geometry3d.Point;
using static Tekla.Structures.Filtering.Categories.PartFilterExpressions;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using WindowsInput;
using WindowsInput.Native;
using System.Text;
using TeklaModel = Tekla.Structures.Model.Model;
using Tekla.Structures.InpParser;
using System.Web.UI.WebControls.WebParts;
using static TheArtOfDevHtmlRenderer.Adapters.RGraphicsPath;


namespace TeklaArtigosOfeliz
{
    public partial class Frm_CriarFase : Form
    {
        public Frm_Inico _formpai;

        public Frm_Pecas FrmPecas { get; }

        public Frm_CriarFase(Frm_Inico formpai)
        {
            InitializeComponent();
            _formpai = formpai;
            label8.Text = "Fase - " + formpai.fase1000;
            label7.Text = "Fase - " + formpai.fase500;
            label6.Text = "Fase - " + formpai.fase;
            label11.Text = _formpai.label11.Text;

        }

        public Frm_CriarFase(Frm_Pecas frmPecas)
        {
            FrmPecas = frmPecas;
        }

        private void FrmCriarFase_Load(object sender, EventArgs e)
        {
            TopMost = true;
            Chb_alw_top.Checked = true;
            listBox1.Items.Clear();
            //foreach (var item in Frm_Inico.str)
            //{
            //    listBox1.Items.Add(item);
            //}

        }

        private void VerificarQuant(ArrayList pecas)
        {
            bool temQuantificacao = true; 

            foreach (TSM.Part part in pecas)
            {
                string quant = string.Empty;
                string artigo = string.Empty;

                part.GetReportProperty("QUANTIFICACAO", ref quant);
                part.GetReportProperty("Artigo", ref artigo);

                if (string.IsNullOrEmpty(quant) && artigo == "Perfil")
                {
                    temQuantificacao = false; 
                    break; 
                }
            }
            if (!temQuantificacao)
            {
                MessageBox.Show(this, "Neste lote, foram encontrados perfis sem Quantificação", " Perfis sem Quantificação !! ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void VerificarArtigoInterno(ArrayList parts)
        {
            //foreach (TSM.Part item in parts)
            //{
            //    string Perfil = item.Profile.ProfileString.ToLower();

            //    if (Perfil.Contains("p0") || Perfil.Contains("p1") || Perfil.Contains("p2") || Perfil.Contains("p3") || Perfil.Contains("p4") || Perfil.Contains("P5") || Perfil.Contains("p6") || 
            //        Perfil.Contains("z1") || Perfil.Contains("z2") || Perfil.Contains("z3") ||
            //        Perfil.Contains("c1") || Perfil.Contains("c2") || Perfil.Contains("c3") || Perfil.ToUpper().StartsWith("SUPEROMEGA"))
            //    {
            //        foreach (TSM.Part part in parts)
            //        {
            //            string ArtigoInterno = null;
            //            part.GetUserProperty("Artigo_interno", ref ArtigoInterno);

            //            string consultaSQL = @"SELECT [Material] FROM [dbo].[Perfilagem3] WHERE [Perfil] = '" + Perfil + "'";

            //            if (string.IsNullOrEmpty(ArtigoInterno))
            //            {
            //                ComunicaBDtekla a = new ComunicaBDtekla();
            //                a.ConectarBD();

            //                var materiais = a.Procurarbd(consultaSQL);

            //                if (materiais != null && materiais.Count > 0)
            //                {
            //                    string materiaisStr = string.Join("\n ", materiais);

            //                    MessageBox.Show("O Perfil: " + Perfil + " Não tem Artigo Interno, \n Apenas existe para os seguintes materiais: \n " + materiaisStr);
            //                }
            //                else
            //                {
            //                    MessageBox.Show("O Perfil: " + Perfil + " Não tem Artigo Interno, \n E não existem materiais associados.");
            //                }

            //                a.DesonectarBD();
            //            }
            //            else
            //            {
            //                //MessageBox.Show("O Perfil: " + Perfil + " Tem Artigo Interno: " + ArtigoInterno);
            //            }
            //        }
            //    }
            //}

            HashSet<string> perfisVerificados = new HashSet<string>(); // P

            foreach (TSM.Part item in parts)
            {
                string Perfil = item.Profile.ProfileString.ToLower();

                if (Perfil.Contains("p0") || Perfil.Contains("p1") || Perfil.Contains("p2") || Perfil.Contains("p3") || Perfil.Contains("p4") || Perfil.Contains("p5") || Perfil.Contains("p6") ||
                    Perfil.Contains("z1") || Perfil.Contains("z2") || Perfil.Contains("z3") ||
                    Perfil.Contains("c1") || Perfil.Contains("c2") || Perfil.Contains("c3") || Perfil.ToUpper().StartsWith("SUPEROMEGA"))
                {
                    if (!perfisVerificados.Contains(Perfil))
                    {
                        string ArtigoInterno = null;
                        item.GetUserProperty("Artigo_interno", ref ArtigoInterno);

                        string consultaSQL = @"SELECT [Material] FROM [dbo].[Perfilagem3] WHERE [Perfil] = '" + Perfil + "'";

                        if (string.IsNullOrEmpty(ArtigoInterno))
                        {
                            ComunicaBDtekla a = new ComunicaBDtekla();
                            a.ConectarBD();

                            var materiais = a.Procurarbd(consultaSQL);

                            if (materiais != null && materiais.Count > 0)
                            {
                                string materiaisStr = string.Join("\n ", materiais);
                                MessageBox.Show(this, "O Perfil: " + Perfil + " Não tem Artigo Interno, \n Apenas existe para os seguintes materiais: \n " + materiaisStr);
                            }
                            else
                            {
                                MessageBox.Show(this, "O Perfil: " + Perfil + " Não tem Artigo Interno, \n E não existem materiais associados.");
                            }

                            a.DesonectarBD();
                        }
                        else
                        {
                            //MessageBox.Show("O Perfil: " + Perfil + " Tem Artigo Interno: " + ArtigoInterno);
                        }

                        // Adiciona o perfil ao conjunto de perfis verificados
                        perfisVerificados.Add(Perfil);
                    }
                }
            }
        }


        private void textBox1_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar);
            if (Convert.ToInt32(e.KeyChar) == 13)
            {
                if (int.TryParse(textBox1.Text, out int valor))
                {
                    string valorFormatado = valor.ToString("D3");

                    label6.Text = "Fase - " + valorFormatado;
                }
                textBox1.Text = "";
            }
        }

        private void textBox2_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar);
            if (Convert.ToInt32(e.KeyChar) == 13)
            {
                if (textBox2.Text.Length == 3 && int.TryParse(textBox2.Text, out int valor))
                {
                    string valorFormatado = valor.ToString("D3");
                    label7.Text = "Fase - " + valorFormatado;
                    textBox2.Text = "";
                }
                else
                {
                    MessageBox.Show(this, "Por favor, insira exatamente 3 dígitos.");
                }                
            }
        }

        private void listBox1_DoubleClick(object sender, EventArgs e)
        {

            if (checkBox1.Checked)
            {
                TSM.Model MODELO = new Model();
                BinaryFilterExpression FILTRO = new BinaryFilterExpression(new TemplateFilterExpressions.CustomString("USERDEFINED.Fase"), StringOperatorType.IS_EQUAL, new StringConstantFilterExpression(listBox1.SelectedItem.ToString()));
                ModelObjectEnumerator Objects = MODELO.GetModelObjectSelector().GetObjectsByFilter(FILTRO);
                ArrayList tudo = new ArrayList();

                while (Objects.MoveNext())
                {
                    TSM.Assembly ASS = Objects.Current as TSM.Assembly;
                    if (ASS != null)
                    {
                        tudo.Add(ASS);
                    }

                }
                TSM.UI.ModelObjectSelector MOS = new TSM.UI.ModelObjectSelector();
                MOS.Select(tudo);
                MODELO.CommitChanges();
            }
            else
            {
                TSM.Model MODELO = new Model();
                BinaryFilterExpression FILTRO = new BinaryFilterExpression(new TemplateFilterExpressions.CustomString("USERDEFINED.Fase"), StringOperatorType.IS_EQUAL, new StringConstantFilterExpression(listBox1.SelectedItem.ToString()));
                ModelObjectEnumerator Objects = MODELO.GetModelObjectSelector().GetObjectsByFilter(FILTRO);
                ArrayList tudo = new ArrayList();

                while (Objects.MoveNext())
                {
                    TSM.Part peca = Objects.Current as TSM.Part;
                    if (peca != null)
                    {
                        ModelObjectEnumerator pararafusos = peca.GetBolts();
                        while (pararafusos.MoveNext())
                        {
                            tudo.Add(pararafusos.Current);
                        }
                        tudo.Add(peca);
                    }

                }
                TSM.UI.ModelObjectSelector MOS = new TSM.UI.ModelObjectSelector();
                MOS.Select(tudo);
                MODELO.CommitChanges();
            }

        }

        private void exc3()
        {
            int ClasseDeExecusao = 0;

            TSM.Model model = new Model();
            model.GetProjectInfo().GetUserProperty("PROJECT_USERFIELD_2", ref ClasseDeExecusao);
            if (ClasseDeExecusao == 2 || ClasseDeExecusao == 3)
            {
                ArrayList pecas = ComunicaTekla.ListadePecasdoConjSelec();
                ComunicaTekla.selectinmodel(pecas);
                new TeklaMacroBuilder.MacroBuilder().Callback("acmd_partnumbers_selected", "", "main_frame").Run();
                ArrayList conjuntos = new ArrayList();
                conjuntos = ComunicaTekla.ListadeconjdaspecasSelec();
                string marcaconjunto = null;

                DataTable dt = new DataTable();
                dt.Columns.Add("ASSEMBLYPOS");
                dt.Columns.Add("GUID");
                dt.Columns.Add("NEWNUMBER");
                DataRow dr = null;

                foreach (Assembly item in conjuntos)
                {
                    item.GetReportProperty("ASSEMBLY_POS", ref marcaconjunto);
                    dr = dt.NewRow();
                    dr["ASSEMBLYPOS"] = marcaconjunto;
                    dr["GUID"] = item.Identifier.GUID;
                    dt.Rows.Add(dr);
                }
                DataView view = dt.DefaultView;
                view.Sort = "ASSEMBLYPOS ASC";
                DataTable sortCONJUNTOS = view.ToTable();
                int id = 1;
                for (int i = 0; i < sortCONJUNTOS.Rows.Count; i++)
                {
                    sortCONJUNTOS.Rows[i][2] = id;
                    if ((i + 1) < sortCONJUNTOS.Rows.Count)
                    {
                        if (sortCONJUNTOS.Rows[i].ItemArray[0].ToString() == sortCONJUNTOS.Rows[i + 1].ItemArray[0].ToString())
                        {
                            id += 1;
                        }
                        else
                        {
                            id = 1;
                        }
                    }
                }
                ArrayList ass = new ArrayList();
                foreach (DataRow item in sortCONJUNTOS.Rows)
                {
                    Tekla.Structures.Identifier ID = new Tekla.Structures.Identifier(item.ItemArray[1].ToString());
                    Assembly Modelassembly = new Model().SelectModelObject(ID) as Assembly;
                    Modelassembly.SetUserProperty("USER_FIELD_2", item.ItemArray[2].ToString());
                    ass.Add(Modelassembly);
                }
                ComunicaTekla.selectinmodel(ass);
            }
        }

        private void exc2(bool selectpartorassembly = false)
        {

            if (label11.Text == new Model().GetProjectInfo().ProjectNumber)
            {
                ArrayList listBox2 = new ArrayList();
                ArrayList LOTES = new ArrayList();
                double loosepart = 0;
                string pintura = "";
                ArrayList pecas = new ArrayList();
                pecas = ComunicaTekla.ListadePecasdoConjSelec();
                TSM.ModelObjectEnumerator modelEnumerator = new TSM.UI.ModelObjectSelector().GetSelectedObjects();

                while (modelEnumerator.MoveNext())
                {
                    TSM.Assembly ass = modelEnumerator.Current as TSM.Assembly;
                    if (ass != null)
                    {
                        foreach (var item in ass.GetSecondaries())
                        {
                            pecas.Add(item);
                        }
                        pecas.Add(ass.GetMainPart());
                    }
                }
                ArrayList pecas500 = new ArrayList();
                ArrayList pecan = new ArrayList();
                foreach (TSM.Part peca in pecas)
                {

                    pintura = "";

                    peca.GetReportProperty("ASSEMBLY.USERDEFINED.pintura", ref pintura);


                    if (pintura.Contains("PINTURA"))
                    {
                        pintura = "";
                    }



                    peca.GetReportProperty("ASSEMBLY.SUPPLEMENT_PART_WEIGHT", ref loosepart);
                    string perfil = peca.Profile.ProfileString.ToLower();
                    bool parafusoescareado = false;

                    List<string> list = new List<string>();

                    ComunicaBDtekla n = new ComunicaBDtekla();
                    n.ConectarBD();
                    list = n.Procurarbd("SELECT [Perfil] FROM [ArtigoTekla].[dbo].[Perfilagem3] where [Perfil]='" + perfil + "'");
                    n.DesonectarBD();

                    if (loosepart == 0 && pintura == "" && (perfil.Contains("pl") || perfil.Contains("cha") || perfil.Contains("z") || perfil.Contains("c1") || perfil.Contains("c2") || perfil.Contains("c3") || perfil.Contains("p3") || perfil.Contains("p2") || perfil.Contains("p1") || perfil.Contains("p5") || perfil.Contains("p6") || perfil.Contains("p0") || perfil.Contains("h60") || perfil.Contains("chg") || perfil.Contains("saida") || perfil.Contains("omega") || perfil.Contains("VRS") || perfil.Contains("pc-") || perfil.Contains("pf-") || perfil.Contains("pw-") || perfil.Contains("bm") || perfil.Contains("wm") || perfil.Contains("max") || perfil.Contains("ca") || list.Count > 0))
                    {
                        if (perfil.Contains("pl") || perfil.Contains("cha"))
                        {
                            string normaparafuso = "";
                            TSM.ModelObjectEnumerator BoltGrupOnPart = peca.GetBolts();

                            while (BoltGrupOnPart.MoveNext())
                            {

                                BoltGrupOnPart.Current.GetReportProperty("BOLT_STANDARD", ref normaparafuso);
                                if (normaparafuso.Contains("7991"))
                                {
                                    parafusoescareado = true;

                                }

                            }
                        }

                        if (parafusoescareado == true)
                        {
                            pecan.Add(peca);
                        }
                        else
                        {
                            pecas500.Add(peca);
                        }

                    }
                    else
                    {
                        pecan.Add(peca);
                    }
                }
                if (pecan.Count != 0)
                {
                    ComunicaTekla.EnviaproPriedadePeca(pecan, "Fase", label6.Text.ToLower().Replace("fase - ", ""));
                    ComunicaTekla.AlteraPrefixo(pecan, int.Parse(label6.Text.ToLower().Replace("fase - ", "")), int.Parse(label8.Text.ToLower().Replace("fase - ", "")));
                    if (!listBox1.Items.Contains(label6.Text.ToLower().Replace("fase - ", "")))
                    {
                        listBox1.Items.Add(label6.Text.ToLower().Replace("fase - ", ""));
                        Frm_Inico.str.Add(label6.Text.ToLower().Replace("fase - ", ""));
                    }
                }
                string lote = null;
                int fase = int.Parse(label7.Text.ToLower().Replace("fase - ", ""));
                if (pecas500.Count != 0)
                {
                    foreach (TSM.Part peca in pecas500)
                    {
                        peca.GetReportProperty("USERDEFINED.lote_number", ref lote);

                        bool found = false;

                        foreach (var item in listBox2)
                        {
                            try
                            {
                                if (item.ToString().Split(',')[0].Contains(lote.ToString()))
                                {
                                    found = true;
                                    break;
                                }
                            }
                            catch (System.Exception)
                            {
                                MessageBox.Show(this, "Erro: Encontraram-se peças sem lotes associados", "ERRO peças sem Lote", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                break;
                            }

                        }
                        if (!found)
                        {
                            if (listBox2.Count == 0)
                            {
                                listBox2.Add(lote + "," + fase);
                            }
                            else
                            {
                                fase = fase + 1;
                                listBox2.Add(lote + "," + fase);
                            }
                        }
                    }

                }
                foreach (TSM.Part peca in pecas500)
                {
                    peca.GetReportProperty("USERDEFINED.lote_number", ref lote);
                    bool found = false;
                    string fase1 = "";
                    foreach (var item in listBox2)
                    {
                        if (item.ToString().Split(',')[0].Contains(lote.ToString()))
                        {
                            fase1 = item.ToString().Split(',')[1].ToString();
                            found = true;
                            break;
                        }
                    }
                    if (found)
                    {
                        ComunicaTekla.EnviaproPriedadePeca(peca, "Fase", fase1);
                        ComunicaTekla.AlteraPrefixo(peca, int.Parse(fase1), int.Parse(label8.Text.ToLower().Replace("fase - ", "")));
                        if (!listBox1.Items.Contains(fase1))
                        {
                            listBox1.Items.Add(fase1);
                            Frm_Inico.str.Add(fase1);
                        }
                    }
                }

                exc3();
                ComunicaTekla.selectinmodel(pecas);

                new TeklaMacroBuilder.MacroBuilder().Callback("acmd_partnumbers_selected", "", "main_frame").Run();

                ComunicaTekla.selectinmodel(ComunicaTekla.ListadeconjdaspecasSelec());

                if (pecas.Count > 0)
                {
                    MessageBox.Show(this, "Foi realizada a alteração da fase para " + pecas.Count + " peças", "Êxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(this, "Por favor. É necessário selecionar pelo menos um conjunto ", "SEM PEÇAS SELECIONADAS", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show(this, "O projeto em uso neste programa difere do projeto atualmente aberto no Tekla.", "Erro de Incompatibilidade de projetos", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            //// NÚMERO DE FICHEIROS NA IMPRESSORA 

            string printerName = "PDFCreator";
            PrintServer ps = new PrintServer();
            PrintQueue pq = ps.GetPrintQueue(printerName);
            if (pq.NumberOfJobs == 0)
            {
                this.Visible = false;
                timer1.Enabled = false;
                Frm_ListaOFeliz f = new Frm_ListaOFeliz(_formpai);
                LBLestado.Text = "Todos os dados foram criados com sucesso";
                f.ShowDialog();
                this.Visible = true;
            }
            else
            {
                LBLestado.Text = "A impressão ainda não foi realizada " + pq.NumberOfJobs.ToString() + " Documentos";
            }
           
        }


        public class AppAbrirTekla
        {
            [DllImport("user32.dll", SetLastError = true)]
            public static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

            [DllImport("user32.dll", SetLastError = true)]
            public static extern IntPtr GetWindowText(IntPtr hWnd, StringBuilder text, int count);

            [DllImport("user32.dll", SetLastError = true)]
            public static extern IntPtr GetForegroundWindow();

            [DllImport("user32.dll")]
            public static extern bool SetForegroundWindow(IntPtr hWnd);

            [DllImport("user32.dll")]
            public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

            const int SW_RESTORE = 5;

            public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

            public void TrazerTeklaParaFrente()
            {
                EnumWindows(new EnumWindowsProc(EnumWindowCallback), IntPtr.Zero);
            }

            private bool EnumWindowCallback(IntPtr hWnd, IntPtr lParam)
            {
                StringBuilder windowTitle = new StringBuilder(256);
                GetWindowText(hWnd, windowTitle, 256);

                if (windowTitle.ToString().StartsWith("Tekla Structures"))
                {
                    ShowWindow(hWnd, SW_RESTORE);
                    SetForegroundWindow(hWnd);
                    SimularTeclas();
                    return false;
                }

                return true;
            }

            private void SimularTeclas()
            {
                var simulator = new InputSimulator();
                simulator.Keyboard.ModifiedKeyStroke(new[] { VirtualKeyCode.CONTROL, VirtualKeyCode.SHIFT }, VirtualKeyCode.F3);
            }
        }

        //public class AppAbrirPrimavera
        //{

        //    [DllImport("user32.dll")]
        //    public static extern bool SetForegroundWindow(IntPtr hWnd);

        //    [DllImport("user32.dll")]
        //    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        //    const int SW_RESTORE = 9;

        //    public void AbrirPrimaveira()
        //    {
        //        try
        //        {

        //            string appPath = @"C:\Program Files (x86)\PRIMAVERA\SG900\Apl\Erp900LE.exe";

        //            if (System.IO.File.Exists(appPath))
        //            {
        //                Process process = Process.Start(appPath);

        //                Thread.Sleep(1000);

        //                IntPtr hWnd = process.MainWindowHandle;

        //                if (hWnd != IntPtr.Zero)
        //                {
        //                    ShowWindow(hWnd, SW_RESTORE);

        //                    SetForegroundWindow(hWnd);
        //                }
        //                else
        //                {
        //                    MessageBox.Show("Não foi possível obter a janela do aplicativo.");
        //                }
        //            }
        //            else
        //            {
        //                MessageBox.Show("O Primavera não foi encontrado no Pc.");
        //            }
        //        }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show("Erro ao tentar abrir o Primavera: " + ex.Message);
        //        }
        //    }
        //}

        public class AppAbrirPrimavera
        {
            [DllImport("user32.dll")]
            public static extern bool SetForegroundWindow(IntPtr hWnd);

            [DllImport("user32.dll")]
            public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

            const int SW_RESTORE = 9;

            public void AbrirPrimaveira()
            {
                try
                {
                    string nomeProcesso = "Erp900LE"; 
                    string appPath = @"C:\Program Files (x86)\PRIMAVERA\SG900\Apl\Erp900LE.exe";

                    Process[] processos = Process.GetProcessesByName(nomeProcesso);

                    if (processos.Length > 0)
                    {
                        Process processoExistente = processos[0];

                        IntPtr hWnd = processoExistente.MainWindowHandle;

                        if (hWnd != IntPtr.Zero)
                        {
                            ShowWindow(hWnd, SW_RESTORE);       
                            SetForegroundWindow(hWnd);          
                        }
                        else
                        {
                            MessageBox.Show("Primavera já está aberto, mas não foi possível aceder à janela principal.");
                        }
                    }
                    else
                    {
                        if (System.IO.File.Exists(appPath))
                        {
                            Process processoNovo = Process.Start(appPath);

                            Thread.Sleep(1000); 

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


        public void CreateXmlFile(string numeroObra)
        {

            string ano = "20" + numeroObra.Substring(0, 2);

            string caminho1 = @"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\1 Obras\";
            string caminho2 = Path.Combine(caminho1, ano, numeroObra, "1.8 Projeto", "1.8.2 Tekla");

            string[] subpastas = Directory.GetDirectories(caminho2);
            if (subpastas.Length == 0)
            {
                Console.WriteLine("Nenhuma subpasta encontrada em " + caminho2);
                return;
            }

            string primeiraPasta = subpastas[0];

            string caminho3 = Path.Combine(primeiraPasta, "attributes");

            Directory.CreateDirectory(caminho1);
            Directory.CreateDirectory(caminho2);
            Directory.CreateDirectory(primeiraPasta);
            Directory.CreateDirectory(caminho3);

            string filePath = Path.Combine(caminho3, $"{numeroObra}.TeklaPowerFabPluginSettings.xml");

            if (File.Exists(filePath))
            {
                return;
            }

            string xmlContent = $@"
            <FabSuiteTeklaDataExchangeSettings xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns=""http://www.fabsuite.com/XML_Schemas/FabSuiteTeklaDataExchangeSettings0100.xsd"">
            <LastSettings>
            <LastAction>Export</LastAction>
            <LastImportSettings>
            <ImportFilename/>
            <ReadStatusOf>DrawingsMainMembers</ReadStatusOf>
            <ImportApprovalStatus>true</ImportApprovalStatus>
            <ImportAssemblyStatus>true</ImportAssemblyStatus>
            <ImportDateIssued>true</ImportDateIssued>
            <ImportShopStatus>true</ImportShopStatus>
            <ImportDateFabricationCompleted>true</ImportDateFabricationCompleted>
            <ImportLoadNumber>true</ImportLoadNumber>
            <ImportLoadStatus>true</ImportLoadStatus>
            <ImportPONumber>true</ImportPONumber>
            <ImportVendor>true</ImportVendor>
            <ImportHeatNumber>true</ImportHeatNumber>
            <ImportDateDue>true</ImportDateDue>
            <ImportDateReceived>true</ImportDateReceived>
            </LastImportSettings>
            <LastExportSettings>
            <ExportFilename>.\Tekla PowerFab\{numeroObra}_L0_F0_R0.zip</ExportFilename>
            <ExportFilenameExtension>.zip</ExportFilenameExtension>
            <AutoGenerateFilename>false</AutoGenerateFilename>
            <ExportDrawings>SelectedFromDrawingList</ExportDrawings>
            <ExportDrawingsOnlySkipAssemblies>false</ExportDrawingsOnlySkipAssemblies>
            <IncludeSinglePartDrawings>true</IncludeSinglePartDrawings>
            <IncludeGeneralArrangementDrawings>false</IncludeGeneralArrangementDrawings>
            <IncludeMultiDrawings>false</IncludeMultiDrawings>
            <ExportDrawingUserDefinedFields>ExportUDAsFromBoth</ExportDrawingUserDefinedFields>
            <ExportPartUserDefinedFields>ExportUDAsFromBoth</ExportPartUserDefinedFields>
            <IncludeBoltsNutsWashers>false</IncludeBoltsNutsWashers>
            <ExportBoltNutWasherUserDefinedFields>DontExportUDAs</ExportBoltNutWasherUserDefinedFields>
            <IncludeStuds>false</IncludeStuds>
            <ExportStudUserDefinedFields>DontExportUDAs</ExportStudUserDefinedFields>
            <CNCFiles>UseCNCFilesFromDirectory</CNCFiles>
            <CNCSettings>standard</CNCSettings>
            <CNCDirectory>\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\1 Obras\{ano}\{numeroObra}\1.9 Gestão de fabrico</CNCDirectory>
            <DrawingFiles>UseDrawingFilesFromDirectory</DrawingFiles>
            <DrawingDirectory>\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\1 Obras\{ano}\{numeroObra}\1.9 Gestão de fabrico</DrawingDirectory>
            <AssemblyFileDirectory>\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\1 Obras\{ano}\{numeroObra}\1.9 Gestão de fabrico</AssemblyFileDirectory>
            <PartFileDirectory>\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\1 Obras\{ano}\{numeroObra}\1.9 Gestão de fabrico</PartFileDirectory>
            <GAFileDirectory>.\01Desenhos\Geral</GAFileDirectory>
            <MultiFileDirectory>.\01Desenhos\Multi</MultiFileDirectory>
            <CompressOutput>true</CompressOutput>
            <OldBoltShapeLogic>false</OldBoltShapeLogic>
            <ConvertPartDelimiterToUnderscore>false</ConvertPartDelimiterToUnderscore>
            </LastExportSettings>
            </LastSettings>
            </FabSuiteTeklaDataExchangeSettings>";

            try
            {
                File.WriteAllText(filePath, xmlContent);
                MessageBox.Show(this, $"Ficheiro XML criado com sucesso na obra. {numeroObra}");
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Erro ao criar o arquivo XML: {ex.Message}");
            }


        }



        // Para o Outlook
        /* private void OpenOutlookAndCreateEmail()
        {            
         string nomeDaObra = string.Empty;
         string ano10 = null;

         try
         {
             MessageBox.Show($"Por favor verifique o email que foi Aberto!", "Verifique o Email", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

             Model modelo = new Model();

             string nomeProjeto = modelo.GetProjectInfo().Name;
             string obra = modelo.GetProjectInfo().ProjectNumber;
             string fase = label6.Text;


             if (fase.Contains("Fase -"))
             {
                 fase = fase.Replace("Fase -", "");
             }

             Outlook.Application outlookApp = new Outlook.Application();

             if (outlookApp != null)
             {
                 Outlook.MailItem mailItem = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);

                 string nomeUsuario = outlookApp.Session.CurrentUser.Name;

                 if (nomeUsuario.Contains(" | O FELIZ Metalomecânica"))
                 {
                     nomeUsuario = nomeUsuario.Replace(" | O FELIZ Metalomecânica", "");
                 }

                 string modelPath = modelo.GetInfo().ModelPath;
                 DirectoryInfo up = new DirectoryInfo(modelPath);
                 string ultimaPasta = up.Name; 

                 string caminho = modelPath + "\\Tekla PowerFab";
                 string nomePastaMaisRecente = string.Empty;
                 string lote = string.Empty;

                 if (Directory.Exists(caminho))
                 {
                     string[] arquivosZip = Directory.GetFiles(caminho, "*.zip");

                     if (arquivosZip.Length > 0)
                     {
                         var arquivoMaisRecente = arquivosZip
                             .OrderByDescending(f => new FileInfo(f).CreationTime)
                             .First();  

                         nomePastaMaisRecente = Path.GetFileName(arquivoMaisRecente);
                         MessageBox.Show("A pasta mais recente criada dentro de 'Tekla PowerFab' é: " + nomePastaMaisRecente);

                     }
                     else
                     {
                         MessageBox.Show("Não há arquivos .zip dentro de 'Tekla PowerFab'.");
                     }
                 }
                 else
                 {
                     MessageBox.Show($"A pasta '{caminho}' não existe.");
                 }

                 if (!string.IsNullOrEmpty(nomePastaMaisRecente))
                 {
                     string pattern = @"L(\d+)";  
                     Match match = Regex.Match(nomePastaMaisRecente, pattern);

                     if (match.Success)
                     {

                         lote = match.Groups[1].Value;
                         MessageBox.Show("Número extraído para lote: " + lote);
                                                                            }
                     else
                     {
                         MessageBox.Show("Não foi possível encontrar o número após 'L_' na pasta.");
                     }
                 }
                 else
                 {
                     Console.WriteLine("Não foi possível determinar o nome da pasta mais recente.");
                 }


                 string linkTexto = ".\\Tekla PowerFab\\" + nomePastaMaisRecente ;                    

                 mailItem.To = "sofia.domingues@ofeliz.com";

                 mailItem.CC = "helder.silva@ofeliz.com; luis.silva@ofeliz.com;";

                 mailItem.Subject = ultimaPasta + "-- PowerFab";

                 mailItem.BodyFormat = Outlook.OlBodyFormat.olFormatRichText;

                 string saudacao = GetSaudacao();  

                 string corpoEmail = "<html><body>";
                 corpoEmail += "<font face = 'Calibri ' size = '3' > <p>" + saudacao + "</font></p>";

                 corpoEmail += "<font face='Calibri ' size='3'><p>Venho por este meio informar, que já foi emitido dentro da pasta da obra em assunto, o PowerFab &nbsp;"                        
                              + "<span style='color:#00B0F0; display:inline-block; margin-right:10px;'><u>"
                              + obra  + "&nbsp Lote " + lote + "&nbsp Fase " + fase + "</font></u> </span></p>";


                 corpoEmail += "<font face = 'Calibri ' size = '3' ><p><b><u> PROCESSO DE FABRICO: </u></b>";

                 corpoEmail += "<font face = 'Calibri ' size = '3' style='color:#5B9BD5;'>"
                            + "<a href='file:///" + caminho.Replace("\\", "/") + "' style='color:#5B9BD5; text-decoration: none;'>" + linkTexto + "</a>" + "</font> ";                            


                 corpoEmail += "<font face = 'Calibri ' size = '3' ><p>  Melhores Cumprimentos,</p> </font> <br>";
                 corpoEmail += "<font face = 'Calibri' size = '3' > <b>" + nomeUsuario + "</b> </Font> <br>";
                 corpoEmail += "<font face = 'Calibri' size = '3' > Construção Metálica | Preparador </Font> <br>";
                 corpoEmail += "<font face = 'Calibri' size = '3' > T + 351 253 080 609 * </font> <br>";
                 corpoEmail += "<font color='red' font face = 'Calibri ' size = '3'> ofeliz.com </font> <br>";
                 corpoEmail += "<p><a href='https://www.ofeliz.com'><img src='cid:imagemOfeliz' alt='Logo Ofeliz' width='127' height='34'></a></p>";
                 corpoEmail += "<i><font color='Light grey' font face = 'Calibri ' size = '1.5'> Alvará Nº 10553 – Pub. *Chamada para a rede fixa nacional. </font> </i><br>";
                 corpoEmail += "<i><font color='green' font face = 'Calibri ' size = '1.5'> Antes de imprimir este e-mail tenha em consideração o meio ambiente. </font> </i><br>";
                 corpoEmail += "</body></html>";


                 mailItem.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
                 mailItem.HTMLBody = corpoEmail;

                 string imagePath = @"C:\Users\carlos.alves\Desktop\email\ofeliz_logo.png";
                 Outlook.Attachment imageAttachment = mailItem.Attachments.Add(imagePath, Outlook.OlAttachmentType.olByValue, 0, "imagemOfeliz");

                 imageAttachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "imagemOfeliz");

                 mailItem.Display();
             }               

         }
         catch (Exception ex)
         {
             MessageBox.Show("Erro ao enviar o e-mail ou ao processar os dados: " + ex.Message);
         }
        }
        */


        public void EnviarPowerFab_OpenEmailPreviewAndCreateEmail()
        {
            string SubjectEnviarPowerFab = string.Empty;
            string lote = string.Empty; 
            string Fase = string.Empty;
            string Revisao = string.Empty;

            try
                {

                Model modelo = new Model();
                string nomeProjeto = modelo.GetProjectInfo().Name;
                string obra = modelo.GetProjectInfo().ProjectNumber;
                string modelPath = modelo.GetInfo().ModelPath;
                DirectoryInfo up = new DirectoryInfo(modelPath);
                string ultimaPasta = up.Name;
                string caminho = modelPath + "\\Tekla PowerFab";
                string nomePastaMaisRecente = string.Empty;

                string imagemOfelizFilePath = @"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\4 Produção\Desenvolvimentos\Ficheiros Temp tekla artigos (Nao Apagar)\ofeliz_logo.png";

                string nomeUsuario = Environment.UserName;

                nomeUsuario = nomeUsuario.Replace('.', ' ');
                nomeUsuario = string.Join(" ", nomeUsuario.Split(' ').Select(p => char.ToUpper(p[0]) + p.Substring(1).ToLower()));

                //if (Directory.Exists(caminho))
                //{
                //    string[] arquivosZip = Directory.GetFiles(caminho, "*.zip");

                //    if (arquivosZip.Length > 0)
                //    {
                //        var arquivoMaisRecente = arquivosZip
                //            .OrderByDescending(f => new FileInfo(f).CreationTime)
                //            .First();

                //        nomePastaMaisRecente = Path.GetFileName(arquivoMaisRecente);
                //        //MessageBox.Show("A pasta mais recente criada dentro de 'Tekla PowerFab' é: " + nomePastaMaisRecente);
                //    }
                //    else
                //    {
                //        MessageBox.Show("Não há arquivos .zip dentro de 'Tekla PowerFab'.");
                //    }

                if (Directory.Exists(caminho))
                {
                    string[] arquivosZip = Directory.GetFiles(caminho, "*.zip");

                    if (arquivosZip.Length > 0)
                    {
                        var arquivoMaisRecente = arquivosZip
                            .OrderByDescending(f => new FileInfo(f).CreationTime)
                            .First();

                        nomePastaMaisRecente = Path.GetFileName(arquivoMaisRecente);
                        DateTime dataCriacaoMaisRecente = new FileInfo(arquivoMaisRecente).CreationTime;
                        DateTime dataAtual = DateTime.Now.Date; 

                        if (dataCriacaoMaisRecente.Date == dataAtual)
                        {
                            //MessageBox.Show("A pasta mais recente criada dentro de 'Tekla PowerFab' é: " + nomePastaMaisRecente + "\nA versão do PowerFab é antiga.");
                        }
                        else
                        {
                            MessageBox.Show(this, "Atenção: A pasta do Poerfab dentro da pasta não corresponde a data de hoje:" + nomePastaMaisRecente);
                        }
                    }
                    else
                    {
                        MessageBox.Show(this, "Não Existe Arquivos .zip dentro de 'Tekla PowerFab'.");
                    }
                }
                else
                {
                    MessageBox.Show(this, $"A pasta '{caminho}' não existe.");
                }
                if (!string.IsNullOrEmpty(nomePastaMaisRecente))
                {
                    string pattern = @"L(\d+).*F(\d+).*R(\d+)";

                    Match match = Regex.Match(nomePastaMaisRecente, pattern);

                    if (match.Success)
                    {
                        lote = match.Groups[1].Value;  
                        Fase = match.Groups[2].Value; 
                        Revisao = match.Groups[3].Value;  

                    }
                    else
                    {
                        MessageBox.Show(this, "Não foi possível localizar os números de Lote, Fase e Revisão na pasta do PowerFab");
                    }
                }
                else
                {
                    MessageBox.Show(this, "Não foi possível determinar o nome da pasta mais recente.");
                }

                ultimaPasta = ultimaPasta.Replace("_", "-");
                SubjectEnviarPowerFab = ultimaPasta + " -- PowerFab";

                string Total = obra + "_L" + lote + "_F" + Fase + ".zip";
                string linkTexto = ".\\Tekla PowerFab\\" + Total;

                string saudacao = GetSaudacao();

                string corpoEmail = "<html><body contenteditable=\"false\">";
                corpoEmail += "<font face = 'Calibri ' size = '3' > <p>" + saudacao + "</font></p>";

                corpoEmail += "<font face='Calibri ' size='3'><p>Venho por este meio informar, que já foi emitido dentro da pasta da obra em assunto, o PowerFab &nbsp;"
                             + "<span style='color:#00B0F0; display:inline-block; margin-right:10px;'><u>"
                             + obra + "&nbsp Lote " + lote + "&nbsp Fase " + Fase + "</font></u> </span></p>";


                corpoEmail += "<font face = 'Calibri ' size = '3' ><p><b><u> PROCESSO DE FABRICO: </u></b>";

                corpoEmail += "<font face = 'Calibri ' size = '3' style='color:#5B9BD5;'>"
                           + "<a href='file:///" + caminho.Replace("\\", "/") + "' style='color:#5B9BD5; text-decoration: none;'>" + linkTexto + "</a>" + "</font> ";


                corpoEmail += "<font face = 'Calibri ' size = '3' > <p> Melhores Cumprimentos,</p> </font> <br>";
                corpoEmail += "<font face = 'Calibri' size = '3' > <b>" + nomeUsuario + "</b> </Font> <br>";
                corpoEmail += "<font face = 'Calibri' size = '3' > Construção Metálica | Preparador </Font> <br>";
                corpoEmail += "<font face = 'Calibri' size = '3' > T + 351 253 080 609 * </font> <br>";
                corpoEmail += "<font color='red' font face = 'Calibri ' size = '3'> ofeliz.com </font> <br>";
                corpoEmail += "<p><a href='https://www.ofeliz.com'><img src='file:///" + imagemOfelizFilePath.Replace("\\", "/") + "' width='127' height='34'></a></p>";

                corpoEmail += "<i><font color='Light grey' font face = 'Calibri ' size = '1.5'> Alvará Nº 10553 – Pub. *Chamada para a rede fixa nacional. </font> </i><br>";
                corpoEmail += "<i><font color='green' font face = 'Calibri ' size = '1.5'> Antes de imprimir este e-mail tenha em consideração o meio ambiente. </font> </i><br>";
                corpoEmail += "</body></html>";


                this.Visible = false;
                Frm_Corpo_de_Texto_Email_Enviar_Powerfab previewForm = new Frm_Corpo_de_Texto_Email_Enviar_Powerfab("Enviar Email do Powerfab", corpoEmail, SubjectEnviarPowerFab, caminho, obra, lote, Fase, linkTexto);
                previewForm.ShowDialog(this);

            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Erro ao abrir a Ferramenta de Recorte ou enviar o e-mail: " + ex.Message);
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

        private void label11_Click(object sender, EventArgs e)
        {        }
             

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

        private void LimparNumeracao()
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
                    }
                }               

                MessageBox.Show(this, "A Numeração de " + pecas.Count + " Peças e " + conj.Count + " Conjuntos, foi limpa com sucesso.", "Êxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Erro: " + ex.Message);
                this.Close();
            }
        }

        private void button488_Click(object sender, EventArgs e)
        {
            _formpai.fase1000 = (int.Parse(_formpai.fase1000) + 1).ToString();
            label8.Text = "Fase - " + _formpai.fase1000;
        }

        private void Button477_Click(object sender, EventArgs e)
        {
            _formpai.fase1000 = (int.Parse(_formpai.fase1000) - 1).ToString();
            label8.Text = "Fase - " + _formpai.fase1000;
        }

        private void button44_Click(object sender, EventArgs e)
        {
            _formpai.fase = (int.Parse(_formpai.fase) - 1).ToString("000");
            label6.Text = "Fase - " + int.Parse(_formpai.fase).ToString("000");
        }

        private void button66_Click(object sender, EventArgs e)
        {
            _formpai.fase = (int.Parse(_formpai.fase) + 1).ToString("000");
            label6.Text = "Fase - " + int.Parse(_formpai.fase).ToString("000");
        }

        private void button200_Click(object sender, EventArgs e)
        {
            _formpai.fase500 = (int.Parse(_formpai.fase500) - 1).ToString();
            label7.Text = "Fase - " + _formpai.fase500;
        }

        private void button190_Click(object sender, EventArgs e)
        {
            _formpai.fase500 = (int.Parse(_formpai.fase500) + 1).ToString();
            label7.Text = "Fase - " + _formpai.fase500;
        }

        private void button77_Click(object sender, EventArgs e)
        {
            ArrayList a = new ArrayList();
            a = ComunicaTekla.ListadePecasdoConjSelec();
            ComunicaTekla.EnviaproPriedadePeca(a, "Fase", label6.Text.Replace("Fase - ", ""));
            ComunicaTekla.AlteraPrefixo(a, int.Parse(label6.Text.Replace("Fase - ", "")));

            MessageBox.Show(this, "Foram alteradas as fases de " + a.Count + " Peças", "Informação", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void guna2Button2_Click(object sender, EventArgs e)
        {
            ArrayList a = new ArrayList();
            //a = ComunicaTekla.ListadePecasdoConjSelec();
            a = ComunicaTekla.ListadePecasSelec();
            ComunicaTekla.EnviaproPriedadePeca(a, "Fase", label7.Text.Replace("Fase - ", ""));
            ComunicaTekla.AlteraPrefixo(a, int.Parse(label7.Text.Replace("Fase - ", "")));

            MessageBox.Show(this, "Foram alteradas as fases de " + a.Count + " Peças", "Informação", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void button211_Click(object sender, EventArgs e)
        {

            //listBox1.Items.Clear();
            //TSM.ModelObjectEnumerator modelEnumerator = new TSM.UI.ModelObjectSelector().GetSelectedObjects();
            //string fase = "";

            //while (modelEnumerator.MoveNext())
            //{
            //    TSM.Assembly ass = modelEnumerator.Current as TSM.Assembly;
            //    if (ass != null)
            //    {
            //        ass.GetMainPart().GetUserProperty("Fase", ref fase);

            //        if (!listBox1.Items.Contains(fase.ToString()))
            //        {
            //            listBox1.Items.Add(fase.ToString());
            //            Frm_Inico.str.Add(fase.ToString());
            //        }

            //    }
            //}
        }

        private void button222_Click(object sender, EventArgs e)
        {
            //ArrayList a = new ArrayList();
            //a = ComunicaTekla.ListadePecasdoConjSelec();
            //try
            //{
            //    ComunicaTekla.EnviaproPriedadePeca(a, "Fase", listBox1.SelectedItem.ToString());
            //    ComunicaTekla.AlteraPrefixo(a, int.Parse(listBox1.SelectedItem.ToString()));
            //    MessageBox.Show("Foram alteradas as fases de " + a.Count + " peças", "Informação", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //}
            //catch (System.Exception)
            //{
            //    MessageBox.Show("Por Favor. Selecione uma fase na Lista de Fases ", "ERRO", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

        private void button2333_Click(object sender, EventArgs e)
        {
            LimparNumeracao();
        }

        private void button100_Click(object sender, EventArgs e)
        {
            if (label11.Text == new Model().GetProjectInfo().ProjectNumber)
            {
                ArrayList pecas = new ArrayList();
                ArrayList conj = new ArrayList();

                pecas = ComunicaTekla.ListadePecasdoConjSelec();
                conj = ComunicaTekla.ListadeConjuntosSelec();

                ComunicaTekla.Artigos(pecas);
                ComunicaTekla.Destinatarioexterno(pecas, true);
                ComunicaTekla.operacoes(pecas, conj);
                ComunicaTekla.Artigo_interno(pecas);
                ComunicaTekla.DestinatarioexternoparaH60(pecas);

                VerificarQuant(pecas);
                VerificarArtigoInterno(pecas);

                if (pecas.Count > 0)
                {
                    MessageBox.Show(this, "Artigos inseridos com sucesso " + pecas.Count + " Peças e " + conj.Count + " Conjuntos", "Êxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(this, "Por favor selecione pelo menos um conjunto ", "SEM CONJUNTOS SELECIONADOS", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show(this, "O projeto em uso neste programa difere do projeto atualmente aberto no Tekla.", "Erro de Incompatibilidade de projetos", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button111_Click(object sender, EventArgs e)
        {
            exc2();
        }

        private void button45_Click(object sender, EventArgs e)
        {
            LBLestado.Text = "Conjuntos selecionados em verificação";
            ArrayList c = ComunicaTekla.ListadePecasdoConjSelec();

            LBLestado.Text = "A gerar desenhos. Por favor, aguarde";
            ComunicaTekla a = new ComunicaTekla();
            ComunicaTekla.selectinmodel(c);
            string fasex = null;
            List<string> l = new List<string>();
            foreach (TSM.Part item in c)
            {
                item.GetReportProperty("Fase", ref fasex);

                if (!l.Contains(fasex))
                {
                    l.Add(fasex);
                }
            }
            l.Distinct();
            string outfase = null;
            foreach (var item in l)
            {
                if (item.ToString().Trim() != "0")
                {
                    outfase += "<p>Fase " + item + "</p>";
                }

            }

            new TeklaMacroBuilder.MacroBuilder().Callback("acmd_partnumbers_selected", "", "main_frame").Run();
            bool b = a.CriaDesenhos(c);
            if (b == true)
            {              
                LBLestado.Text = "Os desenhos foram criados com sucesso";
                MessageBox.Show(this, "Os desenhos foram criados com sucesso", "Criação de Desenhos", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LBLestado.Text = "Os desenhos foram criados com sucesso";

            }
            else
            {
                LBLestado.Text = "Erro na criação de desenhos";
                MessageBox.Show(this, "Possivel erro, o método de seleção." + Environment.NewLine + "Altere o método para a opção ' Selecionar conjuntos '", "Criação de Desenhos", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {

            if (label11.Text == new Model().GetProjectInfo().ProjectNumber)
            {

                if (Directory.GetFiles(@"C:\R\").Length > 0)
                {
                    DialogResult DIAL = MessageBox.Show(this, @"Existe ficheiros na pasta C:\R\ deseja continuar?" + Environment.NewLine + "Se Responder sim o programa ira limpar a pasta e prosseguir", "ALERTA", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (DIAL == DialogResult.Yes)
                    {
                        foreach (string file in Directory.GetFiles(@"C:\R\"))
                        {
                            File.Delete(file);
                        }
                    }
                }
                if (Directory.GetFiles(@"C:\R\").Length == 0)
                {
                    LBLestado.Text = "A selecionar Conjuntos";
                    ArrayList peças = new ArrayList(ComunicaTekla.ListadePecasdoConjSelec());
                    ArrayList conjuntos = new ArrayList(ComunicaTekla.ListadeConjuntosSelec());
                    ArrayList objectos = new ArrayList(peças);
                    //////////////////////////////////////////////////////////soldaura//////////////////////////////////////////////////////////////////////////////////

                    //PROCURA POR DESENHOS DE SOLDADURA 
                    List<string> tudo = new List<string>();
                    foreach (TSM.Assembly ASS in conjuntos)
                    {
                        if (ASS != null)
                        {
                            string CONJ = null;
                            string soldadura = null;
                            ASS.GetReportProperty("ASSEMBLY_POS", ref CONJ);
                            ASS.GetReportProperty("Operacoes_Conj", ref soldadura);
                            if (soldadura == "Opção 2" || soldadura == "Opção 5" || soldadura == "Opção 6" || soldadura == "Opção 16")
                            {
                                tudo.Add(CONJ);
                            }
                        }
                    }
                    IEnumerable dis = tudo.Distinct();

                    // FIM DA PROCURA

                    foreach (string item in dis)
                    {
                        var result = Regex.Split(item, @"\d+$")[0] + "." + Regex.Match(item, @"\d+$").Value;


                        new TeklaMacroBuilder.MacroBuilder()
                            .ValueChange("Drawing_selection", "diaSearchInOptionMenu", "7")
                            .ValueChange("Drawing_selection", "diaDrawingListSearchCriteria", result)
                            .PushButton("diaDrawingListSearch", "Drawing_selection")
                            .TableSelect("Drawing_selection", "dia_draw_select_list", new int[] { 1 })
                            .PopupCallback("acmd_copy_drawing_to_new_sheet", "", "Drawing_selection", "dia_draw_select_list")
                            .ValueChange("Drawing_selection", "diaDrawingListSearchCriteria", result + " - 1")
                            .PushButton("diaDrawingListSearch", "Drawing_selection")
                            .TableSelect("Drawing_selection", "dia_draw_select_list", new int[] { 1 })
                            .Activate("Drawing_selection", "dia_draw_select_list")
                            .Run();
                        DrawingHandler drawingHandler = new DrawingHandler();
                        try
                        {

                            ContainerView sheet = drawingHandler.GetActiveDrawing().GetSheet();
                            if (drawingHandler.GetConnectionStatus())
                            {

                                System.Type[] Types = new System.Type[1];
                                Types.SetValue(typeof(StraightDimension), 0);

                                DrawingObjectEnumerator allDimLines = sheet.GetAllObjects(Types);

                                foreach (StraightDimension line in allDimLines)
                                {

                                    line.Delete();

                                }

                                Types.SetValue(typeof(AngleDimension), 0);

                                allDimLines = sheet.GetAllObjects(Types);

                                foreach (AngleDimension line in allDimLines)
                                {
                                    line.Delete();
                                }
                                Types.SetValue(typeof(RadiusDimension), 0);

                                allDimLines = sheet.GetAllObjects(Types);
                                foreach (RadiusDimension line in allDimLines)
                                {
                                    line.Delete();
                                }
                            }

                            new TeklaMacroBuilder.MacroBuilder()
                                .ValueChange("gr_close_dr_editor_confirm_instance", "gr_close_save_dr_editor_freeze", "1")
                                .PushButton("gr_close_save_dr_editor_yes", "gr_close_dr_editor_confirm_instance").Run();


                            DrawingHandler dh = new DrawingHandler();
                            ViewBase _sheet = dh.GetActiveDrawing().GetSheet();
                            Text text = new Text(_sheet, new Tekla.Structures.Geometry3d.Point(285, 70), "NOTA: Soldar segundo os nossos procedimentos habituais");
                            text.Attributes.LoadAttributes("soldaconfig");
                            text.Insert();
                            dh.SaveActiveDrawing();
                            dh.CloseActiveDrawing();
                        }
                        catch (System.Exception)
                        {
                            MessageBox.Show(this, "Atenção: É necessário numerar o modelo.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                    /////////////////////////////////////////////////////////////////fim da soldadura//////////////////////////////////////////////////////////////////////////


                    LBLestado.Text = "A imprimir ...";
                    ComunicaTekla.imprimepdf(conjuntos, peças, LBLestado);

                    LBLestado.Text = "A selecionar os Objectos";

                    //for (int i = peças.Count - 1; i >= 0; i--)
                    //{
                    //    TSM.Part item = (TSM.Part)peças[i];
                    //    string perfil = item.Profile.ProfileString;

                    //    // Verifica se o perfil contém "VRSM", "BM", "NM" ou "WM"
                    //    if (perfil.Contains("VRSM") || perfil.Contains("BM") || perfil.Contains("NM") || perfil.Contains("WM"))
                    //    {
                    //        peças.RemoveAt(i);
                    //    }
                    //    else
                    //    {
                    //        foreach (BoltGroup parafuso in item.GetBolts())
                    //        {
                    //            if (parafuso != null)
                    //            {
                    //                if (parafuso.PartToBoltTo.Identifier.ID == item.Identifier.ID)
                    //                {
                    //                    objectos.Add(parafuso);
                    //                }
                    //            }
                    //        }
                    //    }
                    //}
                    foreach (TSM.Part item in peças)
                    {
                        foreach (BoltGroup parafuso in item.GetBolts())
                        {
                            if (parafuso != null)
                            {
                                if (parafuso.PartToBoltTo.Identifier.ID == item.Identifier.ID)
                                {
                                    objectos.Add(parafuso);
                                }
                            }
                        }
                    }
                    ComunicaTekla.selectinmodel(objectos);
                    if (!Directory.Exists(Frm_Inico.CaminhoModelo + @"\listas"))
                    {
                        Directory.CreateDirectory(Frm_Inico.CaminhoModelo + @"\listas");
                    }

                    LBLestado.Text = "A criar lista ";
                    TSM.Operations.Operation.CreateReportFromSelected("OFELIZ", @"C:\R\OFELIZ.CSV", "", "", "");
                    TSM.Operations.Operation.CreateReportFromSelected("OFELIZ.csv", @"C:\R\PEÇAS_E_CONJUNTOS.CSV", "", "", "");
                    LBLestado.Text = "A criar CNC ";
                    TSM.Operations.Operation.CreateNCFilesFromSelected("OFELIZ_chapas", @"c:\r\");
                    TSM.Operations.Operation.CreateNCFilesFromSelected("OFELIZ_perfis", @"c:\r\");
                    TSM.Operations.Operation.CreateNCFilesFromSelected("OFELIZ_madres", @"c:\r\");
                    LBLestado.Text = "A converter DXF ";

                    string[] NCfiles = Directory.GetFiles(@"c:\r", "*.nc1", SearchOption.TopDirectoryOnly);
                    List<string> myfiles = new List<string>();
                    foreach (var item in NCfiles)
                    {
                        myfiles.Add(item);
                    }
                    dstv_dxf.CRIAR(myfiles);
                    LBLestado.Text = "DXF's convertidos com sucesso ";
                    timer1.Enabled = true;                                      
                }
                else
                {
                    MessageBox.Show(this, "O projeto em uso neste programa difere do projeto atualmente aberto no Tekla.", "Erro de Incompatibilidade de projetos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }        

        private void button23_Click(object sender, EventArgs e)
        {
            Model m = new Model();
            bool VERIFICACAO = false;
            if (Environment.UserName.ToLower() != "rui.ferreira")
            {
                DialogResult a = MessageBox.Show(this, "O método de seleção do tekla esta em modo de peça?", "ATENÇÃO", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (a == DialogResult.Yes)
                {
                    VERIFICACAO = true;
                }
                else
                {
                    MessageBox.Show(this, "Por favor, Altere o modo de seleção");
                }
            }
            else
            {
                VERIFICACAO = true;
            }

            if (VERIFICACAO)
            {
                if (label11.Text == m.GetProjectInfo().ProjectNumber)
                {

                    LBLestado.Text = "Extraindo lista de parafusos";
                    TSM.Operations.Operation.CreateReportFromSelected("OFELIZPARAFUSOOBRA", @"C:\R\OFELIZ.CSV", "", "", "");

                    this.Visible = false;
                    Frm_Parafusos p = new Frm_Parafusos(_formpai);
                    p.ShowDialog();
                    this.Visible = true;

                    DialogResult resultado2 = MessageBox.Show(this, "Quer importar no Primavera?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (resultado2 == DialogResult.Yes)
                    {
                        MessageBox.Show(this, "Não esquecer de importar no Primavera!", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        AppAbrirPrimavera primaveraHandler = new AppAbrirPrimavera();
                        primaveraHandler.AbrirPrimaveira();
                    }
                    else
                    { }
                }
                else
                {
                    MessageBox.Show(this, "O projeto em uso neste programa difere do projeto atualmente aberto no Tekla.", "Erro de Incompatibilidade de projetos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void button13_Click_1(object sender, EventArgs e)
        {
            if (label11.Text == new Model().GetProjectInfo().ProjectNumber)
            {

                ArrayList peças = new ArrayList(ComunicaTekla.ListadePecasdoConjSelec());
                ArrayList objectos = new ArrayList(peças);
                LBLestado.Text = "A selecionar objectos";
                foreach (TSM.Part item in peças)
                {
                    foreach (BoltGroup parafuso in item.GetBolts())
                    {
                        if (parafuso != null)
                        {
                            if (parafuso.PartToBoltTo.Identifier.ID == item.Identifier.ID)
                            {
                                objectos.Add(parafuso);
                            }
                        }


                    }
                }
                ComunicaTekla.selectinmodel(objectos);
                if (!Directory.Exists(Frm_Inico.CaminhoModelo + @"\listas"))
                {
                    Directory.CreateDirectory(Frm_Inico.CaminhoModelo + @"\listas");
                }
                TSM.Operations.Operation.CreateReportFromSelected("OFELIZ", @"C:\R\OFELIZ.CSV", "", "", "");
                Frm_ListaOFeliz f = new Frm_ListaOFeliz(_formpai);
                LBLestado.Text = "Lista criada com sucesso";
                f.ShowDialog(this);

                this.Visible = true;
                DialogResult resultado2 = MessageBox.Show(this, "Pretende importar os dados para o Primavera?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (resultado2 == DialogResult.Yes)
                {
                    TopMost = false;
                    AppAbrirPrimavera primaveraHandler = new AppAbrirPrimavera();
                    primaveraHandler.AbrirPrimaveira();
                }

            }
            else
            {
                MessageBox.Show(this, "O projeto em uso neste programa difere do projeto atualmente aberto no Tekla.", "Erro de Incompatibilidade de projetos", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }



        public void button1_Click_1(object sender, EventArgs e)
        {
            TopMost = false;
            AppAbrirPrimavera primaveraHandler = new AppAbrirPrimavera();
            primaveraHandler.AbrirPrimaveira();
        }

        

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            string numeroObra = label11.Text;
            CreateXmlFile(numeroObra);
            AppAbrirTekla teklaHandler = new AppAbrirTekla();
            teklaHandler.TrazerTeklaParaFrente();

            DialogResult resultado = MessageBox.Show(this, "O Powerfab foi gerado?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

           if (resultado == DialogResult.Yes)
           {
             EnviarPowerFab_OpenEmailPreviewAndCreateEmail();            
           }
        }

        private void Frm_CriarFase_FormClosed(object sender, FormClosedEventArgs e)
        {
            listBox1.Items.Clear();
        }

        private void guna2Button2_Click_1(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
        }
    }
}


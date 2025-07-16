using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Printing;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tekla.Structures.Drawing;
using Tekla.Structures.Filtering;
using Tekla.Structures.Filtering.Categories;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Model;
using TSM = Tekla.Structures.Model;

namespace TeklaArtigosOfeliz
{
    public partial class Frm_Lista_Pecas_Conjuntos : Form
    {
        DataTable _parafusos = null;
        DataTable dtpecaPerfis = null;
        DataTable dtpecaChapas = null;
        DataTable dtconj = null;
      
        public Frm_Lista_Pecas_Conjuntos()
        {
            InitializeComponent();
            
        }

        private void Lista_Pecas_Conjuntos_Load(object sender, EventArgs e)
        {



        }

        private void button1_Click(object sender, EventArgs e)
        {

            ArrayList Conjuntos = ComunicaTekla.ListadeConjuntosSelec();
            CarregaConjuntos(Conjuntos);

        }

        void LimparApp()
        {
            dtconj = null;
            _parafusos = null;
            dtpecaChapas = null;
            dtpecaPerfis = null;
            dataGridView1.DataSource = null;
            dataGridView2.DataSource = null;
            dataGridView3.DataSource = null;
            dataGridView4.DataSource = null;
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();
            dataGridView3.Rows.Clear();
            dataGridView4.Rows.Clear();
            dtconj = CriadtConj();
            _parafusos = CriadtParafusos();
            dtpecaPerfis = CriadtPecasPerfis();
            dtpecaChapas = CriadtPecasChapas();
        }

        DataTable CriadtParafusos()
        {
            DataTable tabela = new DataTable();
            tabela.Columns.Add("Artigo");
            tabela.Columns.Add("Comprimento");
            tabela.Columns.Add("Quantidade", typeof(int));
            tabela.Columns.Add("Classe");
            tabela.Columns.Add("Norma");
            tabela.Columns.Add("Lote");
            tabela.Columns.Add("Entrega");
            tabela.Columns.Add("Pecas", typeof(ArrayList));
            return tabela;
        }

        DataTable CriadtConj()
        {
            DataTable dtconj = new DataTable();
            dtconj.Columns.Add("CDU_ID").ColumnMapping = MappingType.Hidden;
            dtconj.Columns.Add("CDU_IDTekla");
            dtconj.Columns.Add("CDU_Conjunto");
            dtconj.Columns.Add("CDU_CodigoCliente").ColumnMapping = MappingType.Hidden;
            dtconj.Columns.Add("CDU_NomeCliente").ColumnMapping = MappingType.Hidden;
            dtconj.Columns.Add("CDU_CodigoObra").ColumnMapping = MappingType.Hidden;
            dtconj.Columns.Add("CDU_DescricaoObra").ColumnMapping = MappingType.Hidden;
            dtconj.Columns.Add("CDU_Fase");
            dtconj.Columns.Add("CDU_Lote");
            dtconj.Columns.Add("CDU_Perfil");
            dtconj.Columns.Add("CDU_Artigo");
            dtconj.Columns.Add("CDU_DataInicioProducao").ColumnMapping = MappingType.Hidden;
            dtconj.Columns.Add("CDU_DataEntregaPrevista");
            dtconj.Columns.Add("CDU_DataCriacao").ColumnMapping = MappingType.Hidden;
            dtconj.Columns.Add("CDU_comprimento");
            dtconj.Columns.Add("CDU_Altura");
            dtconj.Columns.Add("CDU_Largura");
            dtconj.Columns.Add("CDU_GrauPreparacao");
            dtconj.Columns.Add("CDU_ClasseExecucao");
            dtconj.Columns.Add("CDU_ReferenciaCliente");
            dtconj.Columns.Add("CDU_Quantidade");
            dtconj.Columns.Add("CDU_Comentarios");
            dtconj.Columns.Add("CDU_PreparacaoChapas", typeof(bool));
            dtconj.Columns.Add("CDU_Armacao", typeof(bool));
            dtconj.Columns.Add("CDU_Soldadura", typeof(bool));
            dtconj.Columns.Add("CDU_Decapagem", typeof(bool));
            dtconj.Columns.Add("CDU_Pintura", typeof(bool));
            dtconj.Columns.Add("CDU_Contagem", typeof(bool));
            dtconj.Columns.Add("CDU_Destinatario");
            dtconj.Columns.Add("Pecas", typeof(List<Assembly>)).ColumnMapping = MappingType.Hidden;
            return dtconj;
        }

        DataTable CriadtPecasChapas()
        {
            DataTable dtpeca = new DataTable();
            dtpeca.Columns.Add("CDU_ID").ColumnMapping = MappingType.Hidden;
            dtpeca.Columns.Add("CDU_IDCabec").ColumnMapping = MappingType.Hidden;
            dtpeca.Columns.Add("CDU_IDTekla");
            dtpeca.Columns.Add("CDU_Peca");
            dtpeca.Columns.Add("CDU_Conjunto");
            dtpeca.Columns.Add("CDU_Classe");
            dtpeca.Columns.Add("CDU_Artigo");
            dtpeca.Columns.Add("CDU_Certificado");
            dtpeca.Columns.Add("CDU_comprimento");
            dtpeca.Columns.Add("CDU_Espessura");
            dtpeca.Columns.Add("CDU_Largura");
            dtpeca.Columns.Add("CDU_Peso");
            dtpeca.Columns.Add("CDU_Area");
            dtpeca.Columns.Add("CDU_ClasseExecucao");
            dtpeca.Columns.Add("CDU_ReferenciaCliente");
            dtpeca.Columns.Add("CDU_Quantidade");
            dtpeca.Columns.Add("CDU_Perfil");
            dtpeca.Columns.Add("CDU_RequesitosEspeciais");
            dtpeca.Columns.Add("CDU_GrauPreparacao");
            dtpeca.Columns.Add("CDU_Tolerancias");
            dtpeca.Columns.Add("CDU_estadoSuperficie");
            dtpeca.Columns.Add("CDU_propriedadesEspeciais");
            dtpeca.Columns.Add("CDU_Cor");
            dtpeca.Columns.Add("CDU_Marca");
            dtpeca.Columns.Add("CDU_Norma");
            dtpeca.Columns.Add("CDU_EsquemaPintura");
            dtpeca.Columns.Add("CDU_DataCriacao");
            dtpeca.Columns.Add("CDU_Comentarios");
            dtpeca.Columns.Add("CDU_Destinatario");
            dtpeca.Columns.Add("CDU_Corte", typeof(bool));
            dtpeca.Columns.Add("CDU_CorteFuracao", typeof(bool));
            dtpeca.Columns.Add("CDU_AnguloA");
            dtpeca.Columns.Add("CDU_AnguloB");
            dtpeca.Columns.Add("CDU_ArtigoOF");
            dtpeca.Columns.Add("Pecas", typeof(List<TSM.Part>));
            return dtpeca;
        }

        DataTable CriadtPecasPerfis()
        {
            DataTable dtpeca = new DataTable();
            dtpeca.Columns.Add("CDU_ID").ColumnMapping = MappingType.Hidden;
            dtpeca.Columns.Add("CDU_IDCabec").ColumnMapping = MappingType.Hidden;
            dtpeca.Columns.Add("CDU_IDTekla");
            dtpeca.Columns.Add("CDU_Peca");
            dtpeca.Columns.Add("CDU_Conjunto");
            dtpeca.Columns.Add("CDU_Classe");
            dtpeca.Columns.Add("CDU_Artigo");
            dtpeca.Columns.Add("CDU_Certificado");
            dtpeca.Columns.Add("CDU_comprimento");
            dtpeca.Columns.Add("CDU_Espessura");
            dtpeca.Columns.Add("CDU_Largura");
            dtpeca.Columns.Add("CDU_Peso");
            dtpeca.Columns.Add("CDU_Area");
            dtpeca.Columns.Add("CDU_ClasseExecucao");
            dtpeca.Columns.Add("CDU_ReferenciaCliente");
            dtpeca.Columns.Add("CDU_Quantidade");
            dtpeca.Columns.Add("CDU_Perfil");
            dtpeca.Columns.Add("CDU_RequesitosEspeciais");
            dtpeca.Columns.Add("CDU_GrauPreparacao");
            dtpeca.Columns.Add("CDU_Tolerancias");
            dtpeca.Columns.Add("CDU_estadoSuperficie");
            dtpeca.Columns.Add("CDU_propriedadesEspeciais");
            dtpeca.Columns.Add("CDU_Cor");
            dtpeca.Columns.Add("CDU_Marca");
            dtpeca.Columns.Add("CDU_Norma");
            dtpeca.Columns.Add("CDU_EsquemaPintura");
            dtpeca.Columns.Add("CDU_DataCriacao");
            dtpeca.Columns.Add("CDU_Comentarios");
            dtpeca.Columns.Add("CDU_Destinatario");
            dtpeca.Columns.Add("CDU_Corte", typeof(bool));
            dtpeca.Columns.Add("CDU_CorteFuracao", typeof(bool));
            dtpeca.Columns.Add("CDU_AnguloA");
            dtpeca.Columns.Add("CDU_AnguloB");
            dtpeca.Columns.Add("CDU_ArtigoOF");
            dtpeca.Columns.Add("Pecas", typeof(List<TSM.Part>));
            return dtpeca;
        }

        internal void CarregaConjuntos(ArrayList Conjuntos)
        {
            LimparApp();

            Model m = new Model();
            ArrayList pecas = new ArrayList();
            string cdu_codigoobra = m.GetProjectInfo().ProjectNumber;
            ComunicaBDprimavera dg = new ComunicaBDprimavera();
            dg.ConectarBD();
            List<string> dados = dg.Procurarbd("SELECT [ERPEntidadeA],[Nome_Cliente],[Descricao] FROM [PRIOFELIZ].[dbo].[MT_View_Obras_Clientes_Descricao] where Codigo = '" + cdu_codigoobra + "'");
            dg.DesonectarBD();

            string cdu_codigocliente = dados[0];
            string cdu_nomecliente = dados[1]; ;
            string cdu_descricaoobra = dados[2]; ;

            if (cdu_nomecliente != "")
            {
                foreach (Assembly Conjunto in Conjuntos)
                {
                    string cdu_perfil = null;
                    Conjunto.GetReportProperty("MAINPART.PROFILE", ref cdu_perfil);
                    if ((cdu_perfil.Contains("NM") || cdu_perfil.Contains("WM") || cdu_perfil.Contains("VRSM"))&& !cdu_perfil.Contains("FWM"))
                    {
                        parafuso(Conjunto);
                    }
                    else
                    {
                        string cdu_id = Guid.NewGuid().ToString();

                        string cdu_fase = null;
                        Conjunto.GetReportProperty("Fase", ref cdu_fase);

                        string cdu_lote = null;
                        Conjunto.GetReportProperty("lote_number", ref cdu_lote);

                        string cdu_idtekla = "";
                        Conjunto.GetReportProperty("MAINPART.USERDEFINED.QUANTIFICACAO", ref cdu_idtekla);
                        cdu_idtekla = cdu_idtekla.Split('-')[0];

                        string cdu_exc3 = "";
                        string cdu_conjunto = null;

                        Conjunto.GetReportProperty("ASSEMBLY_POS", ref cdu_conjunto);
                        cdu_conjunto = "2." + cdu_codigoobra + "." + cdu_conjunto;

                        Conjunto.GetReportProperty("USERDEFINED.USER_FIELD_2", ref cdu_exc3);
                        if (cdu_exc3 != "")
                        {
                            cdu_conjunto = cdu_conjunto + "." + cdu_exc3;
                        }

                        string cdu_artigo = null;
                        Conjunto.GetReportProperty("USERDEFINED.Artigo_interno", ref cdu_artigo);

                        string cdu_datainicioproducao = null;

                        string cdu_dataentregaprevista = null;
                        Conjunto.GetReportProperty("USERDEFINED.lote_data", ref cdu_dataentregaprevista);

                        string cdu_datacriacao = DateTime.Now.ToShortDateString();

                        double cdu_comprimento = 0;
                        Conjunto.GetReportProperty("LENGTH", ref cdu_comprimento);

                        double cdu_altura = 0;
                        Conjunto.GetReportProperty("WIDTH", ref cdu_altura);

                        double cdu_largura = 0;
                        Conjunto.GetReportProperty("HEIGHT", ref cdu_largura);

                        string cdu_graupreparacao = null;
                        Conjunto.GetReportProperty("MAINPART.USERDEFINED.Grau_DE_pre", ref cdu_graupreparacao);

                        string cdu_classeexecucao = null;
                        Conjunto.GetReportProperty("PROJECT.USERDEFINED.PROJECT_USERFIELD_2", ref cdu_classeexecucao);

                        string cdu_referenciacliente = null;
                        Conjunto.GetReportProperty("PROJECT.USERDEFINED.USER_FIELD_1", ref cdu_referenciacliente);

                        int cdu_quantidade = 0;
                        Conjunto.GetReportProperty("NUMBER", ref cdu_quantidade);

                        string cdu_comentarios = null;
                        Conjunto.GetReportProperty("PROJECT.USERDEFINED.comment", ref cdu_comentarios);

                        bool cdu_preparacaochapas = false;

                        bool cdu_armacao = false;

                        bool cdu_soldadura = false;

                        string str = null;
                        Conjunto.GetReportProperty("OperacaoFabrica", ref str);

                        bool cdu_decapagem = false;
                        if (str != null)
                        {
                            if (str.ToUpper().Contains("DECAPADO"))
                            {
                                cdu_decapagem = true;
                            }
                        }

                        bool cdu_pintura = false;
                        if (str != null)
                        {


                            if (str.ToUpper().Contains("PINTADO"))
                            {
                                cdu_pintura = true;
                            }
                        }
                        string CDU_Destinatario = null;
                        Conjunto.GetReportProperty("USERDEFINED.Destinata_ext", ref CDU_Destinatario);

                        double RT = 0; Conjunto.GetReportProperty("SUPPLEMENT_PART_WEIGHT", ref RT);
                        if (RT>0)
                        {
                            cdu_armacao = true;
                            cdu_soldadura = true;
                        }

                        bool cdu_contagem = true;
                        if (cdu_fase!=null)
                        {
                            Pecas(Conjunto, cdu_conjunto, cdu_codigoobra,cdu_id);

                            if (int.Parse(cdu_fase) < 500)
                            {
                                cdu_artigo = "";
                                

                            }
                            else
                            {
                                cdu_preparacaochapas = true;
                                //string cdu = null;
                                //Conjunto.GetReportProperty("MAINPART.PART_POS", ref cdu);
                                //cdu_conjunto = cdu_conjunto + "." + cdu;
                            }

                        }else
                        {
                            MessageBox.Show(this, "Nem todas as peças tem fase" +Environment.NewLine+"P.F. Numerar novamente o modelo e atribuir nova Fase.", "Erro", MessageBoxButtons.OK);
                            dataGridView1.Rows.Clear();
                            dataGridView2.Rows.Clear();
                            dataGridView3.Rows.Clear();
                            dataGridView4.Rows.Clear();
                            break;
                        }

                        DataRow dt = dtconj.AsEnumerable().SingleOrDefault(r => r.Field<string>("cdu_idtekla") == cdu_idtekla && r.Field<string>("cdu_conjunto") == cdu_conjunto && r.Field<string>("cdu_codigocliente") == cdu_codigocliente && r.Field<string>("cdu_nomecliente") == cdu_nomecliente && r.Field<string>("cdu_codigoobra") == cdu_codigoobra && r.Field<string>("cdu_descricaoobra") == cdu_descricaoobra && r.Field<string>("cdu_fase") == cdu_fase && r.Field<string>("cdu_lote") == cdu_lote && r.Field<string>("cdu_perfil") == cdu_perfil && r.Field<string>("cdu_artigo") == cdu_artigo && r.Field<string>("cdu_datainicioproducao") == cdu_datainicioproducao && r.Field<string>("cdu_dataentregaprevista") == cdu_dataentregaprevista && r.Field<string>("cdu_datacriacao") == cdu_datacriacao && r.Field<string>("cdu_graupreparacao") == cdu_graupreparacao && r.Field<string>("cdu_classeexecucao") == cdu_classeexecucao && r.Field<string>("cdu_referenciacliente") == cdu_referenciacliente && r.Field<string>("cdu_comentarios") == cdu_comentarios);

                        if (dt != null)
                        {
                            int d = int.Parse(dt["CDU_Quantidade"].ToString()) + 1;
                            dt.SetField("CDU_Quantidade", d);

                            List<Assembly> uri = dt.Field<List<Assembly>>("Pecas");
                            uri.Add(Conjunto);
                            dt.SetField<List<Assembly>>("Pecas", uri);

                        }
                        else
                        {
                            List<Assembly> Pecas = new List<Assembly>();
                            Pecas.Add(Conjunto);
                            dtconj.Rows.Add(cdu_id, cdu_idtekla, cdu_conjunto, cdu_codigocliente, cdu_nomecliente, cdu_codigoobra, cdu_descricaoobra, cdu_fase, cdu_lote, cdu_perfil, cdu_artigo, cdu_datainicioproducao, cdu_dataentregaprevista, cdu_datacriacao, cdu_comprimento.ToString("0"), cdu_altura.ToString("0"), cdu_largura.ToString("0"), cdu_graupreparacao, cdu_classeexecucao, cdu_referenciacliente, cdu_quantidade, cdu_comentarios, cdu_preparacaochapas, cdu_armacao, cdu_soldadura, cdu_decapagem, cdu_pintura, cdu_contagem, CDU_Destinatario, Pecas);
                        }
                    }
                }
                dataGridView1.DataSource = dtconj;
                dataGridView2.DataSource = dtpecaPerfis;
                dataGridView3.DataSource = dtpecaChapas;
                PreencheParafusos(Conjuntos);
                dataGridView4.DataSource = _parafusos;
                verificaperfil();
            }
        }

        /// <summary>
        /// Preenche a lista de parafusos quando o modelado como Parafuso
        /// </summary>
        /// <param name="Conjuntos"></param>
        void PreencheParafusos(ArrayList Conjuntos)
        {
            ArrayList BOLTS = new ArrayList();
            foreach (Assembly Conjunto in Conjuntos)
            {
                ArrayList pecas = Conjunto.GetSecondaries();
                pecas.Add(Conjunto.GetMainPart());
                foreach (TSM.Part peca in pecas)
                {
                  ModelObjectEnumerator dt = peca.GetBolts();
                    foreach (var item in dt)
                    {
                        BOLTS.Add(item);
                    }
                }
            }

            foreach (var Parafuso in BOLTS.ToArray().Distinct())
            { 

                string Artigo = null;
                string Comprimento = null;
                int Quantidade = 0;
                string Classe = null;
                string Norma = null;
                string lote = null;
                string Entrega = null;
              

                if (Parafuso is TSM.BoltArray)
                {
                    TSM.BoltArray b = Parafuso as TSM.BoltArray;
                    if (b.Bolt)
                    {
                        double comp = 0;
                        b.GetReportProperty("LENGTH", ref comp);
                        b.GetReportProperty("GRADE", ref Classe);
                        Norma = b.BoltStandard;
                        do {if (comp % 5 != 0){comp = comp + 1;}} while (comp % 5 != 0);
                        Comprimento = comp.ToString("0");

                        if (Classe.ToLower().Contains("bum"))
                        {
                            Artigo = "BUMM" + b.BoltSize;
                        }
                        else if (Classe.ToLower().Contains("buq"))
                        {
                            Artigo = "BUQM" + b.BoltSize;
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(Norma))
                            {


                                if (Norma.Contains("SD1"))
                                {
                                    Artigo = "SD1." + b.BoltSize + "X" + comp;
                                }
                                else Artigo = "BM" + b.BoltSize + "X" + comp;

                            }
                            else
                                Artigo = "BM" + b.BoltSize + "X" + comp;

                        }


                        b.GetReportProperty("NUMBER", ref Quantidade);
                        Norma = b.BoltStandard;
                      
                        b.GetReportProperty("ASSEMBLY.MAINPART.USERDEFINED.lote_number", ref lote);
                        b.GetReportProperty("ASSEMBLY.MAINPART.USERDEFINED.lote_data", ref Entrega);
                    }

                }
                else if (Parafuso is TSM.BoltCircle)
                {

                    TSM.BoltCircle b = Parafuso as TSM.BoltCircle;
                    if (b.Bolt)
                    {
                        double comp = 0;
                        b.GetReportProperty("LENGTH", ref comp);
                        b.GetReportProperty("GRADE", ref Classe);
                        do { if (comp % 5 != 0) { comp = comp + 1; } } while (comp % 5 != 0);
                        Comprimento = comp.ToString("0");

                        if (Classe.ToLower().Contains("bum"))
                        {
                            Artigo = "BUMM" + b.BoltSize;
                        }
                        else if (Classe.ToLower().Contains("buq"))
                        {
                            Artigo = "BUQM" + b.BoltSize;
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(Norma))
                            {


                                if (Norma.Contains("SD1"))
                                {
                                    Artigo = "SD1." + b.BoltSize + "X" + comp;
                                }
                                else Artigo = "BM" + b.BoltSize + "X" + comp;

                            }
                            else
                                Artigo = "BM" + b.BoltSize + "X" + comp;

                        }

                        b.GetReportProperty("NUMBER", ref Quantidade);
                        Norma = b.BoltStandard;

                        b.GetReportProperty("ASSEMBLY.MAINPART.USERDEFINED.lote_number", ref lote);
                        b.GetReportProperty("ASSEMBLY.MAINPART.USERDEFINED.lote_data", ref Entrega);
                    }

                }
                else if (Parafuso is TSM.BoltGroup)
                {

                    TSM.BoltGroup b = Parafuso as TSM.BoltGroup;
                    if (b.Bolt)
                    {
                        double comp = 0;
                        b.GetReportProperty("LENGTH", ref comp);
                        b.GetReportProperty("GRADE", ref Classe);
                        do { if (comp % 5 != 0) { comp = comp + 1; } } while (comp % 5 != 0);
                        Comprimento = comp.ToString("0");

                        if (Classe.ToLower().Contains("bum"))
                        {
                            Artigo = "BUMM" + b.BoltSize;
                        }
                        else if (Classe.ToLower().Contains("buq"))
                        {
                            Artigo = "BUQM" + b.BoltSize;
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(Norma))
                            {


                                if (Norma.Contains("SD1"))
                                {
                                    Artigo = "SD1." + b.BoltSize + "X" + comp;
                                }
                                else Artigo = "BM" + b.BoltSize + "X" + comp;

                            }
                            else
                                Artigo = "BM" + b.BoltSize + "X" + comp;

                        }

                        b.GetReportProperty("NUMBER", ref Quantidade);
                        Norma = b.BoltStandard;

                        b.GetReportProperty("ASSEMBLY.MAINPART.USERDEFINED.lote_number", ref lote);
                        b.GetReportProperty("ASSEMBLY.MAINPART.USERDEFINED.lote_data", ref Entrega);
                    }


                }
                else if (Parafuso is TSM.BoltXYList)
                {

                    TSM.BoltXYList b = Parafuso as TSM.BoltXYList;
                    if (b.Bolt)
                    {
                        double comp = 0;
                        b.GetReportProperty("LENGTH", ref comp);
                        b.GetReportProperty("GRADE", ref Classe);
                        do { if (comp % 5 != 0) { comp = comp + 1; } } while (comp % 5 != 0);
                        Comprimento = comp.ToString("0");


                        if (Classe.ToLower().Contains("bum"))
                        {
                            Artigo = "BUMM" + b.BoltSize;
                        }
                        else if (Classe.ToLower().Contains("buq"))
                        {
                            Artigo = "BUQM" + b.BoltSize;
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(Norma))
                            {


                                if (Norma.Contains("SD1"))
                                {
                                    Artigo = "SD1." + b.BoltSize + "X" + comp;
                                }
                                else Artigo = "BM" + b.BoltSize + "X" + comp;

                            }
                            else
                                Artigo = "BM" + b.BoltSize + "X" + comp;

                        }

                        b.GetReportProperty("NUMBER", ref Quantidade);
                    

                        b.GetReportProperty("ASSEMBLY.MAINPART.USERDEFINED.lote_number", ref lote);
                        b.GetReportProperty("ASSEMBLY.MAINPART.USERDEFINED.lote_data", ref Entrega);
                    }


                }
                if (lote==null)
                {
                    lote = "";
                }


                DataRow dt = _parafusos.AsEnumerable().SingleOrDefault(r => r.ItemArray[0].ToString() == Artigo && r.ItemArray[1] + "" == Comprimento && r.ItemArray[3] + "" == Classe && r.ItemArray[5].ToString() == lote);

                if (dt != null)
                {
                    int d = int.Parse(dt["Quantidade"].ToString()) +Quantidade;
                    dt.SetField("Quantidade", d);

                    ArrayList uri = dt.Field<ArrayList>(7);
                    uri.Add(Parafuso);
                    dt.SetField<ArrayList>(7, uri);
                }
                else
                {
                    if (Artigo!=""&&Artigo!=null)
                    {
                        ArrayList Pecas = new ArrayList();
                        Pecas.Add(Parafuso);
                        _parafusos.Rows.Add(Artigo, Comprimento, Quantidade, Classe, Norma, lote, Entrega, Pecas);
                    }
                }
            }
          

        }

        /// <summary>
        /// Preenche a listas de peças
        /// </summary>
        /// <param name="Conjunto"></param>
        void Pecas(Assembly Conjunto, string ReferenciaConjunto, string NumeroDeObra, string CDU_IDCabec)
        {
            ArrayList ListaPecas = new ArrayList();
            ListaPecas = Conjunto.GetSecondaries();
            ListaPecas.Add(Conjunto.GetMainPart());

            foreach (TSM.Part Peca in ListaPecas)
            {

                string CDU_ID = Guid.NewGuid().ToString();

                string CDU_IDTekla = "";
                Peca.GetReportProperty("USERDEFINED.QUANTIFICACAO", ref CDU_IDTekla);
                CDU_IDTekla = CDU_IDTekla.Split('-')[0];
                string CDU_Peca = null;
                Peca.GetReportProperty("PART_POS", ref CDU_Peca);
                CDU_Peca = "2." + NumeroDeObra + "." + CDU_Peca;

                string CDU_Conjunto = ReferenciaConjunto;

                string CDU_Classe = null;
                Peca.GetReportProperty("MATERIAL", ref CDU_Classe);
                string[] str = CDU_Classe.Split('.');
                try { CDU_Classe = str[0].ToString(); } catch (Exception) { }

                string CDU_Certificado = null;
                try { CDU_Certificado = str[1].ToString(); } catch (Exception) { }


                string CDU_Destinatario = null;
                Peca.GetReportProperty("USERDEFINED.Destinata_ext", ref CDU_Destinatario);

                double CDU_comprimento = 0;
                Peca.GetReportProperty("LENGTH", ref CDU_comprimento);

                double CDU_Espessura = 0;
                Peca.GetReportProperty("WIDTH", ref CDU_Espessura);

                double CDU_Largura = 0;
                Peca.GetReportProperty("HEIGHT", ref CDU_Largura);

                double CDU_Peso = 0;
                Peca.GetReportProperty("WEIGHT_NET", ref CDU_Peso);

                double CDU_Area = 0;
                Peca.GetReportProperty("AREA", ref CDU_Area);

                string CDU_ClasseExecucao = null;
                Peca.GetReportProperty("PROJECT.USERDEFINED.PROJECT_USERFIELD_2", ref CDU_ClasseExecucao);

                string CDU_ReferenciaCliente = null;

                int CDU_Quantidade = 1;

                string CDU_Perfil = null;
                Peca.GetReportProperty("PROFILE", ref CDU_Perfil);

                string CDU_RequesitosEspeciais = null;
                Peca.GetReportProperty("USERDEFINED.Requisitos", ref CDU_RequesitosEspeciais);

                string CDU_GrauPreparacao = null;
                Peca.GetReportProperty("USERDEFINED.Grau_DE_pre", ref CDU_GrauPreparacao);

                string CDU_Tolerancias = null;
                string CDU_estadoSuperficie = null;
                string CDU_propriedadesEspeciais = null;
                string CDU_Artigo = null;

                string corte = null;

                bool B = false;
                if (CDU_Perfil.StartsWith("CHA") || CDU_Perfil.StartsWith("PL"))
                {
                    B = true;
                    Peca.GetReportProperty("USERDEFINED.Artigo_interno", ref CDU_Artigo);
                }
                else
                {
                    Peca.GetReportProperty("USERDEFINED.Operacoes", ref corte);
                }
               

                bool CDU_Corte = false;
                bool CDU_CorteFuracao = false;
                if (corte != null)
                {

                    if (corte.Contains("Corte e Furação"))
                    {
                        CDU_CorteFuracao = true;
                    }
                    else if (corte.Contains("Corte"))
                    {
                        CDU_Corte = true;
                    }

                }

                bool C = true;
                if (CDU_Classe.Contains("S235") || CDU_Classe.Contains("S275") || CDU_Classe.Contains("S355"))
                {
                    C = true;
                }

                if (B)
                {
                    if (C)
                    {
                        Peca.GetReportProperty("PROJECT.USERDEFINED.MTTolerancias", ref CDU_Tolerancias);
                        if (CDU_Tolerancias == "")
                        {
                            CDU_Tolerancias = "EN10029 Classe A";
                        }
                        Peca.GetReportProperty("PROJECT.USERDEFINED.MTEEstsuperfchapas", ref CDU_estadoSuperficie);
                        if (CDU_estadoSuperficie == "")
                        {
                            CDU_estadoSuperficie = "EN10163-2 Classe A1";
                        }

                        if (Peca.Name.ToLower().StartsWith("br"))
                        {
                            Peca.GetReportProperty("PROJECT.USERDEFINED.MTPropespbarra", ref CDU_propriedadesEspeciais);
                        }
                        else
                        {
                            Peca.GetReportProperty("PROJECT.USERDEFINED.MTPropespchapa", ref CDU_propriedadesEspeciais);
                        }
                    }
                }
                else
                {
                    if (C)
                    {

                        Peca.GetReportProperty("PROJECT.USERDEFINED.MTEEstsuperfperfis", ref CDU_estadoSuperficie);
                        if (CDU_estadoSuperficie == "")
                        {
                            CDU_estadoSuperficie = "EN10163-3 Classe C1";
                        }
                        if (CDU_Perfil.StartsWith("BR"))
                        {
                            Peca.GetReportProperty("PROJECT.USERDEFINED.MTPropespbarra", ref CDU_propriedadesEspeciais);
                        }
                        else
                        {
                            Peca.GetReportProperty("PROJECT.USERDEFINED.MTPropespperfis", ref CDU_propriedadesEspeciais);
                        }

                    }
                }


                string CDU_Cor = null;
                Peca.GetReportProperty("USERDEFINED.CHAPA_LACADA", ref CDU_Cor);

                string CDU_Marca = null;
                string CDU_Norma = null;

                string CDU_EsquemaPintura = null;
                Peca.GetReportProperty("USERDEFINED.pintura", ref CDU_Cor);

                string CDU_DataCriacao = DateTime.Now.ToShortDateString();

                string CDU_Comentarios = null;
                Peca.GetReportProperty("comment", ref CDU_Comentarios);


//Chapa
//Conjunto
//Perfil
//Ancoragem
//ProdInterno
//ProdExterno




               


                if (CDU_Perfil.StartsWith("PL") || CDU_Perfil.StartsWith("CHA"))
                {
                    DataRow dt = dtpecaChapas.AsEnumerable().SingleOrDefault(r => r.Field<string>("CDU_Peca") == CDU_Peca && r.Field<string>("CDU_IDTekla") == CDU_IDTekla && r.Field<string>("CDU_Conjunto") == CDU_Conjunto && r.Field<string>("CDU_ReferenciaCliente") == CDU_ReferenciaCliente && r.Field<string>("CDU_RequesitosEspeciais") == CDU_RequesitosEspeciais && r.Field<string>("CDU_GrauPreparacao") == CDU_GrauPreparacao && r.Field<string>("CDU_Tolerancias") == CDU_Tolerancias && r.Field<string>("CDU_estadoSuperficie") == CDU_estadoSuperficie && r.Field<string>("CDU_propriedadesEspeciais") == CDU_propriedadesEspeciais && r.Field<string>("CDU_Cor") == CDU_Cor && r.Field<string>("CDU_Marca") == CDU_Marca && r.Field<string>("CDU_Norma") == CDU_Norma && r.Field<string>("CDU_EsquemaPintura") == CDU_EsquemaPintura && r.Field<string>("CDU_DataCriacao") == CDU_DataCriacao && r.Field<string>("CDU_Comentarios") == CDU_Comentarios);

                    string CDU_ArtigoOF = "Chapa";

                    if (dt != null)
                    {

                        int d = int.Parse(dt["CDU_Quantidade"].ToString()) + 1;
                        dt.SetField("CDU_Quantidade", d);

                        double E = double.Parse(dt["CDU_Peso"].ToString()) + CDU_Peso;
                        dt.SetField("CDU_Peso", E.ToString("0.00"));

                        E = double.Parse(dt["CDU_Area"].ToString()) + (CDU_Area / 100000);
                        dt.SetField("CDU_Area", E.ToString("0.00"));

                        List<TSM.Part> uri = dt.Field<List<TSM.Part>>("Pecas");
                        uri.Add(Peca);
                        dt.SetField<List<TSM.Part>>("Pecas", uri);

                    }
                    else
                    {

                        List<TSM.Part> Peca1 = new List<TSM.Part>();
                        Peca1.Add(Peca);
                        dtpecaChapas.Rows.Add(CDU_ID, CDU_IDCabec, CDU_IDTekla, CDU_Peca, CDU_Conjunto, CDU_Classe, CDU_Artigo, CDU_Certificado, CDU_comprimento.ToString("0"), CDU_Espessura.ToString("0"), CDU_Largura.ToString("0"), CDU_Peso.ToString("0.00"), (CDU_Area / 100000).ToString("0.00"), CDU_ClasseExecucao, CDU_ReferenciaCliente, CDU_Quantidade, CDU_Perfil, CDU_RequesitosEspeciais, CDU_GrauPreparacao, CDU_Tolerancias, CDU_estadoSuperficie, CDU_propriedadesEspeciais, CDU_Cor, CDU_Marca, CDU_Norma, CDU_EsquemaPintura, CDU_DataCriacao, CDU_Comentarios, CDU_Destinatario, CDU_Corte, CDU_CorteFuracao, "0.00", "0.00",CDU_ArtigoOF,Peca1);

                    }
                }
                else
                {

                    DataRow dt = dtpecaPerfis.AsEnumerable().SingleOrDefault(r => 
                    r.Field<string>("CDU_Peca") == CDU_Peca && 
                    r.Field<string>("CDU_IDTekla") == CDU_IDTekla && 
                    r.Field<string>("CDU_Conjunto") == CDU_Conjunto && 
                    r.Field<string>("CDU_ReferenciaCliente") == CDU_ReferenciaCliente && 
                    r.Field<string>("CDU_RequesitosEspeciais") == CDU_RequesitosEspeciais && 
                    r.Field<string>("CDU_GrauPreparacao") == CDU_GrauPreparacao && 
                    r.Field<string>("CDU_Tolerancias") == CDU_Tolerancias && 
                    r.Field<string>("CDU_estadoSuperficie") == CDU_estadoSuperficie && 
                    r.Field<string>("CDU_propriedadesEspeciais") == CDU_propriedadesEspeciais && 
                    r.Field<string>("CDU_Cor") == CDU_Cor && 
                    r.Field<string>("CDU_Norma") == CDU_Norma && 
                    r.Field<string>("CDU_EsquemaPintura") == CDU_EsquemaPintura && 
                    r.Field<string>("CDU_DataCriacao") == CDU_DataCriacao && 
                    r.Field<string>("CDU_Comentarios") == CDU_Comentarios);

                    if (CDU_Artigo==null)
                    {
                        Peca.GetAssembly().GetReportProperty("USERDEFINED.Artigo_interno", ref CDU_Artigo);
                    }

                    string CDU_ArtigoOF = "Perfil";

                    if (CDU_Destinatario=="CP")
                    {
                        CDU_ArtigoOF = "ProdInterno";
                    }
                    if (CDU_Destinatario == "DAP")
                    {
                        CDU_ArtigoOF = "ProdExterno";
                    }

                    ComunicaBDtekla db = new ComunicaBDtekla();
                    db.ConectarBD();
                    string peso = null;
                    try
                    {
                        peso = db.Procurarbd("SELECT [Peso] FROM [ArtigoTekla].[dbo].[Perfilagem3] where Perfil = '" + CDU_Perfil + "'")[0];
                    }
                    catch (Exception)
                    {

                        
                    }
                    string marca =null;
                    try
                    {
                         marca = db.Procurarbd("SELECT [Marca] FROM [ArtigoTekla].[dbo].[Perfilagem3] where Perfil = '" + CDU_Perfil + "'")[0];
                    }
                    catch (Exception)
                    {

                        
                    }
               
                    db.DesonectarBD();

                    if (peso!=null)
                    {
                        CDU_Peso = ((CDU_comprimento * double.Parse(peso))/1000) * CDU_Quantidade;
                    }
                    if (marca != null)
                    {
                        CDU_Marca = marca;
                    }


                    if (dt != null)
                    {

                        int d = int.Parse(dt["CDU_Quantidade"].ToString()) + 1;
                        dt.SetField("CDU_Quantidade", d);

                        double E = double.Parse(dt["CDU_Peso"].ToString()) + CDU_Peso;
                        dt.SetField("CDU_Peso", E.ToString("0.00"));

                        E = double.Parse(dt["CDU_Area"].ToString()) + (CDU_Area / 100000);
                        dt.SetField("CDU_Area", E.ToString("0.00"));

                        List<TSM.Part> uri = dt.Field<List<TSM.Part>>("Pecas");
                        uri.Add((TSM.Part)Peca);
                        dt.SetField<List<TSM.Part>>("Pecas", uri);

                    }
                    else
                    {

                        Peca.GetReportProperty("LENGTH", ref CDU_comprimento);
                        double AnguloLadoA = 0; Peca.GetReportProperty("END1_CUT_ANGLE_Z", ref AnguloLadoA);
                        double AnguloLadoB = 0; Peca.GetReportProperty("END2_CUT_ANGLE_Z", ref AnguloLadoB);

                        List<TSM.Part> Peca1 = new List<TSM.Part>();
                        Peca1.Add(Peca);
                        dtpecaPerfis.Rows.Add(CDU_ID, CDU_IDCabec, CDU_IDTekla, CDU_Peca, CDU_Conjunto, CDU_Classe, CDU_Artigo, CDU_Certificado, CDU_comprimento.ToString("0"), CDU_Espessura.ToString("0"), CDU_Largura.ToString("0"), CDU_Peso.ToString("0.00"), (CDU_Area / 100000).ToString("0.00"), CDU_ClasseExecucao, CDU_ReferenciaCliente, CDU_Quantidade, CDU_Perfil, CDU_RequesitosEspeciais, CDU_GrauPreparacao, CDU_Tolerancias, CDU_estadoSuperficie, CDU_propriedadesEspeciais, CDU_Cor, CDU_Marca, CDU_Norma, CDU_EsquemaPintura, CDU_DataCriacao, CDU_Comentarios, CDU_Destinatario, CDU_Corte, CDU_CorteFuracao, AnguloLadoA.ToString("0.00"), AnguloLadoB.ToString("0.00"),CDU_ArtigoOF, Peca1);

                    }
                }

            }

        }

        /// <summary>
        /// Preenche a lista de parafusos quando o modelado como perfil
        /// </summary>
        /// <param name="Conjunto"></param>
        void parafuso(Assembly Conjunto)
        {

            string Artigo = null;
            double Comprimento = 0;
            string _Comprimento = null;
            int Quantidade = 1;
            string Classe = "";
            string Norma = null;
            string lote = null;
            string Entrega = null;
            Conjunto.GetReportProperty("MAINPART.PROFILE", ref Artigo);
            Conjunto.GetReportProperty("LENGTH", ref Comprimento);
            Conjunto.GetReportProperty("MAINPART.MATERIAL", ref Classe);
            Conjunto.GetReportProperty("USERDEFINED.lote_data", ref Entrega);
            Conjunto.GetReportProperty("USERDEFINED.lote_number", ref lote);
            if (Artigo.Contains("VRSM"))
            {
                _Comprimento = Comprimento.ToString("0");
                Norma = "DIN976";
            }
            else
            {
                _Comprimento = "";
                if (Artigo.Contains("NM"))
                {
                    Norma = "DIN934";
                }
                else
                {
                    Norma = "300HV";
                }
            }

            DataRow dt = _parafusos.AsEnumerable().SingleOrDefault(r => r.ItemArray[0].ToString() == Artigo &&
           r.ItemArray[1] + "" == _Comprimento && r.ItemArray[3] + "" == Classe);

            if (dt != null)
            {
                int d = int.Parse(dt["Quantidade"].ToString()) + 1;
                dt.SetField("Quantidade", d);

                ArrayList uri = dt.Field<ArrayList>(7);
                uri.Add(Conjunto);
                dt.SetField<ArrayList>(7, uri);

            }
            else
            {
                ArrayList Pecas = new ArrayList();
                Pecas.Add(Conjunto);
                _parafusos.Rows.Add(Artigo, _Comprimento, Quantidade, Classe, Norma, lote ,Entrega, Pecas);
            }
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView4_CellDoubleClick_1(object sender, DataGridViewCellEventArgs e)
        {
            DataGridView GRID = (DataGridView)sender;
            DataTable bs = (DataTable)GRID.DataSource; // Se convierte el DataSource 



            try
            {
                ComunicaTekla.selectinmodel(new ArrayList(bs.DefaultView.ToTable().Rows[e.RowIndex].Field<List<Assembly>>(bs.Columns.Count - 1).ToArray()));
            }
            catch (Exception)
            {
                try
                {
                    ComunicaTekla.selectinmodel(new ArrayList(bs.DefaultView.ToTable().Rows[e.RowIndex].Field<List<TSM.Part>>(bs.Columns.Count - 1).ToArray()));
                }
                catch (Exception)
                {
                    try
                    {
                        ComunicaTekla.selectinmodel(bs.DefaultView.ToTable().Rows[e.RowIndex].Field<ArrayList>(bs.Columns.Count - 1));
                    }
                    catch (Exception)
                    {

                    }

                }

            }
            finally
            {

            }

        }

        private static void RemoveUnusedColumnsAndRows(DataTable table)
        {

            foreach (var column in table.Columns.Cast<DataColumn>().ToArray())
            {
                if (table.AsEnumerable().All(dr => dr.IsNull(column)))
                    table.Columns[column.ColumnName].ColumnMapping = MappingType.Hidden;
            }
            table.AcceptChanges();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            dataGridView2.DataSource = null;
            dataGridView3.DataSource = null;
            dataGridView4.DataSource = null;

            RemoveUnusedColumnsAndRows(dtconj);
            RemoveUnusedColumnsAndRows(dtpecaPerfis);
            RemoveUnusedColumnsAndRows(dtpecaChapas);
            RemoveUnusedColumnsAndRows(_parafusos);

            dataGridView1.DataSource = dtconj;
            dataGridView2.DataSource = dtpecaPerfis;
            dataGridView3.DataSource = dtpecaChapas;
            dataGridView4.DataSource = _parafusos;
        }

        private void button3_Click(object sender, EventArgs e)
        {







        }

        void verificaperfil()
        {
            List<DataGridView> dt = new List<DataGridView>();
            dt.Add(dataGridView1);
            dt.Add(dataGridView2);
            dt.Add(dataGridView3);
            foreach (DataGridView item in dt)
            {
                foreach (DataGridViewRow Linha in item.Rows)
                {
                    Linha.Cells["CDU_Perfil"].Value = (Linha.Cells["CDU_Perfil"].Value + "").Replace(".", ",");
                    if ((Linha.Cells["CDU_Perfil"].Value + "").StartsWith("CFRHS"))
                    {
                        Linha.Cells["CDU_Perfil"].Value = (Linha.Cells["CDU_Perfil"].Value + "").Replace("HS", "");

                        string[] NovoArtigo = Linha.Cells["CDU_Perfil"].Value.ToString().Split('X');
                        if (NovoArtigo.Length == 2)
                        {
                            Linha.Cells["CDU_Perfil"].Value = NovoArtigo[0] + "X" + NovoArtigo[0].ToString().Replace("CFR", "") + "X" + NovoArtigo[1];
                        }
                    }
                    else if ((Linha.Cells["CDU_Perfil"].Value + "").StartsWith("HFRHS"))
                    {
                        Linha.Cells["CDU_Perfil"].Value = (Linha.Cells["CDU_Perfil"].Value + "").Replace("HS", "");

                        string[] NovoArtigo = Linha.Cells["CDU_Perfil"].Value.ToString().Split('X');
                        if (NovoArtigo.Length == 2)
                        {
                            Linha.Cells["CDU_Perfil"].Value = NovoArtigo[0] + "X" + NovoArtigo[0].ToString().Replace("HFR", "") + "X" + NovoArtigo[1];
                        }
                    }
                    else if ((Linha.Cells["CDU_Perfil"].Value + "").StartsWith("TGRHS"))
                    {
                        Linha.Cells["CDU_Perfil"].Value = (Linha.Cells["CDU_Perfil"].Value + "").Replace("TG", "");

                        string[] NovoArtigo = Linha.Cells["CDU_Perfil"].Value.ToString().Split('X');
                        if (NovoArtigo.Length == 2)
                        {
                            Linha.Cells["CDU_Perfil"].Value = NovoArtigo[0] + "X" + NovoArtigo[0].ToString().Replace("TGR", "") + "X" + NovoArtigo[1];
                        }
                    }
                    else if ((Linha.Cells["CDU_Perfil"].Value + "").StartsWith("CFCHS"))
                    {
                        Linha.Cells["CDU_Perfil"].Value = (Linha.Cells["CDU_Perfil"].Value + "").Replace("HS", "");
                    }
                    else if ((Linha.Cells["CDU_Perfil"].Value + "").StartsWith("HFCHS"))
                    {
                        Linha.Cells["CDU_Perfil"].Value = (Linha.Cells["CDU_Perfil"].Value + "").Replace("HS", "");
                    }
                    else if ((Linha.Cells["CDU_Perfil"].Value + "").StartsWith("TGCHS"))
                    {
                        Linha.Cells["CDU_Perfil"].Value = (Linha.Cells["CDU_Perfil"].Value + "").Replace("HS", "");
                    }
                    else if ((Linha.Cells["CDU_Perfil"].Value + "").StartsWith("PL"))
                    {
                        Linha.Cells["CDU_Perfil"].Value = (Linha.Cells["CDU_Perfil"].Value + "").Replace("PL", "").Split('X')[0];
                    }
                    else if ((Linha.Cells["CDU_Perfil"].Value + "").StartsWith("CHA"))
                    {
                        Linha.Cells["CDU_Perfil"].Value = (Linha.Cells["CDU_Perfil"].Value + "").Replace("CHA", "").Split('X')[0];
                    }
                }
            }
        }

        private void Guardar_Click(object sender, EventArgs e)
        {
            List<string> dbOperations = new List<string>();
            DataTable view = dataGridView1.DataSource as DataTable;
            foreach (DataRow Linha in view.Rows)
            {

                string CDU_ID = Linha.Field<string>("CDU_ID") ;
                string CDU_IDTekla = Linha.Field<string>("CDU_IDTekla");
                string CDU_Conjunto = Linha.Field<string>("CDU_Conjunto");
                string CDU_CodigoCliente = Linha.Field<string>("CDU_CodigoCliente");
                string CDU_NomeCliente = Linha.Field<string>("CDU_NomeCliente");
                string CDU_CodigoObra = Linha.Field<string>("CDU_CodigoObra");
                string CDU_DescricaoObra = Linha.Field<string>("CDU_DescricaoObra");
                string CDU_Fase = Linha.Field<string>("CDU_Fase");
                string CDU_Lote = Linha.Field<string>("CDU_Lote");
                string CDU_Perfil = Linha.Field<string>("CDU_Perfil");
                string CDU_Artigo = Linha.Field<string>("CDU_Artigo");
                string CDU_DataInicioProducao = Linha.Field<string>("CDU_DataInicioProducao");

             


                string[] data = null;
                if (Linha.Field<string>("CDU_DataEntregaPrevista").Contains("/"))
                {
                    data = Linha.Field<string>("CDU_DataEntregaPrevista").Split('/');
                }
                else if (Linha.Field<string>("CDU_DataEntregaPrevista").Contains("-"))
                {
                    data = Linha.Field<string>("CDU_DataEntregaPrevista").Split('-');
                }

                string CDU_DataEntregaPrevista = data[2] + "-" + data[1] + "-" + data[0];

                data = null;

                if (Linha.Field<string>("CDU_DataCriacao").Contains("/"))
                {
                    data = Linha.Field<string>("CDU_DataCriacao").Split('/');
                }
                else if (Linha.Field<string>("CDU_DataCriacao").Contains("-"))
                {
                    data = Linha.Field<string>("CDU_DataCriacao").Split('-');
                }

                string CDU_DataCriacao = data[2] + "-" + data[1] + "-" + data[0];
                string CDU_comprimento = Linha.Field<string>("CDU_comprimento");
                string CDU_Altura =  Linha.Field<string>("CDU_Altura");
                string CDU_Largura = Linha.Field<string>("CDU_Largura");
                string CDU_GrauPreparacao = Linha.Field<string>("CDU_GrauPreparacao");
                string CDU_ClasseExecucao = Linha.Field<string>("CDU_ClasseExecucao");
                string CDU_ReferenciaCliente = Linha.Field<string>("CDU_ReferenciaCliente");
                string CDU_Quantidade = Linha.Field<string>("CDU_Quantidade");
                string CDU_Comentarios = Linha.Field<string>("CDU_Comentarios");
                string CDU_PreparacaoChapas = Linha.Field<bool>("CDU_PreparacaoChapas").ToString();
                string CDU_Armacao = Linha.Field<bool>("CDU_Armacao").ToString();
                string CDU_Soldadura = Linha.Field<bool>("CDU_Soldadura").ToString();
                string CDU_Decapagem = Linha.Field<bool>("CDU_Decapagem").ToString();
                string CDU_Pintura = Linha.Field<bool>("CDU_Pintura").ToString();
                string CDU_Contagem = Linha.Field<bool>("CDU_Contagem").ToString();


                dbOperations.Add("INSERT INTO[dbo].[TDU_INCabecFichaTEcnicaCM](" +
                    "[CDU_ID]," +
                    "[CDU_IDTekla]," +
                    "[CDU_Conjunto]," +
                    "[CDU_CodigoCliente]," +
                    "[CDU_NomeCliente]," +
                    "[CDU_CodigoObra]," +
                    "[CDU_DescricaoObra]," +
                    "[CDU_Fase],[CDU_Lote]," +
                    "[CDU_Perfil]," +
                    "[CDU_Artigo]," +
                    "[CDU_DataInicioProducao]," +
                    "[CDU_DataEntregaPrevista]," +
                    "[CDU_DataCriacao]," +
                    "[CDU_comprimento]," +
                    "[CDU_Altura]," +
                    "[CDU_Largura]," +
                    "[CDU_GrauPreparacao]," +
                    "[CDU_ClasseExecucao]," +
                    "[CDU_ReferenciaCliente]," +
                    "[CDU_Quantidade]," +
                    "[CDU_Comentarios]," +
                    "[CDU_PreparacaoChapas]," +
                    "[CDU_Armacao]," +
                    "[CDU_Soldadura]," +
                    "[CDU_Decapagem]," +
                    "[CDU_Pintura]," +
                    "[CDU_Contagem]" +
                    ")VALUES('"
                                    + CDU_ID
                                    + "','" + CDU_IDTekla
                                    + "','" + CDU_Conjunto
                                    + "','" + CDU_CodigoCliente
                                    + "','" + CDU_NomeCliente
                                    + "','" + CDU_CodigoObra
                                    + "','" + CDU_DescricaoObra
                                    + "','" + CDU_Fase
                                    + "','" + CDU_Lote
                                    + "','" + CDU_Perfil
                                    + "','" + CDU_Artigo
                                    + "','" + CDU_DataInicioProducao
                                    + "','" + CDU_DataEntregaPrevista
                                    + "','" + CDU_DataCriacao
                                    + "','" + CDU_comprimento
                                    + "','" + CDU_Altura
                                    + "','" + CDU_Largura
                                    + "','" + CDU_GrauPreparacao
                                    + "','" + CDU_ClasseExecucao
                                    + "','" + CDU_ReferenciaCliente
                                    + "','" + CDU_Quantidade
                                    + "','" + CDU_Comentarios
                                    + "','" + CDU_PreparacaoChapas
                                    + "','" + CDU_Armacao
                                    + "','" + CDU_Soldadura
                                    + "','" + CDU_Decapagem
                                    + "','" + CDU_Pintura
                                    + "','" + CDU_Contagem
                                    + "')");

            }
            DataTable view1 = dataGridView2.DataSource as DataTable;
            foreach (DataRow Linha in view1.Rows)
            {

                string CDU_ID = Linha.Field<string>("CDU_ID");
                string CDU_IDTekla = Linha.Field<string>("CDU_IDTekla");
                string CDU_Peca = Linha.Field<string>("CDU_Peca");
                string CDU_Conjunto = Linha.Field<string>("CDU_Conjunto");
                string CDU_Classe = Linha.Field<string>("CDU_Classe");
                string CDU_Artigo = Linha.Field<string>("CDU_Artigo");
                string CDU_Certificado = Linha.Field<string>("CDU_Certificado");
                string CDU_comprimento = Linha.Field<string>("CDU_comprimento");
                string CDU_Espessura = Linha.Field<string>("CDU_Espessura");
                string CDU_Largura = Linha.Field<string>("CDU_Largura");
                string CDU_Peso = Linha.Field<string>("CDU_Peso");
                string CDU_Area = Linha.Field<string>("CDU_Area");
                string CDU_ClasseExecucao = Linha.Field<string>("CDU_ClasseExecucao");
                string CDU_ReferenciaCliente = Linha.Field<string>("CDU_ReferenciaCliente");
                string CDU_Quantidade = Linha.Field<string>("CDU_Quantidade");
                string CDU_Perfil = Linha.Field<string>("CDU_Perfil");
                string CDU_RequesitosEspeciais = Linha.Field<string>("CDU_RequesitosEspeciais");
                string CDU_GrauPreparacao = Linha.Field<string>("CDU_GrauPreparacao");
                string CDU_Tolerancias = Linha.Field<string>("CDU_Tolerancias");
                string CDU_estadoSuperficie = Linha.Field<string>("CDU_estadoSuperficie");
                string CDU_propriedadesEspeciais = Linha.Field<string>("CDU_propriedadesEspeciais");
                string CDU_Cor = Linha.Field<string>("CDU_Cor");
                string CDU_Marca = Linha.Field<string>("CDU_Marca");
                string CDU_Norma = Linha.Field<string>("CDU_Norma");
                string CDU_EsquemaPintura = Linha.Field<string>("CDU_EsquemaPintura");

                string[] data = Linha.Field<string>("CDU_DataCriacao").Split('/');
                string CDU_DataCriacao = data[2] + "-" + data[1] + "-" + data[0];

                string CDU_Comentarios = Linha.Field<string>("CDU_Comentarios");
                string CDU_Destinatario = Linha.Field<string>("CDU_Destinatario");
                string CDU_Corte = Linha.Field<bool>("CDU_Corte").ToString();
                string CDU_CorteFuracao = Linha.Field<bool>("CDU_CorteFuracao").ToString();
                string CDU_AnguloA = Linha.Field<string>("CDU_AnguloA");
                string CDU_AnguloB = Linha.Field<string>("CDU_AnguloB");
                string CDU_ArtigoOF = Linha.Field<string>("CDU_ArtigoOF");
                string CDU_IDCabec = Linha.Field<string>("CDU_IDCabec");

                dbOperations.Add("INSERT INTO[dbo].[TDU_INPecasConjFichaTEcnicaCM] (" +
                    "[CDU_ID]," +                                    
                    "[CDU_IDCabec]," +
                    "[CDU_IDTekla]," +
                    "[CDU_Peca]," +
                    "[CDU_Conjunto]," +
                    "[CDU_Classe]," +
                    "[CDU_Artigo]," +
                    "[CDU_Certificado]," +
                    "[CDU_comprimento]," +
                    "[CDU_Espessura]," +
                    "[CDU_Largura]," +
                    "[CDU_Peso]," +
                    "[CDU_Area]," +
                    "[CDU_ClasseExecucao]," +
                    "[CDU_ReferenciaCliente]," +
                    "[CDU_Quantidade]," +
                    "[CDU_Perfil]," +
                    "[CDU_RequesitosEspeciais]," +
                    "[CDU_GrauPreparacao]," +
                    "[CDU_Tolerancias]," +
                    "[CDU_estadoSuperficie]," +
                    "[CDU_propriedadesEspeciais]," +
                    "[CDU_Cor]," +
                    "[CDU_Marca]," +
                    "[CDU_Norma]," +
                    "[CDU_EsquemaPintura]," +
                    "[CDU_DataCriacao]," +
                    "[CDU_Comentarios]," +
                    "[CDU_Destinatario]," +
                    "[CDU_Corte]," +
                    "[CDU_CorteFuracao]," +
                    "[CDU_AnguloA]," +
                    "[CDU_AnguloB]," +
                    "[CDU_ArtigoOF]" +
                    ") VALUES('"
                                    + CDU_ID
                                    + "','" + CDU_IDCabec
                                    + "','" + CDU_IDTekla
                                    + "','" + CDU_Peca
                                    + "','" + CDU_Conjunto
                                    + "','" + CDU_Classe
                                    + "','" + CDU_Artigo
                                    + "','" + CDU_Certificado
                                    + "','" + CDU_comprimento
                                    + "','" + CDU_Espessura
                                    + "','" + CDU_Largura
                                    + "','" + CDU_Peso
                                    + "','" + CDU_Area
                                    + "','" + CDU_ClasseExecucao
                                    + "','" + CDU_ReferenciaCliente
                                    + "','" + CDU_Quantidade
                                    + "','" + CDU_Perfil
                                    + "','" + CDU_RequesitosEspeciais
                                    + "','" + CDU_GrauPreparacao
                                    + "','" + CDU_Tolerancias
                                    + "','" + CDU_estadoSuperficie
                                    + "','" + CDU_propriedadesEspeciais
                                    + "','" + CDU_Cor
                                    + "','" + CDU_Marca
                                    + "','" + CDU_Norma
                                    + "','" + CDU_EsquemaPintura
                                    + "','" + CDU_DataCriacao
                                    + "','" + CDU_Comentarios
                                    + "','" + CDU_Destinatario
                                    + "','" + CDU_Corte
                                    + "','" + CDU_CorteFuracao
                                    + "','" + CDU_AnguloA.Replace(",",".")
                                    + "','" + CDU_AnguloB.Replace(",", ".")
                                    + "','" + CDU_ArtigoOF
                                    + "')");

            }
            DataTable view3 = dataGridView3.DataSource as DataTable;
            foreach (DataRow Linha in view3.Rows)
            {
                string CDU_ID = Linha.Field<string>("CDU_ID");
                string CDU_IDTekla = Linha.Field<string>("CDU_IDTekla");
                string CDU_Peca = Linha.Field<string>("CDU_Peca");
                string CDU_Conjunto = Linha.Field<string>("CDU_Conjunto");
                string CDU_Classe = Linha.Field<string>("CDU_Classe");
                string CDU_Artigo = Linha.Field<string>("CDU_Artigo");
                string CDU_Certificado = Linha.Field<string>("CDU_Certificado");
                string CDU_comprimento = Linha.Field<string>("CDU_comprimento");
                string CDU_Espessura = Linha.Field<string>("CDU_Espessura");
                string CDU_Largura = Linha.Field<string>("CDU_Largura");
                string CDU_Peso = Linha.Field<string>("CDU_Peso");
                string CDU_Area = Linha.Field<string>("CDU_Area");
                string CDU_ClasseExecucao = Linha.Field<string>("CDU_ClasseExecucao");
                string CDU_ReferenciaCliente = Linha.Field<string>("CDU_ReferenciaCliente");
                string CDU_Quantidade = Linha.Field<string>("CDU_Quantidade");
                string CDU_Perfil = Linha.Field<string>("CDU_Perfil");
                string CDU_RequesitosEspeciais = Linha.Field<string>("CDU_RequesitosEspeciais");
                string CDU_GrauPreparacao = Linha.Field<string>("CDU_GrauPreparacao");
                string CDU_Tolerancias = Linha.Field<string>("CDU_Tolerancias");
                string CDU_estadoSuperficie = Linha.Field<string>("CDU_estadoSuperficie");
                string CDU_propriedadesEspeciais = Linha.Field<string>("CDU_propriedadesEspeciais");
                string CDU_Cor = Linha.Field<string>("CDU_Cor");
                string CDU_Marca = Linha.Field<string>("CDU_Marca");
                string CDU_Norma = Linha.Field<string>("CDU_Norma");
                string CDU_EsquemaPintura = Linha.Field<string>("CDU_EsquemaPintura");
                string[] data = Linha.Field<string>("CDU_DataCriacao").Split('/');
                string CDU_DataCriacao = data[2] + "-" + data[1] + "-" + data[0];
                string CDU_Comentarios = Linha.Field<string>("CDU_Comentarios");
                string CDU_Destinatario = Linha.Field<string>("CDU_Destinatario");
                string CDU_Corte = Linha.Field<bool>("CDU_Corte").ToString();
                string CDU_CorteFuracao = Linha.Field<bool>("CDU_CorteFuracao").ToString();
                string CDU_AnguloA = Linha.Field<string>("CDU_AnguloA");
                string CDU_AnguloB = Linha.Field<string>("CDU_AnguloB");
                string CDU_ArtigoOF = Linha.Field<string>("CDU_ArtigoOF");
                string CDU_IDCabec = Linha.Field<string>("CDU_IDCabec");


                dbOperations.Add("INSERT INTO[dbo].[TDU_INPecasConjFichaTEcnicaCM] (" +
                    "[CDU_ID]," +
                    "[CDU_IDCabec]," +
                    "[CDU_IDTekla]," +
                    "[CDU_Peca]," +
                    "[CDU_Conjunto]," +
                    "[CDU_Classe]," +
                    "[CDU_Artigo]," +
                    "[CDU_Certificado]," +
                    "[CDU_comprimento]," +
                    "[CDU_Espessura]," +
                    "[CDU_Largura]," +
                    "[CDU_Peso]," +
                    "[CDU_Area]," +
                    "[CDU_ClasseExecucao]," +
                    "[CDU_ReferenciaCliente]," +
                    "[CDU_Quantidade]," +
                    "[CDU_Perfil]," +
                    "[CDU_RequesitosEspeciais]," +
                    "[CDU_GrauPreparacao]," +
                    "[CDU_Tolerancias]," +
                    "[CDU_estadoSuperficie]," +
                    "[CDU_propriedadesEspeciais]," +
                    "[CDU_Cor]," +
                    "[CDU_Marca]," +
                    "[CDU_Norma]," +
                    "[CDU_EsquemaPintura]," +
                    "[CDU_DataCriacao]," +
                    "[CDU_Comentarios]," +
                    "[CDU_Destinatario]," +
                    "[CDU_Corte]," +
                    "[CDU_CorteFuracao]," +
                    "[CDU_AnguloA]," +
                    "[CDU_AnguloB]," +
                    "[CDU_ArtigoOF]" +
                    ") VALUES('"
                                    + CDU_ID
                                    + "','" + CDU_IDCabec
                                    + "','" + CDU_IDTekla
                                    + "','" + CDU_Peca
                                    + "','" + CDU_Conjunto
                                    + "','" + CDU_Classe
                                    + "','" + CDU_Artigo
                                    + "','" + CDU_Certificado
                                    + "','" + CDU_comprimento
                                    + "','" + CDU_Espessura
                                    + "','" + CDU_Largura
                                    + "','" + CDU_Peso
                                    + "','" + CDU_Area
                                    + "','" + CDU_ClasseExecucao
                                    + "','" + CDU_ReferenciaCliente
                                    + "','" + CDU_Quantidade
                                    + "','" + CDU_Perfil
                                    + "','" + CDU_RequesitosEspeciais
                                    + "','" + CDU_GrauPreparacao
                                    + "','" + CDU_Tolerancias
                                    + "','" + CDU_estadoSuperficie
                                    + "','" + CDU_propriedadesEspeciais
                                    + "','" + CDU_Cor
                                    + "','" + CDU_Marca
                                    + "','" + CDU_Norma
                                    + "','" + CDU_EsquemaPintura
                                    + "','" + CDU_DataCriacao
                                    + "','" + CDU_Comentarios
                                    + "','" + CDU_Destinatario
                                    + "','" + CDU_Corte
                                    + "','" + CDU_CorteFuracao
                                    + "','" + CDU_AnguloA.Replace(",", ".")
                                    + "','" + CDU_AnguloB.Replace(",", ".")
                                    + "','" + CDU_ArtigoOF
                                    + "')");

            }
            ComunicaBDprimavera a = new ComunicaBDprimavera();
            a.dbOperations(dbOperations);


            MessageBox.Show(this, "Dados Importados com sucesso","Importado", MessageBoxButtons.OK, MessageBoxIcon.Information);


        }

        void MoveCorte(string Peca , string Fase ,string NumeroObra)
        {
           // Form1.PastaReservatorioFicheiros;


        }

        void MoveCorteEFuracao(string peca)
        {

        }

        void MoveCP(string peca)
        {

        }

        void MoveArmacao(string peca)
        {
           
        }

    }
}

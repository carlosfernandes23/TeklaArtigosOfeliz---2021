using Microsoft.Reporting.WinForms;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace TeklaArtigosOfeliz
{
    public partial class Frm_Nest : Form
    {
       
        public Frm_Nest()
        {
            InitializeComponent();
        }

        

        private void cARREGARPEÇASTEKLAToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            // OBTER PEÇAS DO TEKLA
            ArrayList PECAS = ComunicaTekla.ListadePecasSelec();
            List<PecaCorte> ListaPecasCorte = new List<PecaCorte>();

            foreach (Tekla.Structures.Model.Part PECA in PECAS)
            {
                string _referencia = null;
                if (PECA !=null)
                {
                    PecaCorte peca = new PecaCorte();
                    PECA.GetReportProperty("PROFILE", ref peca.Perfil);
                    if (!(peca.Perfil.Contains("PL")|| peca.Perfil.Contains("CHA")))
                    {
                        PECA.GetReportProperty("PART_POS", ref _referencia);
                        PECA.GetReportProperty("LENGTH", ref peca.Comprimento);
                   
                        peca.Referencia = _referencia;
                        ListaPecasCorte.Add(peca);
                    }
                }
            }
           
            foreach (var peca in FILTRAR(ListaPecasCorte))
            {
                dataGridView1.Rows.Add(peca.Quantidade,(int)peca.Comprimento, peca.Referencia, peca.Perfil);
            }
            
        }

        internal List<PecaCorte> FILTRAR(List<PecaCorte> ListaPecasCorte)
        {
            List<PecaCorte> newlist = new List<PecaCorte>();

            List<string> repetido=new List<string>();

            foreach (PecaCorte PECA in ListaPecasCorte)
            {
                if (!repetido.Contains(PECA.Referencia))
                {
                    repetido.Add(PECA.Referencia);
                    foreach (PecaCorte PECA1 in ListaPecasCorte)
                    {

                        if (PECA.Referencia == PECA1.Referencia)
                        {

                            PECA.Quantidade += 1;

                        }

                    }

                    PecaCorte p = new PecaCorte();
                    p.Comprimento = PECA.Comprimento;
                    p.Perfil = PECA.Perfil;
                    p.Referencia = PECA.Referencia;
                    p.Quantidade = PECA.Quantidade;
                    newlist.Add(p);
                }

            }
            return newlist;
        }

        private void cALCULARToolStripMenuItem_Click(object sender, EventArgs e)
        {

            


            DataSet1.DataTable1.Rows.Clear();
            DataSet1.DataTable2.Rows.Clear();
            X = 0;
          
            DataTable teste = new DataTable();
            foreach (var item in BuscaPerfisExixtententes())
            {
                teste = otimizar(ApanhaPecas(item), ApanhaPerfis(item), 0);
                
                DataTable table = new DataTable();
                table.Columns.Add("perfil");
                table.Columns.Add("comprimento");
                table.Columns.Add("quantidade");
                table.Columns.Add("peca");
                table.Columns.Add("compriment");
                table.Columns.Add("qunt");
                table.Columns.Add("desperdicio");
                string perfil = null;
                string Comprimento = null;
                string quantidades = null;
                string peca = null;
                string compriment = null;
                string qunt = null;
                string desperdicio = null;


                if (teste != null)
                {
                    DataSet1.DataTable2.Rows.Add(item, Percentagemdedesperdicio);
                    //correr linhas da lista
                    foreach (DataRow item1 in teste.Rows)
                    {

                        perfil = item1.ItemArray[2].ToString();
                        //prenche perfil
                        Comprimento = item1.ItemArray[1].ToString();
                        //prenche quantidade que é sempre 1 :)
                        quantidades = "1";
                        //prenche despedicio
                        desperdicio = item1.ItemArray[4].ToString();
                        //inicia novo contador que serve para defenir qual a referencia de peça a preencher, inicia lista de repetidos.

                        List<string> repetidos = new List<string>();

                        List<string> fList = new List<string>();

                        //add each last column's value to the list
                        List<float> comp = (List<float>)item1[3];
                        List<string> marca = (List<string>)item1[5];
                        List<float> comprimento = new List<float>();
                        List<int> quantidade = new List<int>();
                        List<string> marcapeça = new List<string>();
                        List<string> f = marca.Distinct().ToList();

                        foreach (string marcaq in f)
                        {
                            quantidade.Add(marca.Where(s => (s.Contains(marcaq))).Count());
                            comprimento.Add(comp[marca.IndexOf(marcaq)]);
                            marcapeça.Add(marcaq.Replace("\"", ""));
                        }
                        for (int i = 0; i < f.Count; i++)
                        {

                            compriment = comprimento[i].ToString("0");
                            peca = marcapeça[i];
                            qunt = quantidade[i].ToString("0");
                            table.Rows.Add(perfil, Comprimento, quantidades, peca, compriment, qunt, desperdicio);

                            //limpa variaveis para a proxima linha 
                            perfil = null;
                            Comprimento = null;
                            quantidades = null;
                            peca = null;
                            compriment = null;
                            qunt = null;
                            desperdicio = null;
                        }

                    }

                }
                else
                {
                    DataSet1.DataTable2.Rows.Add(item, "SEM STOCK");
                }

                foreach (DataRow item1 in table.Rows)
                {
                    DataSet1.DataTable1.Rows.Add(item1.ItemArray);
                }
            }
            string[] TobeDistinct = { "BARRA", "ID" };

            DataTable dtDistinct = GetDistinctRecords(DataSet1.DataTable1, TobeDistinct);
            DataSet1.DataTable3.Rows.Clear();
            foreach (DataRow item1 in dtDistinct.Rows)
            {
                if (!string.IsNullOrEmpty(item1.ItemArray[0] + ""))
                {

                    int R = 0;
                    foreach (DataRow row in DataSet1.DataTable1.Rows)
                    {
                        if (row.ItemArray[1].ToString().ToLower() == item1.ItemArray[0].ToString().ToLower() && row.ItemArray[0].ToString().ToLower() == item1.ItemArray[1].ToString().ToLower())
                        {
                            R += 1;
                        }
                    }
                    DataSet1.DataTable3.Rows.Add(R, item1.ItemArray[0], item1.ItemArray[1]);
                }


            }

            reportViewer1.RefreshReport();
        }
  

    //Following function will return Distinct records for Name, City and State column.
    public static DataTable GetDistinctRecords(DataTable dt, string[] Columns)
    {
        DataTable dtUniqRecords = new DataTable();
        dtUniqRecords = dt.DefaultView.ToTable(true, Columns);
        return dtUniqRecords;
    }

    internal DataTable ApanhaPecas(string perfil)
        {
            DataTable pecas = new DataTable();
            pecas.Columns.Add("quantidade");
            pecas.Columns.Add("comprimento");
            pecas.Columns.Add("referencia");

            foreach (DataGridViewRow item in dataGridView1.Rows)
            {
                if (!string.IsNullOrEmpty(item.Cells[3].Value + ""))
                {
                    if ((item.Cells[3].Value + "").ToLower()==perfil.ToLower())
                    {
                        pecas.Rows.Add(item.Cells[0].Value, item.Cells[1].Value, item.Cells[2].Value);
                    }
                }
            }
            return pecas;
        }
       internal DataTable ApanhaPerfis(string perfil)
        {
            DataTable barras = new DataTable();
            barras.Columns.Add("quantidade");
            barras.Columns.Add("comprimento");
            barras.Columns.Add("prioridade");
            barras.Columns.Add("referencia");

            foreach (DataGridViewRow item in dataGridView2.Rows)
            {
                if (!string.IsNullOrEmpty(item.Cells[3].Value + ""))
                {
                    if ((item.Cells[3].Value + "").ToLower() == perfil.ToLower())
                    {
                        barras.Rows.Add(item.Cells[0].Value, item.Cells[1].Value, item.Cells[2].Value, item.Cells[3].Value);
                    }
                }
            }

            return barras;
        }

        List<string> BuscaPerfisExixtententes()
        {
            List<string> lista = new List<string>();
            foreach (DataGridViewRow item in dataGridView1.Rows)
            {
                if (!string.IsNullOrEmpty(item.Cells[3].Value + ""))
                {
                   
                    lista.Add((item.Cells[3].Value + "").Replace(" ",""));
                }
            }
            return lista.Distinct().ToList();
        }
        private void criaexcel(DataTable teste, String outputPath, string Perfil, string numerodeobra="", string nomecliente="", string nomeobra="")
        {

            if (teste != null)
            {

              
                //iniciar excel
                Excel.Application excel = new Excel.Application();
                Excel.Workbook workbook = excel.Workbooks.Open(Environment.CurrentDirectory + "\\template.xls", 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Excel.Worksheet sheet = (Excel.Worksheet)workbook.Worksheets.get_Item(1);
                //prencher o cabeçalho da pagina 
                ((Excel.Range)sheet.Cells[1, 2]).Value = numerodeobra.Replace(";", "|");
                ((Excel.Range)sheet.Cells[2, 2]).Value = nomecliente.Replace(";", "|");
                ((Excel.Range)sheet.Cells[3, 2]).Value = nomeobra.Replace(";", "|");
                ((Excel.Range)sheet.Cells[1, 4]).Value = Percentagemdedesperdicio;
                //iniciar os contadores
                int contador = 5;
                int contador1 = 1;

                //correr linhas da lista
                foreach (DataRow item in teste.Rows)
                {

                    //se o numero for par a cor da celula altera
                    contador1++;
                    if (contador1 % 2 != 0)
                    {

                        ((Excel.Range)sheet.Cells[contador, 1]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                        ((Excel.Range)sheet.Cells[contador, 2]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                        ((Excel.Range)sheet.Cells[contador, 3]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                        ((Excel.Range)sheet.Cells[contador, 5]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                        ((Excel.Range)sheet.Cells[contador, 4]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                        ((Excel.Range)sheet.Cells[contador, 6]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                        ((Excel.Range)sheet.Cells[contador, 7]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);

                    }
                    else
                    {
                        ((Excel.Range)sheet.Cells[contador, 1]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                        ((Excel.Range)sheet.Cells[contador, 2]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                        ((Excel.Range)sheet.Cells[contador, 3]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                        ((Excel.Range)sheet.Cells[contador, 5]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                        ((Excel.Range)sheet.Cells[contador, 4]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                        ((Excel.Range)sheet.Cells[contador, 6]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                        ((Excel.Range)sheet.Cells[contador, 7]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                    }

                    //prencher os dados da barra 
                    //prenche comprimento
                    ((Excel.Range)sheet.Cells[contador, 1]).Value = item.ItemArray[2].ToString();
                    //prenche perfil
                    ((Excel.Range)sheet.Cells[contador, 2]).Value = item.ItemArray[1].ToString();
                    //prenche quantidade que é sempre 1 :)
                    ((Excel.Range)sheet.Cells[contador, 3]).Value = 1;
                    //prenche despedicio
                    ((Excel.Range)sheet.Cells[contador, 7]).Value = item.ItemArray[4];

                    //Neste momento linha tipo ipe160,12000,1,    ,    ,    ,44


                    //inicia novo contador que serve para defenir qual a referencia de peça a preencher, inicia lista de repetidos.
                    int novocontador = 0;
                    List<string> repetidos = new List<string>();
                    //novo ciclo pelos comprimentos em que a barra se dividiu
                    foreach (string novamedida in item.ItemArray[3].ToString().Split(','))
                    {

                        //se o numero for par a cor da celula altera
                        if (contador1 % 2 != 0)
                        {

                            ((Excel.Range)sheet.Cells[contador, 1]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                            ((Excel.Range)sheet.Cells[contador, 2]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                            ((Excel.Range)sheet.Cells[contador, 3]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                            ((Excel.Range)sheet.Cells[contador, 5]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                            ((Excel.Range)sheet.Cells[contador, 4]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                            ((Excel.Range)sheet.Cells[contador, 6]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                            ((Excel.Range)sheet.Cells[contador, 7]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);

                        }
                        else
                        {
                            ((Excel.Range)sheet.Cells[contador, 1]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                            ((Excel.Range)sheet.Cells[contador, 2]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                            ((Excel.Range)sheet.Cells[contador, 3]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                            ((Excel.Range)sheet.Cells[contador, 5]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                            ((Excel.Range)sheet.Cells[contador, 4]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                            ((Excel.Range)sheet.Cells[contador, 6]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                            ((Excel.Range)sheet.Cells[contador, 7]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                        }

                        //prenche a o comprimento da peca directo com o "item"
                        ((Excel.Range)sheet.Cells[contador, 5]).Value = novamedida;

                        //prenche a referencia da peça com o item split o novo contador é para defenir qual a peça a preencher
                        ((Excel.Range)sheet.Cells[contador, 4]).Value = item.ItemArray[5].ToString().Split('#')[novocontador].Replace("\"", "");

                        //Neste momento linha tipo ipe160,12000,1,1811151.15cj1.15c1,3000,    ,44 ou dependendo do contador    ,    ,    ,1811151.15cj1.15c1,3000,    ,    


                        //saber quantas peças tem iguais o soma é o total de peças
                        int soma = 0;
                        string mysr = item.ItemArray[5].ToString().Split('#')[novocontador].Replace(" ", "");

                        foreach (string con in item.ItemArray[5].ToString().Split('#'))
                        {
                            if (mysr == con.Replace(" ", ""))
                            {
                                soma = soma + 1;
                            }
                        }
                        //prenche o total de peças iguais
                        ((Excel.Range)sheet.Cells[contador, 6]).Value = soma.ToString();


                        //Neste momento linha tipo ipe160,12000,1,1811151.15cj1.15c1,3000,2,44 ou dependendo do contador    ,     ,    ,1811151.15cj1.15c1,3000,2,    


                        novocontador++;
                        //se tiver repetidos apaga toda a linha implicito esta que na proxima sobrepõem 
                        if (!repetidos.Any(str => str.Contains(mysr)))
                        {
                            repetidos.Add(mysr);
                            contador += 1;
                        }
                        else
                        {
                            ((Excel.Range)sheet.Cells[contador, 1]).Value = "";
                            ((Excel.Range)sheet.Cells[contador, 2]).Value = "";
                            ((Excel.Range)sheet.Cells[contador, 3]).Value = "";
                            ((Excel.Range)sheet.Cells[contador, 5]).Value = "";
                            ((Excel.Range)sheet.Cells[contador, 4]).Value = "";
                            ((Excel.Range)sheet.Cells[contador, 6]).Value = "";
                            ((Excel.Range)sheet.Cells[contador, 7]).Value = "";
                        }
                    }
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                workbook.CheckCompatibility = false;
                workbook.SaveAs(outputPath);
                workbook.Close(Type.Missing, Type.Missing, Type.Missing);
                excel.Quit();
                Marshal.ReleaseComObject(sheet);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excel);
            }
            else
            {
                MessageBox.Show(this, "ERRO AO CRIAR O PERFIL - " + Perfil);
            }
        }
        private void aTRIBUIRSTOKAUTOMATICOToolStripMenuItem_Click(object sender, EventArgs e)
        {
          
            foreach (string perfil in BuscaPerfisExixtententes())
            {
                if (perfil.ToLower().Contains("heb") || perfil.ToLower().Contains("ipe") || perfil.ToLower().Contains("upn") || perfil.ToLower().Contains("ipn") || perfil.ToLower().Contains("hem") || perfil.ToLower().Contains("upe") || perfil.ToLower().Contains("hea"))
                {
                    dataGridView2.Rows.Add(100, 6100,  0,perfil);
                    dataGridView2.Rows.Add(100, 10100, 0,perfil);
                    dataGridView2.Rows.Add(100, 12100, 0,perfil);
                    dataGridView2.Rows.Add(100, 14100, 0,perfil);
                    dataGridView2.Rows.Add(100, 15100, 0,perfil);
                    dataGridView2.Rows.Add(100, 16100, 0,perfil);                                   
                }                                    
                else                                 
                {                                                                
                    dataGridView2.Rows.Add(100, 6000,  0,perfil);
                    dataGridView2.Rows.Add(100,12000,  0,perfil);
                }
            }
        }

        /// <summary>
        /// classe de nesting propriamente dita pecas(qtd,comp,ref) ||-||-|| barras(qtd comp prior,refer)
        /// </summary>
        /// 
        private static List<propriedadedebarra> PossibleLengths = new List<propriedadedebarra>();
        private static ListaBarraPerfil RELATORIO;
        private string Percentagemdeaproveitamento = null;
        public string Percentagemdedesperdicio = null;
        private string Percentagemsemsucata = null;
        private string Percentagemdedesperdiciosemsucata = null;
        private static decimal ESPESSURASERRA;
        int X = 0;
        public DataTable otimizar(DataTable peças, DataTable barras, decimal _ESPESSURASERRA = 0)
        {
            ESPESSURASERRA = _ESPESSURASERRA;
            PossibleLengths.Clear();
            DataTable resultado = new DataTable();
            resultado.Columns.Add("ID");
            resultado.Columns.Add("barra");
            resultado.Columns.Add("tamanho da barra");
            resultado.Columns.Add("comprimento a cortar", typeof(List<float>));
            resultado.Columns.Add("desperdicio");
            resultado.Columns.Add("marca de peça", typeof(List<string>));

            barras.DefaultView.Sort = barras.Columns[2].ColumnName + " ASC";
            //carregar os cortes que queremos 
            List<Part> DesiredLength = new List<Part>();
            try
            {
                for (int i = 0; i < peças.Rows.Count; i++)
                {
                    int a = Convert.ToInt32(peças.Rows[i].ItemArray[0].ToString());
                    while (a > 0)
                    {

                        DesiredLength.Add(new Part(float.Parse(peças.Rows[i].ItemArray[1].ToString()), peças.Rows[i].ItemArray[2].ToString()));
                        a = a - 1;
                    }
                }
            }
            catch (Exception) { }

            List<ListaBarraPerfil> aproveitamentos = new List<ListaBarraPerfil>();

            //preparar as otimizaçoes
            int prioridades = 0;

            foreach (DataRow item in barras.Rows)
            {
                try
                {
                    int a = 0;
                    int.TryParse(item.ItemArray[2].ToString(), out a);
                    prioridades += a;
                }
                catch (Exception)
                {
                }
            }
            int teste = 2;
            if (prioridades == 0)
            {
                teste = 6;
            }

            for (int i = 0; i < teste; i++)
            {
                //carregar as barras a ser cortadas 
                try
                {
                    for (int a = 0; a < barras.Rows.Count; a++)
                    {
                        int b = Convert.ToInt32(barras.Rows[a].ItemArray[0].ToString());

                        while (b > 0)
                        {
                            PossibleLengths.Add(new propriedadedebarra { Column1 = float.Parse(barras.Rows[a].ItemArray[1].ToString()), Column2 = barras.Rows[a].ItemArray[3].ToString() });
                            b = b - 1;
                        }
                    }
                }
                catch (Exception)
                {

                }

                //enviar os diferentes configuraçoes das otimizaçoes 

                List<Part> DesiredLengths = null;
                if (i == 0)
                {
                    DesiredLengths = new List<Part>(DesiredLength.OrderBy(x => x.OriginalLength));
                }
                else if (i == 1)
                {
                    DesiredLengths = new List<Part>(DesiredLength.OrderBy(x => x.OriginalLength).Reverse());
                }
                else if (i == 2)
                {
                    PossibleLengths = new List<propriedadedebarra>(PossibleLengths.OrderBy(x => x.Column1));
                    DesiredLengths = new List<Part>(DesiredLength.OrderBy(x => x.OriginalLength));
                }
                else if (i == 3)
                {
                    PossibleLengths = new List<propriedadedebarra>(PossibleLengths.OrderBy(x => x.Column1).Reverse());
                    DesiredLengths = new List<Part>(DesiredLength.OrderBy(x => x.OriginalLength));
                }
                else if (i == 4)
                {
                    PossibleLengths = new List<propriedadedebarra>(PossibleLengths.OrderBy(x => x.Column1));
                    DesiredLengths = new List<Part>(DesiredLength.OrderBy(x => x.OriginalLength).Reverse());
                    DesiredLengths = new List<Part>(DesiredLength.OrderBy(x => x.OriginalLength).Reverse());
                    DesiredLengths = new List<Part>(DesiredLength.OrderBy(x => x.OriginalLength).Reverse());
                }
                else if (i == 5)
                {
                    PossibleLengths = new List<propriedadedebarra>(PossibleLengths.OrderBy(x => x.Column1).Reverse());
                    DesiredLengths = new List<Part>(DesiredLength.OrderBy(x => x.OriginalLength).Reverse());
                }

                //curtar as peças retorna as barras com cortes.
                var BarraPerfils = CalculateCuts(DesiredLengths);
                //armazenar o nest para fazer novo aproveitamento.
                if (BarraPerfils != null)
                {
                    float somapercentagem = 0;
                    float somapercentagemsucata = 0;
                    float perfi = 0;
                    float desperdicio = 0;
                    float desperdicioSucata = 0;
                    foreach (var BarraPerfil in BarraPerfils)
                    {
                        try
                        {
                            perfi += BarraPerfil.OriginalLength;
                            desperdicio += BarraPerfil.FreeLength;
                            if (BarraPerfil.FreeLength <= 2000f)
                            {
                                desperdicioSucata += BarraPerfil.FreeLength;
                            }

                        }
                        catch (Exception)
                        {
                        }

                    }

                    if (BarraPerfils.Count != 0)
                    {
                        somapercentagem = (desperdicio / perfi) * 100;
                        somapercentagemsucata = (desperdicioSucata / perfi) * 100;
                        aproveitamentos.Add(new ListaBarraPerfil(BarraPerfils.OrderBy(x => x.FreeLength).ToList(), somapercentagem, BarraPerfils.Last().FreeLength, somapercentagemsucata));
                    }
                }
            }

            try
            {

                var min = aproveitamentos.Min(t => t.Percentagem);
                List<ListaBarraPerfil> conjperfil = new List<ListaBarraPerfil>();
                foreach (ListaBarraPerfil ite in aproveitamentos.OrderBy(t => t.Percentagem).ToList())
                {
                    if (ite.Percentagem == min)
                    {
                        conjperfil.Add(ite);
                    }
                }
                var perfis = conjperfil.OrderBy(t => t.ultimaponta).Reverse().First();
                

                foreach (var BarraPerfil in perfis)
                {
                    X++;
                    resultado.Rows.Add(X, BarraPerfil.BarraPerfilMark, BarraPerfil.OriginalLength,BarraPerfil.Cuts, BarraPerfil.FreeLength,  BarraPerfil.Cutpartmark);
                    Percentagemdeaproveitamento = (100 - (perfis.Percentagem)).ToString("0.00") + " % ";
                    Percentagemdedesperdicio = perfis.Percentagem.ToString("0.00") + "%";
                    Percentagemsemsucata = (100 - (perfis.PercentagemSucata)).ToString("0.00") + " % ";
                    Percentagemdedesperdiciosemsucata = (perfis.PercentagemSucata).ToString("0.00") + " % ";
                }
                RELATORIO = perfis;
            }
            catch (Exception)
            {
                
                return null;
            }
            return resultado;
        }
 
        private static List<BarraPerfil> CalculateCuts(List<Part> desired)
        {
            var BarraPerfils = new List<BarraPerfil>(); //Buffer list

            //passar por cortes
            foreach (Part i in desired)
            {
                bool repeat = true;
                while (repeat)
                {
                    //se não forem encontradas pranchas com comprimento disponivel
                    if (!BarraPerfils.Any(BarraPerfil => BarraPerfil.FreeLength >= i.OriginalLength))
                    {
                        //fazer a prancha
                        try
                        {

                            float comp = PossibleLengths.First().Column1;
                            BarraPerfils.Add(new BarraPerfil(comp, PossibleLengths.First().Column2));

                            bool primeirapecaencontrada = true;
                            for (int a = 0; a <= PossibleLengths.Count - 1; a++)
                            {
                                if (PossibleLengths[a].Column1 == comp && primeirapecaencontrada)
                                {
                                    PossibleLengths.RemoveAt(a);
                                    primeirapecaencontrada = false;
                                }
                            }
                        }
                        catch (Exception)
                        {
                            return null;
                        }
                    }

                    //cortar quando possivel

                    foreach (var BarraPerfil in BarraPerfils.Where(BarraPerfil => BarraPerfil.FreeLength >= (i.OriginalLength + float.Parse(ESPESSURASERRA.ToString()))))
                    {
                        BarraPerfil.Cut.Add(i.OriginalLength + float.Parse(ESPESSURASERRA.ToString()));
                        BarraPerfil.Cutpartmarks.Add(i.Mark);
                        repeat = false;
                        break;
                    }
                    if (repeat)
                    {
                        BarraPerfils.RemoveAt(BarraPerfils.Count - 1);
                    }
                }
            }

            //reduzir os despedicio minimizando o comprimento da prancha
            foreach (var BarraPerfil in BarraPerfils)
            {
                float newLength = BarraPerfil.OriginalLength;
                foreach (propriedadedebarra possibleLength in PossibleLengths)
                {
                    //possibleLength <= BarraPerfil.OriginalLength && (BarraPerfil.OriginalLength - float.Parse(BarraPerfil.FreeLength.ToString())) <= possibleLength

                    if (possibleLength.Column1 <= BarraPerfil.OriginalLength && (BarraPerfil.OriginalLength - float.Parse(BarraPerfil.FreeLength.ToString())) <= possibleLength.Column1)
                    {
                        newLength = possibleLength.Column1;

                    }
                }
                BarraPerfil.OriginalLength = newLength;
            }
            PossibleLengths.Clear();
            return BarraPerfils;
        }

        private void FrmNest_Load(object sender, EventArgs e)
        {
            this.reportViewer1.RefreshReport();
            this.reportViewer1.RefreshReport();
        }


        private void eXCELToolStripMenuItem_Click(object sender, EventArgs e)
        {
            criaexcel(DataSet1.Tables[0], @"c:\r\teste.xls", "");
        }
    }

    class PecaCorte
    {
        //atributos do objeto
        public int Quantidade=0;
        public double Comprimento;
        public string Referencia;
        public string Perfil;
    }
    class Part
    {
        public Part(float length, string mark)
        {
            OriginalLength = length;
            Mark = mark;
        }

        public float OriginalLength;
        public string Mark;
    }
    class ListaBarraPerfil
    {
        public List<BarraPerfil> perfis = new List<BarraPerfil>();
        public float Percentagem;
        public float ultimaponta;
        public float PercentagemSucata;
        public ListaBarraPerfil(List<BarraPerfil> BarraPerfil, float percentagem, float comprimento, float Percentagemsucata)
        {
            perfis = BarraPerfil;
            Percentagem = percentagem;
            ultimaponta = comprimento;
            PercentagemSucata = Percentagemsucata;

        }
        public List<BarraPerfil> ListaBarraPerfils
        {
            get { return perfis; }
        }
        public float PercentagemDesperdicio
        {
            get { return Percentagem; }
        }
        public float Percentagemsucata
        {
            get { return PercentagemSucata; }
        }
        public float Ultimaponta
        {
            get { return ultimaponta; }
        }
        private IEnumerable<BarraPerfil> Events()
        {
            foreach (var item in perfis)
            {
                yield return item;
            }
        }
        public IEnumerator<BarraPerfil> GetEnumerator()
        {
            return Events().GetEnumerator();
        }
    }
    class BarraPerfil
    {
        public BarraPerfil(float length, string Mark)
        {
            OriginalLength = length;
            BarraPerfilMark = Mark;
        }

        public float FreeLength
        {
            get { return OriginalLength - Cuts.Sum(); }
        }

        public float OriginalLength;
        public string BarraPerfilMark;


        public List<float> Cuts = new List<float>();
        public List<string> Cutpartmark = new List<string>();
        public List<float> Cut
        {
            set { Cuts = value; }
            get { return Cuts; }
        }
        public List<string> Cutpartmarks
        {
            set { Cutpartmark = value; }
            get { return Cutpartmark; }
        }
    }
    public class propriedadedebarra
    {
        // obviously you find meaningful names of the 2 properties
        public float columnsfloat;
        public string columnsstring;

        public float Column1 { get; set; }
        public string Column2 { get; set; }
    }
}

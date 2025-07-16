using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tekla.Structures.Model;
using TSM = Tekla.Structures.Model;

//  delete  FROM [PRIOFELIZ].[dbo].[TDU_INLinhasQuantificacaoCM]
//  delete  FROM [PRIOFELIZ].[dbo].[TDU_INCabecQuantificacaoCM] 

namespace TeklaArtigosOfeliz
{
    public partial class FrmQuantificacao_new : Form
    {

        string subempreiteiro = null;
        Frm_Inico formpai = null;

        public FrmQuantificacao_new(Frm_Inico form)
        {
            InitializeComponent();
            formpai = form;
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
            LBLQuantificação.Text = (ProximaQuantificação()).ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                dataGridView1.DataSource = null;
                dataGridView1.Rows.Clear();
            }
            catch (Exception) { }
            ArrayList PECASAQUANTIFICAR = new ArrayList();
            DataTable LISTA = CreateColumns();
            PECASAQUANTIFICAR = PECASPARAQUANTIFICAR(ComunicaTekla.ListadePecasSelec());
            LISTA = BuscaDados(LISTA, PECASAQUANTIFICAR);
            dataGridView1.DataSource = LISTA.DefaultView;
            CorrigeArtigo();
            ValidaArtigo();
        }

        private int ProximaQuantificação()
        {
            int PROXIMAQUANTIFICAÇAO = 0;
            ComunicaBDprimavera A = new ComunicaBDprimavera();
            A.ConectarBD() ;
            List<string> N = A.Procurarbd("SELECT [CDU_IDTekla] FROM[PRIOFELIZ].[dbo].[TDU_INCabecQuantificacaoCM] WHERE CDU_CodigoObra='"+ formpai.label11.Text+ "' UNION ALL  SELECT[CDU_IDTekla] FROM [PRIOFELIZ].[dbo].[TDU_INCabecQuantificacaoCMHistorico] where CDU_CodigoObra = '"+ formpai.label11.Text+"'");
            A.DesonectarBD();

            try
            {
                PROXIMAQUANTIFICAÇAO = N.Max(item => int.Parse(item));
            }
            catch (Exception)
            {
                PROXIMAQUANTIFICAÇAO = 0;
            }
           
            PROXIMAQUANTIFICAÇAO++;
            return PROXIMAQUANTIFICAÇAO;
        }

        private bool validaimportaçao()
        {
            bool valida = true;
            ComunicaBDprimavera A = new ComunicaBDprimavera();
            A.ConectarBD();
            List<string> N = A.Procurarbd("SELECT [CDU_IDTekla] FROM [PRIOFELIZ].[dbo].[TDU_INCabecQuantificacaoCM] WHERE " +
                "CDU_CodigoObra='" + formpai.label11.Text + 
                "' and " +
                "CDU_IDTekla='"+LBLQuantificação.Text+
                "' UNION ALL  " +
                "SELECT [CDU_IDTekla] FROM [PRIOFELIZ].[dbo].[TDU_INCabecQuantificacaoCMHistorico]" +
                " where CDU_CodigoObra = '" + formpai.label11.Text +
                   "' and " +
                "CDU_IDTekla='" + LBLQuantificação.Text +
                "'");
            A.DesonectarBD();

            if (N.Count!=0)
            {
                valida = false;
            }
            return valida;
        }
       
        private void button2_Click(object sender, EventArgs e)
        {

            if (validaimportaçao())
            {
                DialogResult d = MessageBox.Show(this, "Deseja inportar a quantificação?", "Importação", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (d == DialogResult.Yes)
                {
                    LBLQuantificação.Text = (ProximaQuantificação()).ToString();
                    if (PECASENVIAPRIMAVERA())
                    {
                        PECASENVIATEKLA();
                        MessageBox.Show(this, "Concluido", "Importação", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            else
            {
                MessageBox.Show(this, "Já importado", "Importação", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
   

        }

        private DataTable CreateColumns()
        {
            DataTable LISTA = new DataTable();
            LISTA.Columns.Add("Artigo");
            LISTA.Columns.Add("Classe");
            LISTA.Columns.Add("Certificado");
            LISTA.Columns.Add("Quantidade");
            LISTA.Columns.Add("Peso");
            LISTA.Columns.Add("RequesitosEspeciais");
            LISTA.Columns.Add("Tolerancias");
            LISTA.Columns.Add("EstadoSuperficie");
            LISTA.Columns.Add("PropriedadesEspeciais");
            LISTA.Columns.Add("Comprimento");
            LISTA.Columns.Add("Altura");
            LISTA.Columns.Add("Largura");
            LISTA.Columns.Add("AnguloLadoA");
            LISTA.Columns.Add("AnguloLadoB");
            LISTA.Columns.Add("AnguloBLadoA");
            LISTA.Columns.Add("AnguloBLadoB");
            LISTA.Columns.Add("PECAS", typeof(List<TSM.Part>));
            LISTA.Rows.Clear();
            return LISTA;
        }

        private DataTable BuscaDados(DataTable LISTA,ArrayList PECASAQUANTIFICAR)
        {
            foreach (Tekla.Structures.Model.Part item in PECASAQUANTIFICAR)
            {

                string Artigo = item.Profile.ProfileString;
                string Classe = item.Material.MaterialString.Replace(".3,1", "").Replace(".2,1", "").Replace(".2,2", "").Replace(".3,2", "");
                string Certificado = item.Material.MaterialString;
                if (Certificado.Contains("3,1"))
                {
                    Certificado = "3,1";
                }
                else if (Certificado.Contains("2,1"))
                {
                    Certificado = "2,1";
                }
                else if (Certificado.Contains("2,2"))
                {
                    Certificado = "2,2";
                }
                else if (Certificado.Contains("3,2"))
                {
                    Certificado = "3,2";
                }

                int Quantidade = 1;

                double Peso = 0; item.GetReportProperty("PROFILE_WEIGHT_NET", ref Peso);
                double Comprimento = 0; item.GetReportProperty("LENGTH", ref Comprimento);
                double Altura = 0; item.GetReportProperty("HEIGHT", ref Altura);
                double Largura = 0; item.GetReportProperty("WIDTH", ref Largura);
                double AnguloBLadoA = 0; item.GetReportProperty("END1_CUT_ANGLE_Y", ref AnguloBLadoA);
                double AnguloBLadoB = 0; item.GetReportProperty("END2_CUT_ANGLE_Y", ref AnguloBLadoB);
                double AnguloLadoA = 0; item.GetReportProperty("END1_CUT_ANGLE_Z", ref AnguloLadoA);
                double AnguloLadoB = 0; item.GetReportProperty("END2_CUT_ANGLE_Z", ref AnguloLadoB);
                string RequesitosEspeciais = ""; item.GetReportProperty("USERDEFINED.Requisitos", ref RequesitosEspeciais);

                string Tolerancias = "";
                string estadoSuperficie = "";
                string propriedadesEspeciais = "";
                if (Classe.Contains("S235") || Classe.Contains("S275") || Classe.Contains("S355"))
                {
                    item.GetReportProperty("PROJECT.USERDEFINED.MTTolerancias", ref Tolerancias);
                    item.GetReportProperty("PROJECT.USERDEFINED.MTTolerancias", ref estadoSuperficie);
                    item.GetReportProperty("PROJECT.USERDEFINED.MTPropespperfis", ref propriedadesEspeciais);
                    if (string.IsNullOrEmpty(Tolerancias))
                    {
                        Tolerancias = "EN10029 Classe A";
                    }
                    if (string.IsNullOrEmpty(estadoSuperficie))
                    {
                        estadoSuperficie = "EN10163-3 Classe C1";
                    }
                    if (((Artigo.Contains("PL") || Artigo.Contains("CHA")) && item.Name.ToUpper().Contains("BR")) || Artigo.Contains("BR"))
                    {
                        item.GetReportProperty("PROJECT.USERDEFINED.MTPropespbarra", ref propriedadesEspeciais);
                    }
                }

                string IDCabecQuantificacaoCM = ""; item.GetUserProperty("QUANTIFICACAO", ref IDCabecQuantificacaoCM);
                #region
                //ModelObjectEnumerator P = item.GetBooleans();

                //List<CutPlane> cuts = new List<CutPlane>();

                //while (P.MoveNext())
                //{
                //    CutPlane cutPlane = P.Current as CutPlane;

                //    if (cutPlane != null)
                //    {
                //        cuts.Add(cutPlane);
                //    }

                //}




                //Beam b = new Beam();
                //b.Profile.ProfileString = item.Profile.ProfileString;
                //b.StartPoint = (item as Beam).StartPoint;
                //b.EndPoint = (item as Beam).EndPoint;
                //b.Position = item.Position;
                //b.Insert();
                //List<CutPlane> cuts1 = new List<CutPlane>();
                //if (cuts.Count != 0)
                //{
                //    if (AnguloLadoA != 0)
                //    {
                //        cuts1.Add(cuts.AsEnumerable().Where(r => r.Plane.Origin.X == cuts.AsEnumerable().Min(w => w.Plane.Origin.X)).ToList().First());
                //    }
                //    if (AnguloLadoB != 0)
                //    {
                //        cuts1.Add(cuts.AsEnumerable().Where(r => r.Plane.Origin.X == cuts.AsEnumerable().Max(w => w.Plane.Origin.X)).ToList().First());
                //    }
                //}



                //foreach (CutPlane cutPlane in cuts1)
                //{
                //    Fitting plane = new Fitting();
                //    plane.Plane = cutPlane.Plane;
                //    plane.Father = (b as Beam);
                //    plane.Insert();
                //    b.GetReportProperty("LENGTH", ref Comprimento);
                //    b.GetReportProperty("HEIGHT", ref Altura);
                //    b.GetReportProperty("WIDTH", ref Largura);
                //    b.GetReportProperty("END1_CUT_ANGLE_Y", ref AnguloBLadoA);
                //    b.GetReportProperty("END2_CUT_ANGLE_Y", ref AnguloBLadoB);
                //    b.GetReportProperty("END1_CUT_ANGLE_Z", ref AnguloLadoA);
                //    b.GetReportProperty("END2_CUT_ANGLE_Z", ref AnguloLadoB);
                //}
                //b.Delete();
                //Model model = new Model();
                //model.CommitChanges();
                #endregion erro tentativa falhada

                DataRow DATA = LISTA.AsEnumerable().FirstOrDefault(row =>
                row.Field<string>("Artigo") == Artigo &&
                row.Field<string>("Classe") == Classe &&
                row.Field<string>("Certificado") == Certificado &&
                row.Field<string>("RequesitosEspeciais") == RequesitosEspeciais &&
                row.Field<string>("Tolerancias") == Tolerancias &&
                row.Field<string>("EstadoSuperficie") == estadoSuperficie &&
                row.Field<string>("PropriedadesEspeciais") == propriedadesEspeciais &&
                row.Field<string>("Comprimento") == Comprimento.ToString("0.0", System.Threading.Thread.CurrentThread.CurrentCulture) &&
                row.Field<string>("AnguloLadoA") == AnguloLadoA.ToString("0.0", System.Threading.Thread.CurrentThread.CurrentCulture) &&
                row.Field<string>("AnguloLadoB") == AnguloLadoB.ToString("0.0", System.Threading.Thread.CurrentThread.CurrentCulture) &&
                row.Field<string>("AnguloBLadoA") == AnguloBLadoA.ToString("0.0", System.Threading.Thread.CurrentThread.CurrentCulture) &&
                row.Field<string>("AnguloBLadoB") == AnguloBLadoB.ToString("0.0", System.Threading.Thread.CurrentThread.CurrentCulture));

                if (DATA != null)
                {
                    float A = float.Parse(DATA.Field<string>("PESO"), System.Threading.Thread.CurrentThread.CurrentCulture);
                    float B = float.Parse(Peso.ToString("0.00"), System.Threading.Thread.CurrentThread.CurrentCulture);
                    float C = float.Parse((A + B).ToString("0.0"), System.Threading.Thread.CurrentThread.CurrentCulture);
                    DATA.SetField<float>("PESO", C);
                    DATA.SetField<int>("Quantidade", int.Parse(DATA.Field<string>("Quantidade"), CultureInfo.InvariantCulture) + Quantidade);
                    List<TSM.Part> pt = new List<TSM.Part>();
                    foreach (TSM.Part it in DATA.Field<List<TSM.Part>>(16))
                    {
                        pt.Add(it);
                    }
                    pt.Add(item);

                    DATA.SetField<List<TSM.Part>>(16, pt);
                }
                else
                {
                    List<TSM.Part> pt = new List<TSM.Part>();
                    pt.Add(item);

                    LISTA.Rows.Add(Artigo, Classe, Certificado, Quantidade, Peso.ToString("0.0", System.Threading.Thread.CurrentThread.CurrentCulture), RequesitosEspeciais, Tolerancias, estadoSuperficie, propriedadesEspeciais, Comprimento.ToString("0.0", System.Threading.Thread.CurrentThread.CurrentCulture), Altura.ToString("0.0", System.Threading.Thread.CurrentThread.CurrentCulture), Largura.ToString("0.0", System.Threading.Thread.CurrentThread.CurrentCulture), AnguloLadoA.ToString("0.0", System.Threading.Thread.CurrentThread.CurrentCulture), AnguloLadoB.ToString("0.0", System.Threading.Thread.CurrentThread.CurrentCulture), AnguloBLadoA.ToString("0.0", System.Threading.Thread.CurrentThread.CurrentCulture), AnguloBLadoB.ToString("0.0", System.Threading.Thread.CurrentThread.CurrentCulture), pt);
                }

            }
            return LISTA;
        }

        private bool PECASENVIAPRIMAVERA()
        {

            List<string> dbOperations = new List<string>();
            dataGridView1.ReadOnly = true;
            ComunicaBDprimavera a = new ComunicaBDprimavera();
            Model m = new Model();
            string CDU_IDcabec = Guid.NewGuid().ToString();
            string CDU_IDTekla = LBLQuantificação.Text;
            string CDU_CodigoObra = formpai.label11.Text;
            string CDU_DescricaoObra = m.GetProjectInfo().Builder;

            a.ConectarBD();

            string CDU_CodigoCliente = a.Procurarbd("SELECT [ERPEntidadeA] FROM [dbo].[MT_View_Obras_Clientes_Descricao] where [Codigo]='" 
                + formpai.label11.Text + "'")[0];
            string CDU_LocalDescarga = a.Procurarbd("SELECT [CDU_MTLocal] FROM [PRIOFELIZ].[dbo].[MT_View_MTLocalDescarga] where [CDU_MTSubempreiteiro]='" 
                + localdedescagacb.Text + "'")[0];
            subempreiteiro = CDU_LocalDescarga;
            a.DesonectarBD();

            string CDU_NomeCliente = m.GetProjectInfo().Name;
            DateTime CDU_DataInicioProducao = dateTimePicker1.Value;
            DateTime CDU_DataCriacao = DateTime.Now;
 



            dbOperations.Add("INSERT INTO[dbo].[TDU_INCabecQuantificacaoCM] ([CDU_ID],[CDU_IDTekla],[CDU_CodigoObra],[CDU_DescricaoObra],[CDU_CodigoCliente],[CDU_NomeCliente],[CDU_DataInicioProducao],[CDU_DataCriacao],[CDU_LocalDescarga]) VALUES ('"
                                    + CDU_IDcabec
                                    + "','" + CDU_IDTekla            
                                    + "','" + CDU_CodigoObra         
                                    + "','" + CDU_DescricaoObra      
                                    + "','" + CDU_CodigoCliente      
                                    + "','" + CDU_NomeCliente        
                                    + "','" + CDU_DataInicioProducao 
                                    + "','" + CDU_DataCriacao        
                                    + "','" + CDU_LocalDescarga      
                                    + "')");

            ;

            foreach (DataGridViewRow item in dataGridView1.Rows)
            {
              
                string CDU_ID = Guid.NewGuid().ToString();
                string CDU_Artigo = (item.Cells[0].Value+"").Replace(",",".");
                string CDU_Classe = (item.Cells[1].Value + "").Replace(",",".") ;
                string CDU_Certificado = (item.Cells[2].Value + "").Replace(",", ".");
                string CDU_Quantidde = (item.Cells[3].Value + "").Replace(",", ".");
                string CDU_Peso = (item.Cells[4].Value + "").Replace(",", ".");
                string CDU_RequesitosEspeciais = (item.Cells[5].Value + "").Replace(",", ".");
                string CDU_Tolerancias = (item.Cells[6].Value + "").Replace(",", ".");
                string CDU_estadoSuperficie = (item.Cells[7].Value + "").Replace(",", ".");
                string CDU_propriedadesEspeciais = (item.Cells[8].Value + "").Replace(",", ".");
                string CDU_Comprimento = (item.Cells[9].Value + "").Replace(",", ".");
                string CDU_Altura = (item.Cells[10].Value + "").Replace(",", ".");
                string CDU_Largura = (item.Cells[11].Value + "").Replace(",", ".");
                string CDU_AnguloLadoA = (item.Cells[12].Value + "").Replace(",", ".");
                string CDU_AnguloLadoB = (item.Cells[13].Value + "").Replace(",", ".");
                string CDU_AnguloBLadoA = (item.Cells[14].Value + "").Replace(",", ".");
                string CDU_AnguloBLadoB = (item.Cells[15].Value + "").Replace(",", ".");
                string CDU_IDCabecQuantificacaoCM = CDU_IDcabec+"";

                if (CDU_Comprimento!="")
                {
                    dbOperations.Add(@"INSERT INTO [dbo].[TDU_INLinhasQuantificacaoCM]([CDU_ID],[CDU_Artigo],[CDU_Classe],[CDU_Certificado],[CDU_Quantidde],[CDU_Peso] ,[CDU_RequesitosEspeciais],[CDU_Tolerancias],[CDU_estadoSuperficie],[CDU_propriedadesEspeciais],[CDU_Comprimento],[CDU_Altura],[CDU_Largura],[CDU_AnguloLadoA],[CDU_AnguloLadoB],[CDU_AnguloBLadoA],[CDU_AnguloBLadoB],[CDU_IDCabecQuantificacaoCM]) VALUES ('"
                                    + CDU_ID
                                    + "','" + CDU_Artigo
                                    + "','" + CDU_Classe
                                    + "','" + CDU_Certificado
                                    + "','" + CDU_Quantidde
                                    + "','" + CDU_Peso
                                    + "','" + CDU_RequesitosEspeciais
                                    + "','" + CDU_Tolerancias
                                    + "','" + CDU_estadoSuperficie
                                    + "','" + CDU_propriedadesEspeciais
                                    + "','" + CDU_Comprimento
                                    + "','" + CDU_Altura
                                    + "','" + CDU_Largura
                                    + "','" + CDU_AnguloLadoA
                                    + "','" + CDU_AnguloLadoB
                                    + "','" + CDU_AnguloBLadoA
                                    + "','" + CDU_AnguloBLadoB
                                    + "','" + CDU_IDCabecQuantificacaoCM
                                    + "')");
                }
            }

           return a.dbOperations(dbOperations);

        }

        private void PECASENVIATEKLA()
        {
            object teste = dataGridView1.DataSource;
            DataView DataVew = teste as DataView;
            DataTable tg = DataVew.Table.DefaultView.ToTable();
            foreach (DataRow item in tg.Rows)
            {
                foreach (TSM.Part it in item.Field<List<TSM.Part>>("PECAS"))
                {
                    ComunicaTekla.EnviaproPriedadePeca(it, "QUANTIFICACAO", int.Parse(LBLQuantificação.Text).ToString("000") + "-" + dateTimePicker1.Value.ToShortDateString() + "-" + subempreiteiro);
                }
            }
        }

        private ArrayList PECASPARAQUANTIFICAR(ArrayList lista)
        {
            ArrayList PECASAQUANTIFICAR = new ArrayList();
            foreach (TSM.Part part in lista)
            {
                string quant = null;
                part.GetUserProperty("QUANTIFICACAO", ref quant);

                if (string.IsNullOrEmpty(quant))
                {
                    if (part.Name.Contains("BR") || part.Name.Contains("BQ") || part.Profile.ProfileString.Contains("CF") || part.Profile.ProfileString.Contains("HF") || part.Profile.ProfileString.Contains("TGCHS"))
                    {

                        PECASAQUANTIFICAR.Add(part);
                    }
                    else if (part.Profile.ProfileString.Contains("CH") || part.Profile.ProfileString.Contains("CG") || part.Profile.ProfileString.Contains("VRSM") || part.Profile.ProfileString.Contains("NM") || part.Profile.ProfileString.Contains("WM") || part.Profile.ProfileString.Contains("C1") || part.Profile.ProfileString.Contains("C2") || part.Profile.ProfileString.Contains("C3") || part.Profile.ProfileString.Contains("Z") || part.Profile.ProfileString.Contains("PERNO") || part.Profile.ProfileString.Contains("H60") || part.Profile.ProfileString.Contains("PL") || part.Profile.ProfileString.Contains("MAX") || part.Profile.ProfileString.Contains("SUPEROMEGA") || part.Profile.ProfileString.Contains("CONE"))
                    {

                    }
                    else if (part.Material.MaterialString.Contains("C45E") || part.Material.MaterialString.Contains("CK45"))
                    {
                        PECASAQUANTIFICAR.Add(part);
                    }
                    else if (part.Material.MaterialString.Contains("C") || part.Material.MaterialString.Contains("NEO") || part.Material.MaterialString.Contains("TEF"))
                    {

                    }
                    else
                    {

                        PECASAQUANTIFICAR.Add(part);
                    }
                }
            }
            return PECASAQUANTIFICAR;
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            object teste = dataGridView1.DataSource;
            DataView DataVew = teste as DataView;
            ComunicaTekla.selectinmodel(new ArrayList(DataVew.Table.DefaultView.ToTable().Rows[e.RowIndex].Field<List<TSM.Part>>("PECAS")));
        }

        private void ValidaArtigo()
        {
            ComunicaBDprimavera bd = new ComunicaBDprimavera();
            bd.ConectarBD();
            foreach (DataGridViewRow Linha in dataGridView1.Rows)
            {
              List<string> lista = bd.Procurarbd("SELECT [Artigo] FROM [PRIOFELIZ].[dbo].[MT_ViewArtigoUnidades] where [Artigo]='"+ Linha.Cells[0].Value+"'");
                if (lista.Count==0)
                {
                    Linha.DefaultCellStyle.BackColor = Color.Red;
                }
            }
            bd.DesonectarBD();
        }

        private void CorrigeArtigo()
        {

            foreach (DataGridViewRow Linha in dataGridView1.Rows)
            {
                Linha.Cells[0].Value = (Linha.Cells[0].Value + "").Replace(".", ",");
                if ((Linha.Cells[0].Value+"").StartsWith("CFRHS") )
                {
                    Linha.Cells[0].Value = (Linha.Cells[0].Value + "").Replace("HS", "");

                    string[] NovoArtigo = Linha.Cells[0].Value.ToString().Split('X');
                    if (NovoArtigo.Length==2)
                    {
                        Linha.Cells[0].Value = NovoArtigo[0] + "X" + NovoArtigo[0].ToString().Replace("CFR", "") + "X" + NovoArtigo[1];
                    }
                }
                else if ((Linha.Cells[0].Value + "").StartsWith("HFRHS"))
                {
                    Linha.Cells[0].Value = (Linha.Cells[0].Value + "").Replace("HS", "");

                    string[] NovoArtigo = Linha.Cells[0].Value.ToString().Split('X');
                    if (NovoArtigo.Length == 2)
                    {
                        Linha.Cells[0].Value = NovoArtigo[0] + "X" + NovoArtigo[0].ToString().Replace("HFR", "") + "X" + NovoArtigo[1];
                    }
                }
                else if ((Linha.Cells[0].Value + "").StartsWith("TGRHS"))
                {
                    Linha.Cells[0].Value = (Linha.Cells[0].Value + "").Replace("TG", "");

                    string[] NovoArtigo = Linha.Cells[0].Value.ToString().Split('X');
                    if (NovoArtigo.Length == 2)
                    {
                        Linha.Cells[0].Value = NovoArtigo[0] + "X" + NovoArtigo[0].ToString().Replace("TGR", "") + "X" + NovoArtigo[1];
                    }
                }
                else if ((Linha.Cells[0].Value + "").StartsWith("CFCHS"))
                {
                    Linha.Cells[0].Value = (Linha.Cells[0].Value + "").Replace("HS", "");
                }
                else if ((Linha.Cells[0].Value + "").StartsWith("HFCHS"))
                {
                    Linha.Cells[0].Value = (Linha.Cells[0].Value + "").Replace("HS", "");
                }
                else if ((Linha.Cells[0].Value + "").StartsWith("TGCHS"))
                {
                    Linha.Cells[0].Value = (Linha.Cells[0].Value + "").Replace("HS", "");
                }
            }



        }
    }                    
}               
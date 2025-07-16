using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tekla.Structures.Model;

namespace TeklaArtigosOfeliz
{
    public partial class Frm_ListaOFeliz : Form
    {
        Frm_Inico _Formpai;
        public Frm_ListaOFeliz(Frm_Inico formpai)
        {
            InitializeComponent();
            carregadados();
            _Formpai = formpai;
            alteradstv();
            RemoverParafusos();
        }

        private void RemoverParafusos()
        {
            List<DataGridViewRow> rowsToRemove = new List<DataGridViewRow>();

            foreach (DataGridViewRow item in dataGridView1.Rows)
            {
                DataGridViewCell currentcell = item.Cells[9]; // Coluna 9 (contando a partir de 0)

                if (currentcell.Value != null)
                {
                    string cellValue = currentcell.Value.ToString();
                    if (cellValue.Contains("VRSM") || cellValue.Contains("BM") || cellValue.Contains("WM"))
                    {
                        rowsToRemove.Add(item);
                    }
                }
            }

            foreach (DataGridViewRow row in rowsToRemove)
            {
                dataGridView1.Rows.Remove(row);
            }
        }
        private void alteradstv()
        {
            //foreach (DataGridViewRow linha in dataGridView1.Rows)
            //{
            //    if (int.Parse(label5.Text)<3)
            //    {
            //        if (linha.Cells[2].Value!=null)
            //        {
            //            if (linha.Cells[2].Value.ToString() == "Perfil")
            //            {
            //                string FILENAME = linha.Cells[4].Value.ToString().Split('.').Last();
            //                string fileout = linha.Cells[4].Value.ToString().Split('.')[linha.Cells[4].Value.ToString().Split('.').Count() - 2] + "-" + FILENAME;
            //                if (File.Exists(@"C:\R\" + FILENAME + ".nc1"))
            //                {

            //                    string[] ALLFILE = File.ReadAllLines(@"C:\R\" + FILENAME + ".nc1");
            //                    ALLFILE[7] = linha.Cells[8].Value.ToString();
            //                    ALLFILE[3] = fileout;
            //                    ALLFILE[4] = fileout;
            //                    File.WriteAllLines(@"C:\R\" + fileout + ".nc1", ALLFILE, Encoding.Default);

            //                }
            //            }
            //        }
            //    }
            //    else
            //    {
            //        if (linha.Cells[2].Value != null)
            //        {
            //            if (linha.Cells[2].Value.ToString() == "Perfil")
            //            {
            //                string FILENAME = linha.Cells[4].Value.ToString().Split('.').Last();
            //                string fileout = linha.Cells[4].Value.ToString().Split('.')[linha.Cells[4].Value.ToString().Split('.').Count() - 3] + "-" +linha.Cells[4].Value.ToString().Split('.')[linha.Cells[4].Value.ToString().Split('.').Count() - 2] + "-" + FILENAME;
            //                if (File.Exists(@"C:\R\" + FILENAME + ".nc1"))
            //                {
            //                    string[] ALLFILE = File.ReadAllLines(@"C:\R\" + FILENAME + ".nc1");
            //                    ALLFILE[7] = linha.Cells[8].Value.ToString();
            //                    ALLFILE[3] = fileout;
            //                    ALLFILE[4] = fileout;
            //                    File.WriteAllLines(@"C:\R\" + fileout.Replace(".","-") + ".nc1", ALLFILE, Encoding.Default);
            //                }
            //            }
            //        }
            //    }
            //}
        }

        public void carregadados()
        {
            dataGridView1.Rows.Clear();
            string line = null;
            int i = 1;
            StreamReader file = new StreamReader(@"c:\r\OFELIZ.CSV", Encoding.Default, true);
            while ((line = file.ReadLine()) != null)
            {
                if (i == 2)
                {
                    var fields = line.Split(';');
                    label1.Text = fields[1];
                }
                if (i == 3)
                {
                    var fields = line.Split(';');
                    label2.Text = fields[1];
                }
                if (i == 4)
                {
                    var fields = line.Split(';');
                    lbl_numeroobra.Text = fields[1];
                }
                if (i == 5)
                {
                    var fields = line.Split(';');
                    label4.Text = fields[1];
                }
                if (i == 6)
                {
                    var fields = line.Split(';');
                    label5.Text = fields[1];
                }

                if (i > 8)
                {
                    var fields = line.Split(';');
                    dataGridView1.Rows.Add(fields);
                }
                i++;
            }
            file.Close();
            File.Delete(@"C:\R\OFELIZ.CSV");
            for (int a = 0; a < dataGridView1.Rows.Count - 1; a++)
            {
                for (int b = 0; b < dataGridView1.ColumnCount - 1; b++)
                {
                    //remover lixo da lista como por exemplo espaços
                    if (b == 3)
                    {
                        dataGridView1.Rows[a].Cells[3].Value = "2." + lbl_numeroobra.Text + "." + dataGridView1.Rows[a].Cells[0].Value + "." + (a + 1);
                    }
                    else
                    {
                        dataGridView1.Rows[a].Cells[b].Value = dataGridView1.Rows[a].Cells[b].Value.ToString().Trim();
                    }
                    //////////////////////////////////////////////////////////////
                }
            }

            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

            Validadados();



            bool SOLDA = false;

            foreach (DataGridViewRow LINHA in dataGridView1.Rows)
            {
                if (LINHA.Cells[22].Value != null)
                {
                    if (LINHA.Cells[22].Value.ToString() == "Opção 2" || LINHA.Cells[22].Value.ToString() == "Opção 5" || LINHA.Cells[22].Value.ToString() == "Opção 6" || LINHA.Cells[22].Value.ToString() == "Opção 16")
                    {
                        SOLDA = true;
                    }
                }
            }
            try
            {
                if (SOLDA)
                {
                    string EXC = null;
                    new Model().GetProjectInfo().GetUserProperty("PROJECT_USERFIELD_2", ref EXC);
                    if (EXC.Trim() == "")
                        EXC = "2";

                    pdfitext.CriarPlanoSoldadura(int.Parse(dataGridView1.Rows[0].Cells[0].Value.ToString()).ToString("000"), new Model().GetProjectInfo().Builder, new Model().GetProjectInfo().ProjectNumber, new Model().GetProjectInfo().Name, "EXC" + EXC);
                    File.Copy("131 Lista de soldadores - CM.pdf", "C:\\R\\131 Lista de soldadores - CM.pdf");
                }
            }
            catch (Exception)
            {
            }
        }


        void Validadados()
        {
            double somapeso = 0;
            double somaarea = 0;

            string peça = null;
            for (int a = 0; a < dataGridView1.Rows.Count - 1; a++)
            {
                if (a != 0)
                {
                    if (dataGridView1.Rows[a].Cells[4].Value.ToString().Trim() == dataGridView1.Rows[a - 1].Cells[4].Value.ToString().Trim())
                    {
                        peça = peça + dataGridView1.Rows[a - 1].Cells[4].Value.ToString() + Environment.NewLine;
                    }
                }
                //REMOVER '.0' DA QUANTIDADE
                if (dataGridView1.Rows[a].Cells[8].Value.ToString().Trim().Contains(".0"))
                {
                    dataGridView1.Rows[a].Cells[8].Value = dataGridView1.Rows[a].Cells[8].Value.ToString().Replace(".0", "").Replace(" ", "");
                }

                //se celula fase,lote,artigo vazia erro 
                try
                {
                    string r = dataGridView1.Rows[a].Cells[22].Value.ToString();
                }
                catch (Exception)
                {
                    dataGridView1.Rows[a].DefaultCellStyle.BackColor = Color.Red;
                    dataGridView1.Rows[a].Cells[22].Style.BackColor = Color.Blue;
                    
                }

                if (String.IsNullOrEmpty((string)dataGridView1.Rows[a].Cells[0].Value))
                {
                    dataGridView1.Rows[a].DefaultCellStyle.BackColor = Color.Red;
                    dataGridView1.Rows[a].Cells[0].Style.BackColor = Color.Blue;
                }
                if (String.IsNullOrEmpty((string)dataGridView1.Rows[a].Cells[1].Value) || dataGridView1.Rows[a].Cells[1].Value.ToString() == "0")
                {
                    dataGridView1.Rows[a].DefaultCellStyle.BackColor = Color.Red;
                    dataGridView1.Rows[a].Cells[1].Style.BackColor = Color.Blue;
                }
                if (String.IsNullOrEmpty((string)dataGridView1.Rows[a].Cells[2].Value))
                {
                    dataGridView1.Rows[a].DefaultCellStyle.BackColor = Color.Red;
                    dataGridView1.Rows[a].Cells[2].Style.BackColor = Color.Blue;
                }
                //se celula Artigo interno,Destinatário externo vazio erro 
                if (dataGridView1.Rows[a].Cells[2].Value.ToString().ToLower() == "chapa" & String.IsNullOrEmpty((string)dataGridView1.Rows[a].Cells[19].Value))
                {
                    dataGridView1.Rows[a].DefaultCellStyle.BackColor = Color.Red;
                    dataGridView1.Rows[a].Cells[19].Style.BackColor = Color.Blue;
                }
                if (dataGridView1.Rows[a].Cells[2].Value.ToString().ToLower() == "chapa" & String.IsNullOrEmpty((string)dataGridView1.Rows[a].Cells[20].Value))
                {
                    dataGridView1.Rows[a].DefaultCellStyle.BackColor = Color.Red;
                    dataGridView1.Rows[a].Cells[20].Style.BackColor = Color.Blue;
                }

                //altera peso pelo que esta na base de dados 
                if (dataGridView1.Rows[a].Cells[20].Value.ToString().ToLower().Contains("cp") || dataGridView1.Rows[a].Cells[20].Value.ToString().ToLower().Contains("dap"))
                {
                    if (dataGridView1.Rows[a].Cells[19].Value.ToString().Contains('#'))
                    {

                        List<string> ar = new List<string>();
                        string connstr = "SELECT [peso] ,[LarguraUtil],[Marca] FROM [dbo].[Perfilagem3] WHERE [perfil]='" + dataGridView1.Rows[a].Cells[19].Value.ToString().Split('#').Last().Trim() + "' and [Espessura]='" + dataGridView1.Rows[a].Cells[19].Value.ToString().Split('#')[1].Trim() + "'";
                        ComunicaBDtekla b = new ComunicaBDtekla();
                        b.ConectarBD();
                        ar = b.Procurarbd(connstr);
                        b.DesonectarBD();
                        //se a largura util da base de dados for diferente de zero calcula o peso e area, se não for os campos mantem-se os do tekla  
                        if (ar.First().Trim() != "0" && ar.First().Trim() != "")
                        {
                            dataGridView1.Rows[a].Cells[14].Value = (double.Parse(ar.First().Replace(".", ",")) * double.Parse(dataGridView1.Rows[a].Cells[13].Value.ToString().Replace(".", ",")) * double.Parse(dataGridView1.Rows[a].Cells[8].Value.ToString().Replace(".", ",")) / 1000).ToString("0.00").Replace(".", ",");

                            dataGridView1.Rows[a].Cells[15].Value = (double.Parse(ar[1].Trim().Replace(".", ",")) * double.Parse(dataGridView1.Rows[a].Cells[13].Value.ToString().Replace(".", ",")) * double.Parse(dataGridView1.Rows[a].Cells[8].Value.ToString().Replace(".", ",")) / 1000000).ToString("0.00").Replace(".", ",");
                            //preenche a marca da bd
                            dataGridView1.Rows[a].Cells[44].Value = ar[2].Trim();

                        }
                        dataGridView1.Rows[a].Cells[19].Value = dataGridView1.Rows[a].Cells[19].Value.ToString().Split('#').First().Trim();
                    }
                    else
                    {

                        List<string> ar = new List<string>();
                        string connstr = "SELECT [peso] ,[LarguraUtil],[Marca] FROM [dbo].[Perfilagem3] WHERE [perfil]='" + dataGridView1.Rows[a].Cells[9].Value.ToString().Trim() + "'";
                        ComunicaBDtekla b = new ComunicaBDtekla();
                        b.ConectarBD();
                        ar = b.Procurarbd(connstr);
                        b.DesonectarBD();
                        //se a largura util da base de dados for diferente de zero calcula o peso e area, se não for os campos mantem-se os do tekla  
                        try
                        {
                            if (ar.First().Trim() != "0" && ar.Count != 0)
                            {
                                dataGridView1.Rows[a].Cells[14].Value = (double.Parse(ar.First().Replace(".", ",")) * double.Parse(dataGridView1.Rows[a].Cells[13].Value.ToString().Replace(".", ",")) * double.Parse(dataGridView1.Rows[a].Cells[8].Value.ToString().Replace(".", ",")) / 1000).ToString("0.00").Replace(".", ",");
                                if (ar[1].Trim() != "0")
                                {
                                    dataGridView1.Rows[a].Cells[15].Value = (double.Parse(ar[1].Trim().Replace(".", ",")) * double.Parse(dataGridView1.Rows[a].Cells[13].Value.ToString().Replace(".", ",")) * double.Parse(dataGridView1.Rows[a].Cells[8].Value.ToString().Replace(".", ",")) / 1000000).ToString("0.00").Replace(".", ",");
                                }
                                else if (dataGridView1.Rows[a].Cells[9].Value.ToString().ToUpper().Contains("GRADIL"))
                                {
                                    dataGridView1.Rows[a].Cells[15].Value = (double.Parse(dataGridView1.Rows[a].Cells[42].Value.ToString().Replace(".", ",")) * double.Parse(dataGridView1.Rows[a].Cells[13].Value.ToString().Replace(".", ",")) * double.Parse(dataGridView1.Rows[a].Cells[8].Value.ToString().Replace(".", ",")) / 1000000).ToString("0.00").Replace(".", ",");
                                }
                                //preenche a marca da bd
                                dataGridView1.Rows[a].Cells[44].Value = ar[2].Trim();
                            }
                        }
                        catch (Exception) { }
                    }
                }

                //soma peso conjuntos 

                try
                {
                    if (dataGridView1.Rows[a].Cells[22].Value.ToString().ToLower().Contains("opção"))
                    {
                        if (dataGridView1.Rows[a].Cells[14].Value.ToString() != "")
                        {
                            somapeso = somapeso + double.Parse(dataGridView1.Rows[a].Cells[14].Value.ToString().Replace(".", ","));
                            somaarea = somaarea + double.Parse(dataGridView1.Rows[a].Cells[15].Value.ToString().Replace(".", ","));
                        }
                    }
                }
                catch (Exception)
                {

                   
                }
            

                try
                {
                    dataGridView1.Rows[a].Cells[14].Value = double.Parse(dataGridView1.Rows[a].Cells[14].Value.ToString().Replace(".", ",")).ToString("0.00").Replace(".", ",");
                    dataGridView1.Rows[a].Cells[15].Value = double.Parse(dataGridView1.Rows[a].Cells[15].Value.ToString().Replace(".", ",")).ToString("0.00").Replace(".", ",");
                }
                catch (Exception)
                {


                }

            }
            dataGridView1.Refresh();
            if (peça != null)
            {
                MessageBox.Show(this, "Existe linhas de peça duplicada p.f. corrija o erro apos exportar a lista." + Environment.NewLine + peça, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            label3.Text = somapeso.ToString("0.00");
            label6.Text = somaarea.ToString("0.00");
        }

        private void refresh_Click(object sender, EventArgs e)
        {

        }



        private void SaveToCSV(DataGridView DGV, string filename)
        {
            int columnCount = DGV.ColumnCount;
            string columnNames = "";
            string[] output = new string[DGV.RowCount + 7];
            bool criarpastacontagem = false;
            for (int i = 0; i < columnCount; i++)
            {
                columnNames += DGV.Columns[i].HeaderText.ToString() + ";";
            }
            output[0] = "O FELIZ FICHA DE PEÇAS";
            output[1] = "Designação:;" + label1.Text;
            output[2] = "Cliente:;" + label2.Text;
            output[3] = "Nº Obra:;" + lbl_numeroobra.Text;
            output[4] = "Data:;" + label4.Text;
            output[5] = "Classe de Execução:;" + label5.Text;
            output[6] = "Observações;;;;;;;;;;;;;;" + label3.Text + ";" + label6.Text;
            output[7] += columnNames;
            int a = 1;
            for (int i = 8; (i - 8) < DGV.RowCount - 1; i++)
            {
                for (int j = 0; j < columnCount; j++)
                {
                    try
                    {
                        if (DGV.Rows[i - 8].DefaultCellStyle.BackColor != Color.Red)
                        {
                            output[i] += DGV.Rows[i - 8].Cells[j].Value.ToString() + ";";

                            if (DGV.Rows[i - 8].Cells[j].Value.ToString().ToLower().Contains("opção 8"))
                            {
                                criarpastacontagem = true;
                            }


                        }
                        else
                        {
                            if (a == 1)
                            {
                                MessageBox.Show(this, "As linhas a vermelho não foram exportadas", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                a++;
                            }
                        }
                    }
                    catch (Exception)
                    {

                    }
                }
            }

            System.IO.File.WriteAllLines(filename, output, System.Text.Encoding.Default);

            if (criarpastacontagem)
            {
                if (!Directory.Exists(Frm_Inico.PastaReservatorioFicheiros + int.Parse(dataGridView1.Rows[0].Cells[0].Value.ToString()).ToString("000") + "\\20009"))
                {
                    Directory.CreateDirectory(Frm_Inico.PastaReservatorioFicheiros + int.Parse(dataGridView1.Rows[0].Cells[0].Value.ToString()).ToString("000") + "\\20009");
                }
            }

        }

        private void dataGridView1_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            double somapeso = 0;
            double somaarea = 0;
            for (int a = 0; a < dataGridView1.Rows.Count - 1; a++)
            {


                if (dataGridView1.Rows[a].Cells[22].Value.ToString().ToLower().Contains("opção"))
                {
                    if (dataGridView1.Rows[a].Cells[14].Value.ToString() != "")
                    {
                        somapeso = somapeso + double.Parse(dataGridView1.Rows[a].Cells[14].Value.ToString().Replace(".", ","));
                        somaarea = somaarea + double.Parse(dataGridView1.Rows[a].Cells[15].Value.ToString().Replace(".", ","));
                    }
                }

                try
                {
                    dataGridView1.Rows[a].Cells[14].Value = double.Parse(dataGridView1.Rows[a].Cells[14].Value.ToString().Replace(".", ",")).ToString("0.00").Replace(".", ",");
                    dataGridView1.Rows[a].Cells[15].Value = double.Parse(dataGridView1.Rows[a].Cells[15].Value.ToString().Replace(".", ",")).ToString("0.00").Replace(".", ",");
                }
                catch (Exception)
                {


                }

            }
            dataGridView1.Refresh();
            label3.Text = somapeso.ToString("0.00");
            label6.Text = somaarea.ToString("0.00");


        }

        private void ListaOFeliz_Load(object sender, EventArgs e)
        {
            CBunidadenegocio.SelectedIndex = 0;
        }


       
        private void ListaOFeliz_MouseClick(object sender, MouseEventArgs e)
        {

        }

        private void dataGridView1_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            DialogResult a = new DialogResult();
            if (dataGridView1.CurrentCell.ColumnIndex == 2)
            {
                if (dataGridView1.CurrentCell.Value.ToString().Trim().ToLower() == "chapa")
                {
                    dataGridView1.CurrentCell.Value = "Perfil";
                }
                else if (dataGridView1.CurrentCell.Value.ToString().Trim().ToLower() == "perfil")
                {
                    dataGridView1.CurrentCell.Value = "Chapa";
                }
            }
            else if (dataGridView1.CurrentCell.ColumnIndex == 22)
            {
                if (dataGridView1.CurrentCell.Value.ToString().Trim().ToLower() == "corte e furação")
                {
                    dataGridView1.CurrentCell.Value = "Preparação de Chapas";
                }
                else if (dataGridView1.CurrentCell.Value.ToString().Trim().ToLower() == "corte")
                {
                    dataGridView1.CurrentCell.Value = "Corte e Furação";
                }
                else if (dataGridView1.CurrentCell.Value.ToString().Trim().ToLower() == "preparação de chapas")
                {
                    dataGridView1.CurrentCell.Value = "Corte";
                }
            }
            else if (dataGridView1.CurrentCell.ColumnIndex == 19)
            {
                a = MessageBox.Show(this, "Deseja apagar o texto da celula selecionada", "informação", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (a == DialogResult.Yes)
                {
                    dataGridView1.CurrentCell.Value = "";
                    dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[dataGridView1.CurrentCell.ColumnIndex + 1].Value = "";
                }
            }
            else if (dataGridView1.CurrentCell.ColumnIndex == 20)
            {
                a = MessageBox.Show(this, "Deseja apagar o texto da celula selecionada", "informação", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (a == DialogResult.Yes)
                {
                    dataGridView1.CurrentCell.Value = "";
                    dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[dataGridView1.CurrentCell.ColumnIndex - 1].Value = "";
                }
            }
            refresh_Click(sender, e);
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                dataGridView1.ReadOnly = false;
            }
            else
            {
                dataGridView1.ReadOnly = true;
            }
        }

        private void CBunidadenegocio_SelectedIndexChanged(object sender, EventArgs e)
        {
            string[] FILES = Directory.GetFiles("C:\\R", "*.pdf", SearchOption.TopDirectoryOnly);
            foreach (string file in FILES)
            {
                string FILENEW = "c:\\r\\" + CBunidadenegocio.Text + file.Split('\\').Last().Substring(1);

                if (file.Contains("Plano") || file.Contains("Lista"))
                {

                }
                else
                {

                    File.Move(file, FILENEW);

                }
                
            }


            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                try
                {
                    row.Cells[3].Value = CBunidadenegocio.Text + row.Cells[3].Value.ToString().Substring(1);
                }
                catch (Exception)
                {


                }
                try
                {
                    row.Cells[4].Value = CBunidadenegocio.Text + row.Cells[4].Value.ToString().Substring(1);
                }
                catch (Exception)
                {


                }
            }
        }

       

        /////////////////////////////////////////////////////// Notificação e Crianção da Tarefa Soldadura //////////////////////////////////////////////////////////////////////

        //public void InserirTarefaNoBDElias()
        //{
        //    string nomeObra = label1.Text;
        //    string numerodaObra = lbl_numeroobra.Text;
        //    string faseLimpa = int.Parse(dataGridView1.Rows[0].Cells[0].Value.ToString()).ToString("D3");
        //    string fase = faseLimpa.ToString();
        //    string tarefa = $"Processo de Soldadura da Fase {fase}";
        //    string Preparador = "Elias Tinoco";
        //    string Estado = " ";
        //    string observacoes = "";
        //    string prioridades = "8- Processo de soldadura";
        //    int codigodaTarefa = 404;

        //    DateTime dataAtual = DateTime.Now;
        //    DateTime dataInicio = DateTime.Now;

        //    DateTime dataConclusao = dataAtual.AddDays(3); // Adiciona 3 dias à data atual

        //    // Verifica se a data calculada cai no fim de semana
        //    if (dataConclusao.DayOfWeek == DayOfWeek.Saturday)
        //    {
        //        dataConclusao = dataConclusao.AddDays(2); // Se for sábado, vai para segunda-feira
        //    }
        //    else if (dataConclusao.DayOfWeek == DayOfWeek.Friday)
        //    {
        //        dataConclusao = dataConclusao.AddDays(3); // Se for sexta-feira, vai para segunda-feira
        //    }
        //    else if (dataConclusao.DayOfWeek == DayOfWeek.Thursday)
        //    {
        //        dataConclusao = dataConclusao.AddDays(3); // Se for quinta-feira, vai para segunda-feira
        //    }
        //    int Concluido = 0;
        //    DateTime dataConclusaoUser = guna2DateTimePickerdataconclusaouser.Value;



        //    string query = @"
        //                    INSERT INTO dbo.RegistoTarefas
        //                    ([Numero da Obra], [Nome da Obra], Tarefa, Preparador, Estado, Observações, Prioridades, [Codigo da Tarefa], [Data de Inicio], [Data de Conclusão], Concluido,  [Data de Conclusão do user])
        //                    VALUES
        //                    (@NumerodaObra, @NomedaObra, @TAREFA, @PreparadordaTarefa, @Estado, @Observações, @Prioridades, @CodigodadaTarefa, @DataInicio, @DataConclusão, @Concluido, @DataConclusaoUser)";

        //    ComunicaBaseDados BD = new ComunicaBaseDados();

        //    try
        //    {
        //        BD.ConectarBDArtigo();

        //        using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
        //        {
        //            cmd.Parameters.AddWithValue("@NumerodaObra", numerodaObra);
        //            cmd.Parameters.AddWithValue("@NomedaObra", nomeObra);
        //            cmd.Parameters.AddWithValue("@TAREFA", tarefa);
        //            cmd.Parameters.AddWithValue("@PreparadordaTarefa", Preparador);
        //            cmd.Parameters.AddWithValue("@Estado", Estado);
        //            cmd.Parameters.AddWithValue("@Observações", observacoes);
        //            cmd.Parameters.AddWithValue("@Prioridades", prioridades);
        //            cmd.Parameters.AddWithValue("@CodigodadaTarefa", codigodaTarefa);
        //            cmd.Parameters.AddWithValue("@DataInicio", dataInicio);
        //            cmd.Parameters.AddWithValue("@DataConclusão", dataConclusao);
        //            cmd.Parameters.AddWithValue("@Concluido", Concluido);
        //            cmd.Parameters.AddWithValue("@DataConclusaoUser", dataConclusaoUser);

        //            cmd.ExecuteNonQuery();
        //        }

        //        MessageBox.Show("Tarefa inserida com sucesso!");
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Erro ao inserir tarefa: " + ex.Message);
        //    }
        //    finally
        //    {
        //        BD.DesonectarBDArtigo();
        //    }
        //}

        //public void InserirTarefaNoBDHelder()
        //{
        //    string nomeObra = label1.Text;
        //    string numerodaObra = lbl_numeroobra.Text;
        //    string faseLimpa = int.Parse(dataGridView1.Rows[0].Cells[0].Value.ToString()).ToString("D3");
        //    string fase = faseLimpa.ToString();
        //    string tarefa = $"Processo de Soldadura da Fase {fase}";
        //    string Preparador = "Helder Silva";
        //    string Estado = " ";
        //    string observacoes = "";
        //    string prioridades = "8- Processo de soldadura";
        //    int codigodaTarefa = 404;

        //    DateTime dataAtual = DateTime.Now;
        //    DateTime dataInicio = DateTime.Now;

        //    DateTime dataConclusao = dataAtual.AddDays(3); // Adiciona 3 dias à data atual

        //    // Verifica se a data calculada cai no fim de semana
        //    if (dataConclusao.DayOfWeek == DayOfWeek.Saturday)
        //    {
        //        dataConclusao = dataConclusao.AddDays(2); // Se for sábado, vai para segunda-feira
        //    }
        //    else if (dataConclusao.DayOfWeek == DayOfWeek.Friday)
        //    {
        //        dataConclusao = dataConclusao.AddDays(3); // Se for sexta-feira, vai para segunda-feira
        //    }
        //    else if (dataConclusao.DayOfWeek == DayOfWeek.Thursday)
        //    {
        //        dataConclusao = dataConclusao.AddDays(3); // Se for quinta-feira, vai para segunda-feira
        //    }
        //    int Concluido = 0;
        //    DateTime dataConclusaoUser = guna2DateTimePickerdataconclusaouser.Value;



        //    string query = @"
        //                    INSERT INTO dbo.RegistoTarefas
        //                    ([Numero da Obra], [Nome da Obra], Tarefa, Preparador, Estado, Observações, Prioridades, [Codigo da Tarefa], [Data de Inicio], [Data de Conclusão], Concluido,  [Data de Conclusão do user])
        //                    VALUES
        //                    (@NumerodaObra, @NomedaObra, @TAREFA, @PreparadordaTarefa, @Estado, @Observações, @Prioridades, @CodigodadaTarefa, @DataInicio, @DataConclusão, @Concluido, @DataConclusaoUser)";

        //    ComunicaBaseDados BD = new ComunicaBaseDados();

        //    try
        //    {
        //        BD.ConectarBDArtigo();

        //        using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
        //        {
        //            cmd.Parameters.AddWithValue("@NumerodaObra", numerodaObra);
        //            cmd.Parameters.AddWithValue("@NomedaObra", nomeObra);
        //            cmd.Parameters.AddWithValue("@TAREFA", tarefa);
        //            cmd.Parameters.AddWithValue("@PreparadordaTarefa", Preparador);
        //            cmd.Parameters.AddWithValue("@Estado", Estado);
        //            cmd.Parameters.AddWithValue("@Observações", observacoes);
        //            cmd.Parameters.AddWithValue("@Prioridades", prioridades);
        //            cmd.Parameters.AddWithValue("@CodigodadaTarefa", codigodaTarefa);
        //            cmd.Parameters.AddWithValue("@DataInicio", dataInicio);
        //            cmd.Parameters.AddWithValue("@DataConclusão", dataConclusao);
        //            cmd.Parameters.AddWithValue("@Concluido", Concluido);
        //            cmd.Parameters.AddWithValue("@DataConclusaoUser", dataConclusaoUser);

        //            cmd.ExecuteNonQuery();
        //        }

        //        MessageBox.Show("Tarefa inserida com sucesso!");
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Erro ao inserir tarefa: " + ex.Message);
        //    }
        //    finally
        //    {
        //        BD.DesonectarBDArtigo();
        //    }
        //}


        public void InserirTarefaNoBDElias()
        {
            string nomeObra = label1.Text;
            string numerodaObra = lbl_numeroobra.Text;
            string faseLimpa = int.Parse(dataGridView1.Rows[0].Cells[0].Value.ToString()).ToString();
            string fase = faseLimpa.ToString();
            string tarefa = $"Processo de Soldadura da Fase {fase}";
            string Preparador = "Elias Tinoco";
            string Estado = " ";
            string observacoes = "";
            string prioridades = "8- Processo de soldadura";
            int codigodaTarefa = 404;

            DateTime dataAtual = DateTime.Now;
            DateTime dataInicio = DateTime.Now;
            DateTime dataConclusao = dataAtual.AddDays(3); 

            if (dataConclusao.DayOfWeek == DayOfWeek.Saturday)
            {
                dataConclusao = dataConclusao.AddDays(2); // Se for sábado, vai para segunda-feira
            }
            else if (dataConclusao.DayOfWeek == DayOfWeek.Friday)
            {
                dataConclusao = dataConclusao.AddDays(3); // Se for sexta-feira, vai para segunda-feira
            }
            else if (dataConclusao.DayOfWeek == DayOfWeek.Thursday)
            {
                dataConclusao = dataConclusao.AddDays(3); // Se for quinta-feira, vai para segunda-feira
            }
            int Concluido = 0;
            DateTime dataConclusaoUser = guna2DateTimePickerdataconclusaouser.Value;

            string verificaQuery = @"
                SELECT COUNT(*) FROM dbo.RegistoTarefas 
                WHERE [Numero da Obra] = @NumerodaObra 
                AND [Nome da Obra] = @NomedaObra 
                AND Tarefa = @TAREFA 
                AND Preparador = @PreparadordaTarefa
                AND Prioridades = @Prioridades
                AND [Codigo da Tarefa] = @CodigodadaTarefa";

            ComunicaBaseDados BD = new ComunicaBaseDados();

            try
            {
                BD.ConectarBDArtigo();

                using (SqlCommand verificaCmd = new SqlCommand(verificaQuery, BD.GetConnection()))
                {
                    verificaCmd.Parameters.AddWithValue("@NumerodaObra", numerodaObra);
                    verificaCmd.Parameters.AddWithValue("@NomedaObra", nomeObra);
                    verificaCmd.Parameters.AddWithValue("@TAREFA", tarefa);
                    verificaCmd.Parameters.AddWithValue("@PreparadordaTarefa", Preparador);
                    verificaCmd.Parameters.AddWithValue("@Prioridades", prioridades);
                    verificaCmd.Parameters.AddWithValue("@CodigodadaTarefa", codigodaTarefa);

                    int count = (int)verificaCmd.ExecuteScalar(); 

                    if (count > 0)
                    {
                        return; 
                    }
                }
                string query = @"
                                INSERT INTO dbo.RegistoTarefas
                                ([Numero da Obra], [Nome da Obra], Tarefa, Preparador, Estado, Observações, Prioridades, [Codigo da Tarefa], [Data de Inicio], [Data de Conclusão], Concluido, [Data de Conclusão do user])
                                VALUES
                                (@NumerodaObra, @NomedaObra, @TAREFA, @PreparadordaTarefa, @Estado, @Observações, @Prioridades, @CodigodadaTarefa, @DataInicio, @DataConclusão, @Concluido, @DataConclusaoUser )";

                using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                {
                    cmd.Parameters.AddWithValue("@NumerodaObra", numerodaObra);
                    cmd.Parameters.AddWithValue("@NomedaObra", nomeObra);
                    cmd.Parameters.AddWithValue("@TAREFA", tarefa);
                    cmd.Parameters.AddWithValue("@PreparadordaTarefa", Preparador);
                    cmd.Parameters.AddWithValue("@Estado", Estado);
                    cmd.Parameters.AddWithValue("@Observações", observacoes);
                    cmd.Parameters.AddWithValue("@Prioridades", prioridades);
                    cmd.Parameters.AddWithValue("@CodigodadaTarefa", codigodaTarefa);
                    cmd.Parameters.AddWithValue("@DataInicio", dataInicio);
                    cmd.Parameters.AddWithValue("@DataConclusão", dataConclusao);
                    cmd.Parameters.AddWithValue("@Concluido", Concluido);
                    cmd.Parameters.AddWithValue("@DataConclusaoUser ", dataConclusaoUser);

                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Erro ao inserir tarefa: " + ex.Message);
            }
            finally
            {
                BD.DesonectarBDArtigo();
            }
        }



        public void InserirTarefaNoBDHelder()
        {
            string nomeObra = label1.Text;
            string numerodaObra = lbl_numeroobra.Text;
            string faseLimpa = int.Parse(dataGridView1.Rows[0].Cells[0].Value.ToString()).ToString();
            string fase = faseLimpa.ToString();
            string tarefa = $"Processo de Soldadura da Fase {fase}";
            string Preparador = "Helder Silva";
            string Estado = " ";
            string observacoes = "";
            string prioridades = "8- Processo de soldadura";
            int codigodaTarefa = 404;

            DateTime dataAtual = DateTime.Now;
            DateTime dataInicio = DateTime.Now;
            DateTime dataConclusao = dataAtual.AddDays(3);

            if (dataConclusao.DayOfWeek == DayOfWeek.Saturday)
            {
                dataConclusao = dataConclusao.AddDays(2); // Se for sábado, vai para segunda-feira
            }
            else if (dataConclusao.DayOfWeek == DayOfWeek.Friday)
            {
                dataConclusao = dataConclusao.AddDays(3); // Se for sexta-feira, vai para segunda-feira
            }
            else if (dataConclusao.DayOfWeek == DayOfWeek.Thursday)
            {
                dataConclusao = dataConclusao.AddDays(3); // Se for quinta-feira, vai para segunda-feira
            }
            int Concluido = 0;
            DateTime dataConclusaoUser = guna2DateTimePickerdataconclusaouser.Value;

            string verificaQuery = @"
                SELECT COUNT(*) FROM dbo.RegistoTarefas 
                WHERE [Numero da Obra] = @NumerodaObra 
                AND [Nome da Obra] = @NomedaObra 
                AND Tarefa = @TAREFA 
                AND Preparador = @PreparadordaTarefa
                AND Prioridades = @Prioridades
                AND [Codigo da Tarefa] = @CodigodadaTarefa";

            ComunicaBaseDados BD = new ComunicaBaseDados();

            try
            {
                BD.ConectarBDArtigo();

                using (SqlCommand verificaCmd = new SqlCommand(verificaQuery, BD.GetConnection()))
                {
                    verificaCmd.Parameters.AddWithValue("@NumerodaObra", numerodaObra);
                    verificaCmd.Parameters.AddWithValue("@NomedaObra", nomeObra);
                    verificaCmd.Parameters.AddWithValue("@TAREFA", tarefa);
                    verificaCmd.Parameters.AddWithValue("@PreparadordaTarefa", Preparador);
                    verificaCmd.Parameters.AddWithValue("@Prioridades", prioridades);
                    verificaCmd.Parameters.AddWithValue("@CodigodadaTarefa", codigodaTarefa);

                    int count = (int)verificaCmd.ExecuteScalar();

                    if (count > 0)
                    {
                        return;
                    }
                }

                string query = @"
                                INSERT INTO dbo.RegistoTarefas
                                ([Numero da Obra], [Nome da Obra], Tarefa, Preparador, Estado, Observações, Prioridades, [Codigo da Tarefa], [Data de Inicio], [Data de Conclusão], Concluido, [Data de Conclusão do user])
                                VALUES
                                (@NumerodaObra, @NomedaObra, @TAREFA, @PreparadordaTarefa, @Estado, @Observações, @Prioridades, @CodigodadaTarefa, @DataInicio, @DataConclusão, @Concluido, @DataConclusaoUser )";

                using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                {
                    cmd.Parameters.AddWithValue("@NumerodaObra", numerodaObra);
                    cmd.Parameters.AddWithValue("@NomedaObra", nomeObra);
                    cmd.Parameters.AddWithValue("@TAREFA", tarefa);
                    cmd.Parameters.AddWithValue("@PreparadordaTarefa", Preparador);
                    cmd.Parameters.AddWithValue("@Estado", Estado);
                    cmd.Parameters.AddWithValue("@Observações", observacoes);
                    cmd.Parameters.AddWithValue("@Prioridades", prioridades);
                    cmd.Parameters.AddWithValue("@CodigodadaTarefa", codigodaTarefa);
                    cmd.Parameters.AddWithValue("@DataInicio", dataInicio);
                    cmd.Parameters.AddWithValue("@DataConclusão", dataConclusao);
                    cmd.Parameters.AddWithValue("@Concluido", Concluido);
                    cmd.Parameters.AddWithValue("@DataConclusaoUser ", dataConclusaoUser);

                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Erro ao inserir tarefa: " + ex.Message);
            }
            finally
            {
                BD.DesonectarBDArtigo();
            }
        }

        private void button5_Click_2(object sender, EventArgs e)
        {
            foreach (DataGridViewRow item in dataGridView1.Rows)
            {
                try
                {
                    if (item.Cells[9].Value.ToString().ToUpper().Contains("RHS"))
                    {

                        string[] r = null;


                        if (item.Cells[9].Value.ToString().ToUpper().Contains("*"))
                        {
                            r = item.Cells[9].Value.ToString().Split('*');
                        }
                        else if (item.Cells[9].Value.ToString().ToUpper().Contains("\\"))
                        {
                            r = item.Cells[9].Value.ToString().Split('\\');
                        }
                        else if (item.Cells[9].Value.ToString().ToUpper().Contains("X"))
                        {
                            r = item.Cells[9].Value.ToString().Split('X');
                        }

                        double num1 = double.Parse(r[0].Replace("RHS", ""));
                        double num2 = double.Parse(r[1].Replace("RHS", ""));
                        if (num1 > num2)
                        {
                            item.Cells[9].Value = "CFR" + num1 + "X" + num2 + "X" + r[2].ToString();
                        }
                        else
                        {
                            item.Cells[9].Value = "CFR" + num2 + "X" + num1 + "X" + r[2].ToString();
                        }

                    }
                    else if (item.Cells[9].Value.ToString().ToUpper().Contains("SHS"))
                    {

                        string[] r = null;


                        if (item.Cells[9].Value.ToString().ToUpper().Contains("*"))
                        {
                            r = item.Cells[9].Value.ToString().Split('*');
                        }
                        else if (item.Cells[9].Value.ToString().ToUpper().Contains("\\"))
                        {
                            r = item.Cells[9].Value.ToString().Split('\\');
                        }
                        else if (item.Cells[9].Value.ToString().ToUpper().Contains("X"))
                        {
                            r = item.Cells[9].Value.ToString().Split('X');
                        }

                        item.Cells[9].Value = "CFR" + r[0].Replace("SHS", "") + "X" + r[0].Replace("SHS", "") + "X" + r[1].ToString();
                    }
                    else if (item.Cells[9].Value.ToString().ToUpper().Contains("L"))
                    {
                        string[] r = null;

                        if (item.Cells[9].Value.ToString().ToUpper().Contains("\\"))
                        {
                            r = item.Cells[9].Value.ToString().Split('\\');
                        }
                        else if (item.Cells[9].Value.ToString().ToUpper().Contains("*"))
                        {
                            r = item.Cells[9].Value.ToString().Split('*');
                        }

                        if (r.Count() == 2)
                        {
                            item.Cells[9].Value = "L" + r[0].Replace("L", "") + "X" + r[0].Replace("L", "") + "X" + r[1].ToString();
                        }
                        else if (r.Count() == 3)
                        {
                            item.Cells[9].Value = "L" + r[0].Replace("L", "") + "X" + r[1] + "X" + r[2].ToString();
                        }
                    }
                    if (item.Cells[9].Value.ToString().ToUpper().Contains("_"))
                    {
                        item.Cells[9].Value = item.Cells[9].Value.ToString().Replace("_", "");
                    }

                }
                catch (Exception)
                {


                }


            }
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            pdf_cliente m = new pdf_cliente(this);
            m.ShowDialog();
        }

        public bool listacarregadaexc3 = false;

        private void button3_Click_1(object sender, EventArgs e)
        {
            listacarregadaexc3 = true;


            dataGridView1.Rows.Clear();
            DialogResult A = openFileDialog1.ShowDialog();
            if (A == DialogResult.Cancel)
            {

            }
            else
            {
                string line = null;
                int i = 1;
                StreamReader file = new StreamReader(openFileDialog1.FileName, Encoding.Default, true);
                while ((line = file.ReadLine()) != null)
                {
                    if (i == 2)
                    {
                        var fields = line.Split(';');
                        label1.Text = fields[1];
                    }
                    if (i == 3)
                    {
                        var fields = line.Split(';');
                        label2.Text = fields[1];
                    }
                    if (i == 4)
                    {
                        var fields = line.Split(';');
                        lbl_numeroobra.Text = fields[1];
                    }
                    if (i == 5)
                    {
                        var fields = line.Split(';');
                        label4.Text = fields[1];
                    }
                    if (i == 6)
                    {
                        var fields = line.Split(';');
                        label5.Text = fields[1];
                    }

                    if (i > 8)
                    {
                        var fields = line.Split(';');
                        dataGridView1.Rows.Add(fields);
                    }
                    i++;
                }
                file.Close();

                for (int a = 0; a < dataGridView1.Rows.Count - 1; a++)
                {
                    for (int b = 0; b < dataGridView1.ColumnCount - 1; b++)
                    {
                        //remover lixo da lista como por exemplo espaços
                        if (b == 3)
                        {
                            dataGridView1.Rows[a].Cells[3].Value = "2." + lbl_numeroobra.Text + "." + dataGridView1.Rows[a].Cells[0].Value + "." + (a + 1);
                        }
                        else
                        {
                            try
                            {
                                dataGridView1.Rows[a].Cells[b].Value = dataGridView1.Rows[a].Cells[b].Value.ToString().Trim();
                            }
                            catch (Exception) { }


                        }
                        //////////////////////////////////////////////////////////////
                    }
                }

                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                Validadados();
            }
        }

        private void refresh_Click_1(object sender, EventArgs e)
        {

            //remove linhas duplicadas somando peso
            if (Form.ModifierKeys == Keys.Control)
            {
                List<int> linhas = new List<int>();
                string peça = null;
                for (int a = 0; a < dataGridView1.Rows.Count - 1; a++)
                {

                    if (a != 0)
                    {
                        if (dataGridView1.Rows[a].Cells[4].Value.ToString().Trim() == dataGridView1.Rows[a - 1].Cells[4].Value.ToString().Trim())
                        {

                            linhas.Add(a - 1);
                            peça = peça + dataGridView1.Rows[a - 1].Cells[4].Value.ToString() + Environment.NewLine;

                        }
                    }
                }
                DialogResult A = MessageBox.Show(this, "Existe linhas de peça duplicada pretende corrijir a lista." + Environment.NewLine + peça, "Erro", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                if (A == DialogResult.Yes)
                {
                    foreach (var aq in linhas)
                    {
                        int a = aq + 1;

                        dataGridView1.Rows[a].Cells[8].Value = (int.Parse(dataGridView1.Rows[a].Cells[8].Value.ToString().Replace(".0", "")) + int.Parse(dataGridView1.Rows[a - 1].Cells[8].Value.ToString().Replace(".0", "")));

                        dataGridView1.Rows[a].Cells[14].Value = double.Parse(dataGridView1.Rows[a].Cells[14].Value.ToString().Replace(".", ",")) + double.Parse(dataGridView1.Rows[a - 1].Cells[14].Value.ToString().Replace(".", ","));

                        dataGridView1.Rows[a].Cells[15].Value = double.Parse(dataGridView1.Rows[a].Cells[15].Value.ToString().Replace(".", ",")) + double.Parse(dataGridView1.Rows[a - 1].Cells[15].Value.ToString().Replace(".", ","));
                    }

                    int b = 0;
                    foreach (var item in linhas)
                    {
                        dataGridView1.Rows.RemoveAt(item - b);
                        b++;
                    }
                }
            }
            //verificar normal
            else
            {
                bool dapcp = false;
                for (int a = 0; a < dataGridView1.Rows.Count - 1; a++)
                {
                    if (dataGridView1.Rows[a].Cells[20].Value.ToString().ToLower().Contains("cp") || dataGridView1.Rows[a].Cells[20].Value.ToString().ToLower().Contains("dap"))
                    {
                        dapcp = true;
                    }
                    dataGridView1.Rows[a].DefaultCellStyle.BackColor = default(Color);
                    for (int b = 0; b < dataGridView1.ColumnCount - 1; b++)
                    {
                        dataGridView1.Rows[a].Cells[b].Style.BackColor = default(Color);
                    }
                }
                if (dapcp)
                {
                    MessageBox.Show(this, "não é possivel verificar.");
                }
                else
                {
                    Validadados();
                }


            }
        }

       

        private void guna2Button6_Click(object sender, EventArgs e)
        {
            if (label5.Text == "3" || label5.Text == "4")
            {
                this.Visible = false;
                listacarregadaexc3 = true;
                Frm_PecasExc3 f = new Frm_PecasExc3(this, _Formpai);
                f.ShowDialog();
                this.Visible = true;
            }
            else
            {
                this.Visible = false;
                Frm_Pecas f = new Frm_Pecas(this, _Formpai);
                f.ShowDialog();
                this.Visible = true;
            }

            try
            {
                if (int.TryParse(dataGridView1.Rows[0].Cells[0].Value?.ToString(), out int valor))
                {
                    string caminho = Frm_Inico.PastaReservatorioFicheiros + valor.ToString("000") + "\\20005";

                    if (valor >= 1 && valor <= 499)
                    {
                        if (Directory.Exists(caminho))
                        {
                            InserirTarefaNoBDElias();
                            InserirTarefaNoBDHelder();
                        }
                        else
                        {
                        }
                    }
                    else
                    {
                    }
                }
                else
                {
                    MessageBox.Show(this, "Valor inválido na célula. Certifique-se de que o valor seja um número inteiro.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Erro inesperado: " + ex.Message);
            }
        }

        public void button222_Click(object sender, EventArgs e)
        {
            if (Directory.Exists(Frm_Inico.PastaReservatorioFicheiros + int.Parse(dataGridView1.Rows[0].Cells[0].Value.ToString()).ToString("000")))
            {
                SaveToCSV(this.dataGridView1, Frm_Inico.PastaReservatorioFicheiros + int.Parse(dataGridView1.Rows[0].Cells[0].Value.ToString()).ToString("000") + "\\" + lbl_numeroobra.Text + "F" + int.Parse(dataGridView1.Rows[0].Cells[0].Value.ToString()).ToString("000") + ".csv");
            }
            else
            {
                Directory.CreateDirectory(Frm_Inico.PastaReservatorioFicheiros + int.Parse(dataGridView1.Rows[0].Cells[0].Value.ToString()).ToString("000"));
                SaveToCSV(this.dataGridView1, Frm_Inico.PastaReservatorioFicheiros + int.Parse(dataGridView1.Rows[0].Cells[0].Value.ToString()).ToString("000") + "\\" + lbl_numeroobra.Text + "F" + int.Parse(dataGridView1.Rows[0].Cells[0].Value.ToString()).ToString("000") + ".csv");
            }

            foreach (DataGridViewRow LINHA in dataGridView1.Rows)
            {
                if (LINHA.Cells[20].Value != null)
                {
                    string criapasta = Frm_Inico.PastaPartilhada + "\\" + Frm_Inico.ano + "\\" + LINHA.Cells[20].Value.ToString().Trim() + "\\" + lbl_numeroobra.Text.Trim() + "\\" + int.Parse(LINHA.Cells[0].Value.ToString()).ToString("000");
                    if (LINHA.Cells[20].Value.ToString() == "CP" || LINHA.Cells[20].Value.ToString() == "DAP")
                    {
                        if (!Directory.Exists(criapasta))
                        {
                            Directory.CreateDirectory(criapasta);
                        }
                        string fase = int.Parse(dataGridView1.Rows[0].Cells[0].Value.ToString()).ToString("000");
                        Excel(fase);
                    }
                }
            }


            

            MessageBox.Show(this, "lista exportada com sucesso " + Environment.NewLine + "Não esquecer de enviar os Parafusos desta fase com o procedimento de fase 1000", "Exportação", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void Excel(string fase)
        {
            string caminhocsv = Path.Combine(@"\\marconi\OFELIZ\OFELIZ\OFM\2.AN\2.CM\DP\1 Obras", Frm_Inico.ano, lbl_numeroobra.Text, "1.9 Gestão de fabrico", int.Parse(fase).ToString("000") + @"\");
            string caminhocsv2 = Path.Combine(Frm_Inico.PastaPartilhada, Frm_Inico.ano, "CP", lbl_numeroobra.Text, int.Parse(fase).ToString("000"));
            string nomeprojeto = lbl_numeroobra.Text + "F" + int.Parse(fase).ToString("000");

            string destino = Path.Combine(caminhocsv2, nomeprojeto + ".csv");
            string destinoDir = Path.GetDirectoryName(destino);


            try
            {
                File.Copy(caminhocsv + nomeprojeto + ".csv", destino);
            }
            catch (Exception ex)
            {
            }
        }

    }
}

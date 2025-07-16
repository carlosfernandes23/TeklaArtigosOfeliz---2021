using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;

namespace TeklaArtigosOfeliz
{
    public partial class Frm_Parafusos : Form
    {
        Frm_Inico formpai;
        public Frm_Parafusos(Frm_Inico _formpai)
        {
            InitializeComponent();
            formpai = _formpai;
        }

        private void parafusos_Load(object sender, EventArgs e)
        {
            carregadados();
            VerificarColunaClasse();
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
            dataGridView1.Sort(this.dataGridView1.Columns[9], ListSortDirection.Ascending);
            dataGridView1.Sort(this.dataGridView1.Columns[1], ListSortDirection.Ascending);
            List<DateTime> data = new List<DateTime>(); 
            for (int a = 0; a < dataGridView1.Rows.Count - 1; a++)
            {
                for (int b = 0; b < dataGridView1.ColumnCount - 1; b++)
                {
                    //remover lixo da lista como por exemplo espaços
                    if (b == 0)
                    {
                        dataGridView1.Rows[a].Cells[0].Value = formpai.fase1000;
                    }
                    else if (b == 3)
                    {
                        dataGridView1.Rows[a].Cells[3].Value = "2." + lbl_numeroobra.Text + "." + dataGridView1.Rows[a].Cells[0].Value + "." + (a + 1);
                    }
                    else if (b == 4)
                    {
                        dataGridView1.Rows[a].Cells[4].Value = "2." + lbl_numeroobra.Text + "." + dataGridView1.Rows[a].Cells[1].Value + "." + formpai.fase1000+"H"+(a + 1);
                    }
                    else if (b == 18)
                    {
                        try
                        {
                            dataGridView1.Rows[a].Cells[b].Value = dataGridView1.Rows[a].Cells[b].Value.ToString().Replace(".", "/").Replace("-", "/").Replace("_", "/").Trim();
                            data.Add(Convert.ToDateTime(dataGridView1.Rows[a].Cells[b].Value.ToString()));
                        }
                        catch (Exception)
                        {

                          
                        }
                       
                    }
                    else
                    {
                        dataGridView1.Rows[a].Cells[b].Value = dataGridView1.Rows[a].Cells[b].Value.ToString().Trim();
                    }
                    //////////////////////////////////////////////////////////////
                }
            }
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView1.Columns[18].DefaultCellStyle.Format = "dd/MM/yyyy";
            data.Sort();
            try
            {
                dateTimePicker1.Text = data[0].ToShortDateString();
            }
            catch (Exception)
            {

             
            }
            //RESINA();
            alteraQuantidades();
        }

        private void RESINA()
        {
            foreach (DataGridViewRow item in dataGridView1.Rows)
            {
                DataGridViewCell currentcell = item.Cells[47];
                if (currentcell.Value!=null)
                {
                    if (currentcell.Value.ToString() != "")
                    {
                        int DIAMETRO = int.Parse(item.Cells[9].Value.ToString().Replace("D", ""));

                        string ArtigoVarao = currentcell.Value.ToString().Split('#')[2] + DIAMETRO;
                        if (currentcell.Value.ToString().Split('#')[0] == "QUIMICA")
                        {
                            ArtigoVarao = "VRSM" + DIAMETRO;
                        }

                        item.Cells[9].Value = ArtigoVarao;

                        //item.Cells[11].Value = "DIN975";
                        //item.Cells[12].Value = "2.1";

                        item.Cells[19].Value = ArtigoVarao;
                        item.Cells[10].Value = currentcell.Value.ToString().Split('#')[2];
                        item.Cells[13].Value = (double.Parse(currentcell.Value.ToString().Split('#')[3]) + double.Parse(currentcell.Value.ToString().Split('#')[4])).ToString("0");

                        int DIAMETROFURO = 0;

                        if (DIAMETRO < 24)
                        {
                            DIAMETROFURO = DIAMETRO + 2;
                        }
                        else
                        {
                            DIAMETROFURO = DIAMETRO + 4;
                        }




                        double CALCULO = (Math.PI * Math.Pow(((DIAMETROFURO * 0.01) / 2), 2)) * ((double)2 / (double)3) * (double.Parse(currentcell.Value.ToString().Split('#')[4])) * 0.01 * double.Parse(item.Cells[8].Value.ToString());

                        List<string> cell = new List<string>();
                        foreach (DataGridViewCell celula in item.Cells)
                        {
                            cell.Add("" + celula.Value);
                        }
                        dataGridView1.Rows.Add(cell.ToArray());
                        dataGridView1.Rows[dataGridView1.Rows.Count - 2].Cells[8].Value = (CALCULO * 1000).ToString("0");
                        dataGridView1.Rows[dataGridView1.Rows.Count - 2].Cells[10].Value = currentcell.Value.ToString().Split('#')[1];
                        currentcell.Value = "";
                        dataGridView1.Rows[dataGridView1.Rows.Count - 2].Cells[47].Value = "";
                        dataGridView1.Refresh();

                    }
                }
            }
        }
        // Codigo Rui Parafusaria
        /* private void alteraQuantidades()
         {
             foreach (DataGridViewRow item in dataGridView1.Rows)
             {
                 DataGridViewCell currentcell = item.Cells[8];
                 if (currentcell.Value!=null)
                 {
                     if (double.Parse(currentcell.Value.ToString()) <= 150)
                     {
                         currentcell.Value = int.Parse((double.Parse(currentcell.Value.ToString()) + 5).ToString("0"));

                     }
                     else if (double.Parse(currentcell.Value.ToString()) <= 1000)
                     {

                         currentcell.Value = int.Parse((double.Parse(currentcell.Value.ToString()) * ((5.0 / 100.0) + 1)).ToString("0"));

                     }
                     else if (double.Parse(currentcell.Value.ToString()) <= 10000)
                     {
                         currentcell.Value = int.Parse((double.Parse(currentcell.Value.ToString()) * ((2.5 / 100.0) + 1)).ToString("0"));

                     }
                     else
                     {
                         currentcell.Value = int.Parse((double.Parse(currentcell.Value.ToString()) * ((1 / 100.0) + 1)).ToString("0"));

                     }

                     item.Cells[22].Value = "Opção 8";
                     item.Cells[20].Value = "08";
                 }



             }
         }*/

        private void alteraQuantidades()
        {
            foreach (DataGridViewRow item in dataGridView1.Rows)
            {
                DataGridViewCell currentcell = item.Cells[8];

                if (currentcell.Value != null)
                {
                    double valorAtual = double.Parse(currentcell.Value.ToString());

                    // 0 a 10
                    if (valorAtual <= 10)
                    {
                        currentcell.Value = Math.Round(valorAtual + 1).ToString("0");
                    }

                    // Maior que 10 e até 250 (aumento de 5%)
                    else if (valorAtual > 10 && valorAtual <= 250)
                    {
                        currentcell.Value = Math.Round(valorAtual * 1.05).ToString("0");
                    }

                    // Maior que 250 e até 1000 (aumento de 2,5%)
                    else if (valorAtual > 250 && valorAtual <= 1000)
                    {
                        currentcell.Value = Math.Round(valorAtual * 1.025).ToString("0");
                    }

                    // Maior que 1000 e até 10000 (aumento de 2%)
                    else if (valorAtual > 1000 && valorAtual <= 10000)
                    {
                        currentcell.Value = Math.Round(valorAtual * 1.02).ToString("0");
                    }

                    // Maior que 10000 (aumento de 1%)
                    else
                    {
                        currentcell.Value = Math.Round(valorAtual * 1.01).ToString("0");
                    }

                    // Atualiza outras células conforme solicitado
                    item.Cells[22].Value = "Opção 8";
                    item.Cells[20].Value = "08";
                }
            }
        }

       

        //private void button2_Click(object sender, EventArgs e)
        //{
        //    if (!Directory.Exists(Frm_Inico.CaminhoModelo + @"\listas"))
        //    {
        //        Directory.CreateDirectory(Frm_Inico.CaminhoModelo + @"\listas");
        //    }
        //    string Save = null;
        //    if (Directory.Exists(Frm_Inico.PastaReservatorioFicheiros + formpai.fase1000))
        //    {
        //        Save = Frm_Inico.PastaReservatorioFicheiros + formpai.fase1000 + "\\" + lbl_numeroobra.Text + "F" + formpai.fase1000 + ".csv";
        //    }
        //    else
        //    {
        //        Directory.CreateDirectory(Frm_Inico.PastaReservatorioFicheiros + formpai.fase1000);
        //        Save = Frm_Inico.PastaReservatorioFicheiros + formpai.fase1000 + "\\" + lbl_numeroobra.Text + "F" + formpai.fase1000 + ".csv";
        //    }

        //    if (!Directory.Exists(Frm_Inico.PastaPartilhada+"\\"+Frm_Inico.ano+"\\ARM\\"+ lbl_numeroobra.Text+"\\"+ formpai.fase1000))
        //    {
        //        Directory.CreateDirectory(Frm_Inico.PastaPartilhada + "\\" + Frm_Inico.ano + "\\ARM\\" + lbl_numeroobra.Text + "\\" + formpai.fase1000);
        //    }
        //    if (!Directory.Exists(Frm_Inico.PastaPartilhada + "\\" + Frm_Inico.ano + "\\ARM\\" + lbl_numeroobra.Text + "\\" + formpai.fase1000))
        //    {
        //        Directory.CreateDirectory(Frm_Inico.PastaPartilhada + "\\" + Frm_Inico.ano + "\\ARM\\" + lbl_numeroobra.Text + "\\" + formpai.fase1000);
        //    }
        //    if (!Directory.Exists(Frm_Inico.PastaReservatorioFicheiros + formpai.fase1000+"\\20001"))
        //    {
        //        Directory.CreateDirectory(Frm_Inico.PastaReservatorioFicheiros + formpai.fase1000 + "\\20001");
        //    }
        //    if (!Directory.Exists(Frm_Inico.PastaReservatorioFicheiros + formpai.fase1000 + "\\20009"))
        //    {
        //        Directory.CreateDirectory(Frm_Inico.PastaReservatorioFicheiros + formpai.fase1000 + "\\20009");
        //    }
        //    SaveToCSV(dataGridView1, Save);
        //    Tekla.Structures.Model.Model m = new Tekla.Structures.Model.Model();
        //    string numeroobra = m.GetProjectInfo().ProjectNumber;
        //    MessageBox.Show("Lista exportada " + Environment.NewLine + " Fase:" + formpai.fase1000, "Exportação", MessageBoxButtons.OK, MessageBoxIcon.Information);

        //}

        private void SaveToCSV(DataGridView DGV, string filename)
        {
            int columnCount = DGV.ColumnCount;
            string columnNames = "";
            string[] output = new string[DGV.RowCount + 7];

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

        }

       
        private void button9_Click(object sender, EventArgs e)
        {
            if (!Directory.Exists(Frm_Inico.CaminhoModelo + @"\listas"))
            {
                Directory.CreateDirectory(Frm_Inico.CaminhoModelo + @"\listas");
            }
            string Save = null;

            string numeroObra = lbl_numeroobra.Text.Trim();
            string fase = formpai.fase1000.Trim();

            if (Directory.Exists(Frm_Inico.PastaReservatorioFicheiros + fase))
            {
                Save = Frm_Inico.PastaReservatorioFicheiros + fase + "\\" + numeroObra + "F" + fase + ".csv";
            }
            else
            {
                Directory.CreateDirectory(Frm_Inico.PastaReservatorioFicheiros + fase);
                Save = Frm_Inico.PastaReservatorioFicheiros + fase + "\\" + numeroObra + "F" + fase + ".csv";
            }

            string pastaPartilhada = Frm_Inico.PastaPartilhada + "\\" + Frm_Inico.ano + "\\ARM\\" + numeroObra + "\\" + fase;

            if (!Directory.Exists(pastaPartilhada))
            {
                Directory.CreateDirectory(pastaPartilhada);
            }

            if (!Directory.Exists(Frm_Inico.PastaReservatorioFicheiros + fase + "\\20001"))
            {
                Directory.CreateDirectory(Frm_Inico.PastaReservatorioFicheiros + fase + "\\20001");
            }
            if (!Directory.Exists(Frm_Inico.PastaReservatorioFicheiros + fase + "\\20009"))
            {
                Directory.CreateDirectory(Frm_Inico.PastaReservatorioFicheiros + fase + "\\20009");
            }

            SaveToCSV(dataGridView1, Save);

            Tekla.Structures.Model.Model m = new Tekla.Structures.Model.Model();
            string numeroProjeto = m.GetProjectInfo().ProjectNumber.Trim(); 
            MessageBox.Show(this, "Lista exportada " + Environment.NewLine + " Fase:" + fase, "Exportação", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            for (int a = 0; a < dataGridView1.Rows.Count - 1; a++)
            {

                dataGridView1.Rows[a].Cells[18].Value = dateTimePicker1.Text;
            }
        }

        private void VerificarColunaClasse()
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.IsNewRow) continue;

                var valorClasse = row.Cells["Classe"].Value?.ToString();

                if (string.IsNullOrWhiteSpace(valorClasse)) continue;

                if (!valorClasse.StartsWith("8,8") && !valorClasse.StartsWith("10,9"))
                {
                    MessageBox.Show("Existe Normas com a Classe atribuida errada, Por favor revise a lista.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return; 
                }
            }
        }


    }
}

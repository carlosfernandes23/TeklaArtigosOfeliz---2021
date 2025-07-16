using BDTEKLA;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tekla.Technology.Akit;


namespace TeklaArtigosOfeliz
{
    public partial class Frm_ViewBD: Form
    {
        public string connectionString = "Data Source=GALILEU\\PREPARACAO;Initial Catalog=ArtigoTekla;Persist Security Info=True;User ID=SA;Password=preparacao";
        string connect = @"Data Source = GALILEU\PREPARACAO;Initial Catalog = TempoPreparacao; Persist Security Info=True;User ID = sa; Password=preparacao";

        public Frm_ViewBD()
        {
            InitializeComponent();
        }


        private void AtualizarBD_Load(object sender, EventArgs e)
        {
            string userName = Environment.UserName;
            string formattedUserName = string.Join(" ", userName.Split('.').Select(word => char.ToUpper(word[0]) + word.Substring(1).ToLower()));

            SqlConnection connection = new SqlConnection(connect);
            connection.Open();
            SqlCommand iComando = new SqlCommand("SELECT count(*)  FROM [TempoPreparacao].[dbo].[nPreparadores1] where [nome]='" + formattedUserName + "'", connection);

            int FilaAfectada = (int)iComando.ExecuteScalar();
            connection.Close();
            if (FilaAfectada == 0)
            {
                MessageBox.Show(this, "Desculpe mas nao tem permissões para editar a base de dados, P.F. fale com um preparador " + formattedUserName + "OBRIGADO PELA COMPREENSÃO.", "ERRO", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                this.Close();
            }



            ACTUALIZA();
        }

        void ACTUALIZA()
        {
            SqlConnection connection = new SqlConnection(connectionString);
            SqlDataAdapter dataadapter = new SqlDataAdapter("SELECT [Id],[Familia],[Perfil],[Material],[Espessura],[Artigo],[Peso],[LarguraUtil],[Dest],[Marca] FROM dbo.Perfilagem3", connection);
            DataSet ds = new DataSet();
            connection.Open();
            dataadapter.Fill(ds, "TESTE");
            connection.Close();
            dataGridView1.DataSource = ds;
            dataGridView1.DataMember = "TESTE";
        }

        private void aCTUALIZARToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void cRIARARTIGOToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void aPAGARLINHASToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void iMPORTAFICHEIROToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                textBoxFamilia.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString(); 
                textBoxPerfil.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();  
                textBoxMaterial.Text = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString(); 
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
           
                int row = dataGridView1.CurrentCell.RowIndex;
                int a1 = (int)dataGridView1.Rows[row].Cells[0].Value;
                string a2 = dataGridView1.Rows[row].Cells[1].Value.ToString();
                string a3 = dataGridView1.Rows[row].Cells[2].Value.ToString();
                string a4 = dataGridView1.Rows[row].Cells[3].Value.ToString();
                string a5 = dataGridView1.Rows[row].Cells[4].Value.ToString();
                string a6 = dataGridView1.Rows[row].Cells[5].Value.ToString();
                string a7 = dataGridView1.Rows[row].Cells[6].Value.ToString();
                string a8 = dataGridView1.Rows[row].Cells[7].Value.ToString();
                string a9 = dataGridView1.Rows[row].Cells[8].Value.ToString();
                string a10 = dataGridView1.Rows[row].Cells[9].Value.ToString();

               

                Frm_ActualizaBD A = new Frm_ActualizaBD(a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, connectionString);
                A.ShowDialog();
                ACTUALIZA();

                   
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ACTUALIZA();

            if (dataGridView1.Rows.Count > 0)
            {
                int lastRowIndex = dataGridView1.Rows.Count - 2;

                if (dataGridView1.Rows[lastRowIndex].Cells[0].Value != null)
                {
                    int lastValue = (int)dataGridView1.Rows[lastRowIndex].Cells[0].Value;
                    Frm_ActualizaBD A = new Frm_ActualizaBD(lastValue, connectionString);
                    A.ShowDialog();
                    ACTUALIZA();
                }
                else
                {
                    MessageBox.Show(this, "A célula está vazia. Não é possível obter o valor.");
                }
            }
            else
            {
                MessageBox.Show(this, "Não há linhas no DataGridView.");
            }
        }

        private void buttonImportar_Click(object sender, EventArgs e)
        {
            Frm_BDImportaexcel I = new Frm_BDImportaexcel();
            I.ShowDialog();
            ACTUALIZA();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult a = MessageBox.Show(this, "Deseja apagar as linhas selecionadas?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);

            if (a == DialogResult.Yes)
            {
                SqlConnection connection = new SqlConnection(connectionString);
                connection.Open();
                foreach (DataGridViewRow item in dataGridView1.SelectedRows)
                {
                    string comando = "DELETE FROM dbo.Perfilagem3 WHERE id=" + item.Cells[0].Value;
                    SqlCommand c = new SqlCommand(comando, connection);
                    int nlinhas = c.ExecuteNonQuery();

                }
                connection.Close();
                ACTUALIZA();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            FiltrarTarefas();
        }

        private void FiltrarDataGridViewAddTarefas(ComunicaBDtekla BD, string familia, string perfil, string material)
        {
            string query = "SELECT [Id], [Familia], [Perfil], [Material], [Espessura], [Artigo], [Peso], [LarguraUtil], [Dest], [Marca] FROM dbo.Perfilagem3 WHERE 1=1";

            if (!string.IsNullOrEmpty(familia))
            {
                query += " AND [Familia] LIKE @Familia"; 
            }

            if (!string.IsNullOrEmpty(perfil))
            {
                query += " AND [Perfil] LIKE @Perfil";
            }

            if (!string.IsNullOrEmpty(material))
            {
                query += " AND [Material] LIKE @Material"; 
            }

            using (var command = new SqlCommand(query, BD.GetConnection()))
            {
                if (!string.IsNullOrEmpty(familia))
                {
                    command.Parameters.AddWithValue("@Familia", familia + "%"); 
                }

                if (!string.IsNullOrEmpty(perfil))
                {
                    command.Parameters.AddWithValue("@Perfil", perfil + "%"); 
                }

                if (!string.IsNullOrEmpty(material))
                {
                    command.Parameters.AddWithValue("@Material", material + "%"); 
                }

                DataTable dataTable = new DataTable();
                using (var adapter = new SqlDataAdapter(command))
                {
                    adapter.Fill(dataTable);
                }

                dataGridView1.DataSource = dataTable;
                dataGridView1.ReadOnly = true;
                dataGridView1.ClearSelection();
            }
        }

        private void FiltrarTarefas()
        {
            string familia = string.IsNullOrWhiteSpace(textBoxFamilia.Text.Trim()) ? null : textBoxFamilia.Text;
            string perfil = string.IsNullOrWhiteSpace(textBoxPerfil.Text.Trim()) ? null : textBoxPerfil.Text;
            string material = string.IsNullOrWhiteSpace(textBoxMaterial.Text.Trim()) ? null : textBoxMaterial.Text;

            ComunicaBDtekla BD = new ComunicaBDtekla();
            try
            {
                BD.ConectarBD(); 

                FiltrarDataGridViewAddTarefas(BD, familia, perfil, material);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Erro ao conectar à base de dados: " + ex.Message);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            textBoxFamilia.Text = string.Empty;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            textBoxPerfil.Text = string.Empty;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            textBoxMaterial.Text = string.Empty;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            textBoxFamilia.Text = string.Empty;
            textBoxPerfil.Text = string.Empty;
            textBoxMaterial.Text = string.Empty;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            ACTUALIZA();
        }
    }
}

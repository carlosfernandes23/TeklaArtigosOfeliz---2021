using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BDTEKLA
{
    public partial class Frm_BDImportaexcel : Form
    {
        public Frm_BDImportaexcel()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.

            if (result == DialogResult.OK) // Test result.
            {


                string pathName = openFileDialog1.FileName;
                string fileName = System.IO.Path.GetFileNameWithoutExtension(pathName);
                System.Data.DataTable tbContainer = new System.Data.DataTable();
                string strConn = string.Empty;
                string sheetName = fileName;

                FileInfo file = new FileInfo(pathName);
                if (!file.Exists) { throw new Exception("Error, file doesn't exists!"); }
                string extension = file.Extension;
                switch (extension)
                {
                    case ".xls":
                        strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + pathName + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'";
                        break;
                    case ".xlsx":
                        strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + pathName + ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=1;'";
                        break;
                    default:
                        strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + pathName + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'";
                        break;
                }
                OleDbConnection cnnxls = new OleDbConnection(strConn);
                OleDbDataAdapter oda = new OleDbDataAdapter(string.Format("select * from [Folha1$]"), cnnxls);
                oda.Fill(tbContainer);
                dataGridView1.DataSource = tbContainer;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //Data Source=GALILEU\PREPARACAO;Initial Catalog=ArtigoTekla;Persist Security Info=True;User ID=SA;Password=preparacao
            SqlConnection MiConexion = new SqlConnection("Data Source=GALILEU\\PREPARACAO;Initial Catalog=ArtigoTekla;Persist Security Info=True;User ID=SA;Password=preparacao");
            MiConexion.Open();
            //SqlCommand MiComando = new SqlCommand(Query, MiConexion);
           
            foreach (DataGridViewRow item in dataGridView1.Rows)
            {
                SqlCommand iComando = new SqlCommand("SELECT count(*)  FROM [dbo].[Perfilagem3] where [Perfil]='" + item.Cells[1].Value + "'AND[Espessura]='" + item.Cells[2].Value + "'", MiConexion);
                DialogResult seexistir = DialogResult.Yes;
                int FilaAfectada = (int)iComando.ExecuteScalar();
                if (FilaAfectada != 0)
                {
                    seexistir = MessageBox.Show(this, "Foram Detectados registos identicos ao que deseja inserir, deseja importar este registo?", "ERRO", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                }


                if (seexistir == DialogResult.Yes)
                {


                    if (item.Cells[0].Value != null)
                    {
                        SqlCommand MiComando = new SqlCommand("INSERT INTO[dbo].[Perfilagem]([Familia],[Perfil],[Espessura],[Artigo],[Peso],[LarguraUtil],[Dest],[Marca]) VALUES('" + item.Cells[0].Value + "', '" + item.Cells[1].Value + "', '" + item.Cells[2].Value + "', '" + item.Cells[3].Value + "', '" + item.Cells[4].Value.ToString().Replace(",", ".") + "', '" + int.Parse(item.Cells[5].Value.ToString()) + "', '" + item.Cells[6].Value.ToString() + "', '" + item.Cells[7].Value.ToString() + "')", MiConexion);

                        int FilasAfectadas = MiComando.ExecuteNonQuery();

                        if (FilasAfectadas > 0)
                        {

                        }
                        else
                        {
                            MessageBox.Show(this, "Erro na tranzação de dados :-(", "Erro no sistema", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
            MiConexion.Close();
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog2.ShowDialog(); // Show the dialog.

            if (result == DialogResult.OK) // Test result.
            {
                string pathName = openFileDialog2.FileName;
                string fileName = System.IO.Path.GetFileNameWithoutExtension(pathName);
               
                IEnumerable<string> txts  = File.ReadLines(pathName);
                dataGridView1.Columns.Add("1", "");
                dataGridView1.Columns.Add("2", "");
                dataGridView1.Columns.Add("3", "");
                dataGridView1.Columns.Add("4", "");
                dataGridView1.Columns.Add("5", "");
                dataGridView1.Columns.Add("6", "");
                dataGridView1.Columns.Add("7", "");
                dataGridView1.Columns.Add("8", "");
                dataGridView1.Columns.Add("9", "");
                dataGridView1.Columns.Add("10", "");

                foreach (string item in txts)
                {
                    dataGridView1.Rows.Add(item.Split(';').ToArray());
                }
               
            }
        }
    }
}

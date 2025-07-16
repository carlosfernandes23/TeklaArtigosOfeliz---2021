using Microsoft.IdentityModel.Abstractions;
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

namespace BDTEKLA
{
    public partial class Frm_ActualizaBD : Form
    {
        string connectionString = null;
        int idfinal;
        public Frm_ActualizaBD(int lastValue, string connectionString_)
        {
            InitializeComponent();
            idfinal = (lastValue + 1);
            label11.Text = (lastValue + 1).ToString();
            connectionString = connectionString_;
            button1.Text = "Criar";
            this.Text = "Criar Artigo";
        }

        public Frm_ActualizaBD(int id, string familia, string perfil, string material ,string esp, string artigo, string peso, string larutil, string destinatario, string marca, string connectionString_)
        {
            InitializeComponent();
            
            label11.Text = id.ToString().Trim();           
            textBox1.Text = familia.Trim();
            textBox2.Text = perfil.Trim();
            textBox3.Text = esp.Trim();
            textBox4.Text = artigo.Trim();
            textBox5.Text = peso.Trim();
            textBox6.Text = larutil.ToString().Trim();
            textBox7.Text = destinatario.Trim();
            textBox8.Text = marca.Trim();
            textBox9.Text = material.Trim();
            connectionString = connectionString_;
            button1.Text = "Actualiza";
            this.Text = "Actualiza Artigo";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();

            string comando = null;

            if (button1.Text == "Actualiza")
            {
                comando = "UPDATE [dbo].[Perfilagem3] set [Familia]='" + textBox1.Text + "',[Perfil]='" + textBox2.Text + "',[Material]='" + textBox9.Text + "',[Espessura]='" + textBox3.Text + "',[Artigo]='" + textBox4.Text + "',[Peso]=" + textBox5.Text.Replace(',', '.') + ",[LarguraUtil]=" + textBox6.Text + ",[Dest]='" + textBox7.Text + "',[Marca]='" + textBox8.Text + "' WHERE id =" + label11.Text;
            }
            else
            {
                comando = "INSERT INTO [dbo].[Perfilagem3]([Id], [Familia], [Perfil], [Material], [Espessura], [Artigo], [Peso], [LarguraUtil], [Dest], [Marca]) " +
                                 "VALUES('" + idfinal + "', '" + textBox1.Text + "', '" + textBox2.Text + "', '" + textBox9.Text + "', '" + textBox3.Text + "', '" + textBox4.Text + "', '" + textBox5.Text + "', '" + textBox6.Text + "', '" + textBox7.Text + "', '" + textBox8.Text + "')";
            }
            SqlCommand c = new SqlCommand(comando, connection);
            int nlinhas = c.ExecuteNonQuery();
            connection.Close();
            if (nlinhas > 0)
            {
                MessageBox.Show(this, "Dados actualizados com sucesso", "Actualizaçao bd", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(this, "OCURREU UM ERRO A ACTULIZAR A BD. FORAM AFECTADAS " + nlinhas + " o que podera de ter de voltar a actualizalas", "Actualizaçao bd", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            this.Close();
        }

        private void Frm_ActualizaBD_Load(object sender, EventArgs e)
        {

        }
    }
}

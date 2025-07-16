using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TSM = Tekla.Structures.Model;
using Tekla.Structures.Model;
using TSD = Tekla.Structures.Drawing;
using Tekla.Structures.Filtering;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Filtering.Categories;
using Tekla.Structures.Drawing;


namespace TeklaArtigosOfeliz
{
    public partial class Frm_Parametros : Form
    {
        public Frm_Parametros()
        {
            InitializeComponent();
            TopMost = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ComunicaTekla.EnviaproPriedadePeca(ComunicaTekla.ListadePecasSelec(), "comment", txt_commen.Text);
            if (ComunicaTekla.ListadePecasSelec().Count > 0)
            {
                MessageBox.Show(this, "Foi adicionado o comentário " + ComunicaTekla.ListadePecasSelec().Count + " Conjunto", "Êxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(this, "Por favor selecione pelo menos uma peca ", "SEM PEÇAS SELECIONADAS", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            ArrayList lista = new ArrayList(ComunicaTekla.ListadePecasSelec());
            ComunicaTekla.EnviaproPriedadePeca(lista, "Artigo", "");
            ComunicaTekla.EnviaproPriedadePeca(lista, "Artigo_interno", "");
            ComunicaTekla.EnviaproPriedadePeca(lista, "Destinata_ext", "");
            ComunicaTekla.EnviaproPriedadePeca(lista, "Operacoes", "");
            ComunicaTekla.EnviaproPriedadePeca(lista, "forcar_destino", cb_focardep.Text);


            if (lista.Count > 0)
            {
                MessageBox.Show(this, "Foi adicionado o comentário " + ComunicaTekla.ListadePecasSelec().Count + " Conjunto", "Êxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(this, "Por favor selecione pelo menos uma peca ", "SEM PEÇAS SELECIONADAS", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button25_Click(object sender, EventArgs e)
        {
            ComunicaTekla.EnviaproPriedadePeca(ComunicaTekla.ListadePecasSelec(), "Comentarioprep", textBox2.Text);
            if (ComunicaTekla.ListadePecasSelec().Count > 0)
            {
                MessageBox.Show(this, "Foi adicionado o comentário " + ComunicaTekla.ListadePecasSelec().Count + " Conjunto", "Êxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(this, "Por favor selecione pelo menos uma peca ", "SEM PEÇAS SELECIONADAS", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button33_Click(object sender, EventArgs e)
        {
            string req = null;

            Button b = (Button)sender;
            if (b.Name == "button33")
            {
                req = CbRequesitos.Text;
            }

            if (!string.IsNullOrEmpty(req.Trim()))
            {
                if (LblRequesitos.Text == "")
                {
                    LblRequesitos.Text = req;
                }
                else
                {
                    LblRequesitos.Text = LblRequesitos.Text + " | " + req;
                }
                CbRequesitos.Text = "";
            }
        }

        private void button35_Click(object sender, EventArgs e)
        {
            LblRequesitos.Text = "";
        }

        private void button34_Click(object sender, EventArgs e)
        {
            ArrayList PECAS = new ArrayList();
            PECAS = ComunicaTekla.ListadePecasSelec();
            ComunicaTekla.EnviaproPriedadePeca(PECAS, "Requisitos", LblRequesitos.Text);

            MessageBox.Show(this, "Foram adicionados os requisitos a " + PECAS.Count + " peças." + Environment.NewLine + LblRequesitos.Text, "INFORMAÇÃO", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void button37_Click(object sender, EventArgs e)
        {
            ArrayList part = new ArrayList();
            part = ComunicaTekla.ListadePecasSelec();
            CbRequesitos.Items.Clear();

            foreach (TSM.Part item in part)
            {
                comboBox1.Items.Clear();
                string Perfil = item.Profile.ProfileString.ToLower();
                List<string> ar = new List<string>();
                ComunicaBDtekla a = new ComunicaBDtekla();
                a.ConectarBD();
                string familia = "";
                try
                {
                    familia = a.Procurarbd("SELECT [familia] FROM [dbo].[Perfilagem3] WHERE [Perfil]='" + Perfil + "'").First();
                }
                catch (System.Exception)
                {
                    familia = "outros";
                }

                ar = a.Procurarbd("SELECT [Requisito] FROM [dbo].[Requisitos] WHERE [familia]='" + familia + "'");
                a.DesonectarBD();
                for (int i = 0; i < ar.Count; i++)
                {
                    if (!CbRequesitos.Items.Contains(ar[i].ToString().Trim()))
                    {
                        CbRequesitos.Items.Add(ar[i].ToString().Trim());
                    }
                }
            }
        }

        private void button31_Click(object sender, EventArgs e)
        {
            MessageBox.Show(this, "Duplo clic na label “Espessura Interior_Exterior” para carregar as espessuras possíveis do perfil selecionado.", "Informação", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void button32_Click(object sender, EventArgs e)
        {
            ArrayList PECAS = new ArrayList();
            PECAS = ComunicaTekla.ListadePecasSelec();
            List<string> ar = new List<string>();
            try
            {
                TSM.Part item = PECAS.Cast<TSM.Part>().ToList()[0];
                string Perfil = item.Profile.ProfileString.ToLower();
                string especuradaschapas = "0";
                item.GetUserProperty("Esp_chapa", ref especuradaschapas);

                ComunicaBDtekla a = new ComunicaBDtekla();
                a.ConectarBD();
                ar = a.Procurarbd("SELECT [Familia] FROM [dbo].[Perfilagem3] WHERE [Perfil]='" + Perfil + "'");
                a.DesonectarBD();
            }
            catch (System.Exception)
            {

            }
            //se for painel 
            if (ar[0].ToLower().Trim() == "painel")
            {
                //0,4EXT9006|0,4INT9010//////MaterialRevest
                string[] b = comboBox1.Text.Split('_');
                string c = b[0] + "EXT" + TXTralext.Text + "|" + b[1] + "INT" + TXTralint.Text;
                ComunicaTekla.EnviaproPriedadePeca(PECAS, "Esp_chapa", comboBox1.Text);
                ComunicaTekla.EnviaproPriedadePeca(PECAS, "Ralespcor", c);
                ComunicaTekla.EnviaproPriedadePeca(PECAS, "MaterialRevest", CbMaterial.Text);
            }

            //se for chapa perfilada 

            if (ar[0].ToLower().Trim() == "ch perfilada")
            {
                string inte = "";
                string ext = "";
                string c = "";
                if (TXTralint.Text.Trim() != "")
                {
                    inte = TXTralint.Text.Trim() + "_Face 2";
                }
                if (TXTralext.Text.Trim() != "")
                {
                    ext = TXTralext.Text.Trim() + "_Face 1";
                }
                if (inte != "" && ext != "")
                {
                    c = ext + "|" + inte;
                }
                else
                {
                    c = ext + inte;
                }
                ComunicaTekla.EnviaproPriedadePeca(PECAS, "Esp_chapa", comboBox1.Text);
                ComunicaTekla.EnviaproPriedadePeca(PECAS, "Ralespcor", c);
            }
            if (ar[0].ToLower().Trim() == "policarbonato")
            {

                ComunicaTekla.EnviaproPriedadePeca(PECAS, "Esp_chapa", comboBox1.Text);
                ComunicaTekla.EnviaproPriedadePeca(PECAS, "Ralespcor", Cbcorpoli.Text);
            }

            if (ComunicaTekla.ListadePecasSelec().Count > 0)
            {
                MessageBox.Show(this, "Foi adicionada a propriedade " + ComunicaTekla.ListadePecasSelec().Count + " peças", "Êxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(this, "Por favor selecione pelo menos uma peca ", "SEM PEÇAS SELECIONADAS", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button36_Click(object sender, EventArgs e)
        {
            ArrayList part = new ArrayList();
            part = ComunicaTekla.ListadePecasSelec();

            foreach (TSM.Part item in part)
            {
                comboBox1.Items.Clear();
                string Perfil = item.Profile.ProfileString.ToLower();
                string especuradaschapas = "0";
                item.GetUserProperty("Esp_chapa", ref especuradaschapas);
                List<string> ar = new List<string>();
                ComunicaBDtekla a = new ComunicaBDtekla();
                a.ConectarBD();
                ar = a.Procurarbd("SELECT [Espessura], [familia] FROM [dbo].[Perfilagem3] WHERE [Perfil]='" + Perfil + "'");
                a.DesonectarBD();
                for (int i = 0; i < ar.Count; i = i + 2)
                {
                    comboBox1.Visible = true;
                    label10.Visible = true;
                    button31.Visible = true;
                    button32.Visible = true;
                    if (!comboBox1.Items.Contains(ar[i].ToString().Trim()))
                    {
                        comboBox1.Items.Add(ar[i].ToString().Trim());
                    }
                }

                try
                {
                    if (ar[1].Trim().ToLower() == "ch perfilada")
                    {

                        comboBox1.SelectedIndex = 0;
                        CbMaterial.Visible = false;
                        label14.Visible = false;
                        lblpolicarbonato.Visible = false;
                        Cbcorpoli.Visible = false;
                        Cbcorpoli.Visible = false;
                        TXTralext.Visible = true;
                        TXTralint.Visible = true;
                        label12.Visible = true;
                        label13.Visible = true;
                    }
                    else if (ar[1].Trim().ToLower() == "painel")
                    {
                        comboBox1.SelectedIndex = 0;
                        CbMaterial.Visible = true;
                        label14.Visible = true;
                        lblpolicarbonato.Visible = false;
                        Cbcorpoli.Visible = false;
                        Cbcorpoli.Visible = false;
                        TXTralext.Visible = true;
                        TXTralint.Visible = true;
                        label12.Visible = true;
                        label13.Visible = true;

                        CbMaterial.Items.Clear();
                        a.ConectarBD();
                        ar = a.Procurarbd("SELECT [Material] FROM [dbo].[MateriaisSuplementares] WHERE [Familia]='" + ar[1].Trim() + "'");
                        a.DesonectarBD();
                        for (int i = 0; i < ar.Count; i++)
                        {
                            if (!CbMaterial.Items.Contains(ar[i].ToString().Trim()))
                            {
                                CbMaterial.Items.Add(ar[i].ToString().Trim());
                            }
                        }
                        CbMaterial.SelectedIndex = 0;
                    }
                    else if (ar[1].Trim().ToLower() == "policarbonato")
                    {
                        comboBox1.SelectedIndex = 0;
                        CbMaterial.Visible = false;
                        label14.Visible = true;
                        lblpolicarbonato.Visible = true;
                        Cbcorpoli.Visible = true;
                        TXTralext.Visible = false;
                        TXTralint.Visible = false;
                        label12.Visible = false;
                        label13.Visible = false;
                        Cbcorpoli.Items.Clear();
                        a.ConectarBD();
                        ar = a.Procurarbd("SELECT [Cor] FROM [dbo].[corpolicarbonato] WHERE [Familia]='" + ar[1].Trim() + "'");
                        a.DesonectarBD();
                        for (int i = 0; i < ar.Count; i++)
                        {
                            if (!CbMaterial.Items.Contains(ar[i].ToString().Trim()))
                            {
                                Cbcorpoli.Items.Add(ar[i].ToString().Trim());
                            }
                        }
                        Cbcorpoli.SelectedIndex = 0;
                    }
                }
                catch (System.Exception)
                {
                    Cbcorpoli.Visible = false;
                    comboBox1.Visible = false;
                    label10.Visible = false;
                    button31.Visible = false;
                    button32.Visible = false;
                    TXTralext.Visible = false;
                    TXTralint.Visible = false;
                    label12.Visible = false;
                    label13.Visible = false;
                    lblpolicarbonato.Visible = false;
                    label14.Visible = false;
                    CbMaterial.Visible = false;
                }
            }
        }

        private void button57_Click(object sender, EventArgs e)
        {

            ArrayList PECAS = ComunicaTekla.ListadePecasSelec();

            foreach (TSM.Part part in PECAS)
            {
                part.SetUserProperty("comprimentoteste", null);
            }
        }

        private void button56_Click(object sender, EventArgs e)
        {

            ArrayList PECAS = ComunicaTekla.ListadePecasSelec();

            foreach (TSM.Part part in PECAS)
            {

                int compgross = 0;
                double comp = 0;
                string resultado = null;
                if (part.Profile.ProfileString.Contains("CHG"))
                {
                    part.GetReportProperty("ASSEMBLY.WIDTH", ref comp);
                    part.SetUserProperty("comprimentoteste", comp.ToString("0"));
                }
                else
                {
                    part.GetReportProperty("ComprAdicional", ref compgross);
                    part.GetReportProperty("LENGTH", ref comp);

                    resultado = (compgross + comp).ToString("0") + " mm";
                    part.SetUserProperty("comprimentoteste", resultado);
                }

            }

        }

        private void button38_Click(object sender, EventArgs e)
        {
            ArrayList PECAS = new ArrayList();
            PECAS = ComunicaTekla.ListadePecasSelec();
            if (cBCOMPOUTRO.Checked)
            {
                ComunicaTekla.EnviaproPriedadePeca(PECAS, "ComprAdicional", txtcompradici.Text);
            }
            else if (cBCOMPINICFINAL.Checked)
            {
                foreach (TSM.Part part in PECAS)
                {
                    double compgross = 0;
                    double comp = 0;
                    part.GetReportProperty("LENGTH_GROSS", ref compgross);
                    part.GetReportProperty("LENGTH", ref comp);
                    part.SetUserProperty("ComprAdicional", Convert.ToInt32(compgross - comp));
                }
            }
            MessageBox.Show(this, "Foi adicionada a propriedade a " + PECAS.Count + " Peças");
        }

        private void button2_Click(object sender, EventArgs e)
        {

            ComunicaTekla.EnviaproPriedadeConj(ComunicaTekla.ListadeConjuntosSelec(), "USER_FIELD_1", Txt_des_client.Text);


            if (ComunicaTekla.ListadeConjuntosSelec().Count > 0)
            {
                MessageBox.Show(this, "Foi adicionada a Marca Do cliente a " + ComunicaTekla.ListadeConjuntosSelec().Count + " Conjunto", "Êxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(this, "Por favor selecione pelo menos um conjunto ", "SEM CONJUNTOS SELECIONADOS", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {
            string rui = null;
            ArrayList lista = ComunicaTekla.ListadePecasSelec();
            foreach (TSM.Part part in lista)
            {
                part.GetUserProperty("comment", ref rui);
                if (!string.IsNullOrEmpty(rui))
                {
                    MessageBox.Show(this, rui.ToString());
                }
            }
        }

        private void label15_Click(object sender, EventArgs e)
        {
            string rui = null;
            ArrayList lista = ComunicaTekla.ListadePecasSelec();
            foreach (TSM.Part part in lista)
            {
                part.GetUserProperty("Requisitos", ref rui);
                if (!string.IsNullOrEmpty(rui))
                {
                    MessageBox.Show(this, rui.ToString());
                }
            }
        }

        private void Txt_des_client_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == Convert.ToChar(Keys.Enter))
            {
                ComunicaTekla.EnviaproPriedadeConj(ComunicaTekla.ListadeConjuntosSelec(), "USER_FIELD_1", Txt_des_client.Text);


                if (ComunicaTekla.ListadeConjuntosSelec().Count > 0)
                {
                    MessageBox.Show(this, "Foi adicionada a Marca Do cliente a " + ComunicaTekla.ListadeConjuntosSelec().Count + " Conjunto", "Êxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(this, "Por favor selecione pelo menos um conjunto ", "SEM CONJUNTOS SELECIONADOS", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void label2_Click(object sender, EventArgs e)
        {
            string rui = null;
            ArrayList lista = ComunicaTekla.ListadeConjuntosSelec();
            foreach (TSM.Assembly part in lista)
            {
                part.GetUserProperty("USER_FIELD_1", ref rui);
                if (rui != null)
                {
                    MessageBox.Show(this, rui.ToString());
                }
            }
        }

        private void TXTralext_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == Convert.ToChar(Keys.Enter))
            {
                TXTralint.Select();
            }
        }

        private void TXTralint_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == Convert.ToChar(Keys.Enter))
            {
                TXTralext.Select();
            }
        }
        private void cBCOMPINICFINAL_Click(object sender, EventArgs e)
        {
            cBCOMPOUTRO.Checked = false;
        }

        private void cBCOMPOUTRO_Click(object sender, EventArgs e)
        {
            cBCOMPINICFINAL.Checked = false;
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
    }
}

using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Newtonsoft.Json;

namespace TeklaArtigosOfeliz
{
    public partial class Frm_AtualizarEmails: Form
    {
        public Frm_AtualizarEmails()
        {
            InitializeComponent();
            LoadData();
        }

        private void Frm_AtualizarEmails_Load(object sender, EventArgs e)
        {
            comboBoxemailtipo.Items.AddRange(new string[]
          {
                    "Fabrico",
                    "Lotear",
                    "Quantificação",
                    "Aprovisionamentos",
                    "Powerfab",
                    "Revisões Internas",
                    "Revisões Externas",
          });

            listBoxPara.SelectionMode = SelectionMode.One;
            listBoxCC.SelectionMode = SelectionMode.One;

            string user = Environment.UserName;
            if (user == "carlos.alves" || user == "luis.silva" || user == "helder.silva")
            {
                this.Size = new Size(955, 285);
                this.MinimumSize = new Size(955, 285);
                this.MaximumSize = new Size(955, 285);
            }
            else
            {
                this.Size = new Size(440, 285);
                this.MinimumSize = new Size(442, 285);
                this.MaximumSize = new Size(442, 285);
            }
        }           

        private void LoadData()
        {
            string jsonFilePath = @"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\4 Produção\Desenvolvimentos\Ficheiros Temp tekla artigos (Nao Apagar)\Diretor de Obra Base de dados\DiretordeObra.json";

            List<string> nomes = LoadNamesFromJson(jsonFilePath);
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("Nome"); 

            foreach (var nome in nomes)
            {
                dataTable.Rows.Add(nome); 
            }
            dataGridView1.DataSource = dataTable;
            dataGridView1.RowHeadersVisible = false;
            dataGridView1.Columns["Nome"].Width = 170;

        }

        private List<string> LoadNamesFromJson(string filePath)
        {
            string json = File.ReadAllText(filePath);
            return JsonConvert.DeserializeObject<List<string>>(json);
        }

        private void SaveNamesToJson(string filePath, List<string> nomes)
        {
            string json = JsonConvert.SerializeObject(nomes, Formatting.Indented);
            File.WriteAllText(filePath, json);
        }

        private void guna2Button4_Click(object sender, EventArgs e)
        {
            string jsonFilePath = @"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\4 Produção\Desenvolvimentos\Ficheiros Temp tekla artigos (Nao Apagar)\Diretor de Obra Base de dados\DiretordeObra.json";

            string novoNome = textBox1.Text.Trim();
            if (!string.IsNullOrEmpty(novoNome))
            {
                List<string> nomes = LoadNamesFromJson(jsonFilePath);
                nomes.Add(novoNome);
                SaveNamesToJson(jsonFilePath, nomes);
                LoadData();
                textBox1.Clear();
            }
            else
            {
                MessageBox.Show(this, "Por favor, insira um nome válido.");
            }
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            string jsonFilePath = @"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\4 Produção\Desenvolvimentos\Ficheiros Temp tekla artigos (Nao Apagar)\Diretor de Obra Base de dados\DiretordeObra.json";

            if (dataGridView1.SelectedCells.Count > 0)
            {
                int linhaSelecionada = dataGridView1.SelectedCells[0].RowIndex;
                string nomeSelecionado = dataGridView1.Rows[linhaSelecionada].Cells[0].Value.ToString();
                List<string> nomes = LoadNamesFromJson(jsonFilePath);

                if (nomes.Remove(nomeSelecionado))
                {
                    SaveNamesToJson(jsonFilePath, nomes);
                    dataGridView1.Rows.RemoveAt(linhaSelecionada);
                    LoadData();
                }
                else
                {
                    MessageBox.Show(this, "Nome não encontrado na lista.");
                }
            }
            else
            {
                MessageBox.Show(this, "Por favor, selecione uma célula para remover.");
            }
        }

        private void CarregarEmails(string tipo, string ficheiro)
        {
            labelEmail.Text = $"Email {tipo}";

            string basePath = @"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\4 Produção\Desenvolvimentos\Ficheiros Temp tekla artigos (Nao Apagar)\Diretor de Obra Base de dados\";
            string jsonParaPath = Path.Combine(basePath, $"Email{ficheiro}Para.json");
            string jsonCCPath = Path.Combine(basePath, $"Email{ficheiro}CC.json");

            listBoxPara.DataSource = LerEmailsDeJson(jsonParaPath);
            listBoxPara.ClearSelected();
            listBoxCC.DataSource = LerEmailsDeJson(jsonCCPath);
            listBoxCC.ClearSelected();
        }

        private List<string> LerEmailsDeJson(string caminho)
        {
            if (File.Exists(caminho))
            {
                string json = File.ReadAllText(caminho);
                return JsonConvert.DeserializeObject<List<string>>(json) ?? new List<string>();
            }
            return new List<string> { "Para selecionar Diretor de Obra" };
        }

        private void comboBoxemailtipo_SelectedIndexChanged(object sender, EventArgs e)
        {
            listBoxPara.DataSource = null;
            listBoxCC.DataSource = null;

            if (comboBoxemailtipo.Text == "Fabrico")
            {
                CarregarEmails("Fabrico", "Fabrico");
            }
            else if (comboBoxemailtipo.Text == "Lotear")
            {
                CarregarEmails("Lotear", "Lotear");
            }
            else if (comboBoxemailtipo.Text == "Quantificação")
            {
                CarregarEmails("Quantificacao", "Quantificacao");
            }
            else if (comboBoxemailtipo.Text == "Aprovisionamentos")
            {
                CarregarEmails("Aprovisionamentos", "Aprovisionamentos");
            }
            else if (comboBoxemailtipo.Text == "Powerfab")
            {
                CarregarEmails("Powerfab", "Powerfab");
            }
            else if (comboBoxemailtipo.Text == "Revisões Internas")
            {
                CarregarEmails("Revisões Internas", "RevInt");
            }
            else if (comboBoxemailtipo.Text == "Revisões Externas")
            {
                CarregarEmails("Revisões Externas", "RevExt");
            }
        }

        private void AdicionarEmailAoFicheiro(string caminho, string novoEmail)
        {
            List<string> emails = LerEmailsDeJson(caminho);

            if (!emails.Contains(novoEmail, StringComparer.OrdinalIgnoreCase))
            {
                emails.Add(novoEmail);
                string json = JsonConvert.SerializeObject(emails, Formatting.Indented);
                File.WriteAllText(caminho, json);
            }
            else
            {
                MessageBox.Show(this, "Este email já existe na lista.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void RemoverEmailDoFicheiro(string caminho, string emailParaRemover)
        {
            List<string> emails = LerEmailsDeJson(caminho);

            if (emails.Remove(emailParaRemover))
            {
                string json = JsonConvert.SerializeObject(emails, Formatting.Indented);
                File.WriteAllText(caminho, json);
            }
            else
            {
                MessageBox.Show(this, "O email não foi encontrado na lista.");
            }
        }

        private void guna2Button2_Click(object sender, EventArgs e)
        {
            string novoEmail = textBoxEmail.Text.Trim();

            if (string.IsNullOrEmpty(novoEmail))
            {
                MessageBox.Show(this, "Insira um email válido.");
                return;
            }

            string tipoSelecionado = comboBoxemailtipo.Text;

            if (string.IsNullOrEmpty(tipoSelecionado))
            {
                MessageBox.Show(this, "Selecione um tipo de email.");
                return;
            }
            string basePath = @"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\4 Produção\Desenvolvimentos\Ficheiros Temp tekla artigos (Nao Apagar)\Diretor de Obra Base de dados\";

            string ficheiroTipo = "";

            if (comboBoxemailtipo.Text == "Fabrico")
            {
                ficheiroTipo = "Fabrico";
            }
            else if (comboBoxemailtipo.Text == "Lotear")
            {
                ficheiroTipo = "Lotear";
            }
            else if (comboBoxemailtipo.Text == "Quantificação")
            {
                ficheiroTipo = "Quantificacao";
            }
            else if (comboBoxemailtipo.Text == "Aprovisionamentos")
            {
                ficheiroTipo = "Aprovisionamentos";
            }
            else if (comboBoxemailtipo.Text == "Powerfab")
            {
                ficheiroTipo = "Powerfab";
            }
            else if (comboBoxemailtipo.Text == "Revisões Internas")
            {
                ficheiroTipo = "RevInt";
            }
            else if (comboBoxemailtipo.Text == "Revisões Externas")
            {
                ficheiroTipo = "RevExt";
            }

            if (checkBoxPara.Checked)
            {
                string jsonParaPath = Path.Combine(basePath, $"Email{ficheiroTipo}Para.json");
                AdicionarEmailAoFicheiro(jsonParaPath, novoEmail);
            }
            if (checkBoxCC.Checked)
            {
                string jsonCCPath = Path.Combine(basePath, $"Email{ficheiroTipo}CC.json");
                AdicionarEmailAoFicheiro(jsonCCPath, novoEmail);
            }
            CarregarEmails(ficheiroTipo, ficheiroTipo); 
            textBoxEmail.Clear();
            checkBoxPara.Checked = false;
            checkBoxCC.Checked = false;
        }

        private void guna2Button3_Click(object sender, EventArgs e)
        {
            string tipoSelecionado = comboBoxemailtipo.Text;

            if (string.IsNullOrEmpty(tipoSelecionado))
            {
                MessageBox.Show(this, "Selecione um tipo de email.");
                return;
            }

            string ficheiroTipo = "";

            if (comboBoxemailtipo.Text == "Fabrico")
            {
                ficheiroTipo = "Fabrico";
            }
            else if (comboBoxemailtipo.Text == "Lotear")
            {
                ficheiroTipo = "Lotear";
            }
            else if (comboBoxemailtipo.Text == "Quantificação")
            {
                ficheiroTipo = "Quantificacao";
            }
            else if (comboBoxemailtipo.Text == "Aprovisionamentos")
            {
                ficheiroTipo = "Aprovisionamentos";
            }
            else if (comboBoxemailtipo.Text == "Powerfab")
            {
                ficheiroTipo = "Powerfab";
            }
            else if (comboBoxemailtipo.Text == "Revisões Internas")
            {
                ficheiroTipo = "RevInt";
            }
            else if (comboBoxemailtipo.Text == "Revisões Externas")
            {
                ficheiroTipo = "RevExt";
            }

                string basePath = @"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\4 Produção\Desenvolvimentos\Ficheiros Temp tekla artigos (Nao Apagar)\Diretor de Obra Base de dados\";

            bool itemSelecionado = false;

            if (listBoxPara.SelectedItem != null)
            {
                string emailSelecionado = listBoxPara.SelectedItem.ToString();
                string caminho = Path.Combine(basePath, $"Email{ficheiroTipo}Para.json");
                RemoverEmailDoFicheiro(caminho, emailSelecionado);
                itemSelecionado = true;
            }

            if (listBoxCC.SelectedItem != null)
            {
                string emailSelecionado = listBoxCC.SelectedItem.ToString();
                string caminho = Path.Combine(basePath, $"Email{ficheiroTipo}CC.json");
                RemoverEmailDoFicheiro(caminho, emailSelecionado);
                itemSelecionado = true;
            }

            if (!itemSelecionado)
            {
                MessageBox.Show(this, "Selecione um email em 'Para' ou 'CC' para remover.");
                return;
            }

            CarregarEmails(tipoSelecionado, ficheiroTipo);
        }

        private void checkBoxPara_Click(object sender, EventArgs e)
        {
            if (checkBoxPara.Checked)
            {
                checkBoxCC.Checked = false;
            }
        }

        private void checkBoxCC_Click(object sender, EventArgs e)
        {
            if (checkBoxCC.Checked)
            {
                checkBoxPara.Checked = false;
            }
        }
    }
}

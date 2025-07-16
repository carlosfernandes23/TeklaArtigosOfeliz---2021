using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net.Mail;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;
using CefSharp.WinForms;
using CefSharp;
using System.Net.Mime;
using System.IO;
using Tekla.Structures.Model;
using Newtonsoft.Json;
using HtmlAgilityPack;
using HtmlDocument = HtmlAgilityPack.HtmlDocument;
using Microsoft.Web.WebView2.WinForms;
using Microsoft.Web.WebView2.Core;
using Guna.Charts.WinForms;
using Microsoft.Identity.Client;
using System.Net.Http;
using Render;
using System.Web.UI.WebControls;
using com.itextpdf.text.pdf;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using iTextSharp.text.pdf.codec;
using Guna.UI2.WinForms;
using Point = System.Drawing.Point;
using Color = System.Drawing.Color;


namespace TeklaArtigosOfeliz
{
    public partial class Frm_Corpo_de_Texto_Email_Fabrico: Form
    {

        private string corpoEmail;
        private string caminhoImagem;
        private string Subject;
        private string tempImagePath;
        private string userDirectory;
        private string passwordFilePath;
        private string lote;
        private string dataObra;
        private string textbox1;

        public Frm_Corpo_de_Texto_Email_Fabrico(string titulo, string corpoEmail, string Subject, string tempImagePath, string lote, string DataObra, string textbox1)
        {
            InitializeComponent();
            this.Text = titulo;
            this.corpoEmail = corpoEmail;
            this.Subject = Subject;
            this.tempImagePath = tempImagePath;
            this.lote = lote;
            this.dataObra = DataObra;
            this.textbox1 = textbox1;
        }

        private void Corpo_de_Texto_Email_Fabrico_Load(object sender, EventArgs e)
        {
            textBoxAsu.Text = Subject;
            webBrowser1.DocumentText = corpoEmail;
            CarregarDiretorObra();
            CarregarEmailsCC();
            CarregarEmailsPara();
            ObterTamanhodaImagem();

            string user = Environment.UserName;
            labelemailuser.Text = user + "@ofeliz.com";

            Model modelo = new Model();
            string numeroobra = modelo.GetProjectInfo().ProjectNumber;
            label2.Text = numeroobra;
        }

        private void CarregarEmailsCC()
        {
            string caminho = @"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\4 Produção\Desenvolvimentos\Ficheiros Temp tekla artigos (Nao Apagar)\Diretor de Obra Base de dados\EmailFabricoCC.json";

            try
            {
                if (File.Exists(caminho))
                {
                    string json = File.ReadAllText(caminho);

                    var emails = JsonConvert.DeserializeObject<List<string>>(json);

                    listBoxCC.Items.Clear();
                    foreach (var email in emails)
                    {
                        listBoxCC.Items.Add(email);
                    }
                }
                else
                {
                    MessageBox.Show("Ficheiro dos email CC Não encontrado.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao carregar ficheiro: " + ex.Message);
            }
        }

        private void CarregarEmailsPara()
        {
            string caminho = @"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\4 Produção\Desenvolvimentos\Ficheiros Temp tekla artigos (Nao Apagar)\Diretor de Obra Base de dados\EmailFabricoPara.json";

            try
            {
                if (File.Exists(caminho))
                {
                    string json = File.ReadAllText(caminho);

                    var emails = JsonConvert.DeserializeObject<List<string>>(json);

                    listBoxPara.Items.Clear();
                    foreach (var email in emails)
                    {
                        listBoxPara.Items.Add(email);
                    }
                }
                else
                {
                    MessageBox.Show("Ficheiro dos email para Não encontrado.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao carregar ficheiro: " + ex.Message);
            }
        }

        private void CarregarDiretorObra()
        {
            comboBoxDiretorObra.Items.Clear();
            string jsonFilePath = @"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\4 Produção\Desenvolvimentos\Ficheiros Temp tekla artigos (Nao Apagar)\Diretor de Obra Base de dados\DiretordeObra.json";
            List<string> nomes = LoadNamesFromJson(jsonFilePath);
            foreach (var nome in nomes)
            {
                comboBoxDiretorObra.Items.Add(nome);
            }
        }

        private void ObterTamanhodaImagem()
        {
            using (System.Drawing.Image tempImage = System.Drawing.Image.FromFile(tempImagePath))
            {
                label11.Text = tempImage.Width.ToString(); 
                label12.Text = tempImage.Height.ToString(); 
            }
        }
               
        private List<string> LoadNamesFromJson(string filePath)
        {
            try
            {
                string json = File.ReadAllText(filePath);

                List<string> nomes = JsonConvert.DeserializeObject<List<string>>(json);

                return nomes;
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Erro ao carregar o arquivo JSON: " + ex.Message);
                return new List<string>();
            }
        }            
                
        private string GetSaudacao()
        {
            DateTime horaAtual = DateTime.Now;
            if (horaAtual.Hour < 12 || (horaAtual.Hour == 12 && horaAtual.Minute < 30))
            {
                return "Bom Dia, ";
            }
            else
            {
                return "Boa Tarde, ";
            }
        }

        private string GetEmailDiretorObra()
        {
            string emailDiretorObra = string.Empty;

            if (comboBoxDiretorObra.SelectedItem != null)
            {
                string nome = comboBoxDiretorObra.SelectedItem.ToString();
                emailDiretorObra = ConvertNomeToEmail(nome) + "@ofeliz.com";
            }
            else
            {
                string nomeInseridoManualmente = textBoxEmailCC.Text.Trim();
                if (!string.IsNullOrEmpty(nomeInseridoManualmente))
                {
                    emailDiretorObra = ConvertNomeToEmail(nomeInseridoManualmente) + "@ofeliz.com";
                }
                else
                {
                    labelEmailInformação.Text = "Por favor, insira o nome do diretor de obra.";
                }
            }
            return emailDiretorObra;
        }

        private string ConvertNomeToEmail(string nome)
        {
            var palavras = nome.Split(' ');
            return string.Join(".", palavras.Select(p => p.ToLower()));
        }

        private void ResizeImage(string originalFilePath, string newFilePath, int width, int height)
        {
            using (System.Drawing.Image originalImage = System.Drawing.Image.FromFile(originalFilePath))
            {
                using (Bitmap resizedImage = new Bitmap(originalImage, new Size(width, height)))
                {
                    resizedImage.Save(newFilePath, System.Drawing.Imaging.ImageFormat.Png);
                }
            }
        }

        private string GetNextResizedImagePath(string basePath)
        {
            int index = 1;
            string resizedImagePath;

            do
            {
                resizedImagePath = $"{basePath}imagem_redimensionada_{index}.png";
                index++;
            } while (File.Exists(resizedImagePath));

            return resizedImagePath;
        }

        bool IsValidEmail(string email)
        {
            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return addr.Address == email;
            }
            catch
            {
                return false;
            }
        }

        public static async Task<string> GetAccessTokenAsync()
        {
            var clientId = "0f37e406-80fc-4deb-9635-a20c4a22c53e";
            var tenantId = "67345170-c562-4f1e-aef4-cf8d2d06067f";
            var redirectUri = "http://localhost:61658";

            var publicClientApplication = PublicClientApplicationBuilder.Create(clientId)
                .WithAuthority(AzureCloudInstance.AzurePublic, tenantId)
                .WithRedirectUri(redirectUri)
                .Build();

            var tokenCache = publicClientApplication.UserTokenCache;
            string cacheFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "msal_cache.dat");

            tokenCache.SetBeforeAccess(args =>
            {
                args.TokenCache.DeserializeMsalV3(File.Exists(cacheFilePath) ? File.ReadAllBytes(cacheFilePath) : null);
            });

            tokenCache.SetAfterAccess(args =>
            {
                if (args.HasStateChanged)
                {
                    File.WriteAllBytes(cacheFilePath, args.TokenCache.SerializeMsalV3());
                }
            });

            var accounts = await publicClientApplication.GetAccountsAsync();
            var account = accounts.FirstOrDefault();

            if (account == null)
            {
                var result = await publicClientApplication
                    .AcquireTokenInteractive(new[] { "Mail.Send" })
                    .ExecuteAsync();

                return result.AccessToken;
            }

            try
            {
                var result = await publicClientApplication
                    .AcquireTokenSilent(new[] { "Mail.Send" }, account)
                    .ExecuteAsync();

                return result.AccessToken;
            }
            catch (MsalUiRequiredException)
            {
                var result = await publicClientApplication
                    .AcquireTokenInteractive(new[] { "Mail.Send" })
                    .ExecuteAsync();

                return result.AccessToken;
            }
        }

        public async System.Threading.Tasks.Task SendEmailAsync()
        {
            var accessToken = await GetAccessTokenAsync();
            var subject = textBoxAsu.Text;
            string saudacao = GetSaudacao();
            string nomeUsuario = Environment.UserName;
            nomeUsuario = nomeUsuario.Replace('.', ' ');
            nomeUsuario = string.Join(" ", nomeUsuario.Split(' ').Select(p => char.ToUpper(p[0]) + p.Substring(1).ToLower()));
            string imagemOfelizFilePath = @"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\4 Produção\Desenvolvimentos\Ficheiros Temp tekla artigos (Nao Apagar)\ofeliz_logo.png";       
                                       
            string corpoEmail = @"
                                    <html>
                                    <body contenteditable='false'>
                                        <font face='Calibri' size='3'>
                                            <p>" + saudacao + @"</p>
                                        </font>
                                        <font face='Calibri' size='3'>
                                            <p>Material pronto para fabrico (<span style='color:red;'><u>" + textbox1 + @"</u></span>).</p>
                                        </font>
                                        <font face='Calibri' size='3'>
                                            <p><span style='color:red;'><u>Lote " + lote + @"&nbsp;: " + dataObra + @"</u></span></p>
                                        </font>
                                        <img src='file:///" + tempImagePath.Replace("\\", "/") + @"' width='755' />
                                        <font face='Calibri' size='3'>
                                            <p>Melhores Cumprimentos,</p>
                                        </font>
                                        <font face='Calibri' size='3'>
                                            <b>" + nomeUsuario + @"</b><br>
                                            Construção Metálica | Preparador<br>
                                            T + 351 253 080 609 *
                                        </font>
                                        <font color='red' face='Calibri' size='3'>ofeliz.com</font><br>
                                        <p>
                                            <a href='https://www.ofeliz.com'>
                                                <img src='file:///" + imagemOfelizFilePath.Replace("\\", "/") + @"' width='127' height='34' alt='Logo da O Feliz'>
                                            </a>
                                        </p>
                                        <i>
                                            <font color='LightGrey' face='Calibri' size='1.5'>Alvará Nº 10553 – Pub. *Chamada para a rede fixa nacional.</font>
                                        </i><br>
                                        <i>
                                            <font color='green' face='Calibri' size='1.5'>Antes de imprimir este e-mail, tenha em consideração o meio ambiente.</font>
                                        </i><br>
                                    </body>
                                    </html>";

            var ccRecipients = listBoxCC.Items
                                  .Cast<string>()
                                  .Select(email => email.Trim())
                                  .Where(email => !string.IsNullOrEmpty(email))
                                  .Where(email => IsValidEmail(email))
                                  .ToList();

            var toRecipientsList = listBoxPara.Items
                .Cast<string>()
                .Select(email => email.Trim())
                .Where(email => !string.IsNullOrEmpty(email))
                .Where(email => IsValidEmail(email))
                .ToList();

            var ccRecipientsFormatted = ccRecipients
                .Select(email => new { EmailAddress = new { Address = email } })
                .ToArray();

            var toRecipients = toRecipientsList
                .Select(email => new { EmailAddress = new { Address = email } })
                .ToArray();

            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Add("Authorization", $"Bearer {accessToken}");
                client.DefaultRequestHeaders.Add("Accept", "application/json");

                var emailMessage = new
                {
                    Message = new
                    {
                        Subject = subject,
                        Body = new
                        {
                            ContentType = "HTML",
                            Content = corpoEmail
                        },
                        ToRecipients = toRecipients,
                        CcRecipients = ccRecipientsFormatted
                    },
                    SaveToSentItems = true
                };

                var jsonContent = Newtonsoft.Json.JsonConvert.SerializeObject(emailMessage);
                var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

                var response = await client.PostAsync("https://graph.microsoft.com/v1.0/me/sendMail", content);

                if (response.IsSuccessStatusCode)
                {
                    CustomGradientPanel4.Visible = true;
                    labelEmailInformação.ForeColor = Color.White;
                    labelEmailInformação.Location = new Point(60, 8);

                    labelEmailInformação.Text = "E-mail enviado com sucesso!";
                    await System.Threading.Tasks.Task.Delay(3000);
                    this.Close();
                }
                else
                {
                    var errorResponse = await response.Content.ReadAsStringAsync();
                    CustomGradientPanel4.Visible = true;
                    CustomGradientPanel4.FillColor = Color.Maroon;
                    CustomGradientPanel4.FillColor2 = Color.Maroon;
                    CustomGradientPanel4.FillColor3 = Color.Maroon;
                    CustomGradientPanel4.FillColor4 = Color.Maroon;
                    labelEmailInformação.Text = "Erro ao enviar o e-mail!";
                    MessageBox.Show(this, $"Erro ao enviar o e-mail: {response.StatusCode} - {errorResponse}");
                }
            }
        }

        public async System.Threading.Tasks.Task SendEmailAsyncrezize()
        {
            var accessToken = await GetAccessTokenAsync();
            var subject = textBoxAsu.Text;
            string saudacao = GetSaudacao();
            string nomeUsuario = Environment.UserName;
            nomeUsuario = nomeUsuario.Replace('.', ' ');
            nomeUsuario = string.Join(" ", nomeUsuario.Split(' ').Select(p => char.ToUpper(p[0]) + p.Substring(1).ToLower()));
            string imagemOfelizFilePath = @"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\4 Produção\Desenvolvimentos\Ficheiros Temp tekla artigos (Nao Apagar)\ofeliz_logo.png";

            if (string.IsNullOrWhiteSpace(textBoxWidth.Text) || string.IsNullOrWhiteSpace(textBoxheight.Text))
            {
                MessageBox.Show(this, "Por favor, preencha ambos os campos de largura e altura.");
                return; 
            }

            int newWidth = int.Parse(textBoxWidth.Text);
            int newHeight = int.Parse(textBoxheight.Text);                     
            string basePath = @"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\4 Produção\Desenvolvimentos\Ficheiros Temp tekla artigos (Nao Apagar)\Fabric\";
            string resizedImagePath = GetNextResizedImagePath(basePath);

            ResizeImage(tempImagePath, resizedImagePath, newWidth, newHeight);

            string corpoEmail = @"
                                            <html>
                                                <body contenteditable='false'>
                                                    <font face='Calibri' size='3'>
                                                        <p>" + saudacao + @"</p>
                                                    </font>
                                                    <font face='Calibri' size='3'>
                                                        <p>Material pronto para fabrico (<span style='color:red;'><u>" + textbox1 + @"</u></span>).</p>
                                                    </font>
                                                    <font face='Calibri' size='3'>
                                                        <p><span style='color:red;'><u>Lote " + lote + @"&nbsp;: " + dataObra + @"</u></span></p>
                                                    </font>
                                                    <img src='file:///" + resizedImagePath.Replace("\\", "/") + @"' />
                                                    <font face='Calibri' size='3'>
                                                        <p>Melhores Cumprimentos,</p>
                                                    </font>
                                                    <font face='Calibri' size='3'>
                                                        <b>" + nomeUsuario + @"</b><br>
                                                        Construção Metálica | Preparador<br>
                                                        T + 351 253 080 609 *
                                                    </font>
                                                    <font color='red' face='Calibri' size='3'>ofeliz.com</font><br>
                                                    <p>
                                                        <a href='https://www.ofeliz.com'>
                                                            <img src='file:///" + imagemOfelizFilePath.Replace("\\", "/") + @"' width='127' height='34' alt='Logo da O Feliz'>
                                                        </a>
                                                    </p>
                                                    <i>
                                                        <font color='LightGrey' face='Calibri' size='1.5'>Alvará Nº 10553 – Pub. *Chamada para a rede fixa nacional.</font>
                                                    </i><br>
                                                    <i>
                                                        <font color='green' face='Calibri' size='1.5'>Antes de imprimir este e-mail, tenha em consideração o meio ambiente.</font>
                                                    </i><br>
                                                </body>
                                            </html>";

            var ccRecipients = listBoxCC.Items
                                   .Cast<string>()
                                   .Select(email => email.Trim())
                                   .Where(email => !string.IsNullOrEmpty(email))
                                   .Where(email => IsValidEmail(email))
                                   .ToList();

            var toRecipientsList = listBoxPara.Items
                .Cast<string>()
                .Select(email => email.Trim())
                .Where(email => !string.IsNullOrEmpty(email))
                .Where(email => IsValidEmail(email))
                .ToList();

            var ccRecipientsFormatted = ccRecipients
                .Select(email => new { EmailAddress = new { Address = email } })
                .ToArray();

            var toRecipients = toRecipientsList
                .Select(email => new { EmailAddress = new { Address = email } })
                .ToArray();

            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Add("Authorization", $"Bearer {accessToken}");
                client.DefaultRequestHeaders.Add("Accept", "application/json");

                var emailMessage = new
                {
                    Message = new
                    {
                        Subject = subject,
                        Body = new
                        {
                            ContentType = "HTML",
                            Content = corpoEmail
                        },
                        ToRecipients = toRecipients,
                        CcRecipients = ccRecipientsFormatted
                    },
                    SaveToSentItems = true
                };

                var jsonContent = Newtonsoft.Json.JsonConvert.SerializeObject(emailMessage);
                var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

                var response = await client.PostAsync("https://graph.microsoft.com/v1.0/me/sendMail", content);

                if (response.IsSuccessStatusCode)
                {
                    CustomGradientPanel4.Visible = true;                  
                    labelEmailInformação.ForeColor = Color.White;
                    labelEmailInformação.Location = new Point(60, 8);

                    labelEmailInformação.Text = "E-mail enviado com sucesso!";
                    await System.Threading.Tasks.Task.Delay(3000);
                    this.Close();
                }
                else
                {
                    var errorResponse = await response.Content.ReadAsStringAsync();
                    CustomGradientPanel4.Visible = true;
                    CustomGradientPanel4.FillColor = Color.Maroon;
                    CustomGradientPanel4.FillColor2 = Color.Maroon;
                    CustomGradientPanel4.FillColor3 = Color.Maroon;
                    CustomGradientPanel4.FillColor4 = Color.Maroon;
                    labelEmailInformação.Text = "Erro ao enviar o e-mail!";
                    MessageBox.Show(this, $"Erro ao enviar o e-mail: {response.StatusCode} - {errorResponse}");
                }
            }
        }

        private void RedimencionarImagem()
        {
            try
            {
                string saudacao = GetSaudacao();
                string nomeUsuario = Environment.UserName;
                nomeUsuario = nomeUsuario.Replace('.', ' ');
                nomeUsuario = string.Join(" ", nomeUsuario.Split(' ').Select(p => char.ToUpper(p[0]) + p.Substring(1).ToLower()));

                string imagemOfelizFilePath = @"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\4 Produção\Desenvolvimentos\Ficheiros Temp tekla artigos (Nao Apagar)\ofeliz_logo.png";

                if (string.IsNullOrWhiteSpace(textBoxWidth.Text) || string.IsNullOrWhiteSpace(textBoxheight.Text))
                {
                    MessageBox.Show(this, "Por favor, preencha ambos os campos de largura e altura.");
                    return;
                }

                int newWidth = int.Parse(textBoxWidth.Text);
                int newHeight = int.Parse(textBoxheight.Text);

                string basePath = @"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\4 Produção\Desenvolvimentos\Ficheiros Temp tekla artigos (Nao Apagar)\Fabric\";

                string resizedImagePath = GetNextResizedImagePath(basePath);

                ResizeImage(tempImagePath, resizedImagePath, newWidth, newHeight);

                string corpoEmailnovo = @"
                                        <html>
                                        <body contenteditable='false'>
                                            <font face='Calibri' size='3'>
                                                <p>" + saudacao + @"</p>
                                            </font>
                                            <font face='Calibri' size='3'>
                                                <p>Material pronto para fabrico (<span style='color:red;'><u>" + textbox1 + @"</u></span>).</p>
                                            </font>
                                            <font face='Calibri' size='3'>
                                                <p><span style='color:red;'><u>Lote " + lote + @"&nbsp;: " + dataObra + @"</u></span></p>
                                            </font>
                                            <img src='file:///" + resizedImagePath.Replace("\\", "/") + @"' width='" + newWidth + @"' height='" + newHeight + @"'/>                                            <font face='Calibri' size='3'>
                                                <p>Melhores Cumprimentos,</p>
                                            </font>
                                            <font face='Calibri' size='3'>
                                                <b>" + nomeUsuario + @"</b><br>
                                                Construção Metálica | Preparador<br>
                                                T + 351 253 080 609 *
                                            </font>
                                            <font color='red' face='Calibri' size='3'>ofeliz.com</font><br>
                                            <p>
                                                <a href='https://www.ofeliz.com'>
                                                    <img src='file:///" + imagemOfelizFilePath.Replace("\\", "/") + @"' width='127' height='34' alt='Logo da O Feliz'>
                                                </a>
                                            </p>
                                            <i>
                                                <font color='LightGrey' face='Calibri' size='1.5'>Alvará Nº 10553 – Pub. *Chamada para a rede fixa nacional.</font>
                                            </i><br>
                                            <i>
                                                <font color='green' face='Calibri' size='1.5'>Antes de imprimir este e-mail, tenha em consideração o meio ambiente.</font>
                                            </i><br>
                                        </body>
                                        </html>";

                Console.WriteLine(corpoEmailnovo);
                webBrowser1.DocumentText = corpoEmailnovo;
                labelEmailInformação.Text = $"A Imagem foi redimensionada para : {newWidth} X {newHeight}";
            }
            catch (System.Runtime.InteropServices.ExternalException ex)
            {
                MessageBox.Show(this, "Erro GDI+: " + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Ocorreu um erro: " + ex.Message);
            }
        }

        private void AlterarValoresPercentagem()
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(textBoxWidth.Text) && !string.IsNullOrWhiteSpace(textBoxheight.Text))
                {
                    int originalWidth = int.Parse(textBoxWidth.Text);
                    int originalHeight = int.Parse(textBoxheight.Text);

                    int percent = (int)numericUpDown1.Value;

                    int newWidth = (originalWidth * percent) / 100;
                    int newHeight = (originalHeight * percent) / 100;

                    textBoxWidth.Text = newWidth.ToString();
                    textBoxheight.Text = newHeight.ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Erro ao calcular os valores: {ex.Message}");
            }
        }               

        private void ChamarPrint()
        {
            Frm_EnviarEmailparaFabrico f = new Frm_EnviarEmailparaFabrico();
            f.TextBox1Value = textbox1;
            f.ClickButton9();
        }
        
        private void Corpo_de_Texto_Email_Fabrico_FormClosing(object sender, FormClosingEventArgs e)
        {  }

        private void comboBoxDiretorObra_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            string nomeSelecionado = comboBoxDiretorObra.Text.Trim();

            if (!string.IsNullOrEmpty(nomeSelecionado) & !nomeSelecionado.Equals("Miguel Santos", StringComparison.OrdinalIgnoreCase))
            {
                string emailFormatado = nomeSelecionado.ToLower().Replace(" ", ".") + "@ofeliz.com";

                if (!listBoxCC.Items.Contains(emailFormatado))
                {
                    listBoxCC.Items.Add(emailFormatado);
                }
            }
        }

        private void textBoxWidth_Click(object sender, EventArgs e)
        {   }

        private void label11_Click(object sender, EventArgs e)
        {
            label11.Visible = !label11.Visible;
            label12.Visible = !label12.Visible;
            textBoxWidth.Visible = !textBoxWidth.Visible;
            textBoxheight.Visible = !textBoxheight.Visible;
            textBoxWidth.Text = label11.Text;
            textBoxheight.Text = label12.Text;
            label3.Visible = !label3.Visible;
            numericUpDown1.Visible = !numericUpDown1.Visible;
        }

        private void label12_Click(object sender, EventArgs e)
        {
            label11.Visible = !label11.Visible;
            label12.Visible = !label12.Visible;
            textBoxWidth.Visible = !textBoxWidth.Visible;
            textBoxheight.Visible = !textBoxheight.Visible;
            textBoxWidth.Text = label11.Text;
            textBoxheight.Text = label12.Text;
            label3.Visible = !label3.Visible;
            numericUpDown1.Visible = !numericUpDown1.Visible;
        }      

        private void textBoxheight_Enter(object sender, EventArgs e)
        {
           
        }

        private void textBoxWidth_Enter(object sender, EventArgs e)
        {
        }

        private void textBoxheight_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar))
            {
                e.Handled = true; 
            }

            if (e.KeyChar == (char)Keys.Enter)
            {
                RedimencionarImagem();
            }
        }

        private void textBoxWidth_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar))
            {
                e.Handled = true;
            }

            if (e.KeyChar == (char)Keys.Enter)
            {
                RedimencionarImagem();
            }
        }

        private void NumericUpDown1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar))
            {
                e.Handled = true;
            }

            if (e.KeyChar == (char)Keys.Enter)
            {
                AlterarValoresPercentagem();
                RedimencionarImagem();
            }
        }       

        private void guna2Button2_Click(object sender, EventArgs e)
        {
            System.Threading.Tasks.Task.Delay(2000).Wait();
            this.Close();

            System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
            timer.Interval = 100;
            timer.Tick += (s, args) =>
            {
                timer.Stop();
                ChamarPrint();
            };
            timer.Start();
        }

        private async void Buttonenviar_Click(object sender, EventArgs e)
        {
          try
            {
              if (!string.IsNullOrWhiteSpace(comboBoxDiretorObra.Text))
                 {
                        if (string.IsNullOrWhiteSpace(textBoxWidth.Text) && string.IsNullOrWhiteSpace(textBoxheight.Text))
                        {
                            await SendEmailAsync();
                        }
                        else
                        {
                            await SendEmailAsyncrezize();
                        }
              }
              else
              {
                this.Invoke((MethodInvoker)delegate
                {
                  CustomGradientPanel4.Visible = true;
                  CustomGradientPanel4.FillColor = Color.Maroon;
                  CustomGradientPanel4.FillColor2 = Color.Maroon;
                  CustomGradientPanel4.FillColor3 = Color.Maroon;
                  CustomGradientPanel4.FillColor4 = Color.Maroon;

                  labelEmailInformação.ForeColor = Color.White;
                  labelEmailInformação.Location = new Point(10, 8);
                  labelEmailInformação.Text = "Por favor, insira o nome do diretor de obra.";
                });
              }
                }
                catch (Exception ex)
                {
                    this.Invoke((MethodInvoker)delegate
                    {
                        MessageBox.Show(this, $"Erro: {ex.Message}", "Erro ao enviar e-mail", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    });
                }           
        }
                
        private void guna2Button3_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            Frm_AtualizarEmails F = new Frm_AtualizarEmails();
            F.ShowDialog();
            this.Visible = true;
        }

        private void textBoxEmailCC_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;

                string texto = textBoxEmailCC.Text.Trim();

                if (!string.IsNullOrEmpty(texto))
                {
                    var emails = texto.Split(';')
                        .Select(email => email.Trim())
                        .Where(email => !string.IsNullOrEmpty(email) && IsValidEmail(email))
                        .ToList();

                    foreach (var email in emails)
                    {
                        if (!listBoxCC.Items.Contains(email))
                        {
                            listBoxCC.Items.Add(email);
                        }
                    }

                    textBoxEmailCC.Clear();
                }
            }
        }

        private void listBoxCC_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete && listBoxCC.SelectedItem != null)
            {
                listBoxCC.Items.Remove(listBoxCC.SelectedItem);
            }
        }

        private void guna2Button5_Click(object sender, EventArgs e)
        {
            CarregarDiretorObra();
        }

    }

}
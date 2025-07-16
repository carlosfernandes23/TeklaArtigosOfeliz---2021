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
using Microsoft.Identity.Client;
using Microsoft.Office.Interop.Excel;
using System.Net.Http;
using Point = System.Drawing.Point;


namespace TeklaArtigosOfeliz
{
    public partial class Frm_Corpo_de_Texto_Email_Enviar_Powerfab : Form
    {
        private string corpoEmail;
        private string caminhoImagem;
        private string SubjectEnviarPowerFab;
        private string userDirectory;
        private string passwordFilePath;
        private string caminho;
        private string obra;
        private string lote;
        private string Fase;
        private string linkTexto;

        public Frm_Corpo_de_Texto_Email_Enviar_Powerfab(string titulo, string corpoEmail, string Subject, string caminho, string obra , string lote, string Fase, string linkTexto)
        {
            InitializeComponent();
            this.Text = titulo;
            this.corpoEmail = corpoEmail;
            this.SubjectEnviarPowerFab = Subject;
            this.caminho = caminho;
            this.obra = obra;
            this.lote = lote;
            this.Fase = Fase;
            this.linkTexto = linkTexto;
        }

        private void Corpo_de_Texto_Email_Enviar_Powerfab_Load(object sender, EventArgs e)
        {
            CarregarEmailsCC();
            CarregarEmailsPara();
            textBoxAsu.Text = SubjectEnviarPowerFab;
            webBrowser1.DocumentText = corpoEmail;

            string user = Environment.UserName;
            labelemailuser.Text = user + "@ofeliz.com";
        }

        private void CarregarEmailsCC()
        {
            string caminho = @"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\4 Produção\Desenvolvimentos\Ficheiros Temp tekla artigos (Nao Apagar)\Diretor de Obra Base de dados\EmailPowerfabCC.json";

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
            string caminho = @"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\4 Produção\Desenvolvimentos\Ficheiros Temp tekla artigos (Nao Apagar)\Diretor de Obra Base de dados\EmailPowerfabPara.json";

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
                                        <p>Venho por este meio informar, que já foi emitido dentro da pasta da obra em assunto, o PowerFab &nbsp;
                                            <span style='color:#00B0F0; display:inline-block; margin-right:10px;'><u>" + obra + @"&nbsp; Lote " + lote + @"&nbsp; Fase " + Fase + @"</u></span>
                                        </p>
                                    </font>
                                    <font face='Calibri' size='3'>
                                        <p><b><u>PROCESSO DE FABRICO:</u></b></p>
                                        <font face='Calibri' size='3' style='color:#5B9BD5;'>
                                            <a href='file:///" + caminho.Replace("\\", "/") + @"' style='color:#5B9BD5; text-decoration: none;'>" + linkTexto + @"</a>
                                        </font>
                                    </font>
                                    <font face='Calibri' size='3'>
                                        <p>Melhores Cumprimentos,</p>
                                    </font>
                                    <br>
                                    <font face='Calibri' size='3'>
                                        <b>" + nomeUsuario + @"</b>
                                    </font>
                                    <br>
                                    <font face='Calibri' size='3'>Construção Metálica | Preparador</font>
                                    <br>
                                    <font face='Calibri' size='3'>T + 351 253 080 609 *</font>
                                    <br>
                                    <font color='red' face='Calibri' size='3'>ofeliz.com</font>
                                    <br>
                                    <p>
                                        <a href='https://www.ofeliz.com'>
                                            <img src='file:///" + imagemOfelizFilePath.Replace("\\", "/") + @"' width='127' height='34'>
                                        </a>
                                    </p>
                                    <i>
                                        <font color='Light grey' face='Calibri' size='1.5'>Alvará Nº 10553 – Pub. *Chamada para a rede fixa nacional.</font>
                                    </i>
                                    <br>
                                    <i>
                                        <font color='green' face='Calibri' size='1.5'>Antes de imprimir este e-mail, tenha em consideração o meio ambiente.</font>
                                    </i>
                                    <br>
                                </body>
                                </html>";

            var toRecipients = listBoxPara.Items
                .Cast<string>()
                .Select(email => email.Trim())
                .Where(email => !string.IsNullOrEmpty(email) && IsValidEmail(email))
                .Select(email => new { EmailAddress = new { Address = email } })
                .ToArray();

            var ccRecipients = listBoxCC.Items
                .Cast<string>()
                .Select(email => email.Trim())
                .Where(email => !string.IsNullOrEmpty(email) && IsValidEmail(email))
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
                        CcRecipients = ccRecipients
                    },
                    SaveToSentItems = "true"
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

        private async void Buttonenviar_Click(object sender, EventArgs e)
        {
            try
            {
                if (listBoxPara.Items.Count > 0)
                {
                    await SendEmailAsync();
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

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            Frm_AtualizarEmails F = new Frm_AtualizarEmails();
            F.ShowDialog();
            this.Visible = true;
        }

        private bool IsValidEmail(string email)
        {
            try
            {
                var mailAddress = new System.Net.Mail.MailAddress(email);
                return mailAddress.Address == email;
            }
            catch
            {
                return false;
            }
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

    }
    }


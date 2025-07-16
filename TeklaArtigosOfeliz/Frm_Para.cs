using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Identity.Client;
using System.IO;

namespace TeklaArtigosOfeliz
{
    public partial class Frm_Para : Form
    {
        private List<Contact> allContacts = new List<Contact>();

        public Frm_Para()
        {
            InitializeComponent();
            this.Load += Frm_Para_Load;
            this.textBox1.TextChanged += textBox1_TextChanged;
            this.listBox1.DoubleClick += ListBox1_DoubleClick;
        }

        private async void Frm_Para_Load(object sender, EventArgs e)
        {
            try
            {
                listBox1.Items.Clear();
                listBox1.Items.Add("Carregando contatos...");
                allContacts = await BuscarTodosContactosAsync();
                if (allContacts.Count == 0)
                {
                    listBox1.Items.Clear();
                    listBox1.Items.Add("Nenhum contato encontrado.");
                }
                else
                {
                    AtualizarListBox(allContacts);
                }
            }
            catch (Exception ex)
            {
                listBox1.Items.Clear();
                listBox1.Items.Add("Erro ao buscar contatos.");
                MessageBox.Show("Erro ao buscar contatos: " + ex.Message);
            }
        }

        public class EmailAddress
        {
            public string Name { get; set; }
            public string Address { get; set; }
        }

        public class Contact
        {
            public string DisplayName { get; set; }
            public List<EmailAddress> EmailAddresses { get; set; }
        }

        // Autenticação e obtenção do token
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
                    .AcquireTokenInteractive(new[] { "Contacts.Read", "Mail.Send", "People.Read" })
                    .ExecuteAsync();

                return result.AccessToken;
            }

            try
            {
                var result = await publicClientApplication
                    .AcquireTokenSilent(new[] { "Contacts.Read", "Mail.Send", "People.Read" }, account)
                    .ExecuteAsync();

                return result.AccessToken;
            }
            catch (MsalUiRequiredException)
            {
                var result = await publicClientApplication
                    .AcquireTokenInteractive(new[] { "Contacts.Read", "Mail.Send", "People.Read" })
                    .ExecuteAsync();

                return result.AccessToken;
            }
        }

        // Buscar todos os contatos com paginação
        private async Task<List<Contact>> BuscarTodosContactosAsync()
        {
            var contatos = new List<Contact>();
            string url = "https://graph.microsoft.com/v1.0/me/contacts?$select=displayName,emailAddresses&$top=50";

            string accessToken = await GetAccessTokenAsync();
            using (var httpClient = new HttpClient())
            {
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                while (!string.IsNullOrEmpty(url))
                {
                    var response = await httpClient.GetAsync(url);
                    response.EnsureSuccessStatusCode();

                    string jsonResponse = await response.Content.ReadAsStringAsync();
                    using (JsonDocument doc = JsonDocument.Parse(jsonResponse))
                    {
                        var root = doc.RootElement;
                        var values = root.GetProperty("value");
                        foreach (var item in values.EnumerateArray())
                        {
                            var contact = JsonSerializer.Deserialize<Contact>(item.GetRawText());
                            contatos.Add(contact);
                        }

                        url = root.TryGetProperty("@odata.nextLink", out var nextLink) ? nextLink.GetString() : null;
                    }
                }
            }
            return contatos;
        }

        private void AtualizarListBox(List<Contact> contatos)
        {
            listBox1.Items.Clear();

            foreach (var contact in contatos)
            {
                string displayName = contact.DisplayName ?? "(sem nome)";
                string email = "(sem email)";

                if (contact.EmailAddresses != null && contact.EmailAddresses.Count > 0 && contact.EmailAddresses[0] != null)
                {
                    email = contact.EmailAddresses[0].Address ?? "(sem endereço)";
                }

                listBox1.Items.Add($"{displayName} <{email}>");
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            string filtro = textBox1.Text.Trim().ToLower();

            if (string.IsNullOrEmpty(filtro))
            {
                AtualizarListBox(allContacts);
                return;
            }

            var contatosFiltrados = allContacts
                .Where(c => ((c.DisplayName ?? "").ToLower().Contains(filtro)) ||
                            (c.EmailAddresses != null && c.EmailAddresses.Any(em => (em.Address ?? "").ToLower().Contains(filtro))))
                .ToList();

            if (contatosFiltrados.Count == 0)
            {
                listBox1.Items.Clear();
                listBox1.Items.Add("Nenhum contato encontrado.");
            }
            else
            {
                AtualizarListBox(contatosFiltrados);
            }
        }

        private void ListBox1_DoubleClick(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex >= 0 && listBox1.SelectedIndex < allContacts.Count)
            {
                var contatoSelecionado = allContacts[listBox1.SelectedIndex];
                var email = contatoSelecionado.EmailAddresses?.FirstOrDefault()?.Address;

                if (!string.IsNullOrEmpty(email))
                {
                    try
                    {
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo($"mailto:{email}") { UseShellExecute = true });
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Não foi possível abrir o cliente de email: {ex.Message}");
                    }
                }
                else
                {
                    MessageBox.Show("Contato selecionado não possui endereço de email.");
                }
            }
        }
    }
}

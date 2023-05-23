# youjian
using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace Office365EmailSender
{
    class Program
    {
        static async Task Main(string[] args)
        {
            try
            {
                var accessToken = await GetAccessToken("<Your-App-Id>", "<Your-App-Secret>", "<Your-Tenant-ID>");

                var message = new EmailMessage
                {
                    Subject = "Test Email",
                    Body = new EmailBody
                    {
                        Content = "This is a test email sent from Office 365 API.",
                        ContentType = "Text"
                    },
                    ToRecipients = new[]
                    {
                        new EmailRecipient
                        {
                            EmailAddress = new EmailAddress
                            {
                                Address = "recipient@example.com",
                                Name = "Recipient"
                            }
                        }
                    }
                };

                await SendMessage(accessToken, message);

                Console.WriteLine("Email sent successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to send email: " + ex.Message);
            }
        }

        static async Task<string> GetAccessToken(string appId, string appSecret, string tenantId)
        {
            using (var httpClient = new HttpClient())
            {
                var requestUri = $"https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token";

                var requestContent = new StringContent($"client_id={appId}&scope=https%3A%2F%2Fgraph.microsoft.com%2F.default&client_secret={appSecret}&grant_type=client_credentials");

                requestContent.Headers.ContentType = new MediaTypeHeaderValue("application/x-www-form-urlencoded");

                var response = await httpClient.PostAsync(requestUri, requestContent);

                var responseContent = await response.Content.ReadAsStringAsync();

                var tokenResponse = JsonConvert.DeserializeObject<TokenResponse>(responseContent);

                return tokenResponse.AccessToken;
            }
        }

        static async Task SendMessage(string accessToken, EmailMessage message)
        {
            using (var httpClient = new HttpClient())
            {
                var requestUri = "https://graph.microsoft.com/v1.0/me/sendMail";

                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                var requestContent = new StringContent(JsonConvert.SerializeObject(message));

                requestContent.Headers.ContentType = new MediaTypeHeaderValue("application/json");

                var response = await httpClient.PostAsync(requestUri, requestContent);

                if (!response.IsSuccessStatusCode)
                {
                    var responseContent = await response.Content.ReadAsStringAsync();
                    throw new Exception(responseContent);
                }
            }
        }
    }

    class EmailMessage
    {
        [JsonProperty("subject")]
        public string Subject { get; set; }

        [JsonProperty("body")]
        public EmailBody Body { get; set; }

        [JsonProperty("toRecipients")]
        public EmailRecipient[] ToRecipients {get; set;
        }
    }

    class EmailBody
    {
        [JsonProperty("contentType")]
        public string ContentType { get; set; }

        [JsonProperty("content")]
        public string Content { get; set; }
    }

    class EmailRecipient
    {
        [JsonProperty("emailAddress")]
        public EmailAddress EmailAddress { get; set; }
    }

    class EmailAddress
    {
        [JsonProperty("address")]
        public string Address { get; set; }

        [JsonProperty("name")]
        public string Name { get; set; }
    }

    class TokenResponse
    {
        [JsonProperty("access_token")]
        public string AccessToken { get; set; }
    }
}

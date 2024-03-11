using Microsoft.Identity.Client;
using System;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace PowerAppsUserManagement
{
    class Program
    {
        static async Task Main(string[] args)
        {
            string clientId = "YOUR_CLIENT_ID";
            string tenantId = "YOUR_TENANT_ID";
            string environmentId = "YOUR_ENVIRONMENT_ID";
            string username = "USER_EMAIL";
            string password = "USER_PASSWORD"; // Consider using secure storage methods for production code
            string role = "ENVIRONMENT_ADMIN"; // The role to assign to the user, e.g., ENVIRONMENT_ADMIN or USER

            string[] scopes = { "https://api.powerapps.com/.default" };

            var app = PublicClientApplicationBuilder.Create(clientId)
                .WithAuthority(new Uri($"https://login.microsoftonline.com/{tenantId}"))
                .Build();

            try
            {
                var result = await app.AcquireTokenByUsernamePassword(scopes, username, password).ExecuteAsync();
                string accessToken = result.AccessToken;

                HttpClient client = new HttpClient();
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + accessToken);

                string url = $"https://api.powerapps.com/providers/Microsoft.PowerApps/environments/{tenantId}:{environmentId}/users/{username}";

                HttpResponseMessage response = await client.PutAsync(url, null);

                if (response.IsSuccessStatusCode)
                {
                    Console.WriteLine($"User '{username}' added to PowerApps environment successfully.");
                }
                else
                {
                    Console.WriteLine($"Failed to add user '{username}' to PowerApps environment. Status code: {response.StatusCode}");
                }
            }
            catch (MsalException ex)
            {
                Console.WriteLine($"Error authenticating: {ex.Message}");
            }
            catch (HttpRequestException ex)
            {
                Console.WriteLine($"Error making HTTP request: {ex.Message}");
            }
        }
    }
}

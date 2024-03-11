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
            string tenantId = "YOUR_TENANT_ID";
            string environmentId = "YOUR_ENVIRONMENT_ID";
            string adminUsername = "ADMIN_USER_EMAIL";
            string adminPassword = "ADMIN_USER_PASSWORD"; // Consider using secure storage methods for production code
            string userEmailToAdd = "NEW_USER_EMAIL";

            HttpClient client = new HttpClient();

            try
            {
                var token = await GetAccessToken(adminUsername, adminPassword, tenantId);
                if (token != null)
                {
                    client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);

                    string url = $"https://api.powerapps.com/providers/Microsoft.PowerApps/environments/{tenantId}:{environmentId}/users/{userEmailToAdd}";

                    HttpResponseMessage response = await client.PutAsync(url, null);

                    if (response.IsSuccessStatusCode)
                    {
                        Console.WriteLine($"User '{userEmailToAdd}' added to PowerApps environment successfully.");
                    }
                    else
                    {
                        Console.WriteLine($"Failed to add user '{userEmailToAdd}' to PowerApps environment. Status code: {response.StatusCode}");
                    }
                }
                else
                {
                    Console.WriteLine("Failed to authenticate. Check username and password.");
                }
            }
            catch (HttpRequestException ex)
            {
                Console.WriteLine($"Error making HTTP request: {ex.Message}");
            }
        }

        static async Task<string> GetAccessToken(string username, string password, string tenantId)
        {
            try
            {
                var request = new HttpRequestMessage(HttpMethod.Post, $"https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token");
                var content = new StringContent($"grant_type=password&client_id=&username={username}&password={password}&scope=https://api.powerapps.com/.default", Encoding.UTF8, "application/x-www-form-urlencoded");
                request.Content = content;

                var client = new HttpClient();
                var response = await client.SendAsync(request);

                if (response.IsSuccessStatusCode)
                {
                    var responseContent = await response.Content.ReadAsStringAsync();
                    var accessToken = Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(responseContent).access_token;
                    return accessToken;
                }
                else
                {
                    return null;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting access token: {ex.Message}");
                return null;
            }
        }
    }
}

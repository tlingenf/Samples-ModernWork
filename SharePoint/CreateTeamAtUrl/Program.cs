using Microsoft.Identity.Client;
using Microsoft.Net.Http.Headers;
using Microsoft.SharePoint.Client;
using PnP.Framework;
using PnP.Framework.Entities;
using PnP.Framework.Sites;
using System;
using System.Net;
using System.Net.Http;
using System.Reflection.Metadata.Ecma335;
using System.Security;
using System.Threading.Tasks;

namespace CreateTeamAtUrl
{
    class Program
    {
        static void Main(string[] args)
        {
            MainAsync().Wait();
        }

        static async Task MainAsync()
        {
            Console.Write("Enter a number: ");
            var num = Console.ReadLine();

            string siteUrl = string.Format("https://xxxxxxxxxxx.sharepoint.com/teams/test-site-{0}", num);
            string groupTitle = string.Format("Test Site {0}", num);
            string groupAlias = string.Format("inbox-testsite-{0}", num);

            var app = PublicClientApplicationBuilder.Create("xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx")
                .WithTenantId("xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx")
                .WithRedirectUri("http://localhost")
                .Build();

            var authResult = await app.AcquireTokenInteractive(new string[] { "https://xxxxxxxxxxx.sharepoint.com/AllSites.FullControl" }).ExecuteAsync();

            // use the NetworkCredential as an easy way to convert the access token to a SecureString for the PnP.Framework.AuthenticationManager
            var authToken = new NetworkCredential("", authResult.AccessToken).SecurePassword;
            var authManager = new PnP.Framework.AuthenticationManager(authToken);

            using (var rootContext = authManager.GetContext("https://xxxxxxxxxxx.sharepoint.com"))
            {
                var createInfo = new TeamNoGroupSiteCollectionCreationInformation()
                {
                    Title = groupTitle,
                    Description = "Testing Team site creation",
                    Lcid = 1033,
                    Owner = "user@domain.com",
                    Url = siteUrl
                };
                Console.WriteLine("Creating Site");
                var createSiteResponse = await rootContext.CreateSiteAsync(createInfo);
                
                using (var siteContext = authManager.GetContext(createSiteResponse.Url))
                {
                    Console.WriteLine("Ensuring site is not already groupified");
                    var siteInfo = siteContext.Site;
                    siteContext.Load(siteInfo, s => s.GroupId);
                    siteContext.ExecuteQuery();

                    if (siteInfo.GroupId != Guid.Empty)
                    {
                        Console.WriteLine("Group already exists for site");
                    }
                    else
                    {
                        var groupifyInfo = new TeamSiteCollectionGroupifyInformation()
                        {
                            Alias = groupAlias,
                            Description = "Testing Groupify",
                            DisplayName = groupTitle,
                            IsPublic = false
                        };
                        Console.WriteLine("Groupify site");
                        var groupifyResponse = await siteContext.GroupifySiteAsync(groupifyInfo);

                        siteContext.Load(siteInfo, s => s.GroupId);
                        siteContext.ExecuteQuery();

                        Console.WriteLine("Teamify site");

                        // get graph token
                        var graphAuthResult = await app.AcquireTokenSilent(new string[] { "https://graph.microsoft.com/Group.ReadWrite.All" }, authResult.Account).ExecuteAsync();

                        var requestMessage = new HttpRequestMessage(HttpMethod.Put, string.Format("https://graph.microsoft.com/v1.0/groups/{0}/team", siteInfo.GroupId));
                        requestMessage.Content = new StringContent(
                            "{ \"memberSettings\": { \"allowCreatePrivateChannels\": true, \"allowCreateUpdateChannels\": true }, \"messagingSettings\": { \"allowUserEditMessages\": true, \"allowUserDeleteMessages\": true }, \"funSettings\": { \"allowGiphy\": true, \"giphyContentRating\": \"strict\" } }",
                            System.Text.Encoding.UTF8, 
                            "application/json"
                        );
                        var httpClient = new HttpClient();
                        httpClient.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", graphAuthResult.AccessToken);
                        httpClient.DefaultRequestHeaders.Add("Accept", "application/json");
                        var httpResponse = await httpClient.SendAsync(requestMessage);

                        Console.ReadKey();
                    }
                }
            }
        }
    }
}

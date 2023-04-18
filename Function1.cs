using System.Net;
using Azure.Identity;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using SendGrid;
using SendGrid.Helpers.Mail;
using Microsoft.Azure.Cosmos.Table;
using PasswordGenerator;
using System.Security.Cryptography;
using System.Text;
using System.Linq;
using Newtonsoft.Json;

namespace SignatureProject
{

    public class SignatureT
    {
        //Some Global Variables
        private readonly ILogger _logger;
        //Define Logger
        public SignatureT(ILoggerFactory loggerFactory)
        {
            _logger = loggerFactory.CreateLogger<SignatureT>();
        }

        [Function("SignatureT")]

        //Main task to run all the time. Handles get/post
        public async Task<HttpResponseData> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post")] HttpRequestData req, string DisplayName, string PassCode
        )
        {
            //Split values because we had to use | for HTML to work properly.
            DisplayName = DisplayName;

            //Set Security Protocol
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            //Log
            _logger.LogInformation("Security Protocol Initialized");

            //Create connection builder
            var ConfidentialClientApplication = ConfidentialClientApplicationBuilder
              .Create(Constants.AppId)
              .WithTenantId(Constants.TenantId)
              .WithClientSecret(Constants.ClientSecret)
              .Build();

            //Get Secret Credential
            var authCodeCredential = new ClientSecretCredential(Constants.TenantId, Constants.AppId, Constants.ClientSecret);

            //Create GraphService Client
            var _client = new GraphServiceClient(authCodeCredential);

                var users = await _client.Users
                    .Request()
                    .Filter($"startsWith(userPrincipalName, '{DisplayName}')")
                    .Select("id,displayName,jobTitle,department,mail,city,officelocation")
                    .GetAsync();

            var response = req.CreateResponse(HttpStatusCode.OK);

            // Check if there's a user with the given display name
            if (users?.Count > 0 && PassCode == "ENTER PASS CODE HERE")
            {
                var user = users.First();
                var groups = await _client.Users[user.Id].GetMemberGroups(false).Request().PostAsync();

                var result = new
                    {
                        DisplayName = user.DisplayName,
                        JobTitle = user.JobTitle ?? "Not available",
                        Department = user.Department ?? "Not available",
                        Mail = user.Mail ?? "Not available",
                        City = user.City ?? user.OfficeLocation
                    };
                    response.Headers.Add("Content-Type", "application/json");
                    await response.WriteStringAsync(JsonConvert.SerializeObject(result));
                }
                else
                {
                    response.StatusCode = HttpStatusCode.NotFound;
                    await response.WriteStringAsync("User not found");
                }
            return response;
        }
    }
}

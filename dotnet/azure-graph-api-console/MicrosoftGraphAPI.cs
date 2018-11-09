using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.Azure.Services.AppAuthentication;
using Microsoft.Graph;
using System;
using Microsoft.Extensions.Configuration;
using System.Net.Http;
using Newtonsoft.Json;
using System.Diagnostics;
using System.IO;

namespace msgraph
{
    // https://developer.microsoft.com/en-us/graph/graph-explorer?request=users?$filter=startswith(givenName,%27J%27)&method=GET&version=v1.0#
    // https://developer.microsoft.com/en-us/graph/docs/concepts/auth_v2_service
    // https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/resources/onedrive
    public class MicrosoftGraphAPI : IDocumentProvider
    {
        private GraphServiceClient _graphClient;
        private RESTClient _restClient;
        private OneDrive _oneDrive;

        public MicrosoftGraphAPI(string connectionString)
        {
            AzureServiceTokenProvider azureServiceTokenProvider = new AzureServiceTokenProvider(connectionString);

            string microsoftGraphEndpoint = "https://graph.microsoft.com";
            var token = azureServiceTokenProvider.GetAccessTokenAsync(microsoftGraphEndpoint).ConfigureAwait(false).GetAwaiter().GetResult();

            Task authenticationDelegate(System.Net.Http.HttpRequestMessage req)
            {
                req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                return Task.CompletedTask;
            }

            _restClient = new RESTClient(token);
            _graphClient = new GraphServiceClient(new DelegateAuthenticationProvider(authenticationDelegate));
            _oneDrive = new OneDrive(_restClient);
        }

        public MicrosoftGraphAPI(IConfiguration configuration) : this(buildConnectionString(configuration))
        {
        }

        private static string buildConnectionString(IConfiguration configuration)
        {
            string appId = configuration.GetSection("appId").Value;
            string appSecret = configuration.GetSection("appSecret").Value;
            string tenantId =  configuration.GetSection("tenantId").Value;
            var connectionString = $"RunAs=App;AppId={appId};TenantId={tenantId};AppKey={appSecret}";
            return connectionString;
        }

        public User GetUser(String userPrincipalName)
        {
            var currentPage = _graphClient.Users.Request().Filter($"userPrincipalName eq '{userPrincipalName}'").GetAsync().ConfigureAwait(false).GetAwaiter().GetResult();
            var user = currentPage.SingleOrDefault();
            return user;
        }

        public User CreateUser(String userPrincipalName, String displayName, String mailNickName, String password)
        {
            var user = _graphClient.Users.Request().AddAsync(new User
            {
                AccountEnabled = true,
                DisplayName = displayName,
                MailNickname = mailNickName,
                PasswordProfile = new PasswordProfile
                {
                    Password = password
                },
                UserPrincipalName = userPrincipalName
            }).ConfigureAwait(false).GetAwaiter().GetResult();
            return user;
        }

        public IEnumerable<Document> GetDocuments()
        {
            //https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/resources/onedrive
            //https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/driveitem_get_content_format
            //https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/resources/driveitem


            var format = "pdf";
            var user = GetUser("mrs@commentor.dk");
            //var url = $"https://graph.microsoft.com/v1.0/drive/items";
            //var url = $"https://graph.microsoft.com/v1.0/me/drives";
            var urlRoot = $"https://graph.microsoft.com/v1.0/";
            //var url = $"{urlRoot}/drive/root:/Documents/Test Results:/children";
            //var url = $"{urlRoot}/drive/root:/Documents/";
            //var url = $"{urlRoot}/me";
            //var url = $"{urlRoot}/users/{user.Id}/drive"; // OK            
            //var url = $"{urlRoot}/users/{user.Id}/drives"; // OK Will return "OneDrive" as an item
            var url = $"{urlRoot}/users/{user.Id}/drive/root/children"; // OK
            //var url = $"{urlRoot}/users/{user.Id}/drive/items/01IKWSDMYRMAGMO7EIFVFJDOS7ZN5TM3C5"; // OK
            //var url = $"{urlRoot}/users/{user.Id}/drive/items/01IKWSDMYRMAGMO7EIFVFJDOS7ZN5TM3C5/children"; // OK, will give all documents
            //var url = $"{urlRoot}/users/{user.Id}/drive/items/01IKWSDM3CLTXLYGAVUFHLUSXDWE54KW5U/content?format={format}";

            String responseBody = _restClient.getResponseAsString(url);
            Console.WriteLine(responseBody);
            var oData = JsonConvert.DeserializeObject<ODataDocument>(responseBody);
            return oData.Value;
        }

        public byte[] ConvertDocumentToPDF(byte[] inputDocumentBytes, string extension, string userId)
        {
            return _oneDrive.ConvertDocumentToPDF(inputDocumentBytes, extension, userId);
        }

        public IEnumerable<Document> GetRootDocuments(string userId)
        {
            return _oneDrive.GetRootDocuments(userId);
        }

        public Document GetDocumentByPath(string userId, string itemPath)
        {
            return _oneDrive.GetDocumentByPath(userId, itemPath);
        }

        public Document GetDocumentById(string userId, string itemId)
        {
            return _oneDrive.GetDocumentById(userId, itemId);
        }

        public IEnumerable<Document> GetChildDocumentsById(string userId, string itemId)
        {
            return _oneDrive.GetChildDocumentsById(userId, itemId);
        }
        
        public Document UploadSmallDocument(byte[] inputDocumentBytes, string itemPath, string userId)
        {
            return _oneDrive.UploadSmallDocument(inputDocumentBytes, itemPath, userId);
        }

        public Document UploadLargeDocument(byte[] inputDocumentBytes, string itemPath, string userId)
        {
            return _oneDrive.UploadLargeDocument(inputDocumentBytes, itemPath, userId);
        }

        public void DeleteDocumentById(string userId, string itemId)
        {
            _oneDrive.DeleteDocumentById(userId, itemId);
        }        
    }
}

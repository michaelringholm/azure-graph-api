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

namespace azure_ad_console
{
    // https://developer.microsoft.com/en-us/graph/graph-explorer?request=users?$filter=startswith(givenName,%27J%27)&method=GET&version=v1.0#
    // https://developer.microsoft.com/en-us/graph/docs/concepts/auth_v2_service
    // https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/resources/onedrive
    public class MicrosoftGraphAPI : IDocumentProvider
    {
        private GraphServiceClient _graphClient;
        private IConfiguration _configuration;
        private String _token;

        public MicrosoftGraphAPI(IConfiguration configuration)
        {
            _configuration = configuration;
            string connectionString = null;
            string appId = _configuration.GetSection("appId").Value;
            string appSecret = _configuration.GetSection("appSecret").Value;
            string tenantId = _configuration.GetSection("tenantId").Value;
            connectionString = $"RunAs=App;AppId={appId};TenantId={tenantId};AppKey={appSecret}";

            AzureServiceTokenProvider azureServiceTokenProvider = new AzureServiceTokenProvider(connectionString);

            string microsoftGraphEndpoint = "https://graph.microsoft.com";
            _token = azureServiceTokenProvider.GetAccessTokenAsync(microsoftGraphEndpoint).ConfigureAwait(false).GetAwaiter().GetResult();

            Task authenticationDelegate(System.Net.Http.HttpRequestMessage req)
            {
                req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _token);
                return Task.CompletedTask;
            }

            _graphClient = new GraphServiceClient(new DelegateAuthenticationProvider(authenticationDelegate));
        }

        private HttpResponseMessage postResponse(string url, HttpContent httpContent)
        {
            HttpClient client = new HttpClient();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _token);
            HttpResponseMessage response = client.PostAsync(url, httpContent).ConfigureAwait(false).GetAwaiter().GetResult();
            return response;
        } 

        private HttpResponseMessage putResponse(string url, HttpContent httpContent)
        {
            HttpClient client = new HttpClient();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _token);
            HttpResponseMessage response = client.PutAsync(url, httpContent).ConfigureAwait(false).GetAwaiter().GetResult();
            return response;
        }       

        private HttpResponseMessage getResponse(string url)
        {
            HttpClient client = new HttpClient();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _token);
            HttpResponseMessage response = client.GetAsync(url).ConfigureAwait(false).GetAwaiter().GetResult();
            return response;
        }

        private string getResponseAsString(string url)
        {
            var response = getResponse(url);
            Console.Error.WriteLine(response);
            response.EnsureSuccessStatusCode();
            string responseBody = response.Content.ReadAsStringAsync().ConfigureAwait(false).GetAwaiter().GetResult();
            return responseBody;
        }

        private byte[] getResponseAsBytes(string url)
        {
            var response = getResponse(url);
            Console.Error.WriteLine(response);
            response.EnsureSuccessStatusCode();
            var responseBodyBytes = response.Content.ReadAsByteArrayAsync().ConfigureAwait(false).GetAwaiter().GetResult();
            return responseBodyBytes;
        }        

        public User GetUser(String userPrincipalName)
        {            
            var currentPage = _graphClient.Users.Request().Filter($"userPrincipalName eq '{userPrincipalName}'").GetAsync().ConfigureAwait(false).GetAwaiter().GetResult();
            var user = currentPage.SingleOrDefault();
            return user;
        }        

        public User CreateUser(String userPrincipalName, String displayName, String mailNickName, String password) {
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

        public Document GetDocument(string documentId)
        {
            var url = $"https://graph.microsoft.com/v1.0/documents/{documentId}";
            var responseBody = getResponseAsString(url);
            var oData = JsonConvert.DeserializeObject<ODataDocument>(responseBody);            
            return oData.Value.FirstOrDefault();
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
            GetRootDocuments(user.Id);
            GetDocumentByName(user.Id, "7Digital");
            var dummyDocument = GetDocumentByName(user.Id, "Temp/dummy.txt");
            GetDocumentById(user.Id, "01IKWSDM5REL3VRV2EFNHZ6DCBXVN46JCV"); // 01IKWSDMZZH4XLSPOKLNA33GWZD6XF7MC5
            GetChildDocumentsById(user.Id, "01IKWSDM5REL3VRV2EFNHZ6DCBXVN46JCV");
            var buffer = System.Text.Encoding.UTF8.GetBytes("Empty");
            var tempDocumentName = Guid.NewGuid().ToString();
            var tempDocument = UploadSmallDocument(buffer, $"Temp/{tempDocumentName}.txt");            
            var uploadSession = CreateUploadSession(user.Id, tempDocument.id);
            var largeBuffer = System.Text.Encoding.UTF8.GetBytes("ABCDEF");
            var part1 = largeBuffer.Take(3).ToArray();
            UploadDocumentBytes(part1, 0, 2, 3, 6, uploadSession.UploadUrl);
            var part2 = largeBuffer.Skip(3).Take(3).ToArray();
            UploadDocumentBytes(part2, 3, 5, 3, 6, uploadSession.UploadUrl);
            //while(UploadDocumentBytes() == null) ....

            String responseBody = getResponseAsString(url);
            Console.WriteLine(responseBody);
            var oData = JsonConvert.DeserializeObject<ODataDocument>(responseBody);            
            return oData.Value;
        }

        public IEnumerable<Document> GetRootDocuments(string userId)
        {
            var urlRoot = $"https://graph.microsoft.com/v1.0/";
            var url = $"{urlRoot}/users/{userId}/drive/root/children";         
            String responseBody = getResponseAsString(url);
            Console.WriteLine(responseBody);
            var oData = JsonConvert.DeserializeObject<ODataDocument>(responseBody);            
            return oData.Value;
        }        

        public Document GetDocumentById(string userId, string itemId)
        {
            var urlRoot = $"https://graph.microsoft.com/v1.0/";
            var url = $"{urlRoot}/users/{userId}/drive/items/{itemId}";
            String responseBody = getResponseAsString(url);
            Console.WriteLine(responseBody);
            var document = JsonConvert.DeserializeObject<Document>(responseBody);            
            return document;
        }     

        public Document GetDocumentByName(string userId, string itemName)
        {
            var urlRoot = $"https://graph.microsoft.com/v1.0/";
            var url = $"{urlRoot}/users/{userId}/drive/root:/{itemName}";
            String responseBody = getResponseAsString(url);
            Console.WriteLine(responseBody);
            var document = JsonConvert.DeserializeObject<Document>(responseBody);            
            return document;
        }          

        public IEnumerable<Document> GetChildDocumentsById(string userId, string itemId)
        {
            var urlRoot = $"https://graph.microsoft.com/v1.0/";
            var url = $"{urlRoot}/users/{userId}/drive/items/{itemId}/children";
            String responseBody = getResponseAsString(url);
            Console.WriteLine(responseBody);
            var oData = JsonConvert.DeserializeObject<ODataDocument>(responseBody);            
            return oData.Value;
        }                    

        public byte[] ConvertDocumentToPDF(byte[] inputDocumentBytes, string extension) {
            var user = GetUser("mrs@commentor.dk"); // Use a generic OneDrive service user
            var document = UploadSmallDocument(inputDocumentBytes, extension);
            return DownloadAs("pdf", user.Id, document.DocumentId);
        }

        private Document UploadSmallDocument(byte[] inputDocumentBytes, string itemPath)
        {
            try
            {
                MemoryStream memStream = new MemoryStream(inputDocumentBytes);
                var user = GetUser("mrs@commentor.dk");   
                var urlRoot = $"https://graph.microsoft.com/v1.0/";
                var url = $"{urlRoot}/users/{user.Id}/drive/root:/{itemPath}:/content";                
                var byteContent = new ByteArrayContent(inputDocumentBytes);
                byteContent.Headers.ContentType = new MediaTypeHeaderValue("text/plain");
                var response = putResponse(url, byteContent);
                var responseBody = response.Content.ReadAsStringAsync().ConfigureAwait(false).GetAwaiter().GetResult();
                var document = JsonConvert.DeserializeObject<Document>(responseBody);  
                return document;
            }
            catch (ServiceException ex)
            {
                Console.Error.WriteLine(ex);
            }

            return null;
        }   

        private UploadSession CreateUploadSession(string userId, string itemId)
        {
            try
            {
                var user = GetUser("mrs@commentor.dk");   
                var urlRoot = $"https://graph.microsoft.com/v1.0/";
                var url = $"{urlRoot}/users/{userId}/drive/items/{itemId}/createUploadSession";
                var json = "{ \"@microsoft.graph.conflictBehavior\": \"rename\", \"name\": \"temp-large-file.dat\" }";
                var buffer = System.Text.Encoding.UTF8.GetBytes(json);
                var byteContent = new ByteArrayContent(buffer);
                byteContent.Headers.ContentType = new MediaTypeHeaderValue("application/json");
                var response = postResponse(url, byteContent);
                var responseBody = response.Content.ReadAsStringAsync().ConfigureAwait(false).GetAwaiter().GetResult();
                var uploadSession = JsonConvert.DeserializeObject<UploadSession>(responseBody);
                return uploadSession;
            }
            catch (ServiceException ex)
            {
                Console.Error.WriteLine(ex);
            }

            return null;
        }             

        private Document UploadDocumentBytes(byte[] inputDocumentBytes, int fromBytes, int toBytes, int chunkSize, int totalBytes, string uploadSessionUrl)
        {
            MemoryStream memStream = new MemoryStream(inputDocumentBytes);
            var byteContent = new ByteArrayContent(inputDocumentBytes);
            byteContent.Headers.ContentLength = chunkSize;
            byteContent.Headers.ContentRange = new ContentRangeHeaderValue(fromBytes,toBytes,totalBytes);
            var response = putResponse(uploadSessionUrl, byteContent);
            if(response.StatusCode == System.Net.HttpStatusCode.Accepted)
                return null; // Still more to do
            if(response.StatusCode == System.Net.HttpStatusCode.OK) {
                var responseBody = response.Content.ReadAsStringAsync().ConfigureAwait(false).GetAwaiter().GetResult();
                var document = JsonConvert.DeserializeObject<Document>(responseBody);  
                return document;
            }
            throw new Exception("An unknown error occured");
        }

        public byte[] DownloadAs(string format, string userId, string documentId)
        {
            //https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/resources/onedrive
            //https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/driveitem_get_content_format
            //https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/resources/driveitem
                       
            var urlRoot = $"https://graph.microsoft.com/v1.0/";
            var url = $"{urlRoot}/users/{userId}/drive/items/{documentId}/content?format={format}"; // Currently only supports PDF
            
            var responseBodyBytes = getResponseAsBytes(url);
            return responseBodyBytes;
        }        

        public Document CreateDocument(string documentName)
        {
            throw new NotImplementedException();
        }

        public void CreateFolder(string folderName) {

        }

        public void OneDriveUploadLargeFile(string filename, byte[] fileContent)
        {
            try
            {
                using (System.IO.MemoryStream ms = new System.IO.MemoryStream(fileContent))
                {
                    // Describe the file to upload. Pass into CreateUploadSession, when the service works as expected.
                    //var props = new DriveItemUploadableProperties();
                    //props.Name = "_hamilton.png";
                    //props.Description = "This is a pictureof Mr. Hamilton.";
                    //props.FileSystemInfo = new FileSystemInfo();
                    //props.FileSystemInfo.CreatedDateTime = System.DateTimeOffset.Now;
                    //props.FileSystemInfo.LastModifiedDateTime = System.DateTimeOffset.Now;

                    // Get the provider. 
                    // POST /v1.0/drive/items/01KGPRHTV6Y2GOVW7725BZO354PWSELRRZ:/_hamiltion.png:/microsoft.graph.createUploadSession
                    // The CreateUploadSesssion action doesn't seem to support the options stated in the metadata.
                    var uploadSession = _graphClient.Drive.Items["01KGPRHTV6Y2GOVW7725BZO354PWSELRRZ"].
                    ItemWithPath("_hamilton.png").CreateUploadSession().Request().PostAsync().ConfigureAwait(false).GetAwaiter().GetResult();

                    var maxChunkSize = 320 * 1024; // 320 KB - Change this to your chunk size. 5MB is the default.
                    var provider = new ChunkedUploadProvider(uploadSession, _graphClient, ms, maxChunkSize);

                    // Setup the chunk request necessities
                    var chunkRequests = provider.GetUploadChunkRequests();
                    var readBuffer = new byte[maxChunkSize];
                    var trackedExceptions = new List<Exception>();
                    DriveItem itemResult = null;

                    //upload the chunks
                    foreach (var request in chunkRequests)
                    {
                        // Do your updates here: update progress bar, etc.
                        // ...
                        // Send chunk request
                        var result = provider.GetChunkRequestResponseAsync(request, readBuffer, trackedExceptions).ConfigureAwait(false).GetAwaiter().GetResult();

                        if (result.UploadSucceeded)
                        {
                            itemResult = result.ItemResponse;
                        }
                    }

                    // Check that upload succeeded
                    if (itemResult == null)
                    {
                        // Retry the upload
                        // ...
                    }
                }
            }
            catch (Microsoft.Graph.ServiceException e)
            {
                Console.Error.WriteLine("Something happened, check out a trace. Error code: {0}", e.Error.Code);
            }
        }        
    }
}

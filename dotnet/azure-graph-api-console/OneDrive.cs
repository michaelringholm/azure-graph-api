using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using Microsoft.Graph;
using Newtonsoft.Json;

namespace msgraph
{
    public class OneDrive
    {

        private RESTClient _restClient;

        public OneDrive(RESTClient restClient)
        {
            _restClient = restClient;
        }

        public IEnumerable<Document> GetRootDocuments(string userId)
        {
            var urlRoot = $"https://graph.microsoft.com/v1.0/";
            var url = $"{urlRoot}/users/{userId}/drive/root/children";
            String responseBody = _restClient.getResponseAsString(url);
            Console.WriteLine(responseBody);
            var oData = JsonConvert.DeserializeObject<ODataDocument>(responseBody);
            return oData.Value;
        }

        public Document GetDocument(string itemId)
        {
            var url = $"https://graph.microsoft.com/v1.0/documents/{itemId}";
            var responseBody = _restClient.getResponseAsString(url);
            var oData = JsonConvert.DeserializeObject<ODataDocument>(responseBody);
            return oData.Value.FirstOrDefault();
        }

        public Document GetDocumentById(string userId, string itemId)
        {
            var urlRoot = $"https://graph.microsoft.com/v1.0/";
            var url = $"{urlRoot}/users/{userId}/drive/items/{itemId}";
            String responseBody = _restClient.getResponseAsString(url);
            Console.WriteLine(responseBody);
            var document = JsonConvert.DeserializeObject<Document>(responseBody);
            return document;
        }

        public Document GetDocumentByPath(string userId, string itemPath)
        {
            var urlRoot = $"https://graph.microsoft.com/v1.0/";
            var url = $"{urlRoot}/users/{userId}/drive/root:/{itemPath}";
            String responseBody = _restClient.getResponseAsString(url);
            Console.WriteLine(responseBody);
            var document = JsonConvert.DeserializeObject<Document>(responseBody);
            return document;
        }

        public IEnumerable<Document> GetChildDocumentsById(string userId, string itemId)
        {
            var urlRoot = $"https://graph.microsoft.com/v1.0/";
            var url = $"{urlRoot}/users/{userId}/drive/items/{itemId}/children";
            String responseBody = _restClient.getResponseAsString(url);
            Console.WriteLine(responseBody);
            var oData = JsonConvert.DeserializeObject<ODataDocument>(responseBody);
            return oData.Value;
        }

        public byte[] ConvertDocumentToPDF(byte[] inputDocumentBytes, string itemPath, string userId)
        {
            var document = UploadSmallDocument(inputDocumentBytes, itemPath, userId);
            return DownloadAs("pdf", userId, document.id);
        }

        private Document CreateDocument(string documentName)
        {
            throw new NotImplementedException();
        }

        private void CreateFolder(string folderName)
        {
            throw new NotImplementedException();
        }        

        public Document UploadSmallDocument(byte[] inputDocumentBytes, string itemPath, string userId)
        {
            try
            {
                MemoryStream memStream = new MemoryStream(inputDocumentBytes);

                var urlRoot = $"https://graph.microsoft.com/v1.0/";
                var url = $"{urlRoot}/users/{userId}/drive/root:/{itemPath}:/content";
                var byteContent = new ByteArrayContent(inputDocumentBytes);
                byteContent.Headers.ContentType = new MediaTypeHeaderValue("text/plain");
                var response = _restClient.putResponse(url, byteContent);
                var responseBody = response.Content.ReadAsStringAsync().ConfigureAwait(false).GetAwaiter().GetResult();
                if(response.StatusCode != HttpStatusCode.Created || response.StatusCode != HttpStatusCode.OK)
                    throw new Exception(responseBody);
                var document = JsonConvert.DeserializeObject<Document>(responseBody);
                return document;
            }
            catch (ServiceException ex)
            {
                Console.Error.WriteLine(ex);
            }

            return null;
        }

        public Document UploadLargeDocument(byte[] inputDocumentBytes, string itemPath, string userId) {
            var smallBuffer = System.Text.Encoding.UTF8.GetBytes("Empty");
            var newEmptyDoc = UploadSmallDocument(smallBuffer, itemPath, userId);
            var uploadSession = CreateUploadSession(userId, newEmptyDoc.id);

            /*var part1 = largeBuffer.Take(3).ToArray();
            UploadDocumentBytes(part1, 0, 2, 3, 6, uploadSession.UploadUrl);
            var part2 = largeBuffer.Skip(3).Take(3).ToArray();
            UploadDocumentBytes(part2, 3, 5, 3, 6, uploadSession.UploadUrl);*/
            
            int chunkSize = 32768;//1024;
            int fromIndex = 0;
            int toIndex = -1;
            byte[] nextChunk;
            Document largeDoc = null;

            while(largeDoc == null) {
                if((inputDocumentBytes.Length-fromIndex) < chunkSize)
                    chunkSize = inputDocumentBytes.Length-fromIndex;

                toIndex = fromIndex+chunkSize-1;
                nextChunk = new byte[chunkSize];
                Array.Copy(inputDocumentBytes, fromIndex, nextChunk, 0, chunkSize);
                largeDoc = UploadDocumentBytes(nextChunk, fromIndex, toIndex, chunkSize, inputDocumentBytes.Length, uploadSession.UploadUrl);
                fromIndex += chunkSize;                
            }

            return largeDoc;
            //GetNextChunk();
            
        }

        private UploadSession CreateUploadSession(string userId, string itemPath)
        {
            try
            {
                var urlRoot = $"https://graph.microsoft.com/v1.0/";
                var url = $"{urlRoot}/users/{userId}/drive/items/{itemPath}/createUploadSession";
                var json = "{ \"@microsoft.graph.conflictBehavior\": \"rename\", \"name\": \"temp-large-file.dat\" }";
                var buffer = System.Text.Encoding.UTF8.GetBytes(json);
                var byteContent = new ByteArrayContent(buffer);
                byteContent.Headers.ContentType = new MediaTypeHeaderValue("application/json");
                var response = _restClient.postResponse(url, byteContent);
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
            byteContent.Headers.ContentRange = new ContentRangeHeaderValue(fromBytes, toBytes, totalBytes);
            var response = _restClient.putResponse(uploadSessionUrl, byteContent);
            if (response.StatusCode == System.Net.HttpStatusCode.Accepted)
                return null; // Still more to do
            if (response.StatusCode == System.Net.HttpStatusCode.OK)
            {
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

            var urlRoot = $"https://graph.microsoft.com/v1.0";
            var url = $"{urlRoot}/users/{userId}/drive/items/{documentId}/content?format={format}"; // Currently only supports PDF

            var responseBodyBytes = _restClient.getResponseAsBytes(url);
            return responseBodyBytes;
        }
    }
}
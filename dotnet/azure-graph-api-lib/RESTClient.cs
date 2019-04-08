using System;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;

namespace com.opusmagus.azure.graph
{
    public class RESTClient
    {
        private string _token;
        
        public RESTClient(string token) {
            _token = token;
        }

        public HttpResponseMessage postResponse(string url, HttpContent httpContent)
        {
            HttpClient client = new HttpClient();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _token);
            HttpResponseMessage response = client.PostAsync(url, httpContent).ConfigureAwait(false).GetAwaiter().GetResult();
            return response;
        }

        public HttpResponseMessage putResponse(string url, HttpContent httpContent)
        {
            HttpClient client = new HttpClient();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _token);
            HttpResponseMessage response = client.PutAsync(url, httpContent).ConfigureAwait(false).GetAwaiter().GetResult();
            return response;
        }

        public HttpResponseMessage delResponse(string url)
        {
            HttpClient client = new HttpClient();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _token);
            HttpResponseMessage response = client.DeleteAsync(url).ConfigureAwait(false).GetAwaiter().GetResult();
            return response;
        }        

        public HttpResponseMessage getResponse(string url)
        {
            HttpClient client = new HttpClient();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _token);
            HttpResponseMessage response = client.GetAsync(url).ConfigureAwait(false).GetAwaiter().GetResult();
            return response;
        }

        public string getResponseAsString(string url)
        {
            var response = getResponse(url);
            Console.Error.WriteLine(response);
            response.EnsureSuccessStatusCode();
            string responseBody = response.Content.ReadAsStringAsync().ConfigureAwait(false).GetAwaiter().GetResult();
            return responseBody;
        }

        public byte[] getResponseAsBytes(string url)
        {
            var retryAttempts = 3;
            HttpResponseMessage response = null;
            while(retryAttempts > 0) {
                response = getResponse(url);
                Console.Error.WriteLine(response);
                if(!response.IsSuccessStatusCode && response.StatusCode == HttpStatusCode.ServiceUnavailable)
                    continue;
                break;
            }            
            if(response != null) {
                response.EnsureSuccessStatusCode();
                var responseBodyBytes = response.Content.ReadAsByteArrayAsync().ConfigureAwait(false).GetAwaiter().GetResult();
                return responseBodyBytes;            
            }
            throw new Exception("Ending retry loop due to too many 503s");
        }
    }
}
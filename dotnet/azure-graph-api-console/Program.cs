using System;
using System.IO;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;

namespace azure_ad_console
{
    class Program
    {
        static void Main(string[] args)
        {
            // https://github.com/microsoftgraph/aspnet-snippets-sample/blob/master/Graph-ASPNET-46-Snippets/Microsoft%20Graph%20ASPNET%20Snippets/Controllers/UsersController.cs
            Console.WriteLine("Started...");             

            var serviceCollection = new ServiceCollection();
            serviceCollection.AddSingleton<IConfiguration, JSONConfig>();
            serviceCollection.AddSingleton<IDocumentProvider, MicrosoftGraphAPI>();
            var builder = serviceCollection.BuildServiceProvider();

            var msGraph = builder.GetService<IDocumentProvider>();
            var documents = msGraph.GetDocuments();
            //var folders = msGraph.GetFolders();
            var inputBytes = File.ReadAllBytes("..\\resources\\source\\doc.docx");
            byte[] pdfDocBytes = msGraph.ConvertDocumentToPDF(inputBytes, "docx");
            File.WriteAllBytes($"c:\\data\\myfile.pdf", pdfDocBytes);
            //msGraph.ConvertDocumentToPDF($"c:\\data\\myfile.doc", "doc");
            Console.WriteLine("Ended!");
        }
    }
}

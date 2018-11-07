using System;
using System.IO;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;

namespace msgraph
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
            var serviceUser = msGraph.GetUser("pdf@commentor.dk");
            var documents = msGraph.GetDocuments();
            //var folders = msGraph.GetFolders();

            var htmlSimpleInputBytes = File.ReadAllBytes("..\\..\\resources\\source\\simple-html.html");
            byte[] pdfHTMLSimpleDocBytes = msGraph.ConvertDocumentToPDF(htmlSimpleInputBytes, $"Temp/{Guid.NewGuid()}.html", serviceUser.Id);
            File.WriteAllBytes($"..\\..\\resources\\target\\simple-html.pdf", pdfHTMLSimpleDocBytes);

            /*var csvSimpleInputBytes = File.ReadAllBytes("..\\..\\resources\\source\\simple.csv");
            byte[] pdfCSVSimpleDocBytes = msGraph.ConvertDocumentToPDF(csvSimpleInputBytes, $"Temp/{Guid.NewGuid()}.csv", serviceUser.Id);
            File.WriteAllBytes($"..\\..\\resources\\target\\csv-simple.pdf", pdfCSVSimpleDocBytes);*/

            var inputBytes = File.ReadAllBytes("..\\..\\resources\\source\\doc.docx");
            byte[] pdfDocBytes = msGraph.ConvertDocumentToPDF(inputBytes, $"Temp/{Guid.NewGuid()}.docx", serviceUser.Id);
            File.WriteAllBytes($"..\\..\\resources\\target\\doc.pdf", pdfDocBytes);

            var invoiceInputBytes = File.ReadAllBytes("..\\..\\resources\\source\\invoice.xlsx");
            byte[] pdfInvoiceDocBytes = msGraph.ConvertDocumentToPDF(invoiceInputBytes, $"Temp/{Guid.NewGuid()}.xlsx", serviceUser.Id);
            File.WriteAllBytes($"..\\..\\resources\\target\\invoice.pdf", pdfInvoiceDocBytes);
            
            /*var csvInvoiceInputBytes = File.ReadAllBytes("..\\..\\resources\\source\\invoice.csv");
            byte[] pdfCSVInvoiceDocBytes = msGraph.ConvertDocumentToPDF(csvInvoiceInputBytes, $"Temp/{Guid.NewGuid()}.csv", serviceUser.Id);
            File.WriteAllBytes($"..\\..\\resources\\target\\csv-invoice.pdf", pdfCSVInvoiceDocBytes);*/

            var rootDocs = msGraph.GetRootDocuments(serviceUser.Id);            
            var doc1 = msGraph.GetDocumentByPath(serviceUser.Id, "dummy.txt");
            //var doc2 = msGraph.GetDocumentById(serviceUser.Id, "01IKWSDM5REL3VRV2EFNHZ6DCBXVN46JCV"); // 01IKWSDMZZH4XLSPOKLNA33GWZD6XF7MC5
            //var doc3 = msGraph.GetDocumentByPath(serviceUser.Id, "7Digital");
            //var childDocuments = msGraph.GetChildDocumentsById(serviceUser.Id, "01IKWSDM5REL3VRV2EFNHZ6DCBXVN46JCV");
            
            var smallBuffer = System.Text.Encoding.UTF8.GetBytes("Empty");
            var smallDocName = Guid.NewGuid().ToString();
            var smallDoc = msGraph.UploadSmallDocument(smallBuffer, $"Temp/{smallDocName}.txt", serviceUser.Id);

            var largeBuffer = System.IO.File.ReadAllBytes("..\\..\\resources\\source\\doc.docx");
            var largeDocName = Guid.NewGuid().ToString();
            var largeDoc = msGraph.UploadLargeDocument(largeBuffer, $"Temp/{largeDocName}.txt", serviceUser.Id);

            Console.WriteLine("Ended!");
        }
    }
}

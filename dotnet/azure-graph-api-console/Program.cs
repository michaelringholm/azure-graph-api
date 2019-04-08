using System;
using System.Diagnostics;
using System.IO;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;

namespace com.opusmagus.azure.graph
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
            String userPrincipalName = "pdf@commentor.dk";
            var serviceUser = msGraph.GetUser($"{userPrincipalName}");
            var iterations = 5;
            var sw = Stopwatch.StartNew();
            for(var i=0;i<iterations;i++)
                ConvertWordToPDF(msGraph,serviceUser.Id);
            sw.Stop();

            Console.WriteLine($"Ran {iterations} iterations in {sw.ElapsedMilliseconds} ms with an avg of {sw.ElapsedMilliseconds/iterations} ms!");
        }

        private static void ConvertWordToPDF(IDocumentProvider documentProvider, string serviceUserId)
        {
            var inputBytes = File.ReadAllBytes("..\\..\\resources\\source\\doc.docx");
            var guid = Guid.NewGuid();
            byte[] pdfDocBytes = documentProvider.ConvertDocumentToPDF(inputBytes, $"Temp/{guid}.docx", serviceUserId);
            File.WriteAllBytes($"..\\..\\resources\\target\\{guid}.pdf", pdfDocBytes);
        }

        private static void simpleTest(IDocumentProvider msGraph, Microsoft.Graph.User serviceUser)
        {
            var documents = msGraph.GetDocuments(serviceUser.Id);
        }

        private static void extendedTest(IDocumentProvider msGraph, Microsoft.Graph.User serviceUser)
        {
            var documents = msGraph.GetDocuments(serviceUser.Id);
            //var folders = msGraph.GetFolders();

            /*var htmlLargeInputBytes = File.ReadAllBytes("..\\..\\resources\\source\\large-file.pptx");
            byte[] pdfHTMLLargeDocBytes = msGraph.ConvertDocumentToPDF(htmlLargeInputBytes, $"Temp/{Guid.NewGuid()}.pptx", serviceUser.Id);
            File.WriteAllBytes($"..\\..\\resources\\target\\large-file-pptx.pdf", pdfHTMLLargeDocBytes);*/

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
            msGraph.DeleteDocumentById(serviceUser.Id, smallDoc.id);

            var largeBuffer = System.IO.File.ReadAllBytes("..\\..\\resources\\source\\doc.docx");
            var largeDocName = Guid.NewGuid().ToString();
            var largeDoc = msGraph.UploadLargeDocument(largeBuffer, $"Temp/{largeDocName}.txt", serviceUser.Id);
            msGraph.DeleteDocumentById(serviceUser.Id, largeDoc.id);
        }

        private static void RunExtendedTest() {
            //var serviceUser = msGraph.GetUser("pdf@commentor.dk");
            //simpleTest(msGraph, serviceUser);
            //extendedTest(msGraph, serviceUser);
            //msGraph.CreateUser("mrs-demo@commentor.dk", "Michael Demo", "MRSDemoUser", "Test123!!##");
            //var mrsDemo = msGraph.GetUser("mrs-demo@commentor.dk");
            //msGraph.DeleteUser("mrs-demo@commentor.dk");
            //var serviceUser = msGraph.GetUser("nha@commentor.dk");
            //String userPrincipalName = "mrs.commentor@gmail.com";
            //var serviceUser = msGraph.GetUser($"{userPrincipalName.Replace("@", "_")}#EXT#@opusmagus.onmicrosoft.com");
            //msGraph.DeleteUser(serviceUser.Id);
            //msGraph.DeleteUser($"{userPrincipalName}");
        }
    }
}

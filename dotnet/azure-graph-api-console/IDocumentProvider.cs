using System;
using System.Collections.Generic;
using Microsoft.Graph;

namespace azure_ad_console
{
    public interface IDocumentProvider
    {
        IEnumerable<Document> GetDocuments();
        Document CreateDocument(String documentId);
        byte[] ConvertDocumentToPDF(byte[] inputDocumentBytes, string extension);
    }
}
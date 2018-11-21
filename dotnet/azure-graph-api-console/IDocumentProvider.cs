using System;
using System.Collections.Generic;
using Microsoft.Graph;

namespace msgraph
{
    public interface IDocumentProvider
    {
        IEnumerable<Document> GetDocuments(string userId);
        byte[] ConvertDocumentToPDF(byte[] inputDocumentBytes, string extension, string userId);
        User GetUser(String userPrincipalName);
        IEnumerable<Document> GetRootDocuments(string userId);
        Document GetDocumentByPath(string userId, string itemPath);
        Document GetDocumentById(string userId, string itemId);
        IEnumerable<Document> GetChildDocumentsById(string userId, string itemId);
        Document UploadSmallDocument(byte[] inputDocumentBytes, string itemPath, string userId);
        Document UploadLargeDocument(byte[] inputDocumentBytes, string itemPath, string userId);
        void DeleteDocumentById(string userId, string itemId);
    }
}
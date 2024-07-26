using Azure.Identity;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;

namespace SimpleGraphApiDemoApp
{
    internal class GraphServiceClientUtility
    {
        private readonly ILogger _logger;
        private string _tenantId = string.Empty;
        private string _clientId = string.Empty;
        private string _clientSecret = string.Empty;

        private GraphServiceClient _gsc { get; set; }

        public GraphServiceClientUtility(ILogger<GraphServiceClientUtility> logger, string tenantId, string clientId, string clientSecret)
        {
            try
            {
                _logger = logger;

                // The client credentials flow requires that you request the /.default scope, and pre-configure your permissions on the app registration in Azure.
                // An administrator must grant consent to those permissions beforehand.
                var scopes = new[] { "https://graph.microsoft.com/.default" };

                // Values from app registration
                _tenantId = tenantId;
                _clientId = clientId;
                _clientSecret = clientSecret;

                // using Azure.Identity;
                var options = new ClientSecretCredentialOptions
                {
                    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
                };

                // https://learn.microsoft.com/dotnet/api/azure.identity.clientsecretcredential
                var clientSecretCredential = new ClientSecretCredential(_tenantId, _clientId, _clientSecret, options);

                _gsc = new GraphServiceClient(clientSecretCredential, scopes);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.Message);
                throw;
            }
        }


        #region Sample Outlook functions

        public async Task GetNewestEmails(string userEmailAddress, string folderId, int topLimit)
        {
            bool nonCategorizedOnly = false;

            string filter =  nonCategorizedOnly? "not categories/any()" : "categories/any() eq false";   // sample filter for getting categorised/uncategorised mails

            MessageCollectionResponse messages = await _gsc
                    .Users[userEmailAddress]
                    .MailFolders[folderId]
                    .Messages
                    .GetAsync(requestConfiguration =>
                    {
                        requestConfiguration.QueryParameters.Filter = filter;
                        requestConfiguration.QueryParameters.Top = topLimit;
                        requestConfiguration.QueryParameters.Orderby = new List<string> { "receivedDateTime desc" }.ToArray();
                    });

            // You can handle the messages here..


            // Sample code block...
            if (messages != null && messages.Value != null && messages.Value.Count > 0)
            {
                foreach (var msg in messages.Value)
                {
                    string subjectLine = msg.Subject ?? string.Empty;
                    string messageId = msg.Id ?? string.Empty;
                    string body = msg.Body?.Content ?? string.Empty;


                }
            }

        }

        public async Task FindMailBySubject(string userEmailAddress, string folderId, string subject)
        {
            var filterQuery = $"subject eq '{subject.Replace("'", "''").Replace("\"", "\"\"")}'";    //https://learn.microsoft.com/en-us/graph/query-parameters?tabs=csharp

            MessageCollectionResponse messages = await _gsc
                    .Users[userEmailAddress]
                    .MailFolders[folderId]
                    .Messages
                    .GetAsync(requestConfiguration =>
                    {
                        requestConfiguration.QueryParameters.Top = 10;
                        requestConfiguration.QueryParameters.Orderby = new List<string> { "receivedDateTime desc" }.ToArray();
                        requestConfiguration.QueryParameters.Filter = filterQuery;
                    });

            // You can handle the messages here..


            // Sample code block...
            if (messages != null && messages.Value != null && messages.Value.Count > 0)
            {
                foreach (var msg in messages.Value)
                {
                    string messageId = msg.Subject ?? string.Empty;
                    string subjectLine = msg.Id ?? string.Empty;
                    string body = msg.Body?.Content ?? string.Empty;


                }
            }
        }


        public async Task<string> GetAttachmentData(string userEmailAddress, string folderId, string messageId)
        {
            string ret = string.Empty;

            // Get Attachments from the mailitem
            AttachmentCollectionResponse attachments = await _gsc
                        .Users[userEmailAddress]
                        .MailFolders[folderId]
                        .Messages[messageId]
                        .Attachments
                        .GetAsync();

            if (attachments != null && attachments.Value != null)
            {
                foreach (var attachment in attachments.Value)
                {
                    Console.WriteLine("Attachment Name : " + attachment.Name);

                    Attachment att = attachments.Value[0];
                    if (att == null)
                        return ret;

                    bool isInLine = false;
                    if (att.IsInline != null)
                        isInLine = att.IsInline.Value;

                    if (!isInLine)
                    {
                        Console.WriteLine(att?.Name ?? string.Empty);

                        // Download the mailitem to disk and read the data
                        var rawAttachmentInfo = _gsc
                            .Users[userEmailAddress]
                            .Messages[messageId]
                            .Attachments[att.Id]
                            .ToGetRequestInformation();

                        rawAttachmentInfo.URI = new Uri(rawAttachmentInfo.URI.OriginalString + "/$value");

                        using (var attachmentStream = await _gsc.RequestAdapter.SendPrimitiveAsync<Stream>(rawAttachmentInfo))
                        {
                            if (attachmentStream != null)
                            {
                                using (var reader = new StreamReader(attachmentStream, Encoding.UTF8, leaveOpen: true))
                                {
                                   string attachmentContent = reader.ReadToEnd();

                                    // Do your thing here

                                }
                            }
                        };


                    }

                }

            }

            return ret;
        }


        #endregion


    }
}

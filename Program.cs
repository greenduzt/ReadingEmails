using Microsoft.Extensions.Configuration;
using Newtonsoft.Json.Linq;
using ReadingEmails;
using System;
using System.Net.Http.Headers;
using System.Text;
using File = System.IO.File;

class Program
{
    static async Task Main(string[] args)
    {

        var builder = new ConfigurationBuilder().SetBasePath(Directory.GetCurrentDirectory()).AddJsonFile("appsettings.json", optional: false);

        IConfiguration config = builder.Build();
        Configuration con = new Configuration();
        con.ClientID = config["Auth2:ClientId"];
        con.ClientSecret = config["Auth2:ClientSecret"];
        con.TenantID = config["Auth2:TenantId"];
        con.Email = config["Auth2:Email"];

        string attachmentFolderPath = @"D:\Attachments\";
        string tokenEndpoint = $"https://login.microsoftonline.com/{con.TenantID}/oauth2/v2.0/token";

        using (HttpClient client = new HttpClient())
        {
            // Obtain access token using client credentials grant flow
            var requestContent = new FormUrlEncodedContent(new[]
            {
                new KeyValuePair<string, string>("grant_type", "client_credentials"),
                new KeyValuePair<string, string>("client_id", con.ClientID),
                new KeyValuePair<string, string>("client_secret", con.ClientSecret),
                new KeyValuePair<string, string>("scope", "https://graph.microsoft.com/.default"),
            });

            var response = await client.PostAsync(tokenEndpoint, requestContent);
            if (!response.IsSuccessStatusCode)
            {
                Console.WriteLine($"Error obtaining access token: {response.StatusCode}");
                return;
            }

            var tokenResponse = await response.Content.ReadAsStringAsync();
            var accessToken = JObject.Parse(tokenResponse)["access_token"].ToString();

            // Use obtained access token to retrieve unread emails
            using (HttpClient graphClient = new HttpClient())
            {
                graphClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                // Construct the URL for retrieving unread emails
                var getEmailsUrl = $"https://graph.microsoft.com/v1.0/users/{con.Email}/messages?$filter=isRead eq false";

                // Retrieve unread emails
                response = await graphClient.GetAsync(getEmailsUrl);
                if (!response.IsSuccessStatusCode)
                {
                    Console.WriteLine($"Error retrieving unread emails: {response.StatusCode}");
                    return;
                }

                var unreadEmailsResponse = await response.Content.ReadAsStringAsync();
                var unreadEmails = JObject.Parse(unreadEmailsResponse)["value"].ToObject<JArray>();

                // Display sender email, sender name, subject, and received date for each unread email
                foreach (var unreadEmail in unreadEmails)
                {
                    var emailId = unreadEmail["id"].ToString();
                    var from = unreadEmail["from"];
                    var senderName = from["emailAddress"]["name"].ToString();
                    var senderEmail = from["emailAddress"]["address"].ToString();
                    var subject = unreadEmail["subject"].ToString();
                    var receivedDateTime = DateTime.Parse(unreadEmail["receivedDateTime"].ToString());
                    var getAttachmentsUrl = $"https://graph.microsoft.com/v1.0/users/{con.Email}/messages/{emailId}/attachments";

                    Console.WriteLine($"Sender Name: {senderName}");
                    Console.WriteLine($"Sender Email: {senderEmail}");
                    Console.WriteLine($"Subject: {subject}");
                    Console.WriteLine($"Received Date: {receivedDateTime}");
                    Console.WriteLine();

                    // Retrieve attachments for the email
                    response = await graphClient.GetAsync(getAttachmentsUrl);
                    if (!response.IsSuccessStatusCode)
                    {
                        Console.WriteLine($"Error retrieving attachments for email {emailId}: {response.StatusCode}");
                        continue;
                    }

                    var attachmentsResponse = await response.Content.ReadAsStringAsync();
                    var attachments = JObject.Parse(attachmentsResponse)["value"].ToObject<JArray>();

                    // Process each attachment
                    foreach (var attachment in attachments)
                    {
                        var attachmentId = attachment["id"].ToString();
                        var attachmentName = attachment["name"].ToString();
                        var attachmentContentType = attachment["contentType"].ToString();

                        // Download the attachment if it is a PDF
                        if (attachmentContentType == "application/pdf")
                        {
                            var downloadAttachmentUrl = $"https://graph.microsoft.com/v1.0/users/{con.Email}/messages/{emailId}/attachments/{attachmentId}";

                            // Download the attachment
                            response = await graphClient.GetAsync(downloadAttachmentUrl);
                            if (!response.IsSuccessStatusCode)
                            {
                                Console.WriteLine($"Error downloading PDF attachment {attachmentName} for email {emailId}: {response.StatusCode}");
                                continue;
                            }

                            // Save the attachment to the specified folder on the D drive
                            var attachmentFilePath = Path.Combine(attachmentFolderPath, attachmentName);
                            try
                            {
                                using (var fileStream = File.Create(attachmentFilePath))
                                {
                                    var attachmentStream = await response.Content.ReadAsStreamAsync();
                                    await attachmentStream.CopyToAsync(fileStream);
                                    Console.WriteLine($"PDF attachment {attachmentName} downloaded to {attachmentFilePath} for email {emailId}");
                                }

                                // Check if the file exists and has non-zero length
                                if (!File.Exists(attachmentFilePath) || new FileInfo(attachmentFilePath).Length == 0)
                                {
                                    Console.WriteLine($"Error: Downloaded PDF file {attachmentName} appears to be empty or corrupted.");
                                    File.Delete(attachmentFilePath); // Delete the corrupted file
                                    continue;
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Error saving PDF attachment {attachmentName} for email {emailId}: {ex.Message}");
                            }
                        }
                    }

                    var markAsReadUrl = $"https://graph.microsoft.com/v1.0/users/{con.Email}/messages/{emailId}";

                    var markAsReadContent = new StringContent("{\"isRead\": true}", Encoding.UTF8, "application/json");
                    var markAsReadResponse = await graphClient.PatchAsync(markAsReadUrl, markAsReadContent);

                    if (!markAsReadResponse.IsSuccessStatusCode)
                    {
                        Console.WriteLine($"Error marking email as read: {markAsReadResponse.StatusCode}");
                    }
                    else
                    {
                        Console.WriteLine($"Email marked as read: {emailId}");
                    }
                }
            }
        }
    }
}




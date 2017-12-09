using System;
using Microsoft.SharePoint.Client;
using System.Security;
using System.Net.Http.Headers;
using System.Net.Http;
using System.IO;

namespace IMD.Connect.SPO.Provisioning
{
    class ExportWorkflow
    {
        // The API key and root URL for the REST API.
        // TODO: Replace with your API key and root URL.
        static private string apiKey = "5237477280d042db9122296c697bdb2c";
        static private string apiRootUrl = "https://imdtst.nintexo365.com";

        // The SharePoint site and credentials to use with the REST API.
        // TODO: Replace with your site URL, user name, and password.
        static private string spSiteUrl = "https://imdtst.sharepoint.com/crk";
        static private string spUsername = "ranjca@IMDTST.onmicrosoft.com";
        static private string spPassword = "tataplus@1";

        // The workflow to export, and the file path in which to create
        // the export file.
        // TODO: Replace with your workflow identifier and the file path
        // in which to create your export file.
        static private string workflowId = "d7af2a1a-e8a1-49a4-851c-2d50b0d591ed";
        static private string exportPath = "D:/Nintex";
        static private string filepath = "D:/Nintex/approve workflow.nwp";
        //this is for uplode workflow location
        static private string destinationUplode = "Documents";


        static public string GetSPOCookie()
        {
            // If successful, this variable contains an authentication cookie; 
            // otherwise, an empty string.
            string result = String.Empty;
            try
            {
                // Construct a secure string from the provided password.
                // NOTE: For sample purposes only.
                var securePassword = new SecureString();
                foreach (char c in spPassword) { securePassword.AppendChar(c); }

                // Instantiate a new SharePointOnlineCredentials object, using the 
                // specified username and password.
                var spoCredential = new SharePointOnlineCredentials(spUsername, securePassword);
                // If successful, try to authenticate the credentials for the
                // specified site.
                if (spoCredential == null)
                {
                    // Credentials could not be created.
                    result = String.Empty;
                }
                else
                {
                    // Credentials exist, so attempt to get the authentication cookie
                    // from the specified site.
                    result = spoCredential.GetAuthenticationCookie(new Uri(spSiteUrl));
                }
            }
            catch (Exception ex)
            {
                // An exception occurred while either creating the credentials or
                // getting an authentication cookie from the specified site.
                Console.WriteLine(ex.ToString());
                result = String.Empty;
            }

            // Return the result.
            return result;
        }

        static async public void ExportWorkflowToFile()
        {
            // Create a new HTTP client and configure its base address.
            HttpClient client = new HttpClient();
            client.BaseAddress = new Uri(spSiteUrl);

            // Add common request headers for the REST API to the HTTP client.
            client.DefaultRequestHeaders.Accept.Add(
                new MediaTypeWithQualityHeaderValue("application/json"));
            client.DefaultRequestHeaders.Add("Api-Key", apiKey);

            // Get the SharePoint authorization cookie to be used by the HTTP client
            // for the request, and use it for the Authorization request header.
            string spoCookie = GetSPOCookie();
            if (spoCookie != String.Empty)
            {
                var authHeader = new AuthenticationHeaderValue(
                    "cookie",
                    String.Format("{0} {1}", spSiteUrl, spoCookie)
                );
                // Add the defined authentication header to the HTTP client's
                // default request headers.
                client.DefaultRequestHeaders.Authorization = authHeader;
            }
            else
            {
                throw new InvalidOperationException("Cannot define Authentication header for request.");
            }

            // If we're at this point, we're ready to make our request.
            // Note that we're making this call synchronously - you can call the REST API
            // asynchronously, as needed.
            var exportWorkflowUri = String.Format("{0}/api/v1/workflows/packages/{1}",
                apiRootUrl.TrimEnd('/'),
                Uri.EscapeUriString(workflowId));
            HttpResponseMessage response = client.GetAsync(exportWorkflowUri).Result;

            // If we're successful, write an export file from the body of the response.
            if (response.IsSuccessStatusCode)
            {
                // Concatenate the export file name from the response headers.
                string exportFilePath = String.Empty;
                if (response.Content.Headers.ContentDisposition.FileName != null)
                {
                    // Get the suggested file name from the Content-Disposition header.
                    exportFilePath = Path.Combine(exportPath,
                        response.Content.Headers.ContentDisposition.FileName.Trim('"'));
                }
                else
                {
                    // Use a default file name if the suggested file name couldn't be retrieved.
                    exportFilePath = Path.Combine(exportPath, "DefaultWorkflow.nwp");
                }

                // The response body contains a Base64-encoded binary string, which we'll
                // asynchronously retrieve and then write to a new export file.
                byte[] exportFileContent = await response.Content.ReadAsByteArrayAsync();
                System.IO.File.WriteAllBytes(exportFilePath, exportFileContent);
            }
        }


        static public void UplodeWorkflow()
        {
            try
            {
                ClientContext contextSource = new ClientContext(spSiteUrl);
                var fileName = filepath;
                var passWord = new SecureString();
                foreach (var c in spPassword) passWord.AppendChar(c);
                contextSource.Credentials = new SharePointOnlineCredentials(spUsername, passWord);
                var web = contextSource.Web;
                var newFile = new FileCreationInformation
                {
                    Content = System.IO.File.ReadAllBytes(fileName),
                    Url = Path.GetFileName(fileName)
                };
                var docs = web.Lists.GetByTitle(destinationUplode);
                Microsoft.SharePoint.Client.File uploadFile = docs.RootFolder.Files.Add(newFile);
                contextSource.ExecuteQuery();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }


    }
}

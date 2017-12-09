using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using System.Security;
using System.Net.Http.Headers;
using System.Net.Http;
using System.Net;
using System.IO;


namespace IMD.Connect.SPO.Provisioning
{
    class ImportWorkflow
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

        // The list workflow to export, and the name of the destination list for which
        // the new workflow is to be imported.
        // TODO: Replace with your workflow identifier and list title.
        static private string sourceWorkflowId = "cff451cd-ea74-4255-bf70-ddcab358b780";
        static private string destinationListTitle = "CreateCustomLibray";
        static private string docname = "createdocLibrarychk";


        static private string GetSPOCookie()
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

        static async public void CopyWorkflowToList()
        {
            var serverRelativeUrl = "";
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



            ClientContext contextSource = new ClientContext(spSiteUrl);
            List cjList = contextSource.Web.Lists.GetByTitle("Documents");
            // var fileName = filepath;
            var passWord = new SecureString();
            foreach (var c in spPassword) passWord.AppendChar(c);
            contextSource.Credentials = new SharePointOnlineCredentials(spUsername, passWord);
            //ar file = contextSource.Web.Lists.Include(list =>list.

            CamlQuery listItemQuery = CamlQuery.CreateAllItemsQuery(100);
            //listItemQuery.ViewXml = "ViewXml that limits ListItems returned based on a field val";
            // listItemQuery.ViewXml = @"<View Scope='Recursive'><Query><Where><Eq><FieldRef Name='FSObjType' /><Value Type='Lookup'>0</Value></Eq></Where></Query></View>";
            listItemQuery.ViewXml = @"<Where><Eq><FieldRef Name='FileLeafRef'/><Value Type='File'>" + docname + "</Value></Eq></Where>";
            ListItemCollection docItems = cjList.GetItems(listItemQuery);

            contextSource.Load(docItems);
            contextSource.ExecuteQuery();

            foreach (ListItem listItem in docItems)
            {
                Console.WriteLine(listItem["FileRef"]);
                serverRelativeUrl = Convert.ToString(listItem["FileRef"]);
                // Console.WriteLine(listItem["ID"]);
                // listItem.File.Name or Title;  // Does not work, get field not initialized error
            }

            // var absoluteUrl = new Uri(contextSource.Url).GetLeftPart(UriPartial.Authority) + serverRelativeUrl;
            var exportWorkflowUri = "https://imdtst.sharepoint.com/crk/_layouts/15/guestaccess.aspx?docid=1d58ef554fb054124b32f2c9fd4ad1312&authkey=Aax7A8ZZeFvPO6cz3_tbbsw&e=f61a52662ef44f3caf7fbc75097f5ade";

            using (var httpClient = new HttpClient())
            {
                httpClient.DefaultRequestHeaders.TryAddWithoutValidation("Accept", "application/json");
                httpClient.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent", "Mozilla/5.0 (Windows NT 6.2; WOW64; rv:19.0) Gecko/20100101 Firefox/19.0");

                // HttpResponseMessage response = httpClient.GetAsync(absoluteUrl).Result;
                HttpResponseMessage response = httpClient.GetAsync(exportWorkflowUri).Result;
                Console.WriteLine(response.Content.ReadAsStringAsync().Result);
                // Console.ReadKey();

                if (response.IsSuccessStatusCode)
                {
                    // The response body contains a Base64-encoded binary string, which we'll
                    // asynchronously retrieve as a byte array.
                    byte[] exportFileContent = await response.Content.ReadAsByteArrayAsync();

                    // Next, import the exported workflow to the destination list.
                    var importWorkflowUri = String.Format("{0}/api/v1/workflows/packages/?migrate=true&listTitle={1}",
                        apiRootUrl.TrimEnd('/'),
                        Uri.EscapeUriString(destinationListTitle));

                    // Create a ByteArrayContent object to contain the byte array for the exported workflow.
                    var importContent = new ByteArrayContent(exportFileContent);

                    // Send a POST request to the REST resource.
                    HttpResponseMessage importResponse = client.PostAsync(importWorkflowUri, importContent).Result;

                    // Indicate to the console window the success or failure of the operation.
                    if (importResponse.IsSuccessStatusCode)
                    {
                        Console.WriteLine("Successfully imported workflow.");
                    }
                    else
                    {
                        Console.WriteLine("Failed to import workflow.");
                    }

                }
            }
        }
    }

}

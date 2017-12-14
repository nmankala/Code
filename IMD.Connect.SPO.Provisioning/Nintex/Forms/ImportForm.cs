using System;
using System.Text;
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
    class ImportForm
    {

        //public static string apiKey;
        //public static string apiRootUrl;
        //public static string spSiteUrl;
        ////  public static string importFileName;
        //public static string listId;


        // The API key and root URL for the REST API.
        // TODO: Replace with your API key and root URL.

        static private string apiKey = "5237477280d042db9122296c697bdb2c";



        static private string apiRootUrl = "https://imdtst.nintexo365.com";

        //// The SharePoint site and credentials to use with the REST API.
        //// TODO: Replace with your site URL, user name, and password.
        static private string spSiteUrl = "https://imdtst.sharepoint.com/crk";
        static private string spUsername = "ranjca@IMDTST.onmicrosoft.com";
        static private string spPassword = "tataplus@1";

        //// The list form to export, and the name of the destination list for which
        //// the new form is to be imported.
        //// TODO: Replace with the path to your form package (NFP or XML)
        //// and target list title.
        public static string importFileName = "D:/NFForm.nfp";
        static private string listId = "{01F1F142-52DB-4E2E-9C0D-8C156AC1C467}";

      

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

        static public void CopyFormToList()
        {

            try
            {
                //Console.WriteLine("Please Enter the API key");
                //apiKey = Console.ReadLine();
                //Console.WriteLine("Please Enter the API Root Url");
                //apiRootUrl = Console.ReadLine();
                //Console.WriteLine("Please Enter the Site Url");
                //spSiteUrl = Console.ReadLine();
                ////Console.WriteLine("Please Enter Form Path Which you want to import");
                ////  importFileName = Console.ReadLine();
                //Console.WriteLine("Please Enter Form List ID");
                //listId = Console.ReadLine();


                // Create a new HTTP client and configure its base address.
                HttpClient client = new HttpClient();
                client.BaseAddress = new Uri(spSiteUrl);

                // Add common request headers for the REST API to the HTTP client.
                client.DefaultRequestHeaders.Accept.Add(
                    new MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.Add("Api-Key", apiKey);

                // Get the SharePoint authentication cookie to be used by the HTTP client
                // for the request, and use it for the Authorization request header.
                string spoCookie = GetSPOCookie();
                if (spoCookie != String.Empty)
                {
                    var authHeader = new AuthenticationHeaderValue(
                        "cookie",
                        String.Format("{0} {1}", spSiteUrl, spoCookie)
                    );


                    // Add the defined Authorization header to the HTTP client's
                    // default request headers.
                    client.DefaultRequestHeaders.Authorization = authHeader;

                }
                else
                {
                    throw new InvalidOperationException("Cannot define Authorization header for request.");
                }




                // Read the contents of our form into a byte array, so that we can send the 
                // contents as a ByteArrayContent object with the request.
                if (System.IO.File.Exists(importFileName))
                {
                    // Read the file.
                    byte[] exportFileContents = System.IO.File.ReadAllBytes(importFileName);
                    ByteArrayContent saveContent = new ByteArrayContent(exportFileContents);

                    //this block is for testing
                    var testURI = String.Format("{0}/api/v1/forms/{1}",
                        apiRootUrl.TrimEnd('/'),
                        Uri.EscapeUriString(listId));
                    Console.Write(testURI, "\n");

                    // If we're at this point, we're ready to make our request.
                    // Note that we're making this call synchronously - you can call the REST API
                    // asynchronously, as needed.
                    var importFormUri = String.Format("{0}/api/v1/forms/{1}",
                        apiRootUrl.TrimEnd('/'),
                        Uri.EscapeUriString(listId));
                    HttpResponseMessage saveResponse = client.PutAsync(importFormUri, saveContent).Result;

                    if (saveResponse.IsSuccessStatusCode)
                    {
                        Console.WriteLine("Successfully imported form.PlsCheck D:/NFForm.nfp");
                        Console.WriteLine(saveResponse.ReasonPhrase, saveResponse.StatusCode);
                        Console.ReadLine();
                    }
                    else
                    {
                        Console.WriteLine("Failed to import form.");
                        Console.WriteLine(saveResponse.ReasonPhrase, saveResponse.StatusCode);
                        Console.ReadLine();
                    }
                }
                else
                {
                    Console.WriteLine("Failed to get form.Please Check the URL path");
                    Console.ReadLine();
                }
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }


        }
    }
}

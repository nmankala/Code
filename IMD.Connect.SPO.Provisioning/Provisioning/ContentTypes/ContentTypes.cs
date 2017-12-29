using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using System;
using System.IO;

namespace IMD.Connect.SPO.Provisioning
{
    class ContentTypes
    {
        public static void CreateContentTypes()
        {
            try
            {
                //string file = "KMSProvisioningTemplate123.xml";
                //string directory = "D:\\DMS\\EDAMS";
                Console.WriteLine("Please provide Provisioning Template Path:");
                string FilePath = Console.ReadLine();
                FileInfo fileInfo = new FileInfo(FilePath);
                string directory = fileInfo.DirectoryName;
                string file = fileInfo.Name;

                AuthenticationManager authMgr = new AuthenticationManager();
                var provisioningProvider = new XMLFileSystemTemplateProvider(directory, string.Empty);
                var provisioningTemplate = provisioningProvider.GetTemplate(file);
                provisioningTemplate.Connector.Parameters[FileConnectorBase.CONNECTIONSTRING] = directory;
                OfficeDevPnP.Core.Framework.Provisioning.Model.ContentTypeCollection ContentTypes = provisioningTemplate.ContentTypes;

                using (var clientContext = authMgr.GetAppOnlyAuthenticatedContext(IMDConnect.SiteUrl, IMDConnect.ClientID, IMDConnect.ClientSecrete))
                {
                    Console.WriteLine("Content Types are Provisioning.......");
                    foreach (OfficeDevPnP.Core.Framework.Provisioning.Model.ContentType contenttype in ContentTypes)
                    {
                        if(!clientContext.Web.ContentTypeExistsByName(contenttype.Name,false))
                        { 
                            clientContext.Web.CreateContentType(contenttype.Name, contenttype.Id, contenttype.Group);
                            if(contenttype.Name!= "KMA CT")
                            {
                                OfficeDevPnP.Core.Framework.Provisioning.Model.FieldRefCollection cfr = contenttype.FieldRefs;
                                foreach (OfficeDevPnP.Core.Framework.Provisioning.Model.FieldRef fr in cfr)
                                {
                                    clientContext.Web.AddFieldToContentTypeById(contenttype.Id, fr.Id.ToString(), false, false);
                                }
                            }
                        }
                        clientContext.ExecuteQuery();
                    }
                    Console.WriteLine("Content Types Provisioning is done Successfully");
                    Console.ReadLine();
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message.ToString());
                Console.ReadLine();
            }
        }
    }
}

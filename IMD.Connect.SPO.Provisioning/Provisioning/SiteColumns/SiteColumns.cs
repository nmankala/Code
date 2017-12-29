using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using System;
using System.IO;

namespace IMD.Connect.SPO.Provisioning
{
    class SiteColumns
    {
        public static void CreateSiteColumns()
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
                OfficeDevPnP.Core.Framework.Provisioning.Model.FieldCollection SiteColumns = provisioningTemplate.SiteFields;
                using (var clientContext = authMgr.GetAppOnlyAuthenticatedContext(IMDConnect.SiteUrl,IMDConnect.ClientID,IMDConnect.ClientSecrete))
                {
                    Console.WriteLine("Site Columns are Provisioning....");                
                    foreach (OfficeDevPnP.Core.Framework.Provisioning.Model.Field sitecolumn in SiteColumns)
                    {
                            if (!clientContext.Web.FieldExistsById(sitecolumn.SchemaXml.Substring(12, 36),false))
                            {
                                clientContext.Web.Fields.AddFieldAsXml(sitecolumn.SchemaXml, false, AddFieldOptions.AddFieldInternalNameHint);
                            }                        
                    }
                    clientContext.ExecuteQuery();
                    Console.WriteLine("Site Columns Provisioning is done Succcessfully.");
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

using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using OfficeDevPnP.Core.Framework.Provisioning.Providers;

namespace IMD.Connect.SPO.Provisioning
{
    class MasterLists
    {
        public static void CreateMasterLists()
        {
            AuthenticationManager authMgr = new AuthenticationManager();
            //string file = "KMSProvisioningTemplate_LiveData.xml";
            string file = "KMSProvisioningTemplate123.xml";
            string directory = "D:\\DMS\\EDAMS";
            var provisioningProvider = new XMLFileSystemTemplateProvider(directory, string.Empty);
            var provisioningTemplate = provisioningProvider.GetTemplate(file);
            provisioningTemplate.Connector.Parameters[FileConnectorBase.CONNECTIONSTRING] = directory;
            OfficeDevPnP.Core.Framework.Provisioning.Model.ListInstanceCollection listcoll = provisioningTemplate.Lists;
            using (var clientContext = authMgr.GetAppOnlyAuthenticatedContext(IMDConnect.SiteUrl, IMDConnect.ClientID, IMDConnect.ClientSecrete))
            {
                foreach (OfficeDevPnP.Core.Framework.Provisioning.Model.ListInstance list in listcoll)
                {                  
                   List mlist = clientContext.Web.CreateList(ListTemplateType.GenericList, list.Title, true, false, list.Url, true);
                    clientContext.ExecuteQuery();
                    if (list.ContentTypeBindings.Count>2)
                    {
                        OfficeDevPnP.Core.Framework.Provisioning.Model.ContentTypeBindingCollection contenttypecoll = list.ContentTypeBindings;

                        foreach(OfficeDevPnP.Core.Framework.Provisioning.Model.ContentTypeBinding ctb in contenttypecoll)
                        {
                           
                            mlist.AddContentTypeToListById(ctb.ContentTypeId, false, false);
                        }
                    }                  
                    clientContext.ExecuteQuery();
                }
            }

        }
    }




}

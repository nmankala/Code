﻿using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using OfficeDevPnP.Core.Framework.Provisioning.Providers;

namespace IMD.Connect.SPO.Provisioning
{
    class Lists
    {
        public static void CreateLists()
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
                    OfficeDevPnP.Core.Framework.Provisioning.Model.FieldCollection fc = list.Fields;
                    foreach (OfficeDevPnP.Core.Framework.Provisioning.Model.Field fld in fc)
                    {
                       Field objField = mlist.Fields.AddFieldAsXml(fld.SchemaXml,true, AddFieldOptions.DefaultValue); objField.Update();                      
                    }                   
                    OfficeDevPnP.Core.Framework.Provisioning.Model.DataRowCollection ldr = list.DataRows;
                    foreach(OfficeDevPnP.Core.Framework.Provisioning.Model.DataRow dr in ldr)
                    {
                        ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();                   
                        ListItem oListItem = mlist.AddItem(itemCreateInfo);
                        oListItem["Title"] = dr.Values["Title"];
                        oListItem["Asset_x0020_Type"] = dr.Values["Asset Type"];
                        oListItem["Asset_x0020_Version"] = dr.Values["Asset Version"];
                        oListItem.Update();
                        
                    }
                    clientContext.ExecuteQuery();
                }
            }
                
            }
        }


       
  
}

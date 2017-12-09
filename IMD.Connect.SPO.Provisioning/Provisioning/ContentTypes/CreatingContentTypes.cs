using System;
using System.Xml;
using System.Management.Automation.Runspaces;
using Microsoft.SharePoint.Client;
using System.Linq;

namespace IMD.Connect.SPO.Provisioning
{
    class CreatingContentTypes
    {
        public static void ContentTypeCreation(string filePath)
        {
            XmlDocument xmlDoc = new XmlDocument();
            if (System.IO.File.Exists(filePath))
            {
                try
                {
                    xmlDoc.Load(filePath);
                    XmlNode ContentTypes = xmlDoc.SelectSingleNode("/ProvisioningTemplate/ContentTypes");
                    using (var ctx = CommonConnection.CreateClientContext1())
                    {
                        using (var scope = new ConnectionScope(true))
                        {
                            foreach (XmlNode node in ContentTypes.ChildNodes)
                            {
                                if(!ctx.Web.ContentTypeExistsByName(node.Attributes["Name"].Value))
                                {
                                    scope.ExecuteCommand("Add-PnPContentType",
                                                            new CommandParameter("Name", node.Attributes["Name"].Value),
                                                            new CommandParameter("Group", node.Attributes["Group"].Value));
                                    Console.WriteLine("Content Type is created, now adding site columns");
                                    foreach (XmlNode childnode in node.ChildNodes)
                                    {
                                        foreach( XmlNode schildnode in childnode.ChildNodes)
                                        {
                                            Guid fieldId = new Guid(schildnode.Attributes["ID"].Value);
                                            Field fld = ctx.Site.RootWeb.GetFieldById(fieldId);
                                            if(!ctx.Site.RootWeb.FieldExistsByNameInContentType(node.Attributes["Name"].Value, fld.Title))
                                            {
                                                scope.ExecuteCommand("Add-PnPFieldToContentType",
                                                       new CommandParameter("ContentType", node.Attributes["Name"].Value),
                                                       new CommandParameter("Field", fld.Title));
                                            }
                                            else
                                            {
                                                Console.WriteLine("Site Column is already available in Content Type");
                                            }
                                           
                                        }
                                    }                                   
                                }                                  
                                else
                                {
                                    Console.WriteLine(node.Attributes["Name"].Value + " exists in the site");
                                }
                                                                   
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message.ToString());
                }
            }

        }
    }
}

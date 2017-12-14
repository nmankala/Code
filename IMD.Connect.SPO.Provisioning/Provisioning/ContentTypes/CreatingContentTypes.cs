using System;
using System.Xml;
using System.Management.Automation.Runspaces;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace IMD.Connect.SPO.Provisioning
{
    class CreatingContentTypes
    {
        public static void ContentTypeCreation()
        {
            string FilePath = null;
            XmlDocument xmlDoc = new XmlDocument();
            Console.WriteLine("Please provide provisioning XML File");
            FilePath = Console.ReadLine();
            Console.WriteLine("Creating Content Types and Adding Site Columns...");
            if (System.IO.File.Exists(FilePath))
            {
                try
                {
                    xmlDoc.Load(FilePath);
                    XmlNode ContentTypes = xmlDoc.SelectSingleNode("/ProvisioningTemplate/ContentTypes");
                    using (var ctx = CommonConnection.CreateClientContext1())
                    {
                        using (var scope = new ConnectionScope(true))
                        {
                            foreach (XmlNode node in ContentTypes.ChildNodes)
                            {
                                var succeeded = false;
                                if (!ctx.Web.ContentTypeExistsByName(node.Attributes["Name"].Value))
                                {
                                    scope.ExecuteCommand("Add-PnPContentType",
                                                            new CommandParameter("Name", node.Attributes["Name"].Value),
                                                            new CommandParameter("Group", node.Attributes["Group"].Value));
                                    succeeded = true;
                                    Console.WriteLine("The New Content Type " + node.Attributes["Name"].Value + " has been created");
                                    Console.WriteLine("Adding Site Columns to " + node.Attributes["Name"].Value + " Content Type......");
                                }
                                else
                                {
                                    Console.WriteLine("The Content Type " + node.Attributes["Name"].Value + " is already exists in the site");
                                    succeeded = true;
                                    Console.WriteLine("Adding the Site Columns to " + node.Attributes["Name"].Value + " Content Type......");
                                }

                                if (succeeded)
                                {
                                    foreach (XmlNode childnode in node.ChildNodes)
                                    {
                                        foreach (XmlNode schildnode in childnode.ChildNodes)
                                        {
                                            Guid fieldId = new Guid(schildnode.Attributes["ID"].Value);
                                            Field fld = ctx.Site.RootWeb.GetFieldById(fieldId);
                                            if (!ctx.Site.RootWeb.FieldExistsByNameInContentType(node.Attributes["Name"].Value, fld.InternalName))
                                            {
                                                scope.ExecuteCommand("Add-PnPFieldToContentType",
                                                       new CommandParameter("ContentType", node.Attributes["Name"].Value),
                                                       new CommandParameter("Field", fld.InternalName));
                                                Console.WriteLine("The Site Columns "+ fld.Title+ " is added to " + node.Attributes["Name"].Value+ " Content type");
                                            }
                                            else
                                            {
                                                Console.WriteLine("Site Column is already there in  Content Type");
                                            }
                                        }
                                    }
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

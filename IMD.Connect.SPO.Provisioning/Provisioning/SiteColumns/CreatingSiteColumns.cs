using System;
using System.Management.Automation.Runspaces;
using System.Xml;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using System.Linq;

namespace IMD.Connect.SPO.Provisioning
{
    class CreatingSiteColumns
    {
        public static void SiteCoumnsCreation(string filePath)
        {
            XmlDocument xmlDoc = new XmlDocument();
            if (System.IO.File.Exists(filePath))
            {
                try
                {
                    xmlDoc.Load(filePath);
                    XmlNode SiteColumns = xmlDoc.SelectSingleNode("/ProvisioningTemplate/SiteFields");             
                    using (var ctx = CommonConnection.CreateClientContext1())
                    {
                        using (var scope = new ConnectionScope(true))
                        {
                            foreach (XmlNode node in SiteColumns.ChildNodes)
                            {
                                if (!ctx.Web.FieldExistsByName(node.Attributes["Name"].Value))
                                {
                                    if (node.Attributes["Type"].Value == "MMD")
                                        CreateManagedMetaDataSiteColumns(ctx, node.Attributes["DisplayName"].Value, node.Attributes["Name"].Value, node.Attributes["Group"].Value, node.Attributes["MMDValue"].Value);
                                    else
                                        scope.ExecuteCommand("Add-PnPFieldFromXml", new CommandParameter("FieldXml", node.OuterXml));
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
        static void CreateManagedMetaDataSiteColumns(ClientContext cContext, string displayname, string internalname, string group, string mmdvalue)
        {
            Web rootWeb = cContext.Site.RootWeb;
            Field field = rootWeb.Fields.AddFieldAsXml("<Field DisplayName='" + displayname + "' Name='" + internalname + "' ID='{" + Guid.NewGuid() + "}' Group='" + group + "' Type='TaxonomyFieldTypeMulti' />", false, AddFieldOptions.AddFieldInternalNameHint);
            cContext.ExecuteQuery();
            Guid termStoreId = Guid.Empty;
            Guid termSetId = Guid.Empty;
            GetTaxonomyFieldInfo(cContext, mmdvalue, out termStoreId, out termSetId);
            // Retrieve as Taxonomy Field
            TaxonomyField taxonomyField = cContext.CastTo<TaxonomyField>(field);
            taxonomyField.SspId = termStoreId;
            taxonomyField.TermSetId = termSetId;
            taxonomyField.TargetTemplate = String.Empty;
            taxonomyField.AnchorId = Guid.Empty;
            taxonomyField.Update();
            cContext.ExecuteQuery();
            Console.WriteLine("New " + displayname + " Column is created");
        }
        static void GetTaxonomyFieldInfo(ClientContext clientContext, string TermsetName, out Guid termStoreId, out Guid termSetId)
        {
            termStoreId = Guid.Empty;
            termSetId = Guid.Empty;
            TaxonomySession session = TaxonomySession.GetTaxonomySession(clientContext);
            TermStore termStore = session.GetDefaultSiteCollectionTermStore();
            TermSetCollection termSets = termStore.GetTermSetsByName(TermsetName, 1033);
            clientContext.Load(termSets, tsc => tsc.Include(ts => ts.Id));
            clientContext.Load(termStore, ts => ts.Id);
            clientContext.ExecuteQuery();
            termStoreId = termStore.Id;
            termSetId = termSets.FirstOrDefault().Id;
        }
    }
}
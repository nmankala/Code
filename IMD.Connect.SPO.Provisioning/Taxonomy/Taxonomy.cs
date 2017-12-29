using OfficeDevPnP.Core;
using Microsoft.SharePoint.Client.Taxonomy;
using System;
using System.Xml;

namespace IMD.Connect.SPO.Provisioning
{
    class Taxonomy
    {
        public static void CreateTaxnomy()
        {
            AuthenticationManager authMgr = new AuthenticationManager();
            XmlDocument xmlDoc = new XmlDocument();
            Console.WriteLine("Please provide TermStore Name");
            string TermstoreName = Console.ReadLine();
            Console.WriteLine("Please provide TermGroupID Name");
            string TermGroupID = Console.ReadLine();
            Console.WriteLine("Please provide FilePath");
            string file = Console.ReadLine();
            //string TermstoreName = "Taxonomy_tkZRZMBGOb2pGyCJoiwcbQ==";
            //string TermGroupID = "25ab5f63-1180-4786-8b22-7ea9977fb29e";
            //string file = "D:\\DMS\\EDAMS\\IMDTaxonomy.xml";

            xmlDoc.Load(file);
            XmlNode TermSets = xmlDoc.SelectSingleNode("/Group");
            using (var clientContext = authMgr.GetAppOnlyAuthenticatedContext(IMDConnect.SiteUrl, IMDConnect.ClientID, IMDConnect.ClientSecrete))
            {
                Console.WriteLine("IMD Term Store Provisioning.....");
                TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(clientContext);
                TermStore termStore = taxonomySession.TermStores.GetByName(TermstoreName);
                Guid guid = new Guid(TermGroupID);
                TermGroup termGroup = termStore.GetGroup(guid);
                int lcid = 1033;
                foreach (XmlNode node in TermSets.ChildNodes)
                {
                    Guid guid1 = new Guid(node.Attributes["Id"].Value);
                    TermSet termSetColl = termGroup.CreateTermSet(node.Attributes["Name"].Value, guid1, lcid);
                    foreach (XmlNode schildnode in node.ChildNodes)
                    {
                        Guid guid2 = new Guid(schildnode.Attributes["Id"].Value);
                        termSetColl.CreateTerm(schildnode.Attributes["Name"].Value, lcid, guid2);
                    }
                }
                clientContext.ExecuteQuery();
                Console.WriteLine("IMD Term Store is Created Successfully.");
                Console.ReadLine();
            }
        }
    }
}

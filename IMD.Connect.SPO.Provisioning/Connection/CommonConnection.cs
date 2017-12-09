using Microsoft.SharePoint.Client;
using System;
using System.Configuration;
using System.Security;
using System.Net;
using System.Threading;

namespace IMD.Connect.SPO.Provisioning
{
    class CommonConnection
    {
        #region Constructor
        static CommonConnection()
        {
            // Read configuration data
            DevSiteUrl = IMDConnect.SiteUrl;
            AppId = IMDConnect.ClientID;
            AppSecret = IMDConnect.ClientSecrete;

            if (string.IsNullOrEmpty(DevSiteUrl))
            {
                throw new ConfigurationErrorsException("Tenant site Url or Dev site url in App.config are not set up.");
            }

            if (string.IsNullOrEmpty(AppId) || string.IsNullOrEmpty(AppSecret))
            {
                throw new ConfigurationErrorsException("Tenant site Url or Dev site url in App.config are not set up.");
            }

            // Trim trailing slashes
            DevSiteUrl = DevSiteUrl.TrimEnd(new[] { '/' });

        }
        #endregion

        #region Properties
        public static string TenantUrl { get; set; }
        public static string DevSiteUrl { get; set; }
        static string UserName { get; set; }
        static SecureString Password { get; set; }
        static ICredentials Credentials { get; set; }
        static string Realm { get; set; }
        public static string AppId { get; set; }
        public static string AppSecret { get; set; }
        public static string FilePath { get; set; }
        //public static String AzureStorageKey
        //{
        //    get
        //    {
        //        return ConfigurationManager.AppSettings["AzureStorageKey"];
        //    }
        //}
        //public static string WebHookTestUrl
        //{
        //    get
        //    {
        //        return ConfigurationManager.AppSettings["WebHookTestUrl"];
        //    }
        //}
        #endregion

        #region Methods
        public static ClientContext CreateClientContext()
        {
            return CreateContext(DevSiteUrl, Credentials);
        }


        public static ClientContext CreateClientContext1()
        {
            return CreateContext1(DevSiteUrl, Credentials);
        }
        public static ClientContext CreateTenantClientContext()
        {
            return CreateContext(TenantUrl, Credentials);
        }

        public static bool AppOnlyTesting()
        {
            if (!String.IsNullOrEmpty(ConfigurationManager.AppSettings["AppId"]) &&
                !String.IsNullOrEmpty(ConfigurationManager.AppSettings["AppSecret"]) &&
                String.IsNullOrEmpty(ConfigurationManager.AppSettings["SPOCredentialManagerLabel"]) &&
                String.IsNullOrEmpty(ConfigurationManager.AppSettings["SPOUserName"]) &&
                String.IsNullOrEmpty(ConfigurationManager.AppSettings["SPOPassword"]) &&
                String.IsNullOrEmpty(ConfigurationManager.AppSettings["OnPremUserName"]) &&
                String.IsNullOrEmpty(ConfigurationManager.AppSettings["OnPremDomain"]) &&
                String.IsNullOrEmpty(ConfigurationManager.AppSettings["OnPremPassword"]))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private static ClientContext CreateContext(string contextUrl, ICredentials credentials)
        {
            ClientContext context = null;
            if (!String.IsNullOrEmpty(AppId) && !String.IsNullOrEmpty(AppSecret))
            {
                OfficeDevPnP.Core.AuthenticationManager am = new OfficeDevPnP.Core.AuthenticationManager();

                if (new Uri(DevSiteUrl).DnsSafeHost.Contains("spoppe.com"))
                {
                    context = am.GetAppOnlyAuthenticatedContext(DevSiteUrl, SPOnlineConnectionHelper.GetRealmFromTargetUrl(new Uri(DevSiteUrl)), AppId, AppSecret, acsHostUrl: "windows-ppe.net", globalEndPointPrefix: "login");
                }
                else
                {
                    context = am.GetAppOnlyAuthenticatedContext(DevSiteUrl, AppId, AppSecret);
                }
            }
            else
            {
                context = new ClientContext(DevSiteUrl);
                context.Credentials = Credentials;
            }

            context.RequestTimeout = Timeout.Infinite;
            return context;
        }

        private static ClientContext CreateContext1(string contextUrl, ICredentials credentials)
        {
            ClientContext context = null;
            if (!String.IsNullOrEmpty(AppId) && !String.IsNullOrEmpty(AppSecret))
            {
                OfficeDevPnP.Core.AuthenticationManager am = new OfficeDevPnP.Core.AuthenticationManager();

                if (new Uri(DevSiteUrl).DnsSafeHost.Contains("spoppe.com"))
                {
                    context = am.GetAppOnlyAuthenticatedContext(DevSiteUrl, SPOnlineConnectionHelper.GetRealmFromTargetUrl(new Uri(DevSiteUrl)), AppId, AppSecret, acsHostUrl: "windows-ppe.net", globalEndPointPrefix: "login");
                }
                else
                {
                    context = am.GetAppOnlyAuthenticatedContext(DevSiteUrl, AppId, AppSecret);
                }
            }
            else
            {
                context = new ClientContext(DevSiteUrl);
                context.Credentials = Credentials;
            }

            context.RequestTimeout = Timeout.Infinite;
            return context;
        }

        private static SecureString GetSecureString(string input)
        {
            if (string.IsNullOrEmpty(input))
                throw new ArgumentException("Input string is empty and cannot be made into a SecureString", "input");

            var secureString = new SecureString();
            foreach (char c in input.ToCharArray())
                secureString.AppendChar(c);

            return secureString;
        }
        #endregion
    }
}

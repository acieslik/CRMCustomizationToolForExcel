
using System;
using System.Collections.Generic;
using System.ServiceModel.Description;
using System.IO;
using System.Xml;
using System.Xml.Linq;
using System.Security;
using System.Runtime.InteropServices;
using System.ServiceModel;
using System.DirectoryServices.AccountManagement;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Discovery;
using Microsoft.Xrm.Sdk;
using System.Text;
using System.Net;
using System.Text.RegularExpressions;

namespace  DynamicsCRMCustomizationToolForExcel.Controller

{
    public class ServerConnection
    {
        #region Connection Configuration class
        public class Configuration
        {
            public String ServerAddress;
            public String OrganizationName;
            public String UserName;
            public String Password;
            public Uri DiscoveryUri;
            public Uri OrganizationUri;
            public bool Ssl;
            public ClientCredentials DeviceCredentials = null;
            public ClientCredentials Credentials = null;
            public AuthenticationProviderType EndpointType;
            public String UserPrincipalName;
            internal IServiceManagement<IOrganizationService> OrganizationServiceManagement;
            internal SecurityTokenResponse OrganizationTokenResponse;
            internal Int16 AuthFailureCount = 0;

        }
        #endregion

        private Configuration config = new Configuration();
        private OrganizationDetailCollection organizations;
        private static readonly Random RandomInstance = new Random();

        public ServerConnection(bool Ssl, string ServerAddress)
        {
            config.ServerAddress = ServerAddress;
            config.Ssl = Ssl;
            config.Credentials = new ClientCredentials();
            config.Credentials.Windows.ClientCredential = (NetworkCredential)CredentialCache.DefaultCredentials;
            config.Credentials.Windows.ClientCredential = System.Net.CredentialCache.DefaultNetworkCredentials;
           
        }

        public ServerConnection(bool Ssl, string ServerAddress, string user, string password)
        {
            config.UserName = user;
            config.Password = password;
            config.ServerAddress = ServerAddress;
            config.Ssl = Ssl;
            config.Credentials = new ClientCredentials();
            config.Credentials.Windows.ClientCredential.UserName = user;
            config.Credentials.Windows.ClientCredential.Password = password;
        }

        public string OrganizationName
        {
            set {
                config.OrganizationName = value;
                OrganizationDetail orgDetails = FindOrganization(value,organizations);
                config.OrganizationUri = new System.Uri(orgDetails.Endpoints[EndpointType.OrganizationService]);
            }
        }


        public string ServerAddress
        {
            get
            {
                return config.ServerAddress;
            }
        }


        public  OrganizationServiceProxy GetOrganizationProxy()
        {
            if (config.OrganizationServiceManagement != null
                && config.OrganizationTokenResponse != null)
            {
                return new OrganizationServiceProxy(
                    config.OrganizationServiceManagement,
                    config.OrganizationTokenResponse);
            }

            if (config.OrganizationServiceManagement == null)
                throw new ArgumentNullException("serverConfiguration.OrganizationServiceManagement");

            return new OrganizationServiceProxy(
                config.OrganizationServiceManagement,
                config.Credentials);

        }

        public OrganizationDetailCollection GetDiscoveryService()
        {
            if (config.ServerAddress.EndsWith("dynamics.com") || String.IsNullOrWhiteSpace(config.ServerAddress))
            {
                config.Ssl = true;
            }
            // Check if the organization is provisioned in Microsoft Office 365.
            if (config.ServerAddress.EndsWith("dynamics.com", StringComparison.InvariantCultureIgnoreCase) && config.UserName.EndsWith(".onmicrosoft.com"))
            {
                config.DiscoveryUri = new Uri(String.Format("https://disco.{0}/XRMServices/2011/Discovery.svc", config.ServerAddress));
            }
            // One of the Microsoft Dynamics CRM Online data centers.
            else if (config.ServerAddress.EndsWith("dynamics.com", StringComparison.InvariantCultureIgnoreCase))
            {
                config.DiscoveryUri = new Uri(String.Format("https://dev.{0}/XRMServices/2011/Discovery.svc", config.ServerAddress));
            }
            // Check if the server uses Secure Socket Layer (https).
            else if (config.Ssl)
                config.DiscoveryUri = new Uri(String.Format("https://{0}/XRMServices/2011/Discovery.svc", config.ServerAddress));
            else
                config.DiscoveryUri =
                    new Uri(String.Format("http://{0}/XRMServices/2011/Discovery.svc", config.ServerAddress));
              return GetOrganizationAddress();
       }


        //private static DeviceUserName GenerateDeviceUserName()
        //{
        //    DeviceUserName userName = new DeviceUserName();
        //    userName.DeviceName = GenerateRandomString("0123456789abcdefghijklmnopqrstuvqxyz", 24);
        //    userName.DecryptedPassword = GenerateRandomString("0123456789abcdefghijklmnopqrstuvqxyz", 24);

        //    return userName;
        //}

        private static string GenerateRandomString(string characterSet, int count)
        {
            //Create an array of the characters that will hold the final list of random characters
            char[] value = new char[count];

            //Convert the character set to an array that can be randomly accessed
            char[] set = characterSet.ToCharArray();

            lock (RandomInstance)
            {
                //Populate the array with random characters from the character set
                for (int i = 0; i < count; i++)
                {
                    value[i] = set[RandomInstance.Next(0, set.Length)];
                }
            }

            return new string(value);
        }



        //private ClientCredentials loadDevice()
        //{
        //    LiveDevice device = new LiveDevice(){ User = GenerateDeviceUserName(), Version = 1 };
        //    return device.User.ToClientCredentials();
        //}



        //private AuthenticationCredentials GetCredentials()
        //{

        //    AuthenticationCredentials authCredentials = new AuthenticationCredentials();
        //    switch (config.EndpointType)
        //    {

        //        case AuthenticationProviderType.LiveId:
        //            authCredentials.ClientCredentials.UserName.UserName = config.UserName;
        //            authCredentials.ClientCredentials.UserName.Password = config.Password; ;
        //              authCredentials.SupportingCredentials = new AuthenticationCredentials();
        //              authCredentials.SupportingCredentials.ClientCredentials = loadDevice();
        //            //    Microsoft.Crm.Services.Utility.DeviceIdManager.LoadOrRegisterDevice();
        //            break;
        //        default:
        //            authCredentials.ClientCredentials.UserName.UserName = config.UserName;
        //            authCredentials.ClientCredentials.UserName.Password = config.Password; ;
        //            break;
        //    }

        //    return authCredentials;
        //}



        //public IServiceManagement<IOrganizationService> ConnectToOrganization()
        //{
        //    // Set IServiceManagement for the current organization.
        //    IServiceManagement<IOrganizationService> orgServiceManagement =
        //            ServiceConfigurationFactory.CreateManagement<IOrganizationService>(
        //            config.OrganizationUri);
        //    config.OrganizationServiceManagement = orgServiceManagement;

        //    // Set SecurityTokenResponse for the current organization.
        //    if (config.EndpointType != AuthenticationProviderType.ActiveDirectory)
        //    {
        //        // Set the credentials.
        //        AuthenticationCredentials authCredentials = new AuthenticationCredentials();
        //        // If UserPrincipalName exists, use it. Otherwise, set the logon credentials from the configuration.
        //        if (!String.IsNullOrWhiteSpace(config.UserPrincipalName))
        //        {
        //            authCredentials.UserPrincipalName = config.UserPrincipalName;
        //        }
        //        else
        //        {
        //            authCredentials = GetCredentials();
        //        }
        //        AuthenticationCredentials tokenCredentials =
        //            orgServiceManagement.Authenticate(authCredentials);

        //        if (tokenCredentials != null)
        //        {
        //            if (tokenCredentials.SecurityTokenResponse != null)
        //                config.OrganizationTokenResponse = tokenCredentials.SecurityTokenResponse;
        //        }
        //    }
        //    return orgServiceManagement;
        //}
        /// <summary>
        /// Discovers the organizations that the calling user belongs to.
        /// </summary>
        /// <param name="service">A Discovery service proxy instance.</param>
        /// <returns>Array containing detailed information on each organization that 
        /// the user belongs to.</returns>
        public OrganizationDetailCollection DiscoverOrganizations(IDiscoveryService service)
        {
            if (service == null) throw new ArgumentNullException("service");
            RetrieveOrganizationsRequest orgRequest = new RetrieveOrganizationsRequest();
            RetrieveOrganizationsResponse orgResponse =
                (RetrieveOrganizationsResponse)service.Execute(orgRequest);

            return orgResponse.Details;
        }

        /// <summary>
        /// Finds a specific organization detail in the array of organization details
        /// returned from the Discovery service.
        /// </summary>
        /// <param name="orgFriendlyName">The friendly name of the organization to find.</param>
        /// <param name="orgDetails">Array of organization detail object returned from the discovery service.</param>
        /// <returns>Organization details or null if the organization was not found.</returns>
        /// <seealso cref="DiscoveryOrganizations"/>
        public OrganizationDetail FindOrganization(string orgFriendlyName, OrganizationDetailCollection orgDetails)
        {
            if (String.IsNullOrWhiteSpace(orgFriendlyName))
                throw new ArgumentNullException("orgFriendlyName");
            if (orgDetails == null)
                throw new ArgumentNullException("orgDetails");
            OrganizationDetail orgDetail = new OrganizationDetail();

            foreach (OrganizationDetail detail in orgDetails)
            {
                if (String.Compare(detail.FriendlyName, orgFriendlyName,
                    StringComparison.InvariantCultureIgnoreCase) == 0)
                {
                    orgDetail = detail;
                    break;
                }
            }
            return orgDetail;
        }


        #region Protected methods


        /// <summary>
        /// Obtains the web address (Uri) of the target organization.
        /// </summary>
        /// <returns>Uri of the organization service or an empty string.</returns>
        protected virtual OrganizationDetailCollection GetOrganizationAddress()
        {
            using (DiscoveryServiceProxy serviceProxy = GetDiscoveryProxy())
            {
                // Obtain organization information from the Discovery service. 
                if (serviceProxy != null)
                {
                    // Obtain information about the organizations that the system user belongs to.
                    organizations = DiscoverOrganizations(serviceProxy);
                    return organizations;
                }
                else
                    throw new Exception("An invalid server name was specified.");
            }
        }

 

        #endregion Protected methods

       
        private DiscoveryServiceProxy GetDiscoveryProxy()
        {

            IServiceManagement<IDiscoveryService> serviceManagement =
                        ServiceConfigurationFactory.CreateManagement<IDiscoveryService>(
                        config.DiscoveryUri);

            // Get the EndpointType.
            config.EndpointType = serviceManagement.AuthenticationType;

            AuthenticationCredentials authCredentials = new AuthenticationCredentials();

            if (!String.IsNullOrWhiteSpace(config.UserPrincipalName))
            {
                // Try to authenticate the Federated Identity organization with UserPrinicipalName.
                authCredentials.UserPrincipalName = config.UserPrincipalName;

                try
                {
                    AuthenticationCredentials tokenCredentials = serviceManagement.Authenticate(
                        authCredentials);
                    DiscoveryServiceProxy discoveryProxy = new DiscoveryServiceProxy(serviceManagement,
                        tokenCredentials.SecurityTokenResponse);
                    // Checking authentication by invoking some SDK methods.
                    OrganizationDetailCollection orgs = DiscoverOrganizations(discoveryProxy);
                    return discoveryProxy;
                }
                catch (System.ServiceModel.Security.SecurityAccessDeniedException ex)
                {
                    // If authentication failed using current UserPrincipalName, 
                    // request UserName and Password to try to authenticate using user credentials.
                    if (ex.Message.Contains("Access is denied."))
                    {
                        config.AuthFailureCount += 1;
                        authCredentials.UserPrincipalName = String.Empty;
                    }
                    else
                    {
                        throw ex;
                    }
                }
            
            }

            // Resetting credentials in the AuthenicationCredentials.  
            if (config.EndpointType != AuthenticationProviderType.ActiveDirectory)
            {
                //authCredentials = GetCredentials(); 
                 //Try to authenticate with the user credentials.
                AuthenticationCredentials tokenCredentials = serviceManagement.Authenticate(authCredentials);
                return new DiscoveryServiceProxy(serviceManagement,tokenCredentials.SecurityTokenResponse);

           
            }
            // For an on-premises environment.
            return new DiscoveryServiceProxy(serviceManagement, config.Credentials);
        }


    }
}
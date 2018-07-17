using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Deployment.Application; //TOCHECK
using System.Xml.Serialization;

namespace DynamicsCRMCustomizationToolForExcel.Model
{
    [Serializable]
    public class ConnectionConfiguration
    {
        public bool localuser { set; get; }
        public string serverUri{set;get;}
        public string name{set;get;}
        public bool ssl { set; get; }
        public string user{set;get;}
        public bool office365{set;get;}
        public string organization{set;get;}
        //private string SecurePassword;
       //// public string Password
       // {
       //     get
       //     {
       //         if (!String.IsNullOrEmpty(SecurePassword)) 
       //         {
       //             try
       //             {
       //                 byte[] secPass = ProtectedData.Unprotect(Convert.FromBase64String(SecurePassword), null, DataProtectionScope.CurrentUser);
       //                 string base64SecPass = Convert.ToBase64String(secPass);
       //                 byte[] base64Pass = Convert.FromBase64String(base64SecPass);
       //                 return Encoding.Unicode.GetString(base64Pass);
       //             }
       //             catch (Exception) { }
       //         }
       //         return String.Empty;
       //     }
       //     set
       //     {
       //         if (String.IsNullOrEmpty(value))
       //         {
       //             SecurePassword = String.Empty;
       //         }
       //         else
       //         {
       //             try
       //             {
       //                 string base64Pass = Convert.ToBase64String(Encoding.Unicode.GetBytes(value));
       //                 byte[] secPass = ProtectedData.Protect(Convert.FromBase64String(base64Pass), null, DataProtectionScope.CurrentUser);
       //                 SecurePassword = Convert.ToBase64String(secPass);
       //             }
       //             catch (Exception)
       //             {
       //                 SecurePassword = String.Empty;
       //             }
       //         }
       //     }
       // }
    }

    public class ConnectionsData
    {
        private const string FILENAME = "SavedConnection.xml";
        public List<ConnectionConfiguration> ConfigurationList;

        public ConnectionsData()
        {
            if (File.Exists(getConfigFilePath))
            {
                readConnectionData();
            }
            else
            {
                ConfigurationList = new List<ConnectionConfiguration>();
            }
        }

        private string getConfigFilePath
        {
            get
            {
                string StrXmlPath;
                if (ApplicationDeployment.IsNetworkDeployed)
                {
                    StrXmlPath = Path.Combine(ApplicationDeployment.CurrentDeployment.DataDirectory, FILENAME);
                }
                else
                {
                    StrXmlPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, FILENAME);
                }
                return StrXmlPath;
            }
        }

        public void storeConnectionData()
        {
            try
            {
                var serializer = new XmlSerializer(typeof(ConnectionConfiguration[]));

                using (var writer = new StreamWriter(getConfigFilePath))
                {
                    serializer.Serialize(writer, ConfigurationList.ToArray());
                }
            }
            catch (Exception )
            {

            }
        }

        private void readConnectionData()
        {
            try
            {
                XmlSerializer serializer = new XmlSerializer(typeof(ConnectionConfiguration[]));
                using (var reader = new StreamReader(getConfigFilePath))
                {
                    ConnectionConfiguration[] connConfg = (ConnectionConfiguration[])serializer.Deserialize(reader);
                    if (connConfg != null)
                    {
                        ConfigurationList = connConfg.ToList();
                    }
                    else
                    {
                        ConfigurationList = new List<ConnectionConfiguration>();
                    }
                }
            }
            catch (Exception )
            {
                ConfigurationList = new List<ConnectionConfiguration>();
            }
        }

    }
}

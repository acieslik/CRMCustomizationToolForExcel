using Microsoft.Xrm.Tooling.Connector;
using Microsoft.Xrm.Tooling.CrmConnectControl;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace DynamicsCRMCustomizationToolForExcel.AddIn
{
    /// <summary>
    /// Interaction logic for CrmLogin.xaml
    /// </summary>
    public partial class CrmLogin : Window
    {
        #region Vars
        /// <summary>
        /// Microsoft.Xrm.Tooling.Connector services
        /// </summary>
        private CrmServiceClient CrmSvc = null;
        
        /// <summary>
        /// Bool flag to determine if there is a connection 
        /// </summary>
        private bool bIsConnectedComplete = false;
        /// <summary>
        /// CRM Connection Manager component. 
        /// </summary>
        private CrmConnectionManager mgr = null;
        /// <summary>
        ///  This is used to allow the UI to reset without closing 
        /// </summary>
        private bool resetUiFlag = false;
        #endregion

        #region Properties
        /// <summary>
        /// CRM Connection Manager 
        /// </summary>
        public CrmConnectionManager CrmConnectionMgr { get { return mgr; } }
        #endregion

        #region Event
        /// <summary>
        /// Raised when a connection to CRM has completed. 
        /// </summary>
        public event EventHandler ConnectionToCrmCompleted;
        #endregion


        public CrmLogin()
        {
            InitializeComponent();
            //// Should be used for testing only.
            //ServicePointManager.ServerCertificateValidationCallback = (sender, certificate, chain, sslPolicyErrors) =>
            //{
            //    MessageBox.Show("CertError");
            //    return true;
            //};
        }

        /// <summary>
        /// Raised when the window loads for the first time. 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            // This is the setup process for the login control. 
			// The login control uses a class called CrmConnectionManager to manage 
            // the interaction with CRM. This class can also be queried at later points
            // for information about the current connection. 
			// In this case, the login control is referred to as CrmLoginCtrl.			

            // Set off flag. 
            bIsConnectedComplete = false;

            // Initialize the CRM Connection manager. 
            mgr = new CrmConnectionManager();
            
            // Pass a reference to the current UI or container control. This is used to 
            // synchronize UI threads In the login control.            
            mgr.ParentControl = CrmLoginCtrl;

            // If you are using an unmanaged client, say Microsoft Excel, and need to 
            // store the config in the users local directory, you must set this option to true. 
            mgr.UseUserLocalDirectoryForConfigStore = true ;

            // If you are using an unmanaged client, you need to provide the name of 
            // an executable (.exe) to use to create app config key. 
            mgr.HostApplicatioNameOveride = "DynamicsCRMCustomizationToolForExcel.exe";
            
            // This sets the CRM Connection manager for CrmLoginCtrl. 
            CrmLoginCtrl.SetGlobalStoreAccess(mgr);
            
            // There are several modes to the login control UI.
            CrmLoginCtrl.SetControlMode(ServerLoginConfigCtrlMode.FullLoginPanel);
            
            // This wires an event that is raised when the login button is pressed.            
            CrmLoginCtrl.ConnectionCheckBegining += new EventHandler(CrmLoginCtrl_ConnectionCheckBegining);

            // This wires an event that is raised when an error in the connect process occurs. 
            CrmLoginCtrl.ConnectErrorEvent += new EventHandler<ConnectErrorEventArgs>(CrmLoginCtrl_ConnectErrorEvent);

            // This wires an event that is raised when a status event is returned. 
            CrmLoginCtrl.ConnectionStatusEvent += new EventHandler<ConnectStatusEventArgs>(CrmLoginCtrl_ConnectionStatusEvent);

            // This wires an event that is raised when the user clicks the cancel button. 
            CrmLoginCtrl.UserCancelClicked += new EventHandler(CrmLoginCtrl_UserCancelClicked);
            
            // This prompts the user to automatically sign in using the cached credentials
            // when signing for the second time or later.
            if (!mgr.RequireUserLogin())
            {
                if (MessageBox.Show("Credentials already saved in configuration.\nClick Yes to sign in using saved credentials.\nClick No to reset credentials to sign in.", "Auto Sign-In", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    // If RequireUserLogin is false, it means that there has been a successful login here before and the credentials are cached. 
                    CrmLoginCtrl.IsEnabled = false;
                    // When running an auto login,  you need to wire and listen to the events from the connection manager.
                    // Run Auto User Login process, Wire events. 
                    mgr.ServerConnectionStatusUpdate += new EventHandler<ServerConnectStatusEventArgs>(mgr_ServerConnectionStatusUpdate);
                    mgr.ConnectionCheckComplete += new EventHandler<ServerConnectStatusEventArgs>(mgr_ConnectionCheckComplete);
                    // Start the connection process. 
                    mgr.ConnectToServerCheck();

                    // Show the message grid. 
                    CrmLoginCtrl.ShowMessageGrid();
                }
            }
        }

        #region Events

        /// <summary>
        /// Updates from the Auto Login process. 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void mgr_ServerConnectionStatusUpdate(object sender, ServerConnectStatusEventArgs e)
        {
            // The Status event will contain information about the current login process. If connected is false, then there is no connection. 
            // Set the updated status of the loading process. 
            Dispatcher.Invoke(DispatcherPriority.Normal,
                               new System.Action(() =>
                               {
                                   this.Title = string.IsNullOrWhiteSpace(e.StatusMessage) ? e.ErrorMessage : e.StatusMessage;
                               }
                                   ));
        }

        /// <summary>
        /// Complete Event from the Auto Login process
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void mgr_ConnectionCheckComplete(object sender, ServerConnectStatusEventArgs e)
        {
            // The Status event will contain information about the current login process,  if Connected is false, then there is not yet a connection. 
            // Unwire events that we are not using anymore, this prevents issues if the user uses the control after a failed login. 
            ((CrmConnectionManager)sender).ConnectionCheckComplete -= mgr_ConnectionCheckComplete;
            ((CrmConnectionManager)sender).ServerConnectionStatusUpdate -= mgr_ServerConnectionStatusUpdate;

            if (!e.Connected)
            {
                // if its not connected pop the login screen here. 
                if (e.MultiOrgsFound)
                    MessageBox.Show("Unable to sign in to CRM using cached credentials. Org Not found", "Sign In failed");
                else
                    MessageBox.Show("Unable to sign in to CRM using cached credentials", "Sign In failed");

                resetUiFlag = true;
                CrmLoginCtrl.GoBackToLogin();
                // Bad Login Get back on the UI. 
                Dispatcher.Invoke(DispatcherPriority.Normal,
                       new System.Action(() =>
                       {
                           this.Title = "Failed to sign in with cached credentials.";
                           MessageBox.Show(this.Title, "Notification from ConnectionManager", MessageBoxButton.OK, MessageBoxImage.Error);
                           CrmLoginCtrl.IsEnabled = true;
                       }
                        ));
                resetUiFlag = false;
            }
            else
            {
                // On successful sign in, return to the UI 
                if (e.Connected && !bIsConnectedComplete)
                    ProcessSuccess();
            }

        }

        /// <summary>
        ///  Login control connection check. 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CrmLoginCtrl_ConnectionCheckBegining(object sender, EventArgs e)
        {
            bIsConnectedComplete = false;
            Dispatcher.Invoke(DispatcherPriority.Normal,
                               new System.Action(() =>
                               {
                                   this.Title = "Starting Login Process. ";
                                   CrmLoginCtrl.IsEnabled = true;
                               }
                                   ));
        }

        /// <summary>
        /// Login control connection check status event. 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CrmLoginCtrl_ConnectionStatusEvent(object sender, ConnectStatusEventArgs e)
        {
            // We are using the bIsConnectedComplete to check to make sure we only process this call once. 
            if (e.ConnectSucceeded && !bIsConnectedComplete)
                ProcessSuccess();
        }

        /// <summary>
        /// Login control error event. 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CrmLoginCtrl_ConnectErrorEvent(object sender, ConnectErrorEventArgs e)
        {
            //MessageBox.Show(e.ErrorMessage, "Error here");
        }

        /// <summary>
        /// Event raised when the user clicks Cancel in the common login control. 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CrmLoginCtrl_UserCancelClicked(object sender, EventArgs e)
        {
            if (!resetUiFlag)
            {
                this.Close();                
            }
        }

        #endregion

        /// <summary>
        /// This raises and processes Success
        /// </summary>
        private void ProcessSuccess()
        {
            resetUiFlag = true;
            bIsConnectedComplete = true;
            CrmSvc = mgr.CrmSvc;
            CrmLoginCtrl.GoBackToLogin();
            Dispatcher.Invoke(DispatcherPriority.Normal,
               new System.Action(() =>
               {
                   this.Title = "Notification from Parent";
                   CrmLoginCtrl.IsEnabled = true;
               }
                ));

            // Notify Caller that we are done with success. 
            if (ConnectionToCrmCompleted != null)
                ConnectionToCrmCompleted(this, null);

            resetUiFlag = false;
        }

        private void CrmLoginCtrl_Loaded(object sender, RoutedEventArgs e)
        {

        }

    }

    #region system.diagnostics settings for this control

    // Add or merge this section to your app to enable diagnostics on the use of the CRM login control and connection
    /*
  <system.diagnostics>
    <trace autoflush="true" />
    <sources>
      <source name="Microsoft.Xrm.Tooling.Connector.CrmServiceClient"
              switchName="Microsoft.Xrm.Tooling.Connector.CrmServiceClient"
              switchType="System.Diagnostics.SourceSwitch">
        <listeners>
          <add name="console" type="System.Diagnostics.DefaultTraceListener" />
          <remove name="Default"/>
          <add name ="fileListener"/>
        </listeners>
      </source>
      <source name="Microsoft.Xrm.Tooling.CrmConnectControl"
              switchName="Microsoft.Xrm.Tooling.CrmConnectControl"
              switchType="System.Diagnostics.SourceSwitch">
        <listeners>
          <add name="console" type="System.Diagnostics.DefaultTraceListener" />
          <remove name="Default"/>
          <add name ="fileListener"/>
        </listeners>
      </source>

      <source name="Microsoft.Xrm.Tooling.WebResourceUtility"
              switchName="Microsoft.Xrm.Tooling.WebResourceUtility"
              switchType="System.Diagnostics.SourceSwitch">
        <listeners>
          <add name="console" type="System.Diagnostics.DefaultTraceListener" />
          <remove name="Default"/>
          <add name ="fileListener"/>
        </listeners>
      </source>
      
    <!-- WCF DEBUG SOURCES -->
      <source name="System.IdentityModel" switchName="System.IdentityModel">
        <listeners>
          <add name="xml" />
        </listeners>
      </source>
      <!-- Log all messages in the 'Messages' tab of SvcTraceViewer. -->
      <source name="System.ServiceModel.MessageLogging" switchName="System.ServiceModel.MessageLogging" >
        <listeners>
          <add name="xml" />
        </listeners>
      </source>
      <!-- ActivityTracing and propogateActivity are used to flesh out the 'Activities' tab in
           SvcTraceViewer to aid debugging. -->
      <source name="System.ServiceModel" switchName="System.ServiceModel" propagateActivity="true">
        <listeners>
          <add name="xml" />
        </listeners>
      </source>
      <!-- END WCF DEBUG SOURCES -->
    </sources>
    <switches>
      <!-- 
            Possible values for switches: Off, Error, Warining, Info, Verbose
                Verbose:    includes Error, Warning, Info, Trace levels
                Info:       includes Error, Warning, Info levels
                Warning:    includes Error, Warning levels
                Error:      includes Error level
        -->
      <add name="Microsoft.Xrm.Tooling.Connector.CrmServiceClient" value="Verbose" />
      <add name="Microsoft.Xrm.Tooling.CrmConnectControl" value="Verbose"/>
      <add name="Microsoft.Xrm.Tooling.WebResourceUtility" value="Verbose" />
      <add name="System.IdentityModel" value="Verbose"/>
      <add name="System.ServiceModel.MessageLogging" value="Verbose"/>
      <add name="System.ServiceModel" value="Error, ActivityTracing"/>
      
    </switches>
    <sharedListeners>
      <add name="fileListener" type="System.Diagnostics.TextWriterTraceListener" initializeData="LoginControlTesterLog.txt"/>
      <!--<add name="eventLogListener" type="System.Diagnostics.EventLogTraceListener" initializeData="CRM UII"/>-->
      <add name="xml" type="System.Diagnostics.XmlWriterTraceListener" initializeData="CrmToolBox.svclog" />
    </sharedListeners>
  </system.diagnostics>
*/

    #endregion
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel.Channels;
using System.ServiceModel.Description;
using System.Text;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;

namespace MarketingListFactory
{
    public class CrmService
    {
        public static IOrganizationService CreateOrganizationService()
        {
            var _clientCreds = new ClientCredentials();
            _clientCreds.UserName.UserName = "adrian.w@superbuilder.co.uk";
            _clientCreds.UserName.Password = "Adrian789?";
            _clientCreds.Windows.ClientCredential.UserName = "adrian.w@superbuilder.co.uk";
            _clientCreds.Windows.ClientCredential.Password = "Adrian789?";
            var crmOrgSvc = "https://superbuilder.api.crm4.dynamics.com/XRMServices/2011/Organization.svc";
            var _serviceProxy = new OrganizationServiceProxy(new Uri(crmOrgSvc), null, _clientCreds, null);
            _serviceProxy.ServiceConfiguration.CurrentServiceEndpoint.Behaviors.Add(new ProxyTypesBehavior());
            if (_serviceProxy.ServiceConfiguration.CurrentServiceEndpoint.Binding is CustomBinding)
            {
                var cb = (CustomBinding)_serviceProxy.ServiceConfiguration.CurrentServiceEndpoint.Binding;
                cb.SendTimeout = new TimeSpan(0, 10, 0);
                cb.ReceiveTimeout = new TimeSpan(0, 10, 0);
                foreach (var be in cb.Elements)
                {
                    if (be is HttpTransportBindingElement)
                    {
                        ((HttpTransportBindingElement)be).UnsafeConnectionNtlmAuthentication = true;
                    }
                }
            }
            return _serviceProxy;
        }
    }
}

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
using System.Windows.Navigation;
using System.Windows.Shapes;
using MergingAccountsActivity;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace MarketingListFactory
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private IOrganizationService _crmService;

        public MainWindow()
        {
            InitializeComponent();
            _crmService = CrmService.CreateOrganizationService();
            var marketingLists = _crmService.RetrieveMultiple(ConditionsFactory.MarketingLists(new[] { "listname" })).Entities;
            lstMarketingLists.DisplayMemberPath = "Name";
            lstMarketingLists.SelectedValuePath = "ID";
            foreach (var iList in marketingLists)
            {
                lstMarketingLists.Items.Add(new { ID = iList.Id, Name = iList.GetAttributeValue<string>("listname") });
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if(string.IsNullOrWhiteSpace(txtFetchXml.Text))
                throw new Exception("Wrong fetch XML");
            var leads = _crmService.CrmQuery<Entity>(txtFetchXml.Text).Select(p => new LeadProxy(p)).ToList();
            lblTotalLeads.Content = "Leads total: " + leads.Count().ToString();
            var leadsByContact = leads.GroupBy(p => p.ContactId);
            lblTotalContacts.Content = "Contacts total: " + leadsByContact.Count().ToString();
            var addMembersRequest=new AddListMembersListRequest();
            addMembersRequest.ListId = new Guid(lstMarketingLists.SelectedValue.ToString());
            var listMembers=new List<Guid>();
            var counter = 0;
            foreach (var iLeadGroup in leadsByContact)
            {
                if(iLeadGroup.Key==null)
                    continue;
                var lastImportDate= iLeadGroup.Max(p => p.ImportDate);
                LeadProxy leadResult = null;
                if (lastImportDate != null)
                {
                    leadResult = iLeadGroup.FirstOrDefault(p => p.ImportDate == lastImportDate);
                }
                else
                {
                    leadResult = iLeadGroup.FirstOrDefault();
                }
                //new_mailmergeleadtext
                //new_mailmergelead
                var contUpdate = new Entity("contact");
                contUpdate.Id = iLeadGroup.Key.Value;
                contUpdate["new_mailmergeleadtext"] = leadResult.ProjectName;
                contUpdate["new_mailmergelead"] = leadResult.ToRef();
                _crmService.Update(contUpdate);
                listMembers.Add(iLeadGroup.Key.Value);
                counter++;
            }
            
            addMembersRequest.MemberIds = listMembers.ToArray();
            _crmService.Execute(addMembersRequest);
            MessageBox.Show("Marketing list formed");
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtFetchXml.Text))
                throw new Exception("Wrong fetch XML");
            var leads = _crmService.RetrieveMultiple(new FetchExpression(txtFetchXml.Text)).Entities.Select(p => new LeadProxy(p));
            lblTotalLeads.Content = "Leads total: " + leads.Count().ToString();
            var leadsByContact = leads.GroupBy(p => p.ContactId);
            lblTotalContacts.Content = "Contacts total: " + leadsByContact.Count().ToString();
            foreach (var iLeadGroup in leadsByContact)
            {
                var contact = _crmService.Retrieve("contact", iLeadGroup.Key.Value, new ColumnSet(new[] { "firstname", "lastname", "emailaddress1","suffix" }));
                foreach (var iLead in iLeadGroup)
                {
                    var updateLead = new Entity("lead") { Id = iLead.Id };
                    if (string.IsNullOrWhiteSpace(iLead.FirstName))
                        updateLead["firstname"] = contact.GetAttributeValue<string>("firstname");
                    if (string.IsNullOrWhiteSpace(iLead.LastName))
                        updateLead["lastname"] = contact.GetAttributeValue<string>("lastname");
                    if (string.IsNullOrWhiteSpace(iLead.Email))
                        updateLead["emailaddress1"] = contact.GetAttributeValue<string>("emailaddress1");
                    if (updateLead.Attributes.Any())
                    {
                        _crmService.Update(updateLead);
                    }
                }
            }
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            var query=new QueryExpression("contact");
            query.ColumnSet=new ColumnSet(new []{"accountrolecode"});
            query.Criteria.AddCondition("accountrolecode", ConditionOperator.Null);
            var contacts = _crmService.CrmQuery<Entity>(query);
            int counterTotal = contacts.Count();
            int counter=0;
            foreach (var iContact in contacts)
            {
                iContact["accountrolecode"] = new OptionSetValue(100000002);
                _crmService.Update(iContact);
                counter++;
            }
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtFetchXml.Text))
                throw new Exception("Wrong fetch XML");
            var leads = _crmService.RetrieveMultiple(new FetchExpression(txtFetchXml.Text)).Entities.Select(p => new LeadProxy(p));
            lblTotalLeads.Content = "Leads total: " + leads.Count().ToString();
            var leadsByAccount = leads.GroupBy(p => p.AccountId);
            lblTotalContacts.Content = "Accounts total: " + leadsByAccount.Count().ToString();
            foreach (var iLeadGroup in leadsByAccount)
            {
                var account = _crmService.Retrieve("account", iLeadGroup.Key.Value, new ColumnSet(new[] { "emailaddress1", "name"}));
                foreach (var iLead in iLeadGroup)
                {
                    var updateLead = new Entity("lead") {Id = iLead.Id};
                    if (string.IsNullOrWhiteSpace(iLead.Email))
                        updateLead["emailaddress1"] = account.GetAttributeValue<string>("emailaddress1");
                    if (string.IsNullOrWhiteSpace(iLead.AccountName))
                        updateLead["companyname"] = account.GetAttributeValue<string>("name");
                    if (updateLead.Attributes.Any())
                    {
                        _crmService.Update(updateLead);
                    }
                }
            }
            MessageBox.Show("Update completed");
        }

        private void btnMarketingListLeads_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtFetchXml.Text))
                throw new Exception("Wrong fetch XML");
            var leads = _crmService.CrmQuery<Entity>(txtFetchXml.Text).Select(p => new LeadProxy(p)).ToList();
            lblTotalLeads.Content = "Leads total: " + leads.Count().ToString();
            var leadsGrouped = leads.GroupBy(p => p.Email);
            lblTotalContacts.Content = "Contacts total: " + leadsGrouped.Count().ToString();
            var addMembersRequest = new AddListMembersListRequest();
            addMembersRequest.ListId = new Guid(lstMarketingLists.SelectedValue.ToString());
            var listMembers = new List<Guid>();
            var counter = 0;
            foreach (var iLeadGroup in leadsGrouped)
            {
                if (iLeadGroup.Key == null)
                    continue;
                var lastImportDate = iLeadGroup.Max(p => p.ImportDate);
                LeadProxy leadResult = null;
                if (lastImportDate != null)
                {
                    leadResult = iLeadGroup.FirstOrDefault(p => p.ImportDate == lastImportDate);
                }
                else
                {
                    leadResult = iLeadGroup.FirstOrDefault();
                }
                if (leadResult != null) listMembers.Add(leadResult.Id);
                counter++;
            }

            addMembersRequest.MemberIds = listMembers.ToArray();
            _crmService.Execute(addMembersRequest);
            MessageBox.Show("Marketing list formed");
        }
    }
}

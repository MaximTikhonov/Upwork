using MergeExcelDuplicates.Logic.Implementation;
using MergeExcelDuplicates.Logic.Interfaces;
using MergeExcelDuplicates.ProxyObjects;
using MergeExcelDuplicates.VisualInterface;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MergeExcelDuplicates.Logic;

namespace MergeExcelDuplicates
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            var dialogResult = openFileDialog.ShowDialog();
            if (dialogResult != System.Windows.Forms.DialogResult.OK)
                return;
            var eventNotificator = new EventNotificator();
            //eventNotificator.ProcessStarted += eventNotificator_ProcessStarted;
            //eventNotificator.InitialDataRead += eventNotificator_InitialDataRead;
            //eventNotificator.RecordProcessed += eventNotificator_RecordProcessed;
            //Open Excel file and read data            
            Action startProcessing = (/*OpenFileDialog openFileDialog*/) =>
            {
                eventNotificator.StartProcess();
                IDataConnector dataConnector = new DataConnector();
                
                using (var sourceFile = openFileDialog.OpenFile())
                {
                    //var sourceFile = openFileDialog.FileName;
                    var initialRecords = dataConnector.LoadInitialData(sourceFile);
                    var totalRecordsCount = initialRecords.Count();
                    eventNotificator.InitialDataCompleted(totalRecordsCount);
                    //split data to accounts, contacts and leads(projects)                
                    var accounts = new List<ICommonProxy>();
                    var projects = new List<ICommonProxy>();
                    var contacts = new List<ICommonProxy>();
                    int counter = 1;
                    foreach (var iInitialRecord in initialRecords)
                    {
                        var newAccount = iInitialRecord.ToAccount();
                        var newProject = iInitialRecord.ToProject();
                        var newContact = iInitialRecord.ToContact();
                        newContact.AccountId = newAccount.ImportId;
                        newProject.Accountid = newAccount.ImportId;
                        newProject.ContactId = newContact.ImportId;
                        var existingAccount = accounts.GetEqualAccount(newAccount);
                        if (existingAccount != null)
                        {
                            newContact.AccountId = existingAccount.ImportId;
                            newProject.Accountid = existingAccount.ImportId;
                        }
                        else { accounts.Add(newAccount); }
                        if (!newContact.IsEmpty())
                        {
                            var existingContact = contacts.GetEqualContact(newContact);
                            if (existingContact == null)
                            {
                                contacts.Add(newContact);
                            }
                            else
                            {
                                existingContact.Merge(newContact);
                                newProject.ContactId = existingContact.ImportId;
                            }
                        }
                        var existingProject = projects.GetEqualProject(newProject);
                        if (existingProject == null)
                            projects.Add(newProject);
                        eventNotificator.DataProcessed(counter++, totalRecordsCount);
                    }
                    //Save splitted data to 3 files
                    // split to 3 threads
                    dataConnector.CreateFile(accounts.ToArray(), openFileDialog.InitialDirectory, AccountProxy.ColumnsToArray(), "accounts.xlsx");
                    dataConnector.CreateFile(projects.ToArray(), openFileDialog.InitialDirectory, ProjectProxy.ColumnsToArray(), "projects.xlsx");
                    dataConnector.CreateFile(contacts.ToArray(), openFileDialog.InitialDirectory, ContactProxy.ColumnsToArray(), "contacts.xlsx");
                    MessageBox.Show("Import completed");
                }
            };
            //this.BeginInvoke(startProcessing);
            await Task.Run(startProcessing);
        }

        void eventNotificator_RecordProcessed(object sender, ImportEventArgs e)
        {
            //lblProgressInfo.Text = string.Format("Processing {0} in {1}", e.CurrentRecordNum, e.TotalRecords);
            //processingProgress.Value = e.CurrentRecordNum;
        }

        void eventNotificator_ProcessStarted(object sender, EventArgs e)
        {
            button1.Visible = false;
            //infoPanel.Visible = true;
        }

        void eventNotificator_InitialDataRead(object sender, ImportEventArgs e)
        {
            //lblSplittingInfo.Text += "\nRead data completed";
            //lblProgressInfo.Text = string.Format("Processing {0} in {1}",e.CurrentRecordNum, e.TotalRecords);
            //processingProgress.Maximum = e.TotalRecords;
            //processingProgress.Value = 0;
        }
    }
}

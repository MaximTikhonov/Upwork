using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MarketingListFactory.DataLogic.Comparison;
using MarketingListFactory.ProxyObjects.Crm;
using Microsoft.Xrm.Sdk;

namespace MarketingListFactory.ProxyObjects.Excel
{
    public class ProjectExcelProxy : IEntityExcelProxy//, ICompareEntity
    {
        private Dictionary<Columns, string> _rowData = new Dictionary<Columns, string>();
        public enum Columns
        {
            Name,
            ProjectSiteAddress1,
            ProjectSiteAddress2,
            ProjectSiteTown,
            ProjectSiteCounty,
            ProjectSitePostcode,
            FullData,
            ImportDate,
            GovermentReference,
            ShortWorkDescription1,
            ShortWorkDescription2,
            ImportId,
            Accountid,
            ContactId,
            ProjectNo,
            ProjectValue,
            ProjectPlanningRef,
            ProjectDevelopmentType,
            ProjectStage,
            ProjectStatus,
            ProjectSchemeDetails,
            ProjectContractType,
            ProjectProgrammeTiming,
            Email,
            Phone,
            LastName,
            FirstName,
            AccountCrmId,
            ContactCrmId
        }
        private Dictionary<Columns, string> _crmFieldsMapping = new Dictionary<Columns, string>()
        {
            {Columns.ImportId	, "new_excelimportid"},
            {Columns.ProjectSiteAddress1, "subject"},
            {Columns.ProjectSiteAddress2	, "new_topic2"},
            {Columns.ImportDate	, "new_projectimportdate"},
            {Columns.FullData	, "description"},
            {Columns.Email	, "emailaddress1"},
            {Columns.FirstName	, "firstname"},
            {Columns.Phone, "telephone1"},
            {Columns.LastName, "lastname"},
            {Columns.AccountCrmId	, "new_account"},
            {Columns.ContactCrmId	, "new_contact"},
            {Columns.Name	, "new_projectname"}
        };

        public string Name
        {
            get { return _rowData[Columns.Name]; }
            set
            {
                _rowData[Columns.Name] = value;
            }
        }

        public string ProjectSiteAddress1
        {
            get { return _rowData[Columns.ProjectSiteAddress1]; }
            set
            {
                _rowData[Columns.ProjectSiteAddress1] = value;
            }
        }

        public string ProjectSiteAddress2
        {
            get { return _rowData[Columns.ProjectSiteAddress2]; }
            set
            {
                _rowData[Columns.ProjectSiteAddress2] = value;
            }
        }

        public string ProjectSiteTown
        {
            get { return _rowData[Columns.ProjectSiteTown]; }
            set
            {
                _rowData[Columns.ProjectSiteTown] = value;
            }
        }

        public string ProjectSiteCounty
        {
            get { return _rowData[Columns.ProjectSiteCounty]; }
            set
            {
                _rowData[Columns.ProjectSiteCounty] = value;
            }
        }

        public string ProjectSitePostcode
        {
            get { return _rowData[Columns.ProjectSitePostcode]; }
            set
            {
                _rowData[Columns.ProjectSitePostcode] = value;
            }
        }

        public string FullData
        {
            get { return _rowData[Columns.FullData]; }
            set
            {
                _rowData[Columns.FullData] = value;
            }
        }

        public string ImportDate
        {
            get { return _rowData[Columns.ImportDate]; }
            set
            {
                _rowData[Columns.ImportDate] = value;
            }
        }

        public string GovermentReference
        {
            get { return _rowData[Columns.GovermentReference]; }
            set
            {
                _rowData[Columns.GovermentReference] = value;
            }
        }

        public string ShortWorkDescription1
        {
            get { return _rowData[Columns.ShortWorkDescription1]; }
            set
            {
                _rowData[Columns.ShortWorkDescription1] = value;
            }
        }

        public string ShortWorkDescription2
        {
            get { return _rowData[Columns.ShortWorkDescription2]; }
            set
            {
                _rowData[Columns.ShortWorkDescription2] = value;
            }
        }

        public Guid ImportId
        {
            get
            {
                if (_rowData[Columns.ImportId] == null || string.IsNullOrEmpty(_rowData[Columns.ImportId]))
                    return Guid.Empty;
                return new Guid(_rowData[Columns.ImportId]);
            }
            set
            {
                if (value != null)
                    _rowData[Columns.ImportId] = value.ToString();
                else
                    _rowData[Columns.ImportId] = "";
            }
        }

        public Guid Accountid
        {
            get
            {
                if (_rowData[Columns.Accountid] == null || string.IsNullOrEmpty(_rowData[Columns.Accountid]))
                    return Guid.Empty;
                return new Guid(_rowData[Columns.Accountid]);
            }
            set
            {
                if (value != null)
                    _rowData[Columns.Accountid] = value.ToString();
                else
                    _rowData[Columns.Accountid] = "";
            }
        }

        public Guid ContactId
        {
            get
            {
                if (_rowData[Columns.ContactId] == null || string.IsNullOrEmpty(_rowData[Columns.ContactId]))
                    return Guid.Empty;
                return new Guid(_rowData[Columns.ContactId]);
            }
            set
            {
                if (value != null)
                    _rowData[Columns.ContactId] = value.ToString();
                else
                    _rowData[Columns.ContactId] = "";
            }
        }

        public string ProjectNo
        {
            get { return _rowData[Columns.ProjectNo]; }
            set { _rowData[Columns.ProjectNo] = value; }
        }

        public string ProjectValue
        {
            get { return _rowData[Columns.ProjectValue]; }
            set { _rowData[Columns.ProjectValue] = value; }
        }
        public string ProjectPlanningRef
        {
            get { return _rowData[Columns.ProjectPlanningRef]; }
            set { _rowData[Columns.ProjectPlanningRef] = value; }
        }
        public string ProjectDevelopmentType
        {
            get { return _rowData[Columns.ProjectDevelopmentType]; }
            set { _rowData[Columns.ProjectDevelopmentType] = value; }
        }
        public string ProjectStage
        {
            get { return _rowData[Columns.ProjectStage]; }
            set { _rowData[Columns.ProjectStage] = value; }
        }
        public string ProjectStatus
        {
            get { return _rowData[Columns.ProjectStatus]; }
            set { _rowData[Columns.ProjectStatus] = value; }
        }
        public string ProjectSchemeDetails
        {
            get { return _rowData[Columns.ProjectSchemeDetails]; }
            set { _rowData[Columns.ProjectSchemeDetails] = value; }
        }
        public string ProjectProgrammeTiming
        {
            get { return _rowData[Columns.ProjectProgrammeTiming]; }
            set { _rowData[Columns.ProjectProgrammeTiming] = value; }
        }
        public string ProjectContractType
        {
            get { return _rowData[Columns.ProjectContractType]; }
            set { _rowData[Columns.ProjectContractType] = value; }
        }

        public string Email
        {
            get { return _rowData[Columns.Email]; }
            set { _rowData[Columns.Email] = value; }
        }

        public string Phone
        {
            get { return _rowData[Columns.Phone]; }
            set { _rowData[Columns.Phone] = value; }
        }

        public string LastName
        {
            get { return _rowData[Columns.LastName]; }
            set { _rowData[Columns.LastName] = value; }
        }

        public string FirstName
        {
            get { return _rowData[Columns.FirstName]; }
            set { _rowData[Columns.FirstName] = value; }
        }

        public Guid AccountCrmId
        {
            get
            {
                if (!_rowData.ContainsKey(Columns.AccountCrmId) || string.IsNullOrWhiteSpace(_rowData[Columns.AccountCrmId]))
                    return Guid.Empty;
                return new Guid(_rowData[Columns.AccountCrmId]);
            }
            set
            {
                _rowData[Columns.AccountCrmId] = value.ToString();
            }
        }

        public Guid ContactCrmId
        {
            get
            {
                if (!_rowData.ContainsKey(Columns.ContactCrmId) || string.IsNullOrWhiteSpace(_rowData[Columns.ContactCrmId]))
                    return Guid.Empty;
                return new Guid(_rowData[Columns.ContactCrmId]);
            }
            set
            {
                _rowData[Columns.ContactCrmId] = value.ToString();
            }
        }

        public IEnumerable<KeyValuePair<string, string>> GetData()
        {
            return _rowData.Select(p => new KeyValuePair<string, string>(p.Key.ToString(), p.Value));
        }

        public bool EqualsToAccount(AccountExcelProxy targetAccountExcel)
        {
            throw new NotImplementedException();
        }

        public bool EqualsToContact(ContactExcelProxy targetContactExcel)
        {
            throw new NotImplementedException();
        }

        public bool EqualsToProject(ProjectExcelProxy targetProjectExcel)
        {
            return (targetProjectExcel.FullData.Replace(" ", "").ToLower() == this.FullData.Replace(" ", "").ToLower()
                    &&this.Accountid== targetProjectExcel.Accountid
                    &&this.ContactId==targetProjectExcel.ContactId)
                || (this.Accountid == targetProjectExcel.Accountid&&FuzzyStringComparer.IsStringsFuzzyEquals(targetProjectExcel.ProjectSiteAddress1,this.ProjectSiteAddress1,80));
        }

        internal static string[] ColumnsToArray()
        {
            var result = new List<string>();
            foreach (var item in Enum.GetValues(typeof(Columns)))
            {
                result.Add(item.ToString());
            }
            return result.ToArray();
        }

        internal void Merge(ProjectExcelProxy newProjectExcel)
        {
            //if (string.IsNullOrEmpty(this.ProjectSiteTown) && !string.IsNullOrEmpty(newProject.ProjectSiteTown))
            //    this.ProjectSiteTown = newProject.ProjectSiteTown;
            //if (string.IsNullOrEmpty(this.ProjectSitePostcode) && !string.IsNullOrEmpty(newProject.ProjectSitePostcode))
            //    this.ProjectSitePostcode = newProject.ProjectSitePostcode;
            //if (string.IsNullOrEmpty(this.ProjectSiteCounty) && !string.IsNullOrEmpty(newProject.ProjectSiteCounty))
            //    this.ProjectSiteCounty = newProject.ProjectSiteCounty;
            //if (string.IsNullOrEmpty(this.GovermentReference) && !string.IsNullOrEmpty(newProject.GovermentReference))
            //    this.GovermentReference = newProject.GovermentReference;
            //if (string.IsNullOrEmpty(this.ShortWorkDescription1) && !string.IsNullOrEmpty(newProject.ShortWorkDescription1))
            //    this.ShortWorkDescription1 = newProject.ShortWorkDescription1;
            //if (string.IsNullOrEmpty(this.ShortWorkDescription2) && !string.IsNullOrEmpty(newProject.ShortWorkDescription2))
            //    this.ShortWorkDescription2 = newProject.ShortWorkDescription2;
            //if (string.IsNullOrEmpty(this.ProjectValue) && !string.IsNullOrEmpty(newProject.ProjectValue))
            //    this.ProjectValue = newProject.ProjectValue;
            //if (string.IsNullOrEmpty(this.ShortWorkDescription1) && !string.IsNullOrEmpty(newProject.ShortWorkDescription1))
            //    this.ShortWorkDescription1 = newProject.ShortWorkDescription1;
            var currentData = _rowData.ToList();
            foreach (var iData in currentData)
            {
                if (string.IsNullOrWhiteSpace(iData.Value) && !string.IsNullOrEmpty(newProjectExcel._rowData[iData.Key]))
                    _rowData[iData.Key] = newProjectExcel._rowData[iData.Key];
            }
        }

        public LeadCrmProxy ToCrmEntity()
        {
            var targetLead = new Entity("lead");
            foreach (var iRowData in _rowData)
            {
                if(string.IsNullOrEmpty(iRowData.Value))
                    continue;
                if(!_crmFieldsMapping.ContainsKey(iRowData.Key))
                    continue;
                var crmFieldName = _crmFieldsMapping[iRowData.Key];
                if (iRowData.Key == Columns.ImportDate && !string.IsNullOrWhiteSpace(iRowData.Value))
                    targetLead[crmFieldName] = DateTime.Parse(iRowData.Value);
                else if (iRowData.Key == Columns.AccountCrmId && !string.IsNullOrWhiteSpace(iRowData.Value))
                {
                    targetLead[crmFieldName] = new EntityReference("account",new Guid(iRowData.Value));
                }
                else if (iRowData.Key == Columns.ContactCrmId && !string.IsNullOrWhiteSpace(iRowData.Value))
                {
                    targetLead[crmFieldName] = new EntityReference("contact", new Guid(iRowData.Value));
                }
                else
                {
                    targetLead[crmFieldName] = iRowData.Value;
                }
            }
            return new LeadCrmProxy(targetLead);
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Navigation;
using MarketingListFactory.DataLogic.Comparison;
using MarketingListFactory.ProxyObjects.Crm;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;

namespace MarketingListFactory.ProxyObjects.Excel
{
    public class ContactExcelProxy : IEntityExcelProxy//, ICompareEntity
    {
        private Dictionary<Columns, string> _rowData = new Dictionary<Columns, string>();
        public enum Columns
        {
            FirstName,
            SecondName,
            ArchitectWorkPhone,
            ArchitectEmail,
            ArchitectPhone,
            ArchitectRole,
            ArchitectWebsite,
            ArchitectFax,
            ArchitectSex,
            ImportId,
            AccountId,
            ShortFirstName,
            ShortSecondName,
            ImportDate,
            ParentCustomerId

        }
        private string _shortFirstName;

        public string ShortFirstName
        {
            get { return _shortFirstName; }
            set
            {
                if (value != null)
                    _shortFirstName = value.Replace(" ", "").ToLower().Replace(".", "").Replace(",", "");
                else
                    _shortFirstName = value;
                _rowData[Columns.ShortFirstName] = value;
            }
        }

        public string FirstName
        {
            get { return _rowData[Columns.FirstName]; }
            set
            {
                if (value == "0")
                    value = "";
                ShortFirstName = value;

                _rowData[Columns.FirstName] = value;
            }
        }
        private string _shortSecondName;
        private Dictionary<Columns, string> _crmFieldsMapping = new Dictionary<Columns, string>()
        {
            {Columns.ImportId	, "new_excelimportid"},
            {Columns.ArchitectPhone, "telephone1"},
            {Columns.ShortFirstName	, "new_shortfirstname"},
            {Columns.ShortSecondName	, "new_shortsecondname"},
            {Columns.ArchitectWorkPhone	, "business2"},
            {Columns.ArchitectFax	, "fax"},
            {Columns.ArchitectWebsite	, "websiteurl"},
            {Columns.FirstName	, "firstname"},
            {Columns.ArchitectEmail, "emailaddress1"},
            {Columns.ArchitectSex, "suffix"},
            {Columns.SecondName, "lastname"},
            {Columns.ParentCustomerId	, "parentcustomerid"},
            {Columns.ArchitectRole, "accountrolecode"},
            //{Columns.AccountId, "parentcustomerid"}
        };

        public string ShortSecondName
        {
            get { return _shortSecondName; }
            set
            {
                if (value != null)
                    _shortSecondName = value.Replace(" ", "").ToLower().Replace(".", "").Replace(",", "");
                else
                    _shortSecondName = value;
                _rowData[Columns.ShortSecondName] = value;
            }
        }

        public string SecondName
        {
            get { return _rowData[Columns.SecondName]; }
            set
            {
                if (value == "0")
                    value = "";
                ShortSecondName = value;
                _rowData[Columns.SecondName] = value;
            }
        }

        public string ArchitectWorkPhone
        {
            get { return _rowData[Columns.ArchitectWorkPhone]; }
            set
            {
                if (value == "0")
                    value = "";
                _rowData[Columns.ArchitectWorkPhone] = value;
            }
        }

        public string ArchitectEmail
        {
            get { return _rowData[Columns.ArchitectEmail]; }
            set
            {
                if (value == "0")
                    value = "";
                _rowData[Columns.ArchitectEmail] = value;
            }
        }

        public string ArchitectPhone
        {
            get { return _rowData[Columns.ArchitectPhone]; }
            set
            {
                if (value == "0")
                    value = "";
                _rowData[Columns.ArchitectPhone] = value;
            }
        }

        public string ArchitectRole
        {
            get { return _rowData[Columns.ArchitectRole]; }
            set
            {
                if (value == "0")
                    value = "";
                _rowData[Columns.ArchitectRole] = value;
            }
        }

        public string ArchitectWebsite
        {
            get { return _rowData[Columns.ArchitectWebsite]; }
            set
            {
                if (value == "0")
                    value = "";
                _rowData[Columns.ArchitectWebsite] = value;
            }
        }

        public string ArchitectFax
        {
            get { return _rowData[Columns.ArchitectFax]; }
            set
            {
                if (value == "0")
                    value = "";
                _rowData[Columns.ArchitectFax] = value;
            }
        }

        public string ArchitectSex
        {
            get { return _rowData[Columns.ArchitectSex]; }
            set
            {
                if (value == "0")
                    value = "";
                _rowData[Columns.ArchitectSex] = value;
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
        public string ImportDate
        {
            get { return _rowData[Columns.ImportDate]; }
            set
            {
                var resultValue = value;
                if (value.Length == 8)
                {
                    var year = value.Substring(0, 4);
                    var month = value.Substring(4, 2);
                    var day = value.Substring(6, 2);
                    var resultDate = new DateTime(int.Parse(year), int.Parse(month), int.Parse(day));
                    resultValue = resultDate.ToShortDateString();
                }
                _rowData[Columns.ImportDate] = resultValue;
            }
        }
        public Guid AccountId
        {
            get
            {
                if (_rowData[Columns.AccountId] == null || string.IsNullOrEmpty(_rowData[Columns.AccountId]))
                    return Guid.Empty;
                return new Guid(_rowData[Columns.AccountId]);
            }
            set
            {
                if (value != null)
                    _rowData[Columns.AccountId] = value.ToString();
                else
                    _rowData[Columns.AccountId] = "";
            }
        }

        public Guid AccountCrmId
        {
            get
            {
                if (!_rowData.ContainsKey(Columns.ParentCustomerId) || string.IsNullOrWhiteSpace(_rowData[Columns.ParentCustomerId]))
                    return Guid.Empty;
                return new Guid(_rowData[Columns.ParentCustomerId]);
            }
            set
            {
                _rowData[Columns.ParentCustomerId] = value.ToString();
            }
        }

        public IEnumerable<KeyValuePair<string, string>> GetData()
        {
            return _rowData.Select(p => new KeyValuePair<string, string>(p.Key.ToString(), p.Value));
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


        public bool EqualsToAccount(AccountExcelProxy targetAccountExcel)
        {
            throw new NotImplementedException();
        }
        public bool IsEmpty()
        {
            return ((string.IsNullOrWhiteSpace(this.ShortFirstName) || this.ShortFirstName == "0")
                    && (string.IsNullOrWhiteSpace(this.ShortSecondName) || this.ShortSecondName == "0")
                //&& string.IsNullOrWhiteSpace(this.ArchitectPhone)
                //&& string.IsNullOrWhiteSpace(this.ArchitectWorkPhone)
                    && string.IsNullOrWhiteSpace(this.ArchitectEmail));
        }
        public bool EqualsToContact(ContactExcelProxy targetContactExcel)
        {
            if (targetContactExcel.AccountId != this.AccountId)
                return false;
            return (targetContactExcel.ShortFirstName == this.ShortFirstName && targetContactExcel.ShortSecondName == this.ShortSecondName)
                || (targetContactExcel.ShortSecondName == this.ShortSecondName && targetContactExcel.FirstNameCompare(this.ShortFirstName))
                || targetContactExcel.IsEmpty()
                || (targetContactExcel.ShortSecondName == this.ShortSecondName && AccountExcelProxy.PhoneCompare(this.ArchitectPhone, targetContactExcel.ArchitectPhone))
                || (FuzzyStringComparer.IsStringsFuzzyEquals(targetContactExcel.ShortSecondName, this.ShortSecondName, 70));

        }

        private bool FirstNameCompare(string targetFirstName)
        {
            if (string.IsNullOrWhiteSpace(this.ShortFirstName) && string.IsNullOrWhiteSpace(targetFirstName))
                return true;
            if ((!string.IsNullOrWhiteSpace(this.ShortFirstName) && string.IsNullOrWhiteSpace(targetFirstName))
                || (string.IsNullOrWhiteSpace(this.ShortFirstName) && !string.IsNullOrWhiteSpace(targetFirstName)))
                return false;
            return this.ShortFirstName[0] == targetFirstName[0] && (this.ShortFirstName.Length == 1 || targetFirstName.Length == 1);
        }

        public bool EqualsToProject(ProjectExcelProxy targetProjectExcel)
        {
            throw new NotImplementedException();
        }

        internal void Merge(ContactExcelProxy newContactExcel)
        {
            if (this.FirstName.Length == 1)
                this.FirstName = newContactExcel.FirstName;
            if (!this.ArchitectPhone.Contains(newContactExcel.ArchitectPhone) && string.IsNullOrWhiteSpace(newContactExcel.ArchitectPhone))
                this.ArchitectPhone += string.IsNullOrWhiteSpace(this.ArchitectPhone) ? ";" : "" + newContactExcel.ArchitectPhone;
        }

        public ContactCrmProxy ToCrmEntity()
        {
            var targetAccount = new Entity("contact");
            foreach (var iRowData in _rowData)
            {
                if(string.IsNullOrEmpty(iRowData.Value))
                    continue;
                if(!_crmFieldsMapping.ContainsKey(iRowData.Key))
                    continue;
                var crmFieldName = _crmFieldsMapping[iRowData.Key];
                if (iRowData.Key == ContactExcelProxy.Columns.ImportDate && !string.IsNullOrWhiteSpace(iRowData.Value))
                    targetAccount[crmFieldName] = DateTime.Parse(iRowData.Value);
                else if (iRowData.Key == ContactExcelProxy.Columns.ParentCustomerId)
                {
                    targetAccount[crmFieldName] = new EntityReference("account",new Guid(iRowData.Value));
                }
                else if (iRowData.Key==Columns.ArchitectRole)
                {
                    targetAccount[crmFieldName] = new OptionSetValue(100000002);
                }
                else
                {
                    targetAccount[crmFieldName] = iRowData.Value;
                }
            }
            return new ContactCrmProxy(targetAccount);
        }
    }
}

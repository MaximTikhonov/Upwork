using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using MarketingListFactory.DataLogic.Comparison;
using MarketingListFactory.ProxyObjects.Common;
using MarketingListFactory.ProxyObjects.Crm;
using Microsoft.Xrm.Sdk;

namespace MarketingListFactory.ProxyObjects.Excel
{
    public class AccountExcelProxy : IEntityExcelProxy, IAccountProxy//ICompareEntity, 
    {
        private Dictionary<Columns, string> _rowData = new Dictionary<Columns, string>();
        public enum Columns
        {
            Name,
            Address1,
            Address2,
            Address3,
            TownName,
            CountyName,
            PostCode,
            ImportDate,
            ImportId,
            WorkPhone,
            MobilePhone,
            Email,
            ExtraPhones,
            ShortName,
            WebSite
        }
        private Dictionary<Columns, string> _crmFieldsMapping = new Dictionary<Columns, string>()
        {
            {Columns.Name, "name"},
            {Columns.Address1 , "address1_line1"},
            {Columns.Address2, "address1_line2"},
            {Columns.Address3 , "address1_line3"},
            {Columns.TownName, "address1_city"},
            {Columns.CountyName , "address1_county"},
            {Columns.PostCode, "address1_postalcode"},
            {Columns.ImportDate , "new_excelimportdate"},
                        {Columns.ImportId , "new_excelimportid"},
            {Columns.WorkPhone, "telephone1"},
            {Columns.MobilePhone , "address1_telephone1"},
            {Columns.Email, "emailaddress1"},
            {Columns.ExtraPhones, "telephone3"},
            {Columns.ShortName, "new_shortname"},
            {Columns.WebSite , "websiteurl"},
        };
        private static string[] _technicalStrings = { " ltd", " limited", " lld", "-", " llp" };

        private string _shortPostCode;

        public string ShortPostCode
        {
            get { return _shortPostCode; }
        }
        private string _shortName;
        private string _shortWorkPhone;
        private string _shortMobilePhone;

        public AccountExcelProxy()
        {
            _rowData[Columns.ExtraPhones] = "";
        }
        public string ShortName
        {
            get { return _shortName; }
            set
            {
                if (value != null)
                {
                    _shortName = value;
                    foreach (var iTechStrings in _technicalStrings)
                    {
                        _shortName = _shortName.ToLower().Replace(iTechStrings, "");
                    }
                    _shortName = _shortName.Replace(" ", "").Replace(".", "").Replace("limited", "ltd").Replace(",", "")
                        .Replace("and", "&");
                }
                else
                    _shortName = value;

                _rowData[Columns.ShortName] = _shortName;
            }
        }

        public string Name
        {
            get { return _rowData[Columns.Name]; }
            set
            {
                ShortName = value;
                _rowData[Columns.Name] = value;
            }
        }

        public string Address1
        {
            get { return _rowData[Columns.Address1]; }
            set
            {
                _rowData[Columns.Address1] = value;
            }
        }

        public string Address2
        {
            get { return _rowData[Columns.Address2]; }
            set
            {
                _rowData[Columns.Address2] = value;
            }
        }

        public string Address3
        {
            get { return _rowData[Columns.Address3]; }
            set { _rowData[Columns.Address3] = value; }
        }

        public string TownName
        {
            get { return _rowData[Columns.TownName]; }
            set { _rowData[Columns.TownName] = value; }
        }

        public string CountyName
        {
            get { return _rowData[Columns.CountyName]; }
            set { _rowData[Columns.CountyName] = value; }
        }

        public string PostCode
        {
            get { return _rowData[Columns.PostCode]; }
            set
            {
                if (value != null)
                    _shortPostCode = value.Replace(" ", "").ToLower().Replace(".", "");
                else
                    _shortPostCode = value;
                _rowData[Columns.PostCode] = value;
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
        public IEnumerable<KeyValuePair<string, string>> GetData()
        {

            return _rowData.Select(p => new KeyValuePair<string, string>(p.Key.ToString(), p.Value));
        }

        public Guid ImportId
        {
            get
            {
                if (!_rowData.ContainsKey(Columns.ImportId) || _rowData[Columns.ImportId] == null)
                    return Guid.Empty;
                return new Guid(_rowData[Columns.ImportId]);
            }
            set { _rowData[Columns.ImportId] = value.ToString(); }
        }

        public string WorkPhone
        {
            get
            {

                return _rowData[Columns.WorkPhone];
            }
            set
            {
                value = PhoneRectifier(value);
                //if (value != null)
                //{

                //    _shortWorkPhone = value.ToLower().Replace(" ", "").Replace("+", "");
                //}
                _rowData[Columns.WorkPhone] = value;
            }
        }

        private string PhoneRectifier(string inValue)
        {
            var result = "";
            var exisitingPhone = phoneCompare.Match(inValue).Value;
            result = exisitingPhone;
            if (inValue.StartsWith("0"))
            {
                result = "+44" + exisitingPhone.Remove(1, exisitingPhone.Length - 1);
            }
            if (inValue.StartsWith("+"))
                result = "+" + result;
            if (result.Length != 13)
                result = "";
            return result;
        }

        public string MobilePhone
        {
            get { return _rowData[Columns.MobilePhone]; }
            set
            {
                value = PhoneRectifier(value);
                //if (value != null)
                //{

                //    _shortMobilePhone = value.ToLower().Replace(" ", "").Replace("+", "");
                //}
                _rowData[Columns.MobilePhone] = value;
            }
        }

        //public string ShortWorkPhone
        //{
        //    get { return _shortWorkPhone; }
        //}

        //public string ShortMobilePhone
        //{
        //    get { return _shortMobilePhone; }
        //}

        public string Email
        {
            get { return _rowData[Columns.Email]; }
            set
            {
                value = value.ToLower().Replace(" ", "");
                _rowData[Columns.Email] = value;
            }
        }

        public static string[] ColumnsToArray()
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

            return (targetAccountExcel.ShortName == this.ShortName)
                //|| (targetAccount.ShortPostCode == this.ShortPostCode && targetAccount.ShortPostCode.Length > 1
                //    && targetAccount.CountyName == this.CountyName && targetAccount.CountyName.Length > 1
                //    && targetAccount.TownName == this.TownName && targetAccount.TownName.Length > 1)
                    || (!string.IsNullOrWhiteSpace(targetAccountExcel.Email) && targetAccountExcel.Email == this.Email)
                    || (!string.IsNullOrWhiteSpace(targetAccountExcel.WebSite) && targetAccountExcel.WebSite == this.WebSite)
                    || (!string.IsNullOrWhiteSpace(targetAccountExcel.WorkPhone)
                        && (PhoneCompare(this.WorkPhone, targetAccountExcel.WorkPhone)
                            || PhoneCompare(this.ExtraPhones, targetAccountExcel.WorkPhone)))
                    || (!string.IsNullOrWhiteSpace(targetAccountExcel.MobilePhone)
                        && (PhoneCompare(this.MobilePhone, targetAccountExcel.MobilePhone)
                            || PhoneCompare(this.ExtraPhones, targetAccountExcel.MobilePhone)))
                    || (targetAccountExcel.ShortPostCode == this.ShortPostCode && targetAccountExcel.ShortPostCode.Length > 1
                        && FuzzyStringComparer.IsStringsFuzzyEquals(targetAccountExcel.ShortName, this.ShortName, 70));
        }

        public bool EqualsToContact(ContactExcelProxy targetContactExcel)
        {
            throw new NotImplementedException();
        }

        public bool EqualsToProject(ProjectExcelProxy targetProjectExcel)
        {
            throw new NotImplementedException();
        }

        public AccountCrmProxy ToCrmEntity()
        {
            var targetAccount = new Entity("account");
            foreach (var iRowData in _rowData)
            {
                if(string.IsNullOrEmpty(iRowData.Value))
                    continue;
                var crmFieldName = _crmFieldsMapping[iRowData.Key];
                if (iRowData.Key == Columns.ImportDate && !string.IsNullOrWhiteSpace(iRowData.Value))
                    targetAccount[crmFieldName] = DateTime.Parse(iRowData.Value);
                else
                {
                    targetAccount[crmFieldName] = iRowData.Value;
                }
            }
            return new AccountCrmProxy(targetAccount);
        }

        public void Merge(AccountExcelProxy newAccountExcel)
        {
            if (!string.IsNullOrWhiteSpace(newAccountExcel.WorkPhone))
            {
                if (string.IsNullOrWhiteSpace(this.WorkPhone))
                {
                    this.WorkPhone = newAccountExcel.WorkPhone;
                }
                else if (!this.ExtraPhones.Contains(newAccountExcel.WorkPhone))
                    this.ExtraPhones += string.IsNullOrWhiteSpace(this.ExtraPhones) ? "" : ";" + newAccountExcel.WorkPhone;
            }
            if (!string.IsNullOrWhiteSpace(newAccountExcel.MobilePhone))
            {
                if (string.IsNullOrWhiteSpace(this.MobilePhone))
                {
                    this.MobilePhone = newAccountExcel.MobilePhone;
                }
                else if (!this.ExtraPhones.Contains(newAccountExcel.MobilePhone))
                    this.ExtraPhones += string.IsNullOrWhiteSpace(this.ExtraPhones) ? "" : ";" + newAccountExcel.MobilePhone;
            }
            if (!string.IsNullOrWhiteSpace(newAccountExcel.Address1) && string.IsNullOrWhiteSpace(this.Address1))
            {
                this.Address1 = newAccountExcel.Address1;
            }
            if (!string.IsNullOrWhiteSpace(newAccountExcel.Address2) && string.IsNullOrWhiteSpace(this.Address2))
            {
                this.Address2 = newAccountExcel.Address2;
            }
            if (!string.IsNullOrWhiteSpace(newAccountExcel.Address3) && string.IsNullOrWhiteSpace(this.Address3))
            {
                this.Address3 = newAccountExcel.Address3;
            }
            if (!string.IsNullOrWhiteSpace(newAccountExcel.Email) && string.IsNullOrWhiteSpace(this.Email))
            {
                this.Email = newAccountExcel.Email;
            }

        }

        public string ExtraPhones
        {
            get { return _rowData[Columns.ExtraPhones]; }
            set { _rowData[Columns.ExtraPhones] = value; }
        }

        public string WebSite
        {
            get { return _rowData[Columns.WebSite]; }
            set
            {
                if (value != null)
                    value = value.ToLower().Replace(" ", "");
                _rowData[Columns.WebSite] = value;
            }
        }

        public static bool PhoneCompare(string existingValue, string newValue)
        {
            var exisitingPhone = phoneCompare.Match(existingValue).Value;
            var newPhone = phoneCompare.Match(newValue).Value;
            if (newPhone.Length < 6)
                return false;
            return exisitingPhone.Contains(newPhone);
        }
        private static Regex phoneCompare = new Regex(@"\d+");

        public AccountExcelProxy(Entity entity)
        {
            Address1 = entity.GetAttributeValue<string>("address1_line1");
            Address2 = entity.GetAttributeValue<string>("address1_line2");
            Address3 = entity.GetAttributeValue<string>("address1_line3");
            CountyName = entity.GetAttributeValue<string>("address1_county");
            Email = entity.GetAttributeValue<string>("emailaddress1");
            ExtraPhones = entity.GetAttributeValue<string>("telephone3");
            ImportDate = entity.GetAttributeValue<string>("new_excelimportdate");
            ImportId = new Guid(entity.GetAttributeValue<string>("new_excelimportid"));
            MobilePhone = entity.GetAttributeValue<string>("address1_telephone1");
            Name = entity.GetAttributeValue<string>("name");
            PostCode = entity.GetAttributeValue<string>("address1_postalcode");
            TownName = entity.GetAttributeValue<string>("address1_city");
            WebSite = entity.GetAttributeValue<string>("websiteurl");
            WorkPhone = entity.GetAttributeValue<string>("telephone1");
        }
    }
}

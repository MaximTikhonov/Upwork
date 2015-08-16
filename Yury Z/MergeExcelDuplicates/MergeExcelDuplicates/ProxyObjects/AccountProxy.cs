using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using MergeExcelDuplicates.Logic.Comparesion;

namespace MergeExcelDuplicates.ProxyObjects
{
    public class AccountProxy : ICommonProxy
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
        private static string[] _technicalStrings = { " ltd", " limited", " lld", "-", " llp" };

        private string _shortPostCode;

        public string ShortPostCode
        {
            get { return _shortPostCode; }
        }
        private string _shortName;
        private string _shortWorkPhone;
        private string _shortMobilePhone;

        public AccountProxy()
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


        public bool EqualsToAccount(AccountProxy targetAccount)
        {

            return (targetAccount.ShortName == this.ShortName)
                //|| (targetAccount.ShortPostCode == this.ShortPostCode && targetAccount.ShortPostCode.Length > 1
                //    && targetAccount.CountyName == this.CountyName && targetAccount.CountyName.Length > 1
                //    && targetAccount.TownName == this.TownName && targetAccount.TownName.Length > 1)
                    || (!string.IsNullOrWhiteSpace(targetAccount.Email) && targetAccount.Email == this.Email)
                    || (!string.IsNullOrWhiteSpace(targetAccount.WebSite) && targetAccount.WebSite == this.WebSite)
                    || (!string.IsNullOrWhiteSpace(targetAccount.WorkPhone)
                        && (PhoneCompare(this.WorkPhone, targetAccount.WorkPhone)
                            || PhoneCompare(this.ExtraPhones, targetAccount.WorkPhone)))
                    || (!string.IsNullOrWhiteSpace(targetAccount.MobilePhone)
                        && (PhoneCompare(this.MobilePhone, targetAccount.MobilePhone)
                            || PhoneCompare(this.ExtraPhones, targetAccount.MobilePhone)))
                    || (targetAccount.ShortPostCode == this.ShortPostCode && targetAccount.ShortPostCode.Length > 1
                        && FuzzyStringComparer.IsStringsFuzzyEquals(targetAccount.ShortName,this.ShortName,70));
        }

        public bool EqualsToContact(ContactProxy targetContact)
        {
            throw new NotImplementedException();
        }

        public bool EqualsToProject(ProjectProxy targetProject)
        {
            throw new NotImplementedException();
        }

        public void Merge(AccountProxy newAccount)
        {
            if (!string.IsNullOrWhiteSpace(newAccount.WorkPhone))
            {
                if (string.IsNullOrWhiteSpace(this.WorkPhone))
                {
                    this.WorkPhone = newAccount.WorkPhone;
                }
                else if (!this.ExtraPhones.Contains(newAccount.WorkPhone))
                    this.ExtraPhones += string.IsNullOrWhiteSpace(this.ExtraPhones) ? "" : ";" + newAccount.WorkPhone;
            }
            if (!string.IsNullOrWhiteSpace(newAccount.MobilePhone))
            {
                if (string.IsNullOrWhiteSpace(this.MobilePhone))
                {
                    this.MobilePhone = newAccount.MobilePhone;
                }
                else if (!this.ExtraPhones.Contains(newAccount.MobilePhone))
                    this.ExtraPhones += string.IsNullOrWhiteSpace(this.ExtraPhones) ? "" : ";" + newAccount.MobilePhone;
            }
            if (!string.IsNullOrWhiteSpace(newAccount.Address1) && string.IsNullOrWhiteSpace(this.Address1))
            {
                this.Address1 = newAccount.Address1;
            }
            if (!string.IsNullOrWhiteSpace(newAccount.Address2) && string.IsNullOrWhiteSpace(this.Address2))
            {
                this.Address2 = newAccount.Address2;
            }
            if (!string.IsNullOrWhiteSpace(newAccount.Address3) && string.IsNullOrWhiteSpace(this.Address3))
            {
                this.Address3 = newAccount.Address3;
            }
            if (!string.IsNullOrWhiteSpace(newAccount.Email) && string.IsNullOrWhiteSpace(this.Email))
            {
                this.Email = newAccount.Email;
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
    }
}

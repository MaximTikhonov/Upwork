using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MergeExcelDuplicates.ProxyObjects
{
    public class ContactProxy : ICommonProxy
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
            AccountId
        }
        private string _shortFirstName;

        public string ShortFirstName
        {
            get { return _shortFirstName; }
        }
        public string FirstName
        {
            get { return _rowData[Columns.FirstName]; }
            set
            {
                if (value == "0")
                    value = "";
                if (value != null)
                    _shortFirstName = value.Replace(" ", "").ToLower().Replace(".", "").Replace(",", "");
                else
                    _shortFirstName = value;
                _rowData[Columns.FirstName] = value;
            }
        }
        private string _shortSecondName;

        public string ShortSecondName
        {
            get { return _shortSecondName; }
        }
        public string SecondName
        {
            get { return _rowData[Columns.SecondName]; }
            set
            {
                if (value == "0")
                    value = "";
                if (value != null)
                    _shortSecondName = value.Replace(" ", "").ToLower().Replace(".", "").Replace(",", "");
                else
                    _shortSecondName = value;
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


        public bool EqualsToAccount(AccountProxy targetAccount)
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
        public bool EqualsToContact(ContactProxy targetContact)
        {
            if (targetContact.AccountId != this.AccountId)
                return false;
            return (targetContact.ShortFirstName == this.ShortFirstName && targetContact.ShortSecondName == this.ShortSecondName)
                || (targetContact.ShortSecondName == this.ShortSecondName && targetContact.FirstNameCompare(this.ShortFirstName))
                || targetContact.IsEmpty()
                || (targetContact.ShortSecondName == this.ShortSecondName && AccountProxy.PhoneCompare(this.ArchitectPhone, targetContact.ArchitectPhone));

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

        public bool EqualsToProject(ProjectProxy targetProject)
        {
            throw new NotImplementedException();
        }

        internal void Merge(ContactProxy newContact)
        {
            if (this.FirstName.Length == 1)
                this.FirstName = newContact.FirstName;
            if (!this.ArchitectPhone.Contains(newContact.ArchitectPhone) && string.IsNullOrWhiteSpace(newContact.ArchitectPhone))
                this.ArchitectPhone += string.IsNullOrWhiteSpace(this.ArchitectPhone) ? ";" : "" + newContact.ArchitectPhone;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
            ImportId
        }

        private string _shortPostCode;

        public string ShortPostCode
        {
            get { return _shortPostCode; }
        }
        private string _shortName;

        public string ShortName
        {
            get { return _shortName; }
        }
        public string Name
        {
            get { return _rowData[Columns.Name]; }
            set
            {
                if (value != null)
                {
                    _shortName = value.Replace(" ", "").ToLower().Replace(".", "").Replace("limited", "ltd").Replace(",", "")
                        .Replace("and", "&");

                }
                else
                    _shortName = value;

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
            set { _rowData[Columns.ImportDate] = value; }
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
                    || (targetAccount.ShortPostCode == this.ShortPostCode
                        && targetAccount.CountyName == this.CountyName
                        && targetAccount.TownName == this.TownName);
        }

        public bool EqualsToContact(ContactProxy targetContact)
        {
            throw new NotImplementedException();
        }

        public bool EqualsToProject(ProjectProxy targetProject)
        {
            throw new NotImplementedException();
        }
    }
}

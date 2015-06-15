using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MergeExcelDuplicates.ProxyObjects
{
    public class ProjectProxy : ICommonProxy
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
            ContactId
        }

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

        public IEnumerable<KeyValuePair<string, string>> GetData()
        {
            return _rowData.Select(p => new KeyValuePair<string, string>(p.Key.ToString(), p.Value));
        }

        public bool EqualsToAccount(AccountProxy targetAccount)
        {
            throw new NotImplementedException();
        }

        public bool EqualsToContact(ContactProxy targetContact)
        {
            throw new NotImplementedException();
        }

        public bool EqualsToProject(ProjectProxy targetProject)
        {
            return targetProject.FullData.ToLower() == this.FullData.ToLower();
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
    }
}

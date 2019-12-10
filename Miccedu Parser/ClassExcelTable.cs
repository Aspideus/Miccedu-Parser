using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace miccedux
{
    public class ClassExcelTable
    {
        #region VariablesAndObjects

        public string year { get; private set; } = null;

        public string region { get; private set; } = null;

        public string type_organizations { get; private set; } = null;

        public string[] head_indicators { get; set; }

        public List<Organizations> org { get; private set; } = new List<Organizations>();

        public List<Indicators> indic { get; private set; } = new List<Indicators>();

        public List<Criteries> criteries { get; private set; } = new List<Criteries>();

        #endregion

        #region InnerClasses

        public class Criteries
        {
            public string criterion { get; set; }

            public Criteries(string criterion)
            {
                this.criterion = criterion;
            }
        }

        public class Indicators : INotifyPropertyChanged
        {
            public Indicators(int id, bool check, string number, string title, string type_value)
            {
                this.id = id;
                this.check = check;
                this.number = number;
                this.title = title;
                this.type_value = type_value;
                this.criterion_id = -1;
            }

            private bool _check;

            public int id { get; private set; }

            public bool check { get { return _check; } set { _check = value; OnPropertyChanged("check"); } }

            public string number { get; set; }

            public string title { get; set; }

            public string type_value { get; private set; }

            public int criterion_id { get; private set; }

            public string criterion { get; set; } = null;

            public event PropertyChangedEventHandler PropertyChanged;

            public void OnPropertyChanged([CallerMemberName]string prop = "")
            {
                if (PropertyChanged != null)
                    PropertyChanged(this, new PropertyChangedEventArgs(prop));
            }
        }

        public class Organizations : INotifyPropertyChanged
        {
            public Organizations(int id, bool check, string title, string url)
            {
                this.id = id;
                this.check = check;
                this.title = title;
                this.url = url;
            }

            private bool _check;

            public int id { get; private set; }

            public bool check { get { return _check; } set { _check = value; OnPropertyChanged("check"); } }

            public string title { get; set; }

            public string url { get; private set; }

            public List<string> values { get; private set; } = new List<string>();

            public void AddValue(string value)
            {
                values.Add(value);
            }

            public event PropertyChangedEventHandler PropertyChanged;

            public void OnPropertyChanged([CallerMemberName]string prop = "")
            {
                if (PropertyChanged != null)
                    PropertyChanged(this, new PropertyChangedEventArgs(prop));
            }
        }

        #endregion

        #region PublicVoids

        public void Clear()
        {
            org.Clear();
            indic.Clear();
            criteries.Clear();

            year = null;
            region = null;
            type_organizations = null;
        }

        public void SetTypes(string year, string region, string type_organizations)
        {
            if (this.year == null && this.region == null && this.type_organizations == null)
            {
                this.year = year;
                this.region = region;
                this.type_organizations = type_organizations;
            }
        }

        public bool IsAllCheckedOrganizations()
        {
            for (int i = 0; i < org.Count; i++)
            {
                if (org[i].check == false)
                    return false;
            }
            return true;
        }

        public bool IsAllCheckedIndicators()
        {
            for (int i = 0; i < indic.Count; i++)
            {
                if (indic[i].check == false)
                    return false;
            }
            return true;
        }

        public void SetCheckedOrganizations()
        {
            foreach (Organizations o in org)
            {
                if (o.check == false)
                    o.check = true;
            }
        }

        public void SetUncheckedOrganizations()
        {
            foreach (Organizations o in org)
            {
                if (o.check == true)
                    o.check = false;
            }
        }

        public void SetCheckedIndicators()
        {
            foreach (Indicators i in indic)
            {
                if (i.check == false)
                    i.check = true;
            }
        }

        public void SetUncheckedIndicators()
        {
            foreach (Indicators i in indic)
            {
                if (i.check == true)
                    i.check = false;
            }
        }

        public void SetCriteriesForIndicators()
        {
            for (int i = 0; i < indic.Count; i ++)
            {
                if (indic[i].criterion_id > -1 && indic[i].criterion_id < criteries.Count)
                    indic[i].criterion = criteries[indic[i].criterion_id].criterion;
            }
        }

        public bool IsReady()
        {
            if (GetCountChecked(org) > 0 && GetCountChecked(indic) > 0)
                return true;
            else
                return false;
        }

        public int GetCountChecked(List<Organizations> organizations)
        {
            int count = 0;

            for (int i = 0; i < organizations.Count; i ++)
            {
                if (organizations[i].check)
                    count++;
            }

            return count;
        }

        public int GetCountChecked(List<Indicators> indicators)
        {
            int count = 0;

            for (int i = 0; i < indicators.Count; i++)
            {
                if (indicators[i].check)
                    count++;
            }

            return count;
        }

        public bool OragnizationsValuesIsWithoutMdash(Organizations obj)
        {
            for (int j = 0; j < obj.values.Count; j++)
            {
                if (!obj.values[j].Contains("—"))
                {
                    return false;
                }

            }

            return true;
        }

        public int GetCountIndicatorsWithCriterion(List<Indicators> obj, string selected_criterion)
        {
            int count = 0;

            for (int i = 0; i < obj.Count; i++)
            {
                if (obj[i].criterion == selected_criterion)
                    count++;
            }

            return count;
        }

        #endregion
    }
}

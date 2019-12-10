using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;

namespace miccedux
{
    class SClass2 : INotifyPropertyChanged
    {
        #region UsedVariables

        private int _selectedIndex = -1;
        private string _selectedItem = null;
        private double _height;

        #endregion

        #region VariablesWithProperties

        public int last_correct_selected { get; private set; }

        public List<SClass2_1> Cells { get; private set; } = new List<SClass2_1>();

        public bool IsChanged { get; private set; } = false;

        private IEnumerable Items
        {
            get { return Cells; }
        }

        private string NewItem
        {
            set
            {
                if (SelectedItem != null)
                {
                    return;
                }

                if (!string.IsNullOrEmpty(value))
                {
                    Cells.Add(new SClass2_1(value, value));
                    SelectedItem = value;
                }
            }
        }

        public int SelectedIndex
        {
            get { return _selectedIndex; }
            private set
            {
                IsChanged = true;
                _selectedIndex = value;

                if (value > -1)
                    last_correct_selected = value;

                OnPropertyChanged("SelectedIndex");
            }
        }

        public string SelectedItem
        {
            get { return _selectedItem; }
            private set
            {
                _selectedItem = value;
                OnPropertyChanged("SelectedItem");
            }
        }

        public double height 
        { 
            get 
            { 
                return _height; 
            } 
            private set 
            { 
                _height = value; 
                OnPropertyChanged("height"); 
            } 
        }

        #endregion

        #region PublicMethods

        public void IsChangedToFalse()
        {
            IsChanged = false;
        }

        public void Clear()
        {
            Cells.Clear();
            SelectedIndex = -1;
            SelectedItem = null;
            last_correct_selected = -1;
        }

        public void Add(string title, string url)
        {
            Cells.Add(new SClass2_1(title, url));
            height = Cells.Count * 26.09;
        }

        public string[] GetArrayUrls()
        {
            string[] result = new string[Cells.Count];

            for (int i = 0; i < result.Length; i++)
                result[i] = Cells[i].url;

            return result;
        }

        public string GetTypeOrganizations()
        {
            if (last_correct_selected == 0)
                return "В";
            else
                return "К";
        }

        #endregion

        #region PropertyChanged

        protected void OnPropertyChanged(string propertyName)
        {
                var handler = this.PropertyChanged;
                if (handler != null)
                {
                    handler(this, new PropertyChangedEventArgs(propertyName));
                }
            
        }

        public event PropertyChangedEventHandler PropertyChanged;

        #endregion
    }
}

using System;
using System.Windows;
using System.Windows.Controls;

namespace miccedux
{
    public partial class ChangesPage : UserControl
    {
        #region ConstVariables

        private const string using_all_organizations = "SelectAllOrganizations";
        private const string using_all_indicators = "SelectAllIndicators";

        private const string using_organization = "SelectOrganization";
        private const string using_indicator = "SelectIndicator";

        #endregion

        #region Delegates

        private delegate bool _DelegateBool();
        private delegate void _DelegateVoid();

        #endregion

        public ChangesPage()
        {
            InitializeComponent();

            DataContext = Class1.exceltable.indic;

            indicators_data.ItemsSource = Class1.exceltable.indic;
            organizations_data.ItemsSource = Class1.exceltable.org;
            criteries_data.ItemsSource = Class1.exceltable.criteries;

            SetStackPanelsInHeader(ref organizations_use, using_all_organizations, Class1.exceltable.IsAllCheckedOrganizations());
            SetStackPanelsInHeader(ref indicators_use, using_all_indicators, Class1.exceltable.IsAllCheckedIndicators());

            void SetStackPanelsInHeader(ref DataGridTemplateColumn DGTC, string _CheckBox_name, bool IsChecked)
            {
                CheckBox cb = new CheckBox() { IsChecked = IsChecked, Name = _CheckBox_name };
                TextBlock l = new TextBlock() { Text = "Использовать" };
                StackPanel sp = new StackPanel() { Orientation = Orientation.Horizontal };

                cb.Click += CheckBox_Using_All;
                sp.Children.Add(cb);
                sp.Children.Add(l);

                DGTC.Header = sp;
            }
        }

        #region EventHandlerMethods

        private void CheckBox_Using_All(object sender, RoutedEventArgs e)
        {
            bool all_checked = false;
            _DelegateVoid SetChecked = null;
            _DelegateVoid SetUnchecked = null;

            switch ((sender as CheckBox).Name)
            {
                case using_all_organizations:
                    all_checked = Class1.exceltable.IsAllCheckedOrganizations();
                    SetChecked = Class1.exceltable.SetCheckedOrganizations;
                    SetUnchecked = Class1.exceltable.SetUncheckedOrganizations;
                    break;
                case using_all_indicators:
                    all_checked = Class1.exceltable.IsAllCheckedIndicators();
                    SetChecked = Class1.exceltable.SetCheckedIndicators;
                    SetUnchecked = Class1.exceltable.SetUncheckedIndicators;
                    break;
            }

            if (SetChecked != null && SetUnchecked != null)
            {
                if ((sender as CheckBox).IsChecked == true && all_checked == false)
                    SetChecked();
                else if ((sender as CheckBox).IsChecked == false && all_checked == true)
                    SetUnchecked();
            }
        }

        private void CheckBox_Using(object sender, RoutedEventArgs e)
        {
            _DelegateBool func = null;
            UIElementCollection ui_collection = null;
            DataGrid dg = null;
            CheckBox cb = null;

            switch ((sender as CheckBox).Name)
            {
                case using_organization:
                    func = Class1.exceltable.IsAllCheckedOrganizations;
                    dg = organizations_data;
                    ui_collection = (organizations_use.Header as StackPanel).Children;
                    break;
                case using_indicator:
                    func = Class1.exceltable.IsAllCheckedIndicators;
                    dg = indicators_data;
                    ui_collection = (indicators_use.Header as StackPanel).Children;
                    break;
            }

            if (ui_collection != null && func != null && dg != null)
            {
                foreach (UIElement ui in ui_collection)
                {
                    if (ui.GetType() == typeof(CheckBox))
                    {
                        cb = ui as CheckBox;
                        break;
                    }
                }
            }

            if (cb != null)
            {
                if (func() == true && cb.IsChecked == false)
                    cb.IsChecked = true;
                else if (func() == false && cb.IsChecked == true)
                    cb.IsChecked = false;

                if (dg.Items.NeedsRefresh)
                    dg.Items.Refresh();
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            string name_element = (sender as Button).Name;

            if (name_element == ToReturn.Name)
            {
                Class1.Window_Frame.NavigationService.GoBack();
            }
            else if (name_element == AddNewCriterion.Name)
            {
                ref TextBox text_box = ref new_row;
                string new_criterion = text_box.Text;

                if (new_criterion.Length > 0 && !IsRepeatingValues(new_criterion))
                    Class1.exceltable.criteries.Add(new ClassExcelTable.Criteries(new_criterion));
                else
                    MessageBox.Show("Критерий не может повторяться или быть без названия");

                text_box.Text = null;

                criteries_data.Items.Refresh();

                bool IsRepeatingValues(string value)
                {
                    for (int i = 0; i < Class1.exceltable.criteries.Count; i++)
                    {
                        if (Class1.exceltable.criteries[i].criterion == value)
                            return true;
                    }

                    return false;
                }
            }
        }

        private void TabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (((TabControl)sender).SelectedIndex == 1)
            {
                criteries_data.CancelEdit(DataGridEditingUnit.Row);
            }
        }

        private void SelectIndicator_Initialized(object sender, EventArgs e)
        {
            ComboBox cb = (sender as ComboBox);
            cb.ItemsSource = Class1.exceltable.criteries;
            cb.DisplayMemberPath = "criterion";
        }

        private void SelectIndicator_DropDownOpened(object sender, EventArgs e)
        {
            (sender as ComboBox).Items.Refresh();
        }

        private void TextBox_NewCriterion_KeyUp(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
                Button_Click(AddNewCriterion, null);
        }

        #endregion
    }
}

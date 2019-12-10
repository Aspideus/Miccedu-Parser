using miccedux.UserControls;
using System;
using System.ComponentModel;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using HtmlAgilityPack;

namespace miccedux
{
    public partial class MainPage : UserControl
    {
        #region ConstVariables

        const string Combo_Monitorings_Name = "Combo_Monitorings";
        const string Combo_Years_Name = "Combo_Years";
        const string Combo_Regions_Name = "Combo_Regions";
        const string Organizations_name = "Organizations";

        #endregion

        #region ClassObjects

        private BackgroundWorker BW_ToParse = new BackgroundWorker();
        private BackgroundWorker BW_ToWait = new BackgroundWorker();

        private SClass2 monitorings { get; set; } = new SClass2();
        private SClass2 years { get; set; } = new SClass2();
        private SClass2 regions { get; set; } = new SClass2();
        private SClass2 organizations { get; set; } = new SClass2();

        #endregion

        public MainPage()
        {
            InitializeComponent();

            CompareAndChange();

            Combo_Monitorings.DataContext = monitorings;
            Combo_Years.DataContext = years;
            Combo_Regions.DataContext = regions;

            BW_ToParse.WorkerSupportsCancellation = true;
            BackgroundWorker bw = new BackgroundWorker();
            bw.DoWork += SetValues;
            bw.RunWorkerAsync(argument: "Combo_Monitorings");

            void CompareAndChange()
            {
                if (Combo_Monitorings.Name != Combo_Monitorings_Name)
                    Combo_Monitorings.Name = Combo_Monitorings_Name;

                if (Combo_Years.Name != Combo_Years_Name)
                    Combo_Years.Name = Combo_Years_Name;

                if (Combo_Regions.Name != Combo_Regions_Name)
                    Combo_Regions.Name = Combo_Regions_Name;
            }
        }

        #region PrivateMethods

        private HtmlDocument GetResponseDoc(string address)
        {
            try
            {
                var web = new HtmlWeb();
                web.OverrideEncoding = Encoding.Default;
                HtmlDocument _htmlDocument = web.Load(address);
                return _htmlDocument;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private SClass2_1[] ParseValues(string url, string select_nodes, string widget_name)
        {
            HtmlNodeCollection hnc = GetResponseDoc(url)?.DocumentNode?.SelectNodes(select_nodes);
            string current_url = null;
            SClass2 tmp_sclass2 = new SClass2();

            try
            {
                if (widget_name == Combo_Monitorings_Name)
                {
                    foreach (HtmlNode hn in hnc)
                    {
                        current_url = Regex.Replace(hn.Attributes["onclick"]?.Value, "^(.[^'])+ |('|;)", "").Trim();
                        current_url = url + current_url ?? throw new Exception("url cannot be empty");
                        tmp_sclass2.Add(hn.SelectSingleNode("div[1]")?.InnerText ?? current_url, current_url);
                    }
                }
                else if (widget_name == Combo_Years_Name)
                {
                    foreach (HtmlNode hn in hnc)
                    {
                        current_url = hn.Attributes["href"].Value;
                        tmp_sclass2.Add(hn.SelectSingleNode("b")?.InnerText.Trim() ?? current_url, current_url ?? throw new Exception("url cannot be empty"));
                    }
                }
                else if (widget_name == Combo_Regions_Name)
                {
                    foreach (HtmlNode hn in hnc)
                    {
                        current_url = CheckUrl(url, hn.Attributes["href"].Value);
                        tmp_sclass2.Add(hn.InnerText.Trim() ?? current_url, current_url ?? throw new Exception("url cannot be empty"));
                    }
                }
                else if (widget_name == Organizations_name)
                {
                    foreach (HtmlNode hn in hnc)
                    {
                        current_url = CheckUrl(url, hn.Attributes["href"].Value);
                        tmp_sclass2.Add(hn.InnerText.Trim() ?? current_url, current_url ?? throw new Exception("url cannot be empty"));
                    }
                }
            }
            catch (Exception) { }

            return tmp_sclass2.Cells.ToArray();

            string CheckUrl(string address, string new_address)
            {

                return Regex.IsMatch(new_address, "http") ? new_address : Regex.Match(address, @"^(.+\/)").ToString() + new_address;
            }
        }

        private void ToggleWidgets(ComboBox changed_cb, bool NeedToEnable)
        {
            ComboBox[] arr_ComboBoxes = { Combo_Monitorings, Combo_Years, Combo_Regions };
            SClass2[] arr_ItemsSources = { monitorings, years, regions };


            if (NeedToEnable == false)
            {
                int index_current_ComboBox = 0;

                Class1.exceltable.Clear();

                for (int i = 0; i < arr_ComboBoxes.Length; i++)
                {
                    if (arr_ComboBoxes[i] == changed_cb)
                    {
                        index_current_ComboBox = i;
                        break;
                    }
                }

                for (int i = index_current_ComboBox; i < arr_ComboBoxes.Length; i++)
                {

                    if (i != index_current_ComboBox)
                    {
                        arr_ComboBoxes[i].IsEnabled = false;
                        arr_ItemsSources[i].Clear();
                    }
                }

                organizations.Clear();
                ToChangesPage.IsEnabled = false;
            }
            else
            {
                for (int i = 0; i < arr_ItemsSources.Length; i++)
                {
                    if (arr_ItemsSources[i].Cells.Count > 0)
                        arr_ComboBoxes[i].IsEnabled = true;
                }

                if (organizations.Cells.Count > 0 && regions.Cells.Count > 0)
                    ToChangesPage.IsEnabled = true;
            }
        }

        private void SetValues(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            BackgroundWorker bw = BW_ToParse;

            try
            {
                switch (e.Argument as string)
                {
                    case Combo_Monitorings_Name:
                        AddIntoComboBox(monitorings, Combo_Monitorings, ParseValues(Properties.Resources.url_monitorings, "//div[@class='casemonitoring']", (string)e.Argument));
                        break;
                    case Combo_Years_Name:
                        AddIntoComboBox(years, Combo_Years, ParseValues(monitorings.Cells[monitorings.SelectedIndex].url, "//div[@id='arh']/p/a", (string)e.Argument));
                        break;
                    case Combo_Regions_Name:
                        AddIntoComboBox(regions, Combo_Regions, ParseValues(years.Cells[years.SelectedIndex].url, "//table[@id='tregion']/tr/td/p/a", (string)e.Argument));
                        break;
                    case Organizations_name:
                        AddIntoComboBox(organizations, null, ParseValues(regions.Cells[regions.SelectedIndex].url, "//table[@class='an']//a", (string)e.Argument));
                        break;
                }

                e.Cancel = true;

                void AddIntoComboBox(SClass2 itemssource, ComboBox _ComboBox, in SClass2_1[] new_cells)
                {
                    if (bw.CancellationPending == true)
                    {
                        e.Cancel = true;
                        return;
                    }

                    foreach (SClass2_1 s in new_cells)
                    {
                        itemssource.Add(s.title, s.url);
                    }

                    if (itemssource == organizations)
                        SetIndicatorsTitles();

                    Dispatcher.Invoke(new Action(() => { if (bw.CancellationPending == false) { _ComboBox?.Items.Refresh(); ToggleWidgets(null, true); Class1.Window_App.Topmost = true; Class1.Window_App.Focus(); _ComboBox?.Focus(); Class1.Window_App.Topmost = false; } }));

                    void SetIndicatorsTitles()
                    {
                        string[] organizations_urls = organizations.GetArrayUrls();

                        for (int i = 0; i < organizations_urls.Length; i++)
                        {
                            HtmlDocument _HtmlDocument = GetResponseDoc(organizations_urls[i]);

                            if (_HtmlDocument.DocumentNode.SelectSingleNode("//table[@class='napde']") != null)
                            {
                                GetTitlesHead(in _HtmlDocument);
                                GetTitlesRows(in _HtmlDocument);
                                break;
                            }
                        }

                        for (int i = 0; i < organizations.Cells.Count; i++)
                        {
                            Class1.exceltable.org.Add(new ClassExcelTable.Organizations(i, true, organizations.Cells[i].title, organizations.Cells[i].url));
                        }

                        void GetTitlesHead(in HtmlDocument hd)
                        {
                            HtmlNodeCollection td_nodes = hd.DocumentNode.SelectSingleNode("//table[@class='napde']/tr[@class='napr_head']").SelectNodes("td");

                            string[] head_indicators = new string[td_nodes.Count];

                            for (int i = 0; i < td_nodes.Count; i++)
                            {
                                head_indicators[i] = td_nodes[i].InnerText;
                            }

                            Class1.exceltable.head_indicators = head_indicators;
                        }

                        void GetTitlesRows(in HtmlDocument hd)
                        {
                            HtmlNodeCollection[] columns = new HtmlNodeCollection[3];
                            columns[0] = hd.DocumentNode.SelectNodes("//table[@class='napde']/tr[position()>1]/td[1]");
                            columns[1] = hd.DocumentNode.SelectNodes("//table[@class='napde']/tr[position()>1]/td[2]");
                            columns[2] = hd.DocumentNode.SelectNodes("//table[@class='napde']/tr[position()>1]/td[3]");

                            for (int i = 0; i < columns[0].Count; i++)
                            {
                                Class1.exceltable.indic.Add(new ClassExcelTable.Indicators(i, true, columns[0][i].InnerText, columns[1][i].InnerText, columns[2][i].InnerText));
                            }
                        }
                    }
                }
            }
            catch (Exception ex) { MessageBox.Show("SetValues method exception" + Environment.NewLine + ex.Message); };
        }

        #endregion

        #region EventHandlerMethods

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            string name_element = (sender as Button).Name;

            try
            {
                if (name_element == ToChangesPage.Name)
                {
                    Class1.Window_Frame.NavigationService.Navigate(new ChangesPage());
                }
                else if (name_element == ToExcelCollection.Name && Class1.exceltable.IsReady())
                {
                    lock (Class1.exceltable)
                    {
                        Class1.exceltable.SetTypes(years.Cells[years.last_correct_selected].title, regions.Cells[regions.last_correct_selected].title, monitorings.GetTypeOrganizations());
                        Class1.exceltable.SetCriteriesForIndicators();
                        Class1.Window_Frame.NavigationService.Navigate(new ExcelCollection());
                    }
                }
                else if (name_element == ToExcelCollection.Name)
                {
                    MessageBox.Show("Необходимо выбрать хотя бы одну организацию и показатель");
                }
            }
            catch (Exception ex) { MessageBox.Show("Button_Click method exception" + Environment.NewLine + ex.Message); }
        }

        private void ComboBox_DropDownClosed(object sender, EventArgs e)
        {
            ComboBox current_ComboBox = sender as ComboBox;

            try
            {
                if (current_ComboBox.SelectedIndex > -1 && (current_ComboBox.DataContext as SClass2).IsChanged == true)
                {
                    (current_ComboBox.DataContext as SClass2).IsChangedToFalse();
                    Combo_Change_Selection(ref current_ComboBox);

                    TextBox tb = current_ComboBox.Template.FindName("PART_EditableTextBox", current_ComboBox) as TextBox;
                    tb.Select(0, 0);
                    tb.CaretBrush = Brushes.Transparent;
                }
            }
            catch (Exception ex) { MessageBox.Show("ComboBox_DropDownClosed method exception" + Environment.NewLine + ex.Message); };

            void Combo_Change_Selection(ref ComboBox _ComboBox)
            {
                if (_ComboBox.SelectedIndex > -1)
                {
                    ToggleWidgets(_ComboBox, false);

                    if (BW_ToWait.IsBusy)
                        BW_ToWait.CancelAsync();

                    BW_ToWait = new BackgroundWorker();
                    BW_ToWait.WorkerSupportsCancellation = true;
                    BW_ToWait.DoWork += DoWaitAndRun;
                    BW_ToWait.RunWorkerAsync(argument: _ComboBox.Name);

                    void DoWaitAndRun(object sender, DoWorkEventArgs e)
                    {
                        BackgroundWorker _this_bw = BW_ToWait;

                        if (BW_ToParse.IsBusy)
                            BW_ToParse.CancelAsync();

                        while (BW_ToParse.IsBusy)
                        {
                            Thread.Sleep(1000);
                            if (_this_bw.CancellationPending == true)
                            {
                                e.Cancel = true;
                                return;
                            }
                        }

                        BW_ToParse.Dispose();
                        BW_ToParse = new BackgroundWorker();
                        BW_ToParse.WorkerSupportsCancellation = true;
                        BW_ToParse.DoWork += SetValues;

                        switch ((string)e.Argument)
                        {
                            case Combo_Monitorings_Name:
                                BW_ToParse.RunWorkerAsync(argument: Combo_Years_Name);
                                break;
                            case Combo_Years_Name:
                                BW_ToParse.RunWorkerAsync(argument: Combo_Regions_Name);
                                break;
                            case Combo_Regions_Name:
                                BW_ToParse.RunWorkerAsync(argument: Organizations_name);
                                break;
                        }

                        e.Cancel = true;

                    }
                }
            }
        }

        private void ComboBox_KeyDown(object sender, KeyEventArgs e)
        {
            ComboBox current_ComboBox = sender as ComboBox;

            if (e.Key == Key.Enter && current_ComboBox.SelectedIndex > -1)
                current_ComboBox.IsDropDownOpen = false;
            else if (current_ComboBox.IsDropDownOpen == false)
                current_ComboBox.IsDropDownOpen = true;
        }

        private void ComboBox_GotFocus(object sender, RoutedEventArgs e)
        {
            if ((sender as ComboBox).Items.Count > -1)
                (sender as ComboBox).IsDropDownOpen = true;
        }

        private void ComboBox_Loaded(object sender, RoutedEventArgs e)
        {
            ((sender as ComboBox).Template.FindName("PART_EditableTextBox", (sender as ComboBox)) as TextBox).CaretBrush = (SolidColorBrush)Application.Current.Resources["ControlForegroundWhite"];
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            while (Class1.Window_Frame.CanGoBack)
                Class1.Window_Frame.NavigationService.RemoveBackEntry();
        }

        #endregion

        #region OverrideMethods

        protected override void OnPreviewKeyDown(KeyEventArgs e)
        {
            if (e.Key == Key.Down || e.Key == Key.Up)
            {
                e.Handled = true;
                return;
            }

            base.OnPreviewKeyDown(e);
        }

        #endregion
    }
}

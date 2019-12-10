using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using System.Windows.Controls;

using Excel = Microsoft.Office.Interop.Excel;

namespace miccedux.UserControls
{
    public partial class ExcelCollection : UserControl
    {
        #region InnerClasses

        private class Table
        {
            public int first_row { get; set; }
            public int first_column { get; set; }

            public int height { get; set; }
            public int width { get; set; }

            public int second_row { get; set; }
            public int second_column { get; set; }

            public Table(int first_row, int first_column, int height, int width, int second_row, int second_column)
            {
                this.first_row = first_row;
                this.first_column = first_column;
                this.height = height;
                this.width = width;
                this.second_row = second_row;
                this.second_column = second_column;
            }

            public void SetSeconds(int second_row, int second_column)
            {
                this.second_row = second_row;
                this.second_column = second_column;
            }

            public void SetDefaultRange(int first_row, int height, int second_row, int second_column)
            {
                this.first_row = this.second_row + first_row;
                this.height = height;
                this.second_row = this.first_row + this.height + second_row;
                if (second_column != -1)
                    this.second_column = this.first_column + second_column;
            }

            public void SetTransparentRange(int first_column, int width, int second_column, int second_row)
            {
                this.first_column = this.second_column + first_column;
                this.width = width;
                this.second_column = this.first_column + this.width + second_column;
                if (second_row != -1)
                    this.second_row = this.first_row + second_row;
            }
        }

        private class RangeStyle
        {
            public int ColumnWidth;
            public bool WrapText;
            public string FontName;
            public int FontSize;
            public bool FontBold;
            public Color Interior;
            public Excel.Constants HorizontalAligment;
            public Excel.Constants VerticalAligment;
            public Excel.XlLineStyle BorderLineStyle;
            public Excel.XlBorderWeight BorderWeight;

            public RangeStyle(int ColumnWidth, bool WrapText, string FontName, int FontSize, bool FontBold, Color Interior,
                Excel.Constants HorizontalAligment, Excel.Constants VerticalAligment, Excel.XlLineStyle BorderLineStyle, Excel.XlBorderWeight BorderWeight)
            {
                this.ColumnWidth = ColumnWidth;
                this.WrapText = WrapText;
                this.FontName = FontName;
                this.FontSize = FontSize;
                this.FontBold = FontBold;
                this.Interior = Interior;
                this.HorizontalAligment = HorizontalAligment;
                this.VerticalAligment = VerticalAligment;
                this.BorderLineStyle = BorderLineStyle;
                this.BorderWeight = BorderWeight;
            }
        }

        #endregion

        #region PrivateVariablesAndObjects

        private Thread t;

        readonly private RangeStyle CriteriesStyle = new RangeStyle(20, true, "Times New Roman", 14, true, Color.FromArgb(255, 242, 205), Excel.Constants.xlCenter,
            Excel.Constants.xlCenter, Excel.XlLineStyle.xlDouble, Excel.XlBorderWeight.xlMedium);

        readonly private RangeStyle IndicatorsTitlesAndOrganizationsStyle = new RangeStyle(30, true, "Times New Roman", 14, true, Color.FromArgb(229, 229, 229), Excel.Constants.xlCenter,
            Excel.Constants.xlCenter, Excel.XlLineStyle.xlDouble, Excel.XlBorderWeight.xlMedium);

        readonly private RangeStyle IndicatorsStyle = new RangeStyle(20, true, "Times New Roman", 14, true, Color.FromArgb(229, 229, 229), Excel.Constants.xlCenter,
            Excel.Constants.xlCenter, Excel.XlLineStyle.xlDouble, Excel.XlBorderWeight.xlMedium);

        readonly private RangeStyle ValuesStyle = new RangeStyle(30, true, "Times New Roman", 14, true, new Color(), Excel.Constants.xlCenter,
            Excel.Constants.xlCenter, Excel.XlLineStyle.xlDouble, Excel.XlBorderWeight.xlThin);

        #endregion

        public ExcelCollection()
        {
            InitializeComponent();

            t = new Thread(Background_Thread);
            t.IsBackground = true;
            t.Start();

            void Background_Thread()
            {
                try
                {
                    List<ClassExcelTable.Organizations> tmp_orgs = new List<ClassExcelTable.Organizations>(Class1.exceltable.org);
                    List<ClassExcelTable.Indicators> tmp_indics = new List<ClassExcelTable.Indicators>(Class1.exceltable.indic);

                    for (int i = 0; i < tmp_orgs.Count; i++)
                    {
                        if (tmp_orgs[i].check == true)
                        {
                            ParseValues(tmp_orgs[i], tmp_indics);
                        }
                        Dispatcher.Invoke(new Action(() => Progress_Bar.Value = ((i + 1) * 100 / tmp_orgs.Count)));
                    }

                    RemoveEmptyOrUncheckedOrganizations(tmp_orgs);
                    RemoveUncheckedIndicators(tmp_indics);
                    CreateBaseExcel(tmp_orgs, tmp_indics);
                }
                catch (Exception ex) { MessageBox.Show("Background_Thread method exception" + Environment.NewLine + ex.Message); }

                void ParseValues(ClassExcelTable.Organizations org, List<ClassExcelTable.Indicators> indics)
                {
                    HtmlDocument hd = GetResponseDoc(org.url);

                    HtmlNodeCollection td_4 = hd?.DocumentNode.SelectNodes("//table[@class='napde']/tr[position()>1]/td[4]");

                    if (td_4 == null)
                        return;

                    List<string> values = new List<string>();

                    for (int i = 0; i < td_4.Count; i++)
                    {
                        if (indics[i].check == true)
                        {
                            if (!td_4[i].InnerText.Contains("&mdash"))
                            {
                                org.values.Add(td_4[i].InnerText);
                            }
                            else
                            {
                                org.values.Add("—");
                            }
                        }
                    }

                    HtmlDocument GetResponseDoc(string address)
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
                }

                void RemoveEmptyOrUncheckedOrganizations(List<ClassExcelTable.Organizations> orgs)
                {
                    bool empty;

                    for (int i = orgs.Count - 1; i > -1; i--)
                    {
                        if (orgs[i].values.Count == 0)
                        {
                            orgs.RemoveAt(i);
                        }
                        else
                        {
                            empty = true;

                            if (orgs[i].check)
                            {
                                for (int j = 0; j < orgs[i].values.Count; j++)
                                {
                                    if (orgs[i].values[j] != "")
                                        empty = false;
                                }
                            }

                            if (empty)
                                orgs.RemoveAt(i);
                        }
                    }
                }

                void RemoveUncheckedIndicators(List<ClassExcelTable.Indicators> indics)
                {
                    for (int i = indics.Count - 1; i > -1; i--)
                    {
                        if (indics[i].check == false)
                            indics.RemoveAt(i);
                    }
                }
            }
        }

        #region PrivateMethods

        private void CreateBaseExcel(List<ClassExcelTable.Organizations> orgs, List<ClassExcelTable.Indicators> indics)
        {
            Class1.ExcelAlive = true;
            string[] table_types = { "Обычный", "Без значений в скобках", "Таблица топов" };

            Dispatcher.Invoke(new Action(() =>
            {
                status_block.Text = "Формирование документа...";
                Progress_Bar.Value = 0;
                Progress_Bar.Maximum = table_types.Length * 2;
            }));

            Excel.Application xlApp = new Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("На вашем ПК не установлен Excel !!");
                return;
            }

            Excel.Workbook xlWorkBook;
            List<Excel.Worksheet> list_worksheets = new List<Excel.Worksheet>();

            object misValue = System.Reflection.Missing.Value;
            xlWorkBook = xlApp.Workbooks.Add(misValue);

            string region = Class1.exceltable.region;
            string type_organizations = Class1.exceltable.type_organizations;
            string year = Class1.exceltable.year;

            for (int i = 0; i < table_types.Length; i++)
            {
                list_worksheets.Add(NewWorksheet(table_types[i], false, i));
                Dispatcher.Invoke(new Action(() => Progress_Bar.Value++));
            }

            //транспонированные
            for (int i = 0; i < table_types.Length; i++)
            {
                list_worksheets.Add(NewWorksheet($"Т {table_types[i]}", true, i));
                Dispatcher.Invoke(new Action(() => Progress_Bar.Value++));
            }

            Excel.Worksheet NewWorksheet(string name, bool transparent, int type)
            {
                Excel.Worksheet xlWorkSheet = null;

                if (list_worksheets.Count > 0)
                     xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
                else
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                xlWorkSheet.Name = name;

                SetHead();
                AddRangeTitlesAndOrganizations(out int RowOrColumn);
                AddIndicators(in RowOrColumn);
                AddValues(in RowOrColumn);

                return xlWorkSheet;

                void AddRangeTitlesAndOrganizations(out int index)
                {
                    Table table_size = new Table(2, 1, 1, Class1.exceltable.head_indicators.Length - 1 + Class1.exceltable.GetCountChecked(orgs), 0, 0);

                    object[,] arr = new object[table_size.height, table_size.width];

                    for (int i = 0; i < Class1.exceltable.head_indicators.Length; i++)
                    {
                        arr[0, i] = Class1.exceltable.head_indicators[i];
                    }

                    for (int i = 0; i < orgs.Count; i++)
                    {
                        arr[0, i + Class1.exceltable.head_indicators.Length - 1] = orgs[i].title;
                    }

                    if (transparent)
                        arr = TransparentArray(arr);

                    table_size.SetSeconds(arr.GetLength(0) + table_size.first_row - 1, arr.GetLength(1) + table_size.first_column - 1);

                    Excel.Range range_indicators_titles = xlWorkSheet.get_Range((Excel.Range)xlWorkSheet.Cells[table_size.first_row, table_size.first_column], (Excel.Range)xlWorkSheet.Cells[table_size.second_row, table_size.second_column]);
                    SetStyleRange(ref range_indicators_titles, IndicatorsTitlesAndOrganizationsStyle);

                    range_indicators_titles.Value = arr;

                    if (transparent)
                        index = table_size.second_column;
                    else
                        index = table_size.second_row;
                }

                void AddIndicators(in int index)
                {
                    Table table_size;

                    if (transparent)
                        table_size = new Table(2, 1, 0, 0, 4, index);
                    else
                        table_size = new Table(1, 1, 0, 0, index, 3);

                    object[,] arr;
                    int arr_index = 0;

                    for (int k = 0; k < Class1.exceltable.criteries.Count + 1; k++)
                    {
                        string criterion = null;

                        if (k < Class1.exceltable.criteries.Count)
                        {
                            criterion = Class1.exceltable.criteries[k].criterion;
                        }

                        if (transparent)
                        {
                            table_size.SetTransparentRange(1, Class1.exceltable.GetCountIndicatorsWithCriterion(indics, criterion) + 1, -1, -1);
                            arr = new object[table_size.width, table_size.second_row];
                        }
                        else
                        {
                            table_size.SetDefaultRange(1, Class1.exceltable.GetCountIndicatorsWithCriterion(indics, criterion) + 1, -1, -1);
                            arr = new object[table_size.height, table_size.second_column];
                        }

                        arr_index = 0;

                        for (int i = 0; i < indics.Count; i++)
                        {
                            if (indics[i].criterion == criterion)
                            {
                                arr[++arr_index, 0] = indics[i].number;
                                arr[arr_index, 1] = indics[i].title;
                                arr[arr_index, 2] = indics[i].type_value;
                            }
                        }

                        Excel.Range[] criteries_ranges = null;

                        if (transparent)
                        {
                            arr = TransparentArray(arr);
                            criteries_ranges = new Excel.Range[] { xlWorkSheet.get_Range((Excel.Range)xlWorkSheet.Cells[table_size.first_row, table_size.first_column], (Excel.Range)xlWorkSheet.Cells[table_size.second_row, table_size.first_column]),
                                xlWorkSheet.get_Range((Excel.Range)xlWorkSheet.Cells[table_size.second_row + 1, table_size.first_column], (Excel.Range)xlWorkSheet.Cells[table_size.second_row + orgs.Count, table_size.first_column]) };
                        }
                        else
                        {
                            criteries_ranges = new Excel.Range[] { xlWorkSheet.get_Range((Excel.Range)xlWorkSheet.Cells[table_size.first_row, table_size.first_column], (Excel.Range)xlWorkSheet.Cells[table_size.first_row, table_size.second_column]),
                                xlWorkSheet.get_Range((Excel.Range)xlWorkSheet.Cells[table_size.first_row, table_size.second_column + 1], (Excel.Range)xlWorkSheet.Cells[table_size.first_row, table_size.second_column + orgs.Count]) };
                        }

                        Excel.Range range_indicators_titles = xlWorkSheet.get_Range((Excel.Range)xlWorkSheet.Cells[table_size.first_row, table_size.first_column], (Excel.Range)xlWorkSheet.Cells[table_size.second_row, table_size.second_column]);
                        SetStyleRange(ref range_indicators_titles, IndicatorsStyle);
                        range_indicators_titles.NumberFormat = "@";
                        range_indicators_titles.Value = arr;
                        SetCriterion(ref criteries_ranges);

                        void SetCriterion(ref Excel.Range[] range_criterions)
                        {
                            for (int i = 0; i < range_criterions.Length; i ++)
                            {
                                range_criterions[i].Merge(misValue);
                                range_criterions[i].Value = criterion ?? "Без критерия";
                                SetStyleRange(ref range_criterions[i], CriteriesStyle);
                            }
                        }
                    }
                }

                void AddValues(in int index)
                {
                    Table table_size;

                    if (transparent)
                        table_size = new Table(5, 1, 0, 0, 3, index);
                    else
                        table_size = new Table(1, 4, 0, 0, index, 3);

                    object[,] arr;

                    for (int k = 0; k < Class1.exceltable.criteries.Count + 1; k++)
                    {
                        string criterion = null;

                        if (k < Class1.exceltable.criteries.Count)
                            criterion = Class1.exceltable.criteries[k].criterion.ToString();

                        if (transparent)
                        {
                            table_size.SetTransparentRange(2, Class1.exceltable.GetCountIndicatorsWithCriterion(indics, criterion), -1, orgs.Count - 1);
                            arr = new object[table_size.width, orgs.Count];
                        }
                        else
                        {
                            table_size.SetDefaultRange(2, Class1.exceltable.GetCountIndicatorsWithCriterion(indics, criterion), -1, orgs.Count - 1);
                            arr = new object[table_size.height, orgs.Count];
                        }

                        StandartTable(arr, criterion);

                        if (transparent)
                            arr = TransparentArray(arr);

                        if (type == 1)
                            RemoveParentheses(arr);

                        Excel.Range range_indicators_titles = xlWorkSheet.get_Range((Excel.Range)xlWorkSheet.Cells[table_size.first_row, table_size.first_column], (Excel.Range)xlWorkSheet.Cells[table_size.second_row, table_size.second_column]);
                        SetStyleRange(ref range_indicators_titles, ValuesStyle);

                        range_indicators_titles.Value = arr;
                    }

                    void StandartTable(object[,] arr, string criterion)
                    {
                        object[,] tmp_arr = null;

                        if (type == 2)
                        {
                            tmp_arr = TableOfTops();
                        }

                        int arr_row_index = 0;
                        int arr_column_index = 0;

                        for (int i = 0; i < orgs.Count; i++)
                        {
                            for (int j = 0; j < indics.Count; j++)
                            {
                                if (indics[j].criterion == criterion)
                                {
                                    if (type != 2)
                                        arr[arr_row_index, arr_column_index] = orgs[i].values[j];
                                    else
                                        arr[arr_row_index, arr_column_index] = tmp_arr[j, i];

                                    arr_row_index++;
                                }
                            }
                            arr_column_index++;
                            arr_row_index = 0;
                        }
                    }

                    object[,] TableOfTops()
                    {
                        object[,] tmp_arr = new object[indics.Count, orgs.Count];
                        int index_indic = 0;

                        for (int i = 0; i < indics.Count; i++)
                        {
                            double[] values_indexOftop = new double[orgs.Count];
                            int _index = 0;

                            for (int k = 0; k < orgs.Count; k++)
                            {
                                if (Class1.exceltable.OragnizationsValuesIsWithoutMdash(orgs[k]) != true)
                                {
                                    values_indexOftop[_index] = 0;
                                    if (orgs[k].values.Count > i)
                                        double.TryParse(Regex.Replace(orgs[k].values[i] ?? null, @"\(.+\)", "").Trim(), out values_indexOftop[_index]);
                                    _index++;
                                }
                            }

                            List<double> tmp_list = new List<double>();

                            for (int n = 0; n < values_indexOftop.Length; n++)
                            {
                                if (!tmp_list.Contains(values_indexOftop[n]) && values_indexOftop[n] != 0)
                                    tmp_list.Add(values_indexOftop[n]);
                            }

                            tmp_list.Sort((a, b) => -1 * a.CompareTo(b));

                            for (int k = 0; k < _index; k++)
                            {
                                if (values_indexOftop[k] != 0)
                                {
                                    tmp_arr[index_indic, k] = tmp_list.IndexOf(values_indexOftop[k]) + 1;
                                }
                            }
                            index_indic++;
                        }
                        return tmp_arr;
                    }
                }

                object[,] TransparentArray(object[,] arr)
                {
                    object[,] tmp_arr = new object[arr.GetLength(1), arr.GetLength(0)];

                    for (int i = 0; i < tmp_arr.GetLength(0); i++)
                    {
                        for (int j = 0; j < tmp_arr.GetLength(1); j++)
                        {
                            tmp_arr[i, j] = arr[j, i];
                        }
                    }

                    return tmp_arr;
                }

                void RemoveParentheses(object[,] arr)
                {
                    for (int i = 0; i < arr.GetLength(0); i++)
                    {
                        for (int j = 0; j < arr.GetLength(1); j++)
                        {
                            arr[i, j] = Regex.Replace(arr[i, j] as string, @"\(.+\)", "");
                        }
                    }
                }

                void SetStyleRange(ref Excel.Range range, RangeStyle r_style)
                {
                    if (r_style.ColumnWidth != new int())
                        range.ColumnWidth = r_style.ColumnWidth;
                    range.WrapText = r_style.WrapText;
                    range.Font.Name = r_style.FontName;
                    range.Font.Size = r_style.FontSize;
                    range.Font.Bold = r_style.FontBold;
                    if (r_style.Interior != new Color())
                        range.Interior.Color = r_style.Interior;
                    range.HorizontalAlignment = r_style.HorizontalAligment;
                    range.VerticalAlignment = r_style.VerticalAligment;
                    range.Borders.LineStyle = r_style.BorderLineStyle;
                    range.Borders.Weight = r_style.BorderWeight;
                }

                void SetHead()
                {
                    Excel.Range range_head = xlWorkSheet.get_Range((Excel.Range)xlWorkSheet.Cells[1, 1], (Excel.Range)xlWorkSheet.Cells[1, 3]);
                    SetStyleRange(ref range_head, CriteriesStyle);
                    range_head.Merge(misValue);
                    range_head.Value = region + " " + year + " " + type_organizations;

                }
            }

            Dispatcher.Invoke(new Action(() => Progress_Bar.Value++));

            string file_path = Directory.GetCurrentDirectory() + @"\" + $"{region}_{year}_{type_organizations}.xlsx";

            try
            {
                xlWorkBook.SaveAs(file_path, Excel.XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, misValue,
                misValue, Excel.XlSaveAsAccessMode.xlNoChange, misValue, misValue, misValue, misValue, misValue);

                MessageBox.Show("Файл успешно создан, путь к файлу: " + file_path);
            }
            catch (Exception) { }

            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            foreach (Excel.Worksheet ws in list_worksheets)
            {
                Marshal.ReleaseComObject(ws);
            }

            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            Class1.ExcelAlive = false;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (t == null || (t != null && !t.IsAlive))
            {
                Class1.Window_Frame.NavigationService.GoBack();
            }
            else
            {
                MessageBox.Show("В данный момент выполняется загрузка и формирование EXCEL документа");
            }
        }

        #endregion
    }
}

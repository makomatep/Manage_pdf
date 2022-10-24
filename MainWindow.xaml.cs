using System;
using Init_Data;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Microsoft.Win32;
using System.Data;
using System.Windows.Threading;
using System.ComponentModel;
using System.Text.RegularExpressions;
using Aspose.Pdf.Facades;
using Microsoft.WindowsAPICodePack.Dialogs;
using System.Drawing;



namespace Tayse_Manage_Pdf
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        public string Str_Folder_Save { get; set; }
        public string Str_Folder_Pdf { get; set; }
        public string Str_Data { get; set; }

        public string Str_Global_File_Pdf { get; set; }
        public string Str_Global_File_Excel { get; set; }
        public string CityName { get; set; }

        public int NbreTotal { get; set; }
        public int Int_Combo_Template { get; set; }

        public List<DataGrid_Init_10> List_Excel_File = new List<DataGrid_Init_10>();
        public List<Combo_InitList> List_Excel_Acumatica_File = new List<Combo_InitList>();
        public List<Combo_InitList> List_Tempon = new List<Combo_InitList>();
        public List<Combo_InitList> List_Order_Pdf = new List<Combo_InitList>();
        public List<Combo_InitList> List_Operation = new List<Combo_InitList>();
        public List<Combo_InitList> List_Combo_Template = new List<Combo_InitList>();
        public List<Combo_InitList> List_File_Created = new List<Combo_InitList>();
        public List<DataGrid_Init_10> List_Pdf_File = new List<DataGrid_Init_10>();

        public List<DataGrid_Init_10> List_Data_Pdf_File = new List<DataGrid_Init_10>();
        public List<DataGrid_Init_10> List_Data_Excel_File = new List<DataGrid_Init_10>();

        public List<DataGrid_Init_10> Display_Data10 = new List<DataGrid_Init_10>();
        public List<DataGrid_Init_10> Display_Data10_X_ALL = new List<DataGrid_Init_10>();
        public List<DataGrid_Init_10> Display_Data10_01 = new List<DataGrid_Init_10>();
        public List<DataGrid_Init_10> Display_Data10_All = new List<DataGrid_Init_10>();

        public MainWindow()
        {
            InitializeComponent();
            LabelHeader.Width = 1400;
            Label_Pdf_Hide.Width = 980;
            Label_Pdf_Hide.Content = "";
            // ------
            Manage_Folder();
            // ------------------------------------

            Label_Data_Hide.Content = "Ongoing process. Please wait ...";
            Label_Data_Hide.Width = 1345;
            Data_Init_Menu();
        }

        private void Manage_Folder()
        {
            Str_Folder_Save = "C:\\WayFair_CG_Save";
            Str_Folder_Pdf = "C:\\WayFair_CG_Pdf";
            if (!Directory.Exists(Str_Folder_Save))
            {
                Directory.CreateDirectory(Str_Folder_Save);
            }
            if (!Directory.Exists(Str_Folder_Pdf))
            {
                Directory.CreateDirectory(Str_Folder_Pdf);
            }
            TextBox_Location_Folder.Text = Str_Folder_Pdf;
            TextBox_Location_Folder_Sample.Text = Str_Folder_Pdf + "\\location_PdfFiles_Merge_date1_HourMinute" + System.Environment.NewLine + System.Environment.NewLine + "Example: " + Str_Folder_Pdf + "\\TEXAS_PdfFiles_Merge_10112022_1212.pdf ";
        }

        /// <summary>
        /// 
        /// </summary>
        private void Data_Init_Menu()
        {
            LabelHeader.Content = "GENERATE PDF FILES - IMPORT AND SPLIT PREPARATION";
            Btn_Menu_Display.Visibility = System.Windows.Visibility.Hidden;
            TabControl_Data.SelectedIndex = 1;
        }

        /// <summary>
        /// 
        /// </summary>
        private void Data_Init_Create_Pdf()
        {
            try
            {
                LabelHeader.Content = "Generate PDF file fom Excel File";
                Label_Data_Message.Content = "--- Select the File having the PDF file ...";
                GroupBox_Progress.Visibility = System.Windows.Visibility.Hidden;
                Label_Data_Hide.Visibility = System.Windows.Visibility.Hidden;
                GroupBox_Progress.Visibility = System.Windows.Visibility.Hidden;
                List_Data_Pdf_File.Clear();
                List_Data_Excel_File.Clear();
                DataGrid_Generate_Pdf.ItemsSource = null;
                DataGrid_Data_Pdf.ItemsSource = null;
                DataGrid_Data_Excel.ItemsSource = null;
                // ----------------------------------------------------------------
                string root = @Str_Folder_Save;
                // If directory does not exist, create it. 
                if (!Directory.Exists(root))
                {
                    Directory.CreateDirectory(root);
                }
                else
                {
                    string[] files = Directory.GetFiles(root);
                    foreach (string file in files)
                    {
                        if (file != "")
                        {
                            File.Delete(file);
                        }
                    }
                }
                // ---------------------------------------------
                TabControl_Data.SelectedIndex = 3;
            }
            catch (Exception e)
            {
                Console.WriteLine("Exception: " + e.Message);
            }
        }


        /// <summary>
        /// 
        /// </summary>
        private void MainWindow_Closing(object sender, CancelEventArgs e)
        {
            // this.Owner.Visibility = Visibility.Visible;
        }

        /// <summary>
        /// 
        /// </summary>
        void DataGrid_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = (e.Row.GetIndex() + 1).ToString();
        }

        /// <summary>
        /// 
        /// </summary>
        private void Btn_Data_Close_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Did you want to close the Software ?", "Generate File", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                App.Current.Shutdown();
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private void Btn_Generate_Close_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Did you want to close the Software ?", "Generate File", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                App.Current.Shutdown();
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private void Btn_Data_Display_Click(object sender, RoutedEventArgs e)
        {
            Display_Data_Global();
        }

        /// <summary>
        /// 
        /// </summary>
        private void Btn_Data_Reset_Click(object sender, RoutedEventArgs e)
        {
            Data_Init_Create_Pdf();
        }

        /// <summary>
        /// 
        /// </summary>
        private void Display_Data()
        {
            try
            {
                // ----------------------------------------------------------------------
                Label_Data_Hide.Visibility = System.Windows.Visibility.Hidden;
                GroupBox_Progress.Visibility = System.Windows.Visibility.Hidden;

                if (List_Data_Pdf_File.Count() == 0)
                {
                    MessageBox.Show("You have to select the PDF file", "PDF File");
                    Label_Data_Message.Content = "--- Load PDF files ...";
                    return;
                }
                if (List_Data_Pdf_File.Count() == 0)
                {
                    MessageBox.Show("You have to select the Excel file", "Excel File");
                    Label_Data_Message.Content = "--- Load Excel files ...";
                    return;
                }
                // -------------------------------------------------
                List_Pdf_File.Clear();
                Label_Data_Hide.Visibility = System.Windows.Visibility.Visible;
                Label_Data_Hide.Content = " ... Ongoing process. Please wait ...";
                TabControl_Data.SelectedIndex = 1;
                MessageBox.Show("Contnue", "Excel File");
                // -----------------------------------------------------------------
                Import_Excel_File_From_Acumatica();
                Split_Pdf_File();

                GroupBox_Progress.Visibility = System.Windows.Visibility.Hidden;
                // -----------------------------------------------
                TabControl_Data.SelectedIndex = 2;
            }
            catch (Exception e)
            {
                Console.WriteLine("Exception: " + e.Message);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private void Btn_Generate_Return_Click(object sender, RoutedEventArgs e)
        {
            Data_Init_Create_Pdf();
        }


        /// <summary>
        /// 
        /// </summary>
        private void Proggress_Bar(int Nbre, string Str_Title)
        {
            try
            {
                GroupBox_Progress.Visibility = System.Windows.Visibility.Visible;
                int Percentage = 0;
                double nbr1 = (((Nbre * 100) / NbreTotal));
                Percentage = (int)Math.Ceiling(nbr1);
                if (Percentage >= 100) { Percentage = 100; };
                System.Threading.Thread.Sleep(2);
                LabelMessage2.Content = Str_Title;
                LabelCountMessage2.Content = Nbre.ToString() + " / " + NbreTotal.ToString();
                ProgressBarMessage2_Label.Content = Percentage.ToString() + " %";
                ProgressBarMessage2_ProgressBar.Value = Percentage;
                System.Windows.Forms.Application.DoEvents();
            }
            catch (Exception e)
            {
                Console.WriteLine("Exception: " + e.Message);
            }
        }


        /// <summary>
        /// 
        /// </summary>
        private void Btn_Generate_Pdf_Click(object sender, RoutedEventArgs e)
        {

        }


        /// <summary>
        /// 
        /// </summary>
        private void Btn_Menu_Close_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Did you want to close the Software ?", "Generate File", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                App.Current.Shutdown();
            }
        }


        /// <summary>
        /// 
        /// </summary>
        private void Btn_Menu_Create_Click(object sender, RoutedEventArgs e)
        {
            Data_Init_Create_Pdf();

        }


        /// <summary>
        /// 
        /// </summary>
        private void Btn_Data_Pdf_Create_Click(object sender, RoutedEventArgs e)
        {
            //try
            //{
            OpenFileDialog openfile = new OpenFileDialog();
            openfile.DefaultExt = ".pdf";
            openfile.Filter = "(.pdf)|*.pdf";
            var browsefile = openfile.ShowDialog();
            string Str_Tempon_file = "";
            if (browsefile == true)
            {
                Str_Tempon_file = openfile.FileName;
                if (Str_Tempon_file != "")
                {
                    // ------------------------------------------
                    // --------- Check if the Filelready selected
                    bool exist = true;
                    Mouse.OverrideCursor = Cursors.Wait;
                    var Query10 = from List10 in List_Data_Pdf_File
                                  where List10.Name01.ToLower() == Str_Tempon_file.ToLower()
                                  select List10;
                    foreach (var List20 in Query10)
                    {
                        exist = false;
                        break;
                    }
                    if (exist == false)
                    {
                        Mouse.OverrideCursor = null;
                        MessageBox.Show("... This file was already selected ....", "Pdf File");
                        return;
                    }
                    // Generate file
                    // -+---------------------------------------------------
                    List_Data_Pdf_File.Add(new DataGrid_Init_10()
                    {
                        Name01 = Str_Tempon_file,
                        Name02 = "",
                        Name03 = "",
                        Name04 = "",
                        Name05 = "",
                        Name06 = "",
                        Name07 = "",
                        Name08 = "",
                        Name09 = "",
                        Name10 = "",
                        NameId = 0,
                    });
                    DataGrid_Data_Pdf.ItemsSource = null;
                    DataGrid_Data_Pdf.Visibility = System.Windows.Visibility.Visible;
                    var NewList = List_Data_Pdf_File.OrderBy(x => x.Name01).ToList();
                    DataGrid_Data_Pdf.ItemsSource = NewList;
                    DataGrid_Data_Pdf.IsReadOnly = true;
                    Mouse.OverrideCursor = null;
                }
                else
                {
                }
            }
            Mouse.OverrideCursor = null;
        }

        /// <summary>
        /// 
        /// </summary>
        private void Btn_Data_Return_Click(object sender, RoutedEventArgs e)
        {
            TabControl_Data.SelectedIndex = 1;
        }

        /// <summary>
        /// 
        /// </summary>
        private void Btn_Data_Excel_Create_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openfile = new OpenFileDialog();
            openfile.DefaultExt = ".xlsx";
            openfile.Filter = "(.xlsx)|*.xlsx";
            var browsefile = openfile.ShowDialog();
            string Str_Tempon_file = "";
            Str_Global_File_Excel = "";
            if (browsefile == true)
            {
                Str_Tempon_file = openfile.FileName;
                if (Str_Tempon_file != "")
                {
                    Str_Global_File_Excel = Str_Tempon_file;
                    List_Tempon.Clear();
                    Mouse.OverrideCursor = Cursors.Wait;
                    Str_Data = "";
                    string strfilename = openfile.FileName;
                    int nb = 0;
                    string value10 = "";
                    //dtExcel.TableName = "MyExcelData";
                    string SourceConstr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strfilename + ";Extended Properties= 'Excel 8.0;HDR=Yes;IMEX=1'";
                    Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel.Workbook wb = excel.Workbooks.Open(strfilename);
                    // load Get worksheet names
                    foreach (Microsoft.Office.Interop.Excel.Worksheet sh in wb.Worksheets)
                    {
                        nb = nb + 1;
                        value10 = sh.Name;
                        List_Tempon.Add(new Combo_InitList() { Id = 1, Name = sh.Name });
                    }
                    ComboBox_Super_WorkSheet.IsEnabled = false;
                    if (nb == 1)
                    {
                        // ------------------------------------------
                        // --------- Check if the Filelready selected
                        bool exist = true;
                        Mouse.OverrideCursor = Cursors.Wait;
                        var Query10 = from List10 in List_Data_Excel_File
                                      where List10.Name01.ToLower() == strfilename.ToLower()
                                      select List10;
                        foreach (var List20 in Query10)
                        {
                            exist = false;
                            break;
                        }
                        if (exist == false)
                        {
                            Mouse.OverrideCursor = null;
                            MessageBox.Show(".... This file was already selected .....", "Pdf File");
                            return;
                        }
                        //    // Generate file
                        //    // -+---------------------------------------------------
                        List_Data_Excel_File.Add(new DataGrid_Init_10()
                        {
                            Name01 = strfilename,
                            Name02 = value10,
                            Name03 = "",
                            Name04 = "",
                            Name05 = "",
                            Name06 = "",
                            Name07 = "",
                            Name08 = "",
                            Name09 = "",
                            Name10 = "",
                            NameId = 0,
                        });
                        DataGrid_Data_Excel.ItemsSource = null;
                        DataGrid_Data_Excel.Visibility = System.Windows.Visibility.Visible;
                        var NewList = List_Data_Excel_File.OrderBy(x => x.Name01).ToList();
                        DataGrid_Data_Excel.ItemsSource = NewList;
                        DataGrid_Data_Excel.IsReadOnly = true;
                        Mouse.OverrideCursor = null;
                        TabControl_Data.SelectedIndex = 3;
                        return;
                    }
                    else
                    {
                        ComboBox_Super_WorkSheet.IsEnabled = true;
                        List_Tempon.Add(new Combo_InitList() { Id = 0, Name = ".. Select data" });
                    }

                    var newList = List_Tempon.OrderBy(x => x.Name).ToList();
                    ComboBox_Super_WorkSheet.ItemsSource = newList;
                    ComboBox_Super_WorkSheet.SelectedIndex = 0; ;
                    TextBox_Super_FileName.Text = strfilename;
                    wb.Close(false);
                    excel.Quit();
                    Mouse.OverrideCursor = null;
                    TabControl_Data.SelectedIndex = 6;
                    TabControl_Pdf.SelectedIndex = 0;
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private void Display_Data_Global()
        {
            try
            {
                // ----------------------------------------------------------------------
                Label_Data_Hide.Visibility = System.Windows.Visibility.Hidden;
                GroupBox_Progress.Visibility = System.Windows.Visibility.Hidden;
                List_File_Created.Clear();
                if (List_Data_Pdf_File.Count() == 0)
                {
                    MessageBox.Show("You have to select the PDF file", "PDF File");
                    Label_Data_Message.Content = "--- Load PDF files ...";
                    return;
                }
                if (List_Data_Excel_File.Count() == 0)
                {
                    MessageBox.Show("You have to select the Excel file", "Excel File");
                    Label_Data_Message.Content = "--- Load Excel files ...";
                    return;
                }
                // -------------------------------------------------
                Int_Combo_Template = 0;
                List_Pdf_File.Clear();
                Label_Data_Hide.Visibility = System.Windows.Visibility.Visible;
                Label_Data_Hide.Content = " ... Ongoing process. Please wait ...";
                TabControl_Data.SelectedIndex = 3;
                MessageBox.Show("Contnue", "Generate Files");
                LabelHeader.Content = "GENERATE PDF FILES - SPLIT PROCESS";
                // -----------------------------------------------------------------------------------
                Import_Excel_File_From_Acumatica();
                Split_Pdf_File();
                // Split_Pdf_File_New();
                Control_Acumatica_Pdf();
                Update_Data();
                Final_Data();
                GroupBox_Progress.Visibility = System.Windows.Visibility.Hidden;
                // -----------------------------------------------
                Generate_Combo_Operation();
                LabelHeader.Content = "GENERATE PDF FILES - MERGE PROCESS";
                Mouse.OverrideCursor = null;
                TabControl_Data.SelectedIndex = 4;
            }
            catch (Exception e)
            {
                Console.WriteLine("Exception: " + e.Message);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private void Control_Acumatica_Pdf()
        {
            try
            {
                Mouse.OverrideCursor = Cursors.Wait;
                var Query10 = from List10 in Display_Data10
                              orderby List10.Name01
                              select List10;
                foreach (var List20 in Query10)
                {
                    string str = List20.Name01;
                    int T_nbr = GlobalVar.ControlInt(List20.Name02);
                    int Nbr_Pdf = 0;
                    // -----------------------------------------------
                    List<Combo_InitList> List_Work = new List<Combo_InitList>();
                    List_Work = (from List10 in Display_Data10_01
                                 orderby List10.Name01
                                 group List10 by List10.Name01 into groupOrder
                                 select new Combo_InitList()
                                 {
                                     Id = 0,
                                     Name = groupOrder.Key
                                 }).ToList();

                    var Query100 = from List80 in List_Work
                                   where List80.Name == str
                                   select List80;
                    foreach (var List90 in Query100)
                    {
                        Nbr_Pdf = Display_Data10_01.Where(x => x.Name01 == List90.Name).Count();
                    }
                    // -------------------------------------------------
                    List20.Name03 = GlobalVar.Number_With_Dash(Nbr_Pdf);
                    if (T_nbr > 0)
                    {
                        List20.Name04 = GlobalVar.Number_With_Dash(T_nbr - Nbr_Pdf);
                        if (T_nbr != Nbr_Pdf)
                        {
                            List20.Name05 = "Pb";
                            List20.Name07 = "No Label";
                        }
                    }
                }
                Display_Data10_X_ALL.Clear();
                Display_Data10_X_ALL = Display_Data10.ToList();
                Dg_Generate_Product();
                // -----------------------------------------------------------------------------------
                Create_ComboBox_Pdf();
                Mouse.OverrideCursor = null;
            }
            catch (Exception e)
            {
                Console.WriteLine("Exception: " + e.Message);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private void Dg_Generate_Product()
        {
            DataGrid_Generate_Product.ItemsSource = null;
            DataGrid_Generate_Product.Visibility = System.Windows.Visibility.Visible;
            var NewList = Display_Data10.OrderBy(x => x.Name01).ToList();
            DataGrid_Generate_Product.ItemsSource = NewList;
            DataGrid_Generate_Product.IsReadOnly = true;
            // -----
            int Nbre01 = 0;
            int Nbre02 = 0;
            var Query100 = from List80 in Display_Data10
                           select List80;
            foreach (var List90 in Query100)
            {
                Nbre01 = Nbre01 + GlobalVar.ControlInt(List90.Name02);
                Nbre02 = Nbre02 + GlobalVar.ControlInt(List90.Name03);
            }
            Label_Nbr_Template.Content = GlobalVar.Number_With_Dash(Nbre01).ToString();
            Label_Nbr_Pdf.Content = GlobalVar.Number_With_Dash(Nbre02).ToString();
            Label_Nbr_Diff.Content = (Nbre01 - Nbre02).ToString();
        }

        /// <summary>
        /// 
        /// </summary>
        private void Create_ComboBox_Pdf()
        {
            int nbr = 0;
            var Query100 = from List80 in Display_Data10
                           select List80;
            foreach (var List90 in Query100)
            {
                if (List90.Name05 != "")
                {
                    nbr = nbr + 1;
                }
            }
            // ------------------------------------------------
            ComboBox_Data_Template.ItemsSource = null;
            ComboBox_Data_Template.IsEnabled = false;
            List_Combo_Template.Clear();
            if (nbr == 0)
            {
                List_Combo_Template.Add(new Combo_InitList() { Id = 0, Name = "... All items ..." });
            }
            else
            {
                List_Combo_Template.Add(new Combo_InitList() { Id = 0, Name = "... Display all items ..." });
                List_Combo_Template.Add(new Combo_InitList() { Id = 1, Name = "Display Selected items" });
                List_Combo_Template.Add(new Combo_InitList() { Id = 2, Name = "Display No Label Data" });
                ComboBox_Data_Template.IsEnabled = true;
            }
            var newList = List_Combo_Template.OrderBy(x => x.Id).ToList();
            ComboBox_Data_Template.ItemsSource = newList;
            ComboBox_Data_Template.SelectedIndex = 0; ;
        }

        /// <summary>
        ///  
        /// </summary>
        private void Generate_Combo_Operation()
        {
            try
            {
                ComboBox_Data_Operation.ItemsSource = null;
                List_Operation.Clear();
                List_Operation.Add(new Combo_InitList() { Id = 0, Name = "... Display all items ..." });
                List_Operation.Add(new Combo_InitList() { Id = 1, Name = "Display Selected items" });
                List_Operation.Add(new Combo_InitList() { Id = 2, Name = "Display not selected" });
                var newList = List_Operation.OrderBy(x => x.Id).ToList();
                ComboBox_Data_Operation.ItemsSource = newList;
                ComboBox_Data_Operation.SelectedIndex = 0; ;
                // ------------------------------------------------------
                ComboBox_Data_Order.ItemsSource = null;
                List_Order_Pdf.Clear();
                List_Order_Pdf.Add(new Combo_InitList() { Id = 1, Name = "Product SKU" });
                List_Order_Pdf.Add(new Combo_InitList() { Id = 2, Name = "Location" });
                var newList2 = List_Order_Pdf.OrderBy(x => x.Id).ToList();
                ComboBox_Data_Order.ItemsSource = newList2;
                ComboBox_Data_Order.SelectedIndex = 0; ;
            }
            catch (Exception e)
            {
                Console.WriteLine("Exception: " + e.Message);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private void Import_Excel_File_From_Acumatica()
        {
            try
            {
                //  ----------------------------------------------------------------------------------------------
                if (List_Data_Excel_File.Count() == 0)
                {
                    MessageBox.Show("You have to select the Excel file", "Excel File");
                    Label_Data_Message.Content = "--- Load Excel files ...";
                    return;
                }
                Mouse.OverrideCursor = Cursors.Wait;
                List_Excel_File.Clear();
                var Query20 = from List10 in List_Data_Excel_File
                              select List10;
                foreach (var List20 in Query20)
                {
                    Str_Global_File_Excel = List20.Name01;
                    Str_Data = List20.Name02;
                    // DataRow workRow;
                    System.Data.DataTable dtExcel = new System.Data.DataTable();
                    dtExcel.TableName = "MyExcelData";
                    string SourceConstr10 = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + Str_Global_File_Excel + "';Extended Properties= 'Excel 8.0;HDR=Yes;IMEX=1'";
                    using (OleDbConnection con = new OleDbConnection(SourceConstr10))
                    using (OleDbDataAdapter data = new OleDbDataAdapter("Select * from [" + Str_Data + "$]", con))
                    // using (OleDbDataAdapter data = new OleDbDataAdapter("Select * from [" + Str_Data + "$]", con))
                    {
                        data.Fill(dtExcel);
                    }
                    // -------------------------------------------------------------------------------------------
                    List_Excel_File = (from DataRow List10 in dtExcel.Rows
                                       select new DataGrid_Init_10()
                                       {
                                           Name01 = List10[0].ToString(), // Inventory ID
                                           Name02 = List10[1].ToString(), // Qty
                                           Name03 = List10[2].ToString(), // Location
                                           Name04 = List10[3].ToString(), // Location ID
                                           Name05 = "", // List10[3].ToString(), // Description
                                           Name06 = "",
                                           Name07 = "",
                                           Name08 = "",
                                           Name09 = "",
                                           Name10 = "",
                                           NameId = 0
                                       }).ToList();
                }
                // -----------------------------------------------------------------------
                // List_Excel_File
                List_Excel_File.RemoveAll(i => i.Name01 == "");
                // List_Excel_File.RemoveAll(i => i.Name03 == "0");
                // -------------------------------------------------------------------------
                List_Excel_Acumatica_File.Clear();
                List_Excel_Acumatica_File = (from List10 in List_Excel_File
                                             orderby List10.Name01
                                             group List10 by List10.Name01 into groupOrder
                                             select new Combo_InitList()
                                             {
                                                 Id = 0,
                                                 Name = groupOrder.Key
                                             }).ToList();

                var Query0 = from List11 in List_Excel_Acumatica_File
                             orderby List11.Name
                             select List11;
                foreach (var List50 in Query0)
                {
                    int Nbre = 0;
                    var Query28 = from List10 in List_Excel_File
                                  where List10.Name01 == List50.Name
                                  select List10;
                    foreach (var List20 in Query28)
                    {
                        Nbre = Nbre + GlobalVar.ControlInt(List20.Name02);
                    }
                    List50.Id = Nbre;
                }

                Display_Data10.Clear();
                var Query10 = from List10 in List_Excel_Acumatica_File
                              orderby List10.Name
                              select List10;
                foreach (var List20 in Query10)
                {
                    Display_Data10.Add(new DataGrid_Init_10()
                    {
                        Name01 = List20.Name,
                        Name02 = List20.Id.ToString(),
                        Name03 = "",
                        Name04 = "",
                        Name05 = "",
                        Name06 = "",
                        Name07 = "",
                        Name08 = "",
                        Name09 = "",
                        Name10 = "",
                        NameId = List20.Id,
                    });
                }
                // ----------------------------------------------------------------
                Display_Data10_X_ALL.Clear();
                Display_Data10_X_ALL = Display_Data10.ToList();
                Dg_Generate_Product();
                Mouse.OverrideCursor = null;
            }
            catch (Exception e)
            {
                Console.WriteLine("Exception: " + e.Message);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private void Split_Pdf_File()
        {
            try
            {
                // https://www.c-sharpcorner.com/article/splitting-pdf-file-in-c-sharp-using-itextsharp/
                if (List_Data_Pdf_File.Count() == 0)
                {
                    MessageBox.Show("You have to select the Excel file", "Excel File");
                    Label_Data_Message.Content = "--- Load Excel files ...";
                    return;
                }
                Mouse.OverrideCursor = Cursors.Wait;
                GroupBox_Progress.Visibility = System.Windows.Visibility.Hidden;
                string outputPath = @Str_Folder_Save;
                Display_Data10_01.Clear();
                int Nbre_Add = 0;
                int Nbr_items = List_Data_Pdf_File.Count();
                // create PdfFileEditor object
                var Query20 = from List10 in List_Data_Pdf_File
                              select List10;
                foreach (var List20 in Query20)
                {
                    // -------------------------------------------------------
                    Nbre_Add = Nbre_Add + 1;
                    string Str01 = "";
                    if (Nbre_Add < 10)
                    {
                        Str01 = "0" + Nbre_Add.ToString();
                    }
                    else
                    {
                        Str01 = Nbre_Add.ToString();
                    }
                    Str_Global_File_Pdf = List20.Name01;
                    string pdfFilePath = @Str_Global_File_Pdf;
                    int interval = 1;
                    int pageNameSuffix = 0;
                    // Intialize a new PdfReader instance with the contents of the source Pdf file:
                    PdfReader reader = new PdfReader(pdfFilePath);

                    FileInfo file = new FileInfo(pdfFilePath);
                    string pdfFileName = file.Name.Substring(0, file.Name.LastIndexOf(".")) + "-";
                    NbreTotal = reader.NumberOfPages;
                    int NbreLine = 0;
                    for (int pageNumber = 1; pageNumber <= reader.NumberOfPages; pageNumber += interval)
                    {
                        // -------------------------------
                        NbreLine = NbreLine + 1;
                        Proggress_Bar(NbreLine, Str_Global_File_Pdf + "      ---  Pdf Files  " + Nbre_Add.ToString() + " / " + Nbr_items.ToString());
                        // -----------------------------------
                        pageNameSuffix++;
                        string newPdfFileName = string.Format(pdfFileName + Str01 + "{0}", pageNameSuffix);
                        // ---------------------------------------------------------------------------
                        string Str_Fle = newPdfFileName + ".pdf"; ;
                        string Str_Product = "";
                        //// ----------------------------------------------------------------------------
                        // Tratemen done on Disk --- Source document
                        SplitAndSaveInterval(pdfFilePath, outputPath, pageNumber, interval, newPdfFileName);
                        string str_20 = outputPath + "\\" + newPdfFileName + ".pdf";
                        PdfReader reader1 = new PdfReader(@str_20);
                        int intPageNum = reader1.NumberOfPages;
                        string[] words;
                        string line;
                        for (int i = 1; i <= intPageNum; i++)
                        {
                            string text = PdfTextExtractor.GetTextFromPage(reader1, i, new LocationTextExtractionStrategy());
                            words = text.Split('\n');
                            for (int j = 0, len = words.Length; j < len; j++)
                            {
                                line = Encoding.UTF8.GetString(Encoding.UTF8.GetBytes(words[j]));
                                if (line.ToLower().Contains("part"))
                                {
                                    Str_Product = line.Replace("PART", "");
                                    Str_Product = Str_Product.Replace("#", "");
                                    Str_Product = Str_Product.Replace(":", "");
                                    Str_Product = Str_Product.TrimStart();
                                }
                            }
                        }
                        string Str_05 = "X";
                        CheckBox_Data.IsChecked = true;
                        // -+---------------------------------------------------
                        Display_Data10_01.Add(new DataGrid_Init_10()
                        {
                            Name01 = Str_Product,
                            Name02 = Str_Fle,
                            Name03 = "",
                            Name04 = "",
                            Name05 = Str_05,
                            Name06 = "",
                            Name07 = "",
                            Name08 = "",
                            Name09 = "",
                            Name10 = "",
                            NameId = 0,
                        });
                        // ----------------------------------------------------------
                    }
                }
                // ---------------------------------------------
                var Query0 = from List11 in Display_Data10
                             orderby List11.Name01
                             select List11;
                foreach (var List50 in Query0)
                {
                    int Nbre_Data = List50.NameId;
                    int Nbre = 0;
                    var Query30 = from List10 in Display_Data10_01
                                  where List10.Name01.TrimStart().Replace(" ", "").ToLower() == List50.Name01.TrimStart().Replace(" ", "").ToLower()
                                  select List10;
                    foreach (var List20 in Query30)
                    {
                        string str_18 = outputPath + "\\" + List20.Name02;
                        string str_28 = outputPath + "\\" + "New_" + List20.Name02;
                        //   Add_Text(str_18, str_28, List20.Name01);

                        if (Nbre <= Nbre_Data)
                        {
                            var Query80 = from List80 in List_Excel_File
                                          where List80.Name05 == ""
                                          & List80.Name01.TrimStart().Replace(" ", "").ToLower() == List20.Name01.TrimStart().Replace(" ", "").ToLower()
                                          orderby List80.Name01
                                          select List80;
                            foreach (var List88 in Query80)
                            {
                                List88.Name05 = "ok";
                                List20.Name03 = List88.Name03;
                                break;
                            }

                            List20.Name04 = "Ok";
                        }
                        Nbre = Nbre + 1;
                    }
                }
                List<DataGrid_Init_10> T_Temp = new List<DataGrid_Init_10>();
                T_Temp = Display_Data10_01.OrderBy(x => x.Name01).ToList();
                Display_Data10_01.Clear();
                Display_Data10_01 = T_Temp.ToList();

                // ---------------------------------------------
                DataGrid_Result();
                Display_Data10_All.Clear();
                Display_Data10_All = Display_Data10_01.ToList();
                Mouse.OverrideCursor = null;
            }
            catch (Exception e)
            {
                Console.WriteLine("Exception: " + e.Message);
            }
        }


        /// <summary>
        /// 
        /// </summary>
        private void SplitAndSaveInterval(string pdfFilePath, string outputPath, int startPage, int interval, string pdfFileName)
        {
            try
            {
                using (PdfReader reader = new PdfReader(pdfFilePath))
                {
                    Document document = new Document();
                    PdfCopy copy = new PdfCopy(document, new FileStream(outputPath + "\\" + pdfFileName + ".pdf", FileMode.Create));
                    document.Open();

                    for (int pagenumber = startPage; pagenumber < (startPage + interval); pagenumber++)
                    {
                        if (reader.NumberOfPages >= pagenumber)
                        {
                            // List_Pdf_File
                            copy.AddPage(copy.GetImportedPage(reader, pagenumber));
                        }
                        else
                        {
                            break;
                        }
                    }

                    document.Close();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Exception: " + e.Message);
            }
            finally
            {
                Console.WriteLine("Executing finally block.");
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private void Btn_Super_Return_Click(object sender, RoutedEventArgs e)
        {
            TabControl_Data.SelectedIndex = 3;
        }

        /// <summary>
        /// 
        /// </summary>
        private void ComboBox_Data_Operation_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var row = sender as ComboBox;
            if (row != null)
            {
                int Selected_ID = 0;

                Combo_InitList Value1 = (Combo_InitList)ComboBox_Data_Operation.SelectedItem;
                if (Value1 != null)
                {
                    Display_Data10_01.Clear();
                    Display_Data10_01 = Display_Data10_All.ToList();
                    Selected_ID = Value1.Id;
                    switch (Selected_ID)
                    {
                        case 0:
                            break;
                        case 1:
                            Display_Data10_01.RemoveAll(i => i.Name04 == "");
                            break;
                        case 2:
                            Display_Data10_01.RemoveAll(i => i.Name04 != "");
                            break;
                    }
                    // ------------------------------------------------------------------------
                    List<DataGrid_Init_10> T_Temp = new List<DataGrid_Init_10>();
                    T_Temp = Display_Data10_01.OrderBy(x => x.Name01).ToList();
                    Display_Data10_01.Clear();
                    Display_Data10_01 = T_Temp.ToList();

                    DataGrid_Result();
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private void Add_Text(string oldFile, string newFile, string TextToAdd)
        {
            // open the reader
            PdfReader reader = new PdfReader(oldFile);
            iTextSharp.text.Rectangle size = reader.GetPageSizeWithRotation(1);
            Document document = new Document(size);
            // Document document = new Document();
            // open the writer
            FileStream fs = new FileStream(newFile, FileMode.Create, FileAccess.Write);
            // FileStream fs1 = new FileStream(oldFile, FileMode.Create, FileAccess.Write);
            PdfWriter writer = PdfWriter.GetInstance(document, fs);
            document.Open();
            // the pdf content
            PdfContentByte cb = writer.DirectContent;
            // select the font properties
            BaseFont bf = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetColorFill(BaseColor.DARK_GRAY);
            cb.SetFontAndSize(bf, 8);
            // write the text in the pdf content
            cb.BeginText();
            //Create the first block with different font style
            string text = TextToAdd; // "Some random blablablabla...";
                                     // put the alignment and coordinates here
                                     //float x, y;
                                     //int ts = 18;
                                     //x = (40 + ts + 440) / 2;
                                     //y = 40 + ts + 24;
                                     //  cb.ShowTextAligned(500, text, x, y, 0);
                                     //  cb.ShowTextAligned(1, text, 520, 640, 0);
                                     // cb.ShowTextAligned(1, text, 520, 640, 0);
                                     //Chunk Test10 = new Chunk(text, FontFactory.GetFont("dax-black"));
                                     //Test10.SetUnderline(0.5f, -1.5f);

            //cb.ShowTextAligned(1, Test10.ToString(), 520, 700, 0);
            // cb.ShowTextAligned(1, text, 520, 640, 0);
            // ---------------------------------------------------------------------------------------
            cb.EndText();
            cb.BeginText();
            //  text = "Other random blabla...";
            // put the alignment and coordinates here
            // cb.ShowTextAligned(2, text, 100, 200, 0);
            // cb.ShowTextAligned(400, text, 200, 200, 1);
            cb.ShowTextAligned(400, text, 175, 185, 1);

            // cb.ShowTextAligned(450, text, 0,450, 0);
            cb.EndText();
            // create the new page and add it to the pdf
            PdfImportedPage page = writer.GetImportedPage(reader, 1);
            cb.AddTemplate(page, 0, 0);
            // close the streams and voilá the file should be changed :)
            document.Close();
            fs.Close();
            writer.Close();
            reader.Close();
            reader.Dispose();
            // fs1.Close();
            // --------------------------------------------------------
            if (File.Exists(oldFile))
            {

                File.Delete(oldFile);
            }

            // Create a FileInfo  
            System.IO.FileInfo fi = new System.IO.FileInfo(newFile);
            // Check if file is there  
            if (fi.Exists)
            {
                // Move file with a new name. Hence renamed.  
                fi.MoveTo(@oldFile);
            }

        }


        /// <summary>
        /// 
        /// </summary>
        private void Btn_Menu_Display_Click(object sender, RoutedEventArgs e)
        {
            Data_Generate_Old();


        }

        public static string ExtractTextFromPdf(string path)
        {
            using (PdfReader reader = new PdfReader(path))
            {
                StringBuilder text = new StringBuilder();

                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    text.Append(PdfTextExtractor.GetTextFromPage(reader, i));
                }

                return text.ToString();
            }
        }


        /// <summary>
        /// 
        /// </summary>
        private void Update_Data()
        {
            try
            {
                Mouse.OverrideCursor = Cursors.Wait;
                int NbreLine = 0;
                NbreTotal = List_Excel_File.Count();
                GroupBox_Progress.Visibility = System.Windows.Visibility.Hidden;

                var Query100 = from List100 in List_Excel_File
                               orderby List100.Name01
                               select List100;
                foreach (var List105 in Query100)
                {
                    NbreLine = NbreLine + 1;
                    Proggress_Bar(NbreLine, "Manage Location");

                    int Nbr = GlobalVar.ControlInt(List105.Name02);
                    string Str_Location = List105.Name03;
                    int Count_01 = 0;
                    var Query80 = from List80 in Display_Data10_01
                                  where List80.Name01.Replace(" ", "").ToLower() == List105.Name01.Replace(" ", "").ToLower()
                                  select List80;
                    foreach (var List88 in Query80)
                    {
                        List88.Name03 = Str_Location;
                        Count_01 = Count_01 + 1;
                        if (Count_01 > Nbr)
                        {
                            return;
                        }
                    }
                }
                // -------------------------------------------
                List<DataGrid_Init_10> T_Temp = new List<DataGrid_Init_10>();
                T_Temp = Display_Data10_01.OrderBy(x => x.Name01).ToList();
                Display_Data10_01.Clear();
                Display_Data10_01 = T_Temp.ToList();
                DataGrid_Result();
                // ------------------------------------------------------------
                NbreLine = 0;
                NbreTotal = Display_Data10.Count();

                var Query8 = from List80 in Display_Data10
                             select List80;
                foreach (var List88 in Query8)
                {
                    NbreLine = NbreLine + 1;
                    Proggress_Bar(NbreLine, "Update Location on Template File");
                    string Str_99 = List88.Name01;
                    string Str_loc = "";
                    var Query1 = from List1 in List_Excel_File
                                 where List1.Name01.Replace(" ", "").ToLower() == Str_99.Replace(" ", "").ToLower()
                                 select List1;
                    foreach (var List10 in Query1)
                    {
                        if (Str_loc != "")
                        {
                            Str_loc = Str_loc + ", " + List10.Name03;
                        }
                        else
                        {
                            Str_loc = Str_loc + List10.Name03;
                        }
                    }
                    List88.Name06 = Str_loc;
                }
                Dg_Generate_Product();

                // -----------------------------------------------------------
                NbreLine = 0;
                NbreTotal = Display_Data10_01.Count();
                GroupBox_Progress.Visibility = System.Windows.Visibility.Hidden;
                string outputPath = @Str_Folder_Save;
                var Query800 = from List80 in Display_Data10_01
                               select List80;
                foreach (var List88 in Query800)
                {
                    NbreLine = NbreLine + 1;
                    Proggress_Bar(NbreLine, "Update Pdf by including Location".ToUpper());
                    string T_Product = List88.Name01;
                    string T_Pdf = List88.Name02;
                    string T_Location = List88.Name03;
                    // --------------------------------------
                    string str_20 = outputPath + "\\" + T_Pdf;
                    PdfReader reader1 = new PdfReader(@str_20);
                    int intPageNum = reader1.NumberOfPages;
                    string[] words;
                    string line;
                    for (int i = 1; i <= intPageNum; i++)
                    {
                        string text = PdfTextExtractor.GetTextFromPage(reader1, i, new LocationTextExtractionStrategy());
                        words = text.Split('\n');
                        int ts = words.Length;
                        for (int j = 0, len = words.Length; j < len; j++)
                        {
                            line = Encoding.UTF8.GetString(Encoding.UTF8.GetBytes(words[j]));
                            string ttt = line.ToString();
                            if (line.ToLower().Contains("part"))
                            {
                                int aaa = i;
                                string Str_Product = line.Replace("PART", "");
                                Str_Product = Str_Product.Replace("#", "");
                                Str_Product = Str_Product.Replace(":", "");
                                Str_Product = Str_Product.TrimStart();
                                // --------------------------------------------------------
                                reader1.Dispose();
                                reader1.Close();
                                //if (T_Product.ToLower().Replace(" ", "") == T_Pdf.ToLower().Replace(" ", ""))
                                //{
                                string str_28 = outputPath + "\\" + "New_" + T_Pdf;
                                string Str_Mew_Data = Str_Product;
                                // -------------------------------------------------------------------------
                                Chunk chunk = new Chunk(T_Location, FontFactory.GetFont("dax-black"));

                                chunk.SetUnderline(0.5f, -1.5f);
                                // Add_Text(str_20, str_28, T_Location);
                                Add_Text(str_20, str_28, chunk.ToString());
                            }
                        }
                    }
                }
                Mouse.OverrideCursor = null;
            }
            catch (Exception e)
            {
                Console.WriteLine("Exception: " + e.Message);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private void DataGrid_Data_Pdf_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                var row = sender as DataGridRow;
                if (row != null)
                {
                    var item = row.DataContext as DataGrid_Init_10;
                    if (item != null)
                    {
                        string Str_01 = "";
                        foreach (var List10 in List_Data_Pdf_File)
                        {
                            if (List10.Name01 == item.Name01)
                            {
                                if (MessageBox.Show(List10.Name01 + System.Environment.NewLine + System.Environment.NewLine + "Did you want to delete this file ?", "Management of Pdf File".ToUpper(), MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                                {
                                    Str_01 = List10.Name01;
                                    break;
                                }
                            }
                        }
                        // --------------------------------------------------------------------------------
                        if (Str_01 != "")
                        {
                            if (MessageBox.Show("Did you want to delete this file ?", "Management of Pdf File".ToUpper(), MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                            {
                                List_Data_Pdf_File.RemoveAll(a => a.Name01 == Str_01);
                            }
                        }
                        DataGrid_Data_Pdf.ItemsSource = null;
                        DataGrid_Data_Pdf.Visibility = System.Windows.Visibility.Visible;
                        var NewList = List_Data_Pdf_File.OrderBy(x => x.Name01).ToList();
                        DataGrid_Data_Pdf.ItemsSource = NewList;
                        DataGrid_Data_Pdf.IsReadOnly = true;
                        Mouse.OverrideCursor = null;
                    }
                }
            }
            catch (System.Exception excep)
            {
                MessageBox.Show(excep.Message);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private void DataGrid_Data_Excel_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                var row = sender as DataGridRow;
                if (row != null)
                {
                    var item = row.DataContext as DataGrid_Init_10;
                    if (item != null)
                    {
                        string Str_01 = "";
                        foreach (var List10 in List_Data_Excel_File)
                        {
                            if (List10.Name01 == item.Name01)
                            {
                                if (MessageBox.Show(List10.Name01 + System.Environment.NewLine + System.Environment.NewLine + "Did you want to delete this file ?", "Management of Excel File".ToUpper(), MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                                {
                                    Str_01 = List10.Name01;
                                    break;
                                }
                            }
                        }
                        // --------------------------------------------------------------------------------
                        if (Str_01 != "")
                        {
                            if (MessageBox.Show("Did you want to delete this file ?", "Management of Excel File", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                            {
                                List_Data_Excel_File.RemoveAll(a => a.Name01 == Str_01);
                            }
                        }
                        DataGrid_Data_Excel.ItemsSource = null;
                        DataGrid_Data_Excel.Visibility = System.Windows.Visibility.Visible;
                        var NewList = List_Data_Excel_File.OrderBy(x => x.Name01).ToList();
                        DataGrid_Data_Excel.ItemsSource = NewList;
                        DataGrid_Data_Excel.IsReadOnly = true;
                        Mouse.OverrideCursor = null;
                    }
                }
            }
            catch (System.Exception excep)
            {
                MessageBox.Show(excep.Message);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private void ComboBox_Data_Order_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            List<DataGrid_Init_10> Display_Temp = new List<DataGrid_Init_10>();
            Display_Temp = Display_Data10_01.ToList();
            var row = sender as ComboBox;
            if (row != null)
            {
                int Selected_ID = 0;
                Combo_InitList Value1 = (Combo_InitList)ComboBox_Data_Order.SelectedItem;
                if (Value1 != null)
                {
                    Selected_ID = Value1.Id;
                    switch (Selected_ID)
                    {
                        case 1:
                            Display_Temp = Display_Data10_01.OrderBy(x => x.Name01).ToList();
                            break;
                        case 2:
                            Display_Temp = Display_Data10_01.OrderBy(x => x.Name03).ToList();
                            break;
                    }
                    Display_Data10_01.Clear();
                    Display_Data10_01 = Display_Temp.ToList();
                    DataGrid_Result();
                }
            }
        }


        /// <summary>
        /// 
        /// </summary>
        private void ComboBox_Data_Template_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            List<DataGrid_Init_10> Display_Temp = new List<DataGrid_Init_10>();
            Display_Temp = Display_Data10_X_ALL.ToList();
            Int_Combo_Template = 0;
            // --------------------------------------------------------------------------------
            var row = sender as ComboBox;
            if (row != null)
            {
                int Selected_ID = 0;
                Combo_InitList Value1 = (Combo_InitList)ComboBox_Data_Template.SelectedItem;
                if (Value1 != null)
                {
                    Selected_ID = Value1.Id;
                    switch (Selected_ID)
                    {
                        case 1:
                            Int_Combo_Template = 1;
                            Display_Temp.RemoveAll(i => i.Name05 != "");
                            break;
                        case 2:
                            Int_Combo_Template = 2;
                            Display_Temp.RemoveAll(i => i.Name05 == "");
                            break;
                    }
                }
            }
            // --------------------------------------------------------------------------------
            Display_Data10.Clear();
            Display_Data10 = Display_Temp.ToList();
            Dg_Generate_Product();
        }

        /// <summary>
        /// 
        /// </summary>
        private void Brn_Result_Excel_Template_Click(object sender, RoutedEventArgs e)
        {
            var desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            DateTime time = DateTime.Now;
            string format = "MMddyyyy";
            string date1 = time.ToString(format);
            string Str_Infos = "";
            switch (Int_Combo_Template)
            {
                case 1:
                    Str_Infos = "All_Data_with_Labels";
                    break;
                case 2:
                    Str_Infos = "All_Data_With_No_Labels";
                    break;
                default:
                    Str_Infos = "All_Data";
                    break;
            }
            String HourMinute = DateTime.Now.ToString("HHmm");
            string fileName = "";
            string Str_Titre = "";
            fileName = "TemplateData_" + Str_Infos + "_Print_" + date1 + "_" + HourMinute + ".xlsx";
            var Location = System.IO.Path.Combine(desktopFolder, fileName);
            // ---------------------------------------------------
            string filename1 = Location;
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
            dlg.DefaultExt = Location;
            dlg.FileName = fileName;
            dlg.Filter = "Excel workbook (.xlsx)|*.xlsx";
            bool Suite10 = false;
            if (dlg.ShowDialog() == true)
            {
                filename1 = dlg.FileName;
                Suite10 = true;
            }
            if (Suite10 == true)
            {
                Mouse.OverrideCursor = Cursors.Wait;
                List<DataGrid_Init_10> List_XXX_Temp = new List<DataGrid_Init_10>();
                List_XXX_Temp = Display_Data10.ToList();
                DataTable Dt_Print = new DataTable();
                Dt_Print = GlobalVar.Convert_ListofObject_To_DataTable(List_XXX_Temp);
                Dt_Print.Columns.Remove("Name05");
                Dt_Print.Columns.Remove("Name08");
                Dt_Print.Columns.Remove("Name09");
                Dt_Print.Columns.Remove("Name10");
                Dt_Print.Columns.Remove("NameId");
                DataTable Dt_Print1 = new DataTable();
                Dt_Print1.Columns.Add(new DataColumn("Name01", Type.GetType("System.String")));
                Dt_Print1.Columns.Add(new DataColumn("Name02", Type.GetType("System.String")));
                Dt_Print1.Columns.Add(new DataColumn("Name03", Type.GetType("System.String")));
                Dt_Print1.Columns.Add(new DataColumn("Name04", Type.GetType("System.String")));
                Dt_Print1.Columns.Add(new DataColumn("Name05", Type.GetType("System.String")));
                Dt_Print1.Columns.Add(new DataColumn("Name06", Type.GetType("System.String")));

                DataRow workRow;
                workRow = Dt_Print1.NewRow();
                workRow["Name01"] = "Product SKU";
                workRow["Name02"] = "Qty Temp";
                workRow["Name03"] = "Qty Label";
                workRow["Name04"] = "Difference";
                workRow["Name05"] = "Location";
                workRow["Name06"] = "Observation";

                Dt_Print1.Rows.Add(workRow);
                foreach (DataRow row in Dt_Print.Rows)
                {
                    workRow = Dt_Print1.NewRow();
                    workRow["Name01"] = row["Name01"].ToString();
                    workRow["Name02"] = row["Name02"].ToString();
                    workRow["Name03"] = row["Name03"].ToString();
                    workRow["Name04"] = row["Name04"].ToString();
                    workRow["Name05"] = row["Name06"].ToString();
                    workRow["Name06"] = row["Name07"].ToString();
                    Dt_Print1.Rows.Add(workRow);
                }
                // -------------------------------------------------------------------------
                GlobalVar.ExportToExcel(Dt_Print1, filename1, 1, 1, Dt_Print1.Rows.Count, Dt_Print1.Columns.Count, Str_Titre);
                Mouse.OverrideCursor = null;
                MessageBox.Show(".... Export complete ....", "Data Template");
            }
        }




        /// <summary>
        /// 
        /// </summary>
        private void CheckBox_Data_Checked(object sender, RoutedEventArgs e)
        {
            if (CheckBox_Data.IsChecked == true)
            {
                Display_Data10_01.Where(w => w.Name01 != "").ToList().ForEach(a => a.Name05 = "X");
            }
            else
            {
                Display_Data10_01.Where(w => w.Name01 != "").ToList().ForEach(a => a.Name05 = "");
            }
            DataGrid_Result();
        }

        /// <summary>
        /// 
        /// </summary>
        private void CheckBox_Data_Unchecked(object sender, RoutedEventArgs e)
        {
            if (CheckBox_Data.IsChecked == false)
            {
                Display_Data10_01.Where(w => w.Name01 != "").ToList().ForEach(a => a.Name05 = "");
            }
            else
            {
                Display_Data10_01.Where(w => w.Name01 != "").ToList().ForEach(a => a.Name05 = "X");
            }
            DataGrid_Result();
        }




        /// <summary>
        /// 
        /// </summary>
        private void DataGrid_Generate_Pdf_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                var row = sender as DataGridRow;
                if (row != null)
                {

                    var item = row.DataContext as DataGrid_Init_10;
                    if (item != null)
                    {
                        string Str_Data = item.Name02;
                        foreach (var List10 in Display_Data10_01)
                        {
                            if (List10.Name02 == Str_Data)
                            {
                                if (List10.Name02 != "")
                                {
                                    if (List10.Name05 == "")
                                    {
                                        List10.Name05 = "X";
                                    }
                                    else
                                    {
                                        List10.Name05 = "";
                                    }
                                }
                                break;
                            }
                        }
                        DataGrid_Result();
                    }
                }
            }
            catch (System.Exception excep)
            {
                MessageBox.Show(excep.Message);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private void Btn_Generate_Pdf_Click_1(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Did you want to Generate the new Pdf file ?", "Data Management", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                CityName = "";
                string value1 = Microsoft.VisualBasic.Interaction.InputBox("Enter the Location where the orders will be ship",
                                           "Location Nane : ", "", -1, -1);
                if (value1 != "")
                {
                    CityName = value1.Replace(" ", "").Replace("-", "").Replace("_", "").ToUpper();
                    //{
                    // --------------------------------------------------------
                    Mouse.OverrideCursor = Cursors.Wait;
                    string root = @Str_Folder_Pdf;
                    string Text_Dir = @Str_Folder_Pdf + "\\" + "Text_File";
                    string Text_root = @Text_Dir;
                    // If directory does not exist, create it. 
                    if (!Directory.Exists(root))
                    {
                        Directory.CreateDirectory(root);
                    }
                    if (!Directory.Exists(Text_root))
                    {
                        Directory.CreateDirectory(Text_root);
                    }
                    // -----------------------------------------------------------------------
                    var desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                    DateTime time = DateTime.Now;
                    string format = "MMddyyyy";
                    string date1 = time.ToString(format);
                    String HourMinute = DateTime.Now.ToString("HHmm");
                    string FileName_Created = "";

                    FileName_Created = root + "\\" + CityName + "_" + "PdfFiles_Merge_" + date1 + "_" + HourMinute + ".pdf";
                    string Txt_FileName_Created = Text_root + "\\" + CityName + "_" + "PdfFiles_Merge_" + date1 + "_" + HourMinute + ".txt";
                    // -----------------------------------------------------------------------------------------
                    // Check if file already exists. If yes, delete it.     
                    if (File.Exists(Txt_FileName_Created))
                    {
                        File.Delete(Txt_FileName_Created);
                    }
                    List<Combo_InitList> List_Text = new List<Combo_InitList>();
                    var Query40 = from List10 in Display_Data10_01
                                  where List10.Name05 != ""
                                  select List10;
                    foreach (var List20 in Query40)
                    {
                        List_Text.Add(new Combo_InitList()
                        {
                            Name = List20.Name02,
                            Id = 0
                        });
                    }
                    if (File.Exists(Txt_FileName_Created))
                    {
                        using (StreamReader file = new StreamReader(Txt_FileName_Created))
                        {
                            int counter = 0;
                            string ln;
                            while ((ln = file.ReadLine()) != null)
                            {
                                List_Text.Add(new Combo_InitList()
                                {
                                    Name = ln,
                                    Id = 0
                                });
                                counter++;
                            }
                            file.Close();
                        }
                        File.Delete(Txt_FileName_Created);
                    }
                    FileStream fs1 = new FileStream(Txt_FileName_Created, FileMode.OpenOrCreate, FileAccess.Write);
                    StreamWriter writer10 = new StreamWriter(fs1);
                    writer10.WriteLine("Last Date Update: " + DateTime.Now.ToShortDateString() + " - Time: " + DateTime.Now.ToString("HH:mm"));
                    if (List_Text.Count() > 0)
                    {
                        foreach (var List10 in List_Text)
                        {
                            writer10.WriteLine(List10.Name);
                        }
                    }
                    writer10.Close();
                    //Create the document which will contain the combined PDF's
                    Document document = new Document();
                    //Create a writer for de document
                    string Merge_Pdf_Files = FileName_Created;
                    PdfCopy writer = new PdfCopy(document, new FileStream(Merge_Pdf_Files, FileMode.Create));
                    if (writer == null)
                    {
                        return;
                    }
                    //Open the document
                    document.Open();
                    var Query20 = from List10 in Display_Data10_01
                                  select List10;
                    foreach (var List20 in Query20)
                    {
                        //  Read the PDF file
                        string outputPath = @Str_Folder_Save;
                        string str_20 = outputPath + "\\" + List20.Name02;
                        using (PdfReader reader = new PdfReader(str_20))
                        {
                            //Add the file to the combined one
                            writer.AddDocument(reader);
                        }
                    }
                    //Finally close the document and writer
                    writer.Close();
                    document.Close();
                    Mouse.OverrideCursor = null;
                    // ----------------------------------------------------------------------
                    TextBox_Save.Text = "Number of pdf created: " + Label_Nbr_To_Generate.Content + System.Environment.NewLine + System.Environment.NewLine + "Pdf file Created:" + System.Environment.NewLine + FileName_Created;
                    Mouse.OverrideCursor = null;
                    TabControl_Data.SelectedIndex = 6;
                    TabControl_Pdf.SelectedIndex = 1;

                    // --------------------------------------------------------------

                    //TabControl_Data.SelectedIndex = 3;

                    //MessageBox.Show("... File Created ...", "Import Data");
                    //Data_Init_Create_Pdf();
                }
                else
                {
                    MessageBox.Show("...ERROR ..." + System.Environment.NewLine + "You have to enter the City name", "Pdf Data");
                }
                Mouse.OverrideCursor = null;
            }
        }

        private void Btn_Save_Continue_Click(object sender, RoutedEventArgs e)
        {
            Data_Init_Create_Pdf();
        }

        /// <summary>
        /// 
        /// </summary>v
        private void Brn_Result_Excel_Data_Click(object sender, RoutedEventArgs e)
        {
            var desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            DateTime time = DateTime.Now;
            string format = "MMddyyyy";
            string date1 = time.ToString(format);
            //string Str_Infos = "";
            //switch (Int_Combo_Template)
            //{
            //    case 1:
            //        Str_Infos = "All_Data_with_Labels";
            //        break;
            //    case 2:
            //        Str_Infos = "All_Data_With_No_Labels";
            //        break;
            //    default:
            //        Str_Infos = "All_Data";
            //        break;
            //}
            String HourMinute = DateTime.Now.ToString("HHmm");
            string fileName = "";
            string Str_Titre = "";
            //fileName = "PdfSata_" + Str_Infos + "_Print_" + date1 + "_" + HourMinute + ".xlsx";
            fileName = "PdfData_Print_" + date1 + "_" + HourMinute + ".xlsx";
            var Location = System.IO.Path.Combine(desktopFolder, fileName);
            // ---------------------------------------------------
            string filename1 = Location;
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
            dlg.DefaultExt = Location;
            dlg.FileName = fileName;
            dlg.Filter = "Excel workbook (.xlsx)|*.xlsx";
            bool Suite10 = false;
            if (dlg.ShowDialog() == true)
            {
                filename1 = dlg.FileName;
                Suite10 = true;
            }
            if (Suite10 == true)
            {
                Mouse.OverrideCursor = Cursors.Wait;
                List<DataGrid_Init_10> List_XXX_Temp = new List<DataGrid_Init_10>();
                List_XXX_Temp = Display_Data10_01.ToList();
                DataTable Dt_Print = new DataTable();
                Dt_Print = GlobalVar.Convert_ListofObject_To_DataTable(List_XXX_Temp);
                Dt_Print.Columns.Remove("Name04");
                Dt_Print.Columns.Remove("Name05");
                Dt_Print.Columns.Remove("Name06");
                Dt_Print.Columns.Remove("Name07");
                Dt_Print.Columns.Remove("Name08");
                Dt_Print.Columns.Remove("Name09");
                Dt_Print.Columns.Remove("Name10");
                Dt_Print.Columns.Remove("NameId");
                DataTable Dt_Print1 = new DataTable();
                Dt_Print1.Columns.Add(new DataColumn("Name01", Type.GetType("System.String")));
                Dt_Print1.Columns.Add(new DataColumn("Name02", Type.GetType("System.String")));
                Dt_Print1.Columns.Add(new DataColumn("Name03", Type.GetType("System.String")));

                DataRow workRow;
                workRow = Dt_Print1.NewRow();
                workRow["Name01"] = "Product SKU";
                workRow["Name02"] = "Pdf Files";
                workRow["Name03"] = "Location";

                Dt_Print1.Rows.Add(workRow);
                foreach (DataRow row in Dt_Print.Rows)
                {
                    workRow = Dt_Print1.NewRow();
                    workRow["Name01"] = row["Name01"].ToString();
                    workRow["Name02"] = row["Name02"].ToString();
                    workRow["Name03"] = row["Name03"].ToString();
                    Dt_Print1.Rows.Add(workRow);
                }
                // -------------------------------------------------------------------------
                GlobalVar.ExportToExcel(Dt_Print1, filename1, 1, 1, Dt_Print1.Rows.Count, Dt_Print1.Columns.Count, Str_Titre);
                Mouse.OverrideCursor = null;
                MessageBox.Show(".... Export complete ....", "Pdf Generate Files");
            }
        }


        /// <summary>
        private void Data_Generate_Old()
        {
            string outputPath = @Str_Folder_Save;
            string tfile = "Wayfair_Carton_Labels_SPOS_11752773_2-0114";
            string str_20 = "C:\\delete" + "\\" + tfile + ".pdf";
            PdfReader reader1 = new PdfReader(@str_20);
            int intPageNum = reader1.NumberOfPages;
            string[] words;
            string line;
            for (int i = 1; i <= intPageNum; i++)
            {
                string text = PdfTextExtractor.GetTextFromPage(reader1, i, new LocationTextExtractionStrategy());
                words = text.Split('\n');
                int ts = words.Length;
                for (int j = 0, len = words.Length; j < len; j++)
                {
                    line = Encoding.UTF8.GetString(Encoding.UTF8.GetBytes(words[j]));
                    string ttt = line.ToString();
                    if (line.ToLower().Contains("part"))
                    {

                        int aaa = i;
                        string Str_Product = line.Replace("PART", "");
                        Str_Product = Str_Product.Replace("#", "");
                        Str_Product = Str_Product.Replace(":", "");
                        Str_Product = Str_Product.TrimStart();
                        // --------------------------------------------------------
                        reader1.Dispose();
                        reader1.Close();
                        string str_28 = outputPath + "\\" + "New_" + tfile + ".pdf";
                        string Str_Mew_Data = Str_Product;
                        // -------------------------------------------------------------------------
                        Add_Text_New(str_20, str_28, Str_Mew_Data);

                    }
                }
            }
            MessageBox.Show("... File Created ...", "Import Data");
        }

        /// <summary>
        /// 
        /// </summary>
        private void Add_Text_New(string oldFile, string newFile, string TextToAdd)
        {

            // open the reader
            PdfReader reader = new PdfReader(oldFile);
            iTextSharp.text.Rectangle size = reader.GetPageSizeWithRotation(1);
            Document document = new Document(size);
            // Document document = new Document();
            // open the writer
            FileStream fs = new FileStream(newFile, FileMode.Create, FileAccess.Write);
            // FileStream fs1 = new FileStream(oldFile, FileMode.Create, FileAccess.Write);
            PdfWriter writer = PdfWriter.GetInstance(document, fs);
            document.Open();
            // the pdf content
            PdfContentByte cb = writer.DirectContent;
            // select the font properties
            BaseFont bf = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetColorFill(BaseColor.DARK_GRAY);
            cb.SetFontAndSize(bf, 8);
            // write the text in the pdf content
            cb.BeginText();
            //Create the first block with different font style
            string text = TextToAdd; // "Some random blablablabla...";
                                     // put the alignment and coordinates here
                                     //float x, y;
                                     //int ts = 18;
                                     //x = (40 + ts + 440) / 2;
                                     //y = 40 + ts + 24;
                                     //  cb.ShowTextAligned(500, text, x, y, 0);
                                     //  cb.ShowTextAligned(1, text, 520, 640, 0);
                                     // cb.ShowTextAligned(1, text, 520, 640, 0);
                                     //Chunk Test10 = new Chunk(text, FontFactory.GetFont("dax-black"));
                                     //Test10.SetUnderline(0.5f, -1.5f);

            //cb.ShowTextAligned(1, Test10.ToString(), 520, 700, 0);
            //  cb.ShowTextAligned(1, text, 520, 640, 0);
            //cb.ShowTextAligned(400, text, 200, 200, 1);
            // ---------------------------------------------------------------------------------------
            cb.EndText();
            cb.BeginText();
            //  text = "Other random blabla...";
            // put the alignment and coordinates here
            // cb.ShowTextAligned(2, text, 100, 200, 0);
            // Setting font of the text PdfFont 

            // cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "This text is left aligned new", 200, 200, 0);

            // cb.ShowTextAligned(400, "( 102-20-45}hauteur, 0);
            // Setting font of the text PdfFont 
            var veryNiceBaseFont = BaseFont.CreateFont(text, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetColorFill(new BaseColor(0x00, 0x57, 0x94)); // Certain BLUE RGB
            cb.SetFontAndSize(veryNiceBaseFont, 11f);

            //var bfCourier = BaseFont.CreateFont(BaseFont.COURIER, BaseFont.CP1252, false);
            //var courier = new iTextSharp.text.Font(bfCourier, 50, iTextSharp.text.Font.ITALIC, BaseColor.BLACK);
            text = "Here is an amazing text that I add on top of the template!";
            cb.SetCharacterSpacing(0.8f);
            // cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 79, 582, 0)

            // cb.ShowTextAligned(400, text30, 175, 185, 0);
            cb.ShowTextAligned(400, text, 175, 185, 0);
            // cb.ShowTextAligned(400, "( 105-20-35 )", 175, 185, 0);

            // BaseFont bfTimes = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, false);
            //font times = new Font(bfTimes, 12, Font.ITALIC, Color.RED);

            cb.EndText();
            // create the new page and add it to the pdf
            PdfImportedPage page = writer.GetImportedPage(reader, 1);
            cb.AddTemplate(page, 0, 0);
            // close the streams and voilá the file should be changed :)
            document.Close();
            fs.Close();
            writer.Close();
            reader.Close();
            reader.Dispose();
            // fs1.Close();
            // --------------------------------------------------------
            if (File.Exists(oldFile))
            {

                File.Delete(oldFile);
            }

            // Create a FileInfo  
            System.IO.FileInfo fi = new System.IO.FileInfo(newFile);
            // Check if file is there  
            if (fi.Exists)
            {
                // Move file with a new name. Hence renamed.  
                fi.MoveTo(@oldFile);
            }

        }


        /// <summary>
        /// 
        /// </summary>
        private void Split_Pdf_File_New()
        {
            try
            {
                // https://www.c-sharpcorner.com/article/splitting-pdf-file-in-c-sharp-using-itextsharp/
                if (List_Data_Pdf_File.Count() == 0)
                {
                    MessageBox.Show("You have to select the Excel file", "Excel File");
                    Label_Data_Message.Content = "--- Load Excel files ...";
                    return;
                }
                Mouse.OverrideCursor = Cursors.Wait;
                GroupBox_Progress.Visibility = System.Windows.Visibility.Hidden;
                string outputPath = @Str_Folder_Save;
                Display_Data10_01.Clear();
                int Nbre_Add = 0;
                int Nbr_items = List_Data_Pdf_File.Count();
                // create PdfFileEditor object
                var Query20 = from List10 in List_Data_Pdf_File
                              select List10;
                foreach (var List20 in Query20)
                {
                    // -------------------------------------------------------
                    Nbre_Add = Nbre_Add + 1;
                    string Str01 = "";
                    if (Nbre_Add < 10)
                    {
                        Str01 = "0" + Nbre_Add.ToString();
                    }
                    else
                    {
                        Str01 = Nbre_Add.ToString();
                    }
                    Str_Global_File_Pdf = List20.Name01;
                    string pdfFilePath = @Str_Global_File_Pdf;
                    int interval = 1;
                    int pageNameSuffix = 0;
                    // Intialize a new PdfReader instance with the contents of the source Pdf file:
                    PdfReader reader = new PdfReader(pdfFilePath);

                    FileInfo file = new FileInfo(pdfFilePath);
                    string pdfFileName = file.Name.Substring(0, file.Name.LastIndexOf(".")) + "-";
                    NbreTotal = reader.NumberOfPages;
                    int NbreLine = 0;
                    for (int pageNumber = 1; pageNumber <= reader.NumberOfPages; pageNumber += interval)
                    {
                        // -------------------------------
                        NbreLine = NbreLine + 1;
                        Proggress_Bar(NbreLine, Str_Global_File_Pdf + "      ---  Pdf Files  " + Nbre_Add.ToString() + " / " + Nbr_items.ToString());
                        // -----------------------------------
                        pageNameSuffix++;
                        string newPdfFileName = string.Format(pdfFileName + Str01 + "{0}", pageNameSuffix);
                        // ---------------------------------------------------------------------------
                        string Str_Fle = newPdfFileName + ".pdf"; ;
                        string Str_Product = "";
                        //// ----------------------------------------------------------------------------
                        // Tratement done on Disk --- Source document
                        //  SplitAndSaveInterval(pdfFilePath, outputPath, pageNumber, interval, newPdfFileName);
                        // -------------------------------------------------------------------------------
                        List_File_Created.Add(new Combo_InitList()
                        {
                            Name = Str_Fle,
                            Id = pageNumber,
                        });



                        // -----------------------------------------------
                        // string str_20 = outputPath + "\\" + newPdfFileName + ".pdf";
                        string str_20 = Str_Fle;
                        PdfReader reader1 = new PdfReader(@str_20);
                        int intPageNum = reader1.NumberOfPages;
                        string[] words;
                        string line;
                        for (int i = 1; i <= intPageNum; i++)
                        {
                            string text = PdfTextExtractor.GetTextFromPage(reader1, i, new LocationTextExtractionStrategy());
                            words = text.Split('\n');
                            for (int j = 0, len = words.Length; j < len; j++)
                            {
                                line = Encoding.UTF8.GetString(Encoding.UTF8.GetBytes(words[j]));
                                if (line.ToLower().Contains("part"))
                                {
                                    Str_Product = line.Replace("PART", "");
                                    Str_Product = Str_Product.Replace("#", "");
                                    Str_Product = Str_Product.Replace(":", "");
                                    Str_Product = Str_Product.TrimStart();
                                    // --------------------------- Create 
                                }
                            }
                        }
                        string Str_05 = "X";
                        CheckBox_Data.IsChecked = true;
                        // -+---------------------------------------------------
                        Display_Data10_01.Add(new DataGrid_Init_10()
                        {
                            Name01 = Str_Product,
                            Name02 = Str_Fle,
                            Name03 = "",
                            Name04 = "",
                            Name05 = Str_05,
                            Name06 = "",
                            Name07 = "",
                            Name08 = "",
                            Name09 = "",
                            Name10 = "",
                            NameId = 0,
                        });
                        // ----------------------------------------------------------
                    }
                }
                // ---------------------------------------------
                var Query0 = from List11 in Display_Data10
                             orderby List11.Name01
                             select List11;
                foreach (var List50 in Query0)
                {
                    int Nbre_Data = List50.NameId;
                    int Nbre = 0;
                    var Query30 = from List10 in Display_Data10_01
                                  where List10.Name01.TrimStart().Replace(" ", "").ToLower() == List50.Name01.TrimStart().Replace(" ", "").ToLower()
                                  select List10;
                    foreach (var List20 in Query30)
                    {
                        string str_18 = outputPath + "\\" + List20.Name02;
                        string str_28 = outputPath + "\\" + "New_" + List20.Name02;
                        //   Add_Text(str_18, str_28, List20.Name01);

                        if (Nbre <= Nbre_Data)
                        {
                            var Query80 = from List80 in List_Excel_File
                                          where List80.Name05 == ""
                                          & List80.Name01.TrimStart().Replace(" ", "").ToLower() == List20.Name01.TrimStart().Replace(" ", "").ToLower()
                                          orderby List80.Name01
                                          select List80;
                            foreach (var List88 in Query80)
                            {
                                List88.Name05 = "ok";
                                List20.Name03 = List88.Name03;
                                break;
                            }

                            List20.Name04 = "Ok";
                        }
                        Nbre = Nbre + 1;
                    }
                }
                List<DataGrid_Init_10> T_Temp = new List<DataGrid_Init_10>();
                T_Temp = Display_Data10_01.OrderBy(x => x.Name01).ToList();
                Display_Data10_01.Clear();
                Display_Data10_01 = T_Temp.ToList();

                // ---------------------------------------------
                DataGrid_Result();
                Display_Data10_All.Clear();
                Display_Data10_All = Display_Data10_01.ToList();
                Mouse.OverrideCursor = null;
            }
            catch (Exception e)
            {
                Console.WriteLine("Exception: " + e.Message);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private void DataGrid_Result()
        {

            Btn_Generate_Pdf.Visibility = System.Windows.Visibility.Hidden;
            int Nbre_To_Generate = Display_Data10_01.Where(x => x.Name05 != "").Count();
            Label_Nbr_To_Generate.Content = GlobalVar.Number_With_Dash(Nbre_To_Generate).ToString();
            if (Display_Data10_01.Where(x => x.Name05 != "").Count() > 0)
            {
                Btn_Generate_Pdf.Visibility = System.Windows.Visibility.Visible;
            }
            // --------------------------------------------------

            // ---------------------------------------------

            // -------------------------------------------------------------------------------------------------------------------------------     

            DataGrid_Generate_Pdf.ItemsSource = null;
            DataGrid_Generate_Pdf.Visibility = System.Windows.Visibility.Visible;
            //var NewList = Display_Data05_01.OrderBy(x => x.Name01).ToList();
            DataGrid_Generate_Pdf.ItemsSource = Display_Data10_01;
            DataGrid_Generate_Pdf.IsReadOnly = true;

        }


        /// <summary>
        /// 
        /// </summary>
        private void Final_Data()
        {
            var Query0 = from List11 in Display_Data10_01
                         select List11;
            foreach (var List50 in Query0)
            {
                if (List50.Name03.Replace("-","") == "")
                {
                    string Str_Loc = "";

                    var Query80 = from List80 in List_Excel_File
                                  where List80.Name01.TrimStart().Replace(" ", "").ToLower() == List50.Name01.TrimStart().Replace(" ", "").ToLower()
                                  select List80;
                    foreach (var List88 in Query80)
                    {
                        Str_Loc = List88.Name04;
                        break;
                    }

                    List50.Name03 = Str_Loc;
                }
            }
            DataGrid_Result();
        }
    }
}

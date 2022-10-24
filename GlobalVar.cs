
using Init_Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;

public static class GlobalVar
{
    static int _globalValue;


    /// <summary>
    /// 
    /// </summary>
    public static List<DataGrid_Init_05> Spilt_Data(string Str_Data)
    {
        List<DataGrid_Init_05> result = new List<DataGrid_Init_05>();
        string Test = Str_Data;
        string Str_Data01 = "";
        int T_Test = Test.Length;
        int Nbr = 0;
        string Val_01 = "";
        string Val_02 = "";
        string T_Val_01 = "";
        string T_Val_02 = "";
        int T_Long = T_Test;
        for (int i = 1; i <= T_Long; i++)
        {

            string Str01 = Test.Substring(i - 1, 1);
            if (Str01 == "*")
            {
                Nbr = Nbr + 1;
                switch (Nbr)
                {
                    case 1:
                        Val_01 = Str_Data01;
                        T_Val_01 = Val_01;
                        Str_Data01 = "";
                        break;
                    case 2:
                        Val_02 = Str_Data01;
                        T_Val_02 = Val_02;
                        Str_Data01 = "";
                        Nbr = 0;
                        break;
                }
                if (Val_02 != "")
                {
                    result.Add(new DataGrid_Init_05()
                    {
                        Name01 = Val_01,
                        Name02 = Val_02,
                        Name03 = "",
                        Name04 = "",
                        Name05 = "",
                        NameId = 0
                    });
                    Val_01 = "";
                    Val_02 = "";
                }
            }
            else
            {
                Str_Data01 = Str_Data01 + Str01;
                if (i == T_Long)
                {
                    result.Add(new DataGrid_Init_05()
                    {
                        Name01 = Val_01,
                        Name02 = Str_Data01,
                        Name03 = "",
                        Name04 = "",
                        Name05 = "",
                        NameId = 0
                    });
                }
            }
        }


        return result;
    }

    /// <summary>
    /// 
    /// </summary>
    public static string ReadtextFile_ForBase_ConnectionString()
    {
        string result = "";
        string Str_Text_File = "\\\\vmhost2\\_WMS_Share\\Tayse_WMS_Config.txt";
        string Str_Data = "";
        using (StreamReader file = new StreamReader(Str_Text_File))
        {
            int counter = 0;
            string ln;

            while ((ln = file.ReadLine()) != null)
            {
                //Console.WriteLine(ln);
                Str_Data = ln;
                counter++;
            }
            file.Close();
            // Console.WriteLine($ "File has {counter} lines.");
        }
        if (Str_Data != "")
        {
            result = Str_Data;
        }
        return result;
    }


    /// <summary>
    /// 
    /// </summary>
    public static string Generate_Aisle_From_Location(string Str_Location, int ProductType_ID)
    {
        string result = "";
        if (ProductType_ID == 0) { ProductType_ID = 1; };
        if (Str_Location != "")
        {
            string T_Location = Str_Location.Replace("-", "").Replace(" ", "");
            int T_Taille = T_Location.Length;
            switch (ProductType_ID)
            {
                case 1:
                    if (T_Taille == 8)
                    {
                        if (GlobalVar.ControlInt(T_Location.Substring(0, 3)) > 0)
                        {
                            result = T_Location.Substring(0, 3);
                        }
                        else
                        {
                            result = "Special";
                        }
                    }
                    else
                    {
                        result = "Special";
                    }
                    break;
                case 4:
                    if (T_Taille == 9)
                    {
                        if (T_Location.ToLower().StartsWith("d"))
                        {
                            if (GlobalVar.ControlInt(T_Location.Substring(1, 3)) > 0)
                            {
                                result = T_Location.Substring(1, 3);
                            }
                            else
                            {
                                result = "Special";
                            }
                        }
                        else
                        {
                            result = "Special";
                        }
                    }
                    else
                    {
                        result = "Special";
                    }
                    break;
            }
        }
        return result;
    }


   
    /// <summary>
    /// 
    /// </summary>
    public static string ControlDecimal_Data(string Str_Text)
    {
        string result = "";
        if (Str_Text != "")
        {
            int Len10 = Str_Text.Length;
            result = Str_Text.Substring(0, Len10 - 4);
        }
        return result;
    }

    /// <summary>
    /// 
    /// </summary>
    public static string ControlText(string Str_Text, int T_Len)
    {
        string result = "";
        if (Str_Text != "")
        {
            if (Str_Text.Length > T_Len)
            {
                result = Str_Text.Substring(0, (T_Len - 1));
            }
            else
            {
                result = Str_Text;
            }

        }
        return result;
    }

    /// <summary>
    /// 
    /// </summary>
    public static int GlobalValue
    {
        get
        {
            return _globalValue;
        }
        set
        {
            _globalValue = value;
        }
    }

    /// <summary>
    /// 
    /// </summary>
    public static string Change_Date(string T_Date)
    {
        string Result = "";
        if (T_Date != "")
        {
            Result = Convert.ToDateTime(T_Date).ToShortDateString();
        }
        else
        {
            Result = DateTime.Now.ToShortDateString();
        }
        return Result;
    }

    /// <summary>
    /// 
    /// </summary>
    public static string Write_in_Text(string filepath)
    {
        string Result = "";
        List<Combo_InitList> List_Text = new List<Combo_InitList>();
        if (File.Exists(filepath))
        {
            using (StreamReader file = new StreamReader(filepath))
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
            File.Delete(filepath);
        }
        FileStream fs1 = new FileStream(filepath, FileMode.OpenOrCreate, FileAccess.Write);
        StreamWriter writer = new StreamWriter(fs1);
        writer.WriteLine("Last Date Update: " + DateTime.Now.ToShortDateString() + " - Time: " + DateTime.Now.ToString("HH:mm"));
        if (List_Text.Count() > 0)
        {
            foreach (var List10 in List_Text)
            {
                writer.WriteLine(List10.Name);
            }
        }
        writer.Close();
        return Result;
    }

    /// <summary>
    /// 
    /// </summary>
    public static string Specify_Aisle_From_Location(string Str_Location)
    {
        string Result = "";
        if (Str_Location != "")
        {
            int T_Len = Str_Location.ToLower().Replace(" ", "").Replace("-", "").Length;
            string Str_Aisle = "";
            if (T_Len == 9)
            {
                Str_Aisle = Str_Location.Substring(0, 4);
                if ((Str_Aisle.ToLower().Substring(0, 1) == "d") && (GlobalVar.ControlInt(Str_Aisle.ToLower().Substring(1, 3)) > 0))
                {
                    Result = Str_Aisle;
                }
            }
            else
            {
                if (T_Len == 8)
                {
                    Str_Aisle = Str_Location.Substring(0, 3);
                    if (GlobalVar.ControlInt(Str_Location.Substring(0, 3)) > 0)
                    {
                        Result = Str_Aisle;
                    }
                }
            }
        }
        else
        {
            Result = "";
        }
        return Result;
    }

   
    /// <summary>
    /// 
    /// </summary>
    public static string RetrieveData_In_List_Combo(int Str_Id, List<Combo_InitList> List_Data)
    {
        string Result = "";
        var queryFile0 = from List10 in List_Data
                         where List10.Id == Str_Id
                         select List10;
        foreach (var List0 in queryFile0)
        {
            Result = List0.Name;
        }
        return Result;
    }

    

    /// <summary>
    /// 
    /// </summary>
    public static string Aisle_On_Location(string Str_Location)
    {
        string result = "";
        if (Str_Location != "")
        {
            if (GlobalVar.ControlInt(Str_Location.Substring(0, 3)) > 0)
            {
                result = Str_Location.Substring(0, 3);
            }
            else
            {
                result = Str_Location;
            }
        }
        return result;
    }



    /// <summary>
    /// 
    /// </summary>
    public static int Turn_Location_To_Bin(string Str_Location)
    {
        int result = 0;
        result = GlobalVar.ControlInt(Str_Location.Substring(0, 3));
        return result;
    }


    /// <summary>
    /// 
    /// </summary>
    public static string Extract_ProductSKU_On_Location_To_Size(string Str_ProductSKU, string Str_Location)
    {
        string result = "";
        int Str_Bin = GlobalVar.ControlInt(Str_Location.Substring(0, 3));
        if (Str_Bin > 0)
        {
            int Length_Location = Str_ProductSKU.Length;
            if (Length_Location >= 11 && Length_Location <= 15)
            {
                result = Str_ProductSKU.Substring(8, Length_Location - 8).Replace(" ", "");
            }
        }
        return result;
    }


    /// <summary>
    /// 
    /// </summary>
    public static string NumberWithoutZero(int Str_Num)
    {
        string result = "";
        if (Str_Num > 0)
        {
            result = Str_Num.ToString();
        }
        return result;
    }


    /// <summary>
    /// 
    /// </summary>
    public static string Date_Normal(string Str_Date)
    {
        string result = "";
        if (Str_Date != "")
        {
            int Taille = Str_Date.Trim().Length;
            for (int i = 0; i < Taille; i++)
            {
                string Str_01 = Str_Date.Substring(i, 1);
                if (Str_01 == " ")
                {
                    break;
                }
                else
                {
                    result = result + Str_01;
                }
                Console.WriteLine(i);
            }
        }
        return result;
    }


    /// <summary>
    /// 
    /// </summary>
    public static int DateToNumber(string Str_Date)
    {
        int result = 0;
        if (Str_Date != null)
        {
            if (Str_Date != "")
            {
                bool suite = false;
                if (Str_Date.Length == 9)
                {
                    if (GlobalVar.ControlInt(Str_Date.Substring(0, 1)) > 0)
                    {
                        suite = true;
                    }
                }
                if (Str_Date.Length == 10)
                {
                    if (GlobalVar.ControlInt(Str_Date.Substring(0, 2)) > 0)
                    {
                        suite = true;
                    }
                }
                if (Str_Date.Length == 8)
                {
                    if (GlobalVar.ControlInt(Str_Date.Substring(0, 1)) > 0)
                    {
                        suite = true;
                    }
                }
                if (suite == true)
                {
                    System.DateTime GL_DateforCalculation = new System.DateTime(2000, 1, 1, 0, 0, 0);
                    System.TimeSpan diffResult = Convert.ToDateTime(Str_Date).Subtract(GL_DateforCalculation);
                    result = diffResult.Days;
                }
            }
        }
        return result;
    }

    /// <summary>
    /// 
    /// </summary>
    public static string NumberToDate(int T_Date)
    {
        string result = "";
        if (T_Date != 0)
        {
            string iDate = "01/01/2000";
            DateTime oDate = Convert.ToDateTime(iDate).AddDays(T_Date);
            result = oDate.ToShortDateString();
        }
        return result;
    }

    /// <summary>
    /// 
    /// </summary>
    public static string EmptyNumber(int Str_Data)
    {
        string result = "";
        if (Str_Data > 0)
        {
            result = Str_Data.ToString();
        }
        return result;
    }


    /// <summary>
    /// 
    /// </summary>
    public static string Update_Data_InDataBase(string strsql, string ConnectionString)
    {
        string Result = "";
        using (SqlConnection connection = new SqlConnection(ConnectionString))
        using (var cmd = new SqlCommand())
        {
            cmd.CommandText = strsql;
            cmd.Connection = connection;
            connection.Open();
            int numberDeleted = cmd.ExecuteNonQuery();
            connection.Close();
        }
        return Result;
    }

    /// <summary>
    /// 
    /// </summary>
    public static string NumberTostring(int Str_Date)
    {
        string result = "";
        if (Str_Date > 0)
        {
            result = Str_Date.ToString();
        }
        return result;
    }

    /// <summary>
    /// 
    /// </summary>
    public static string CheckString(string Data1, int Str_Taille)
    {
        string Result = "";
        if (Data1 != "")
        {
            int T_01 = Data1.Length;
            int Max = Str_Taille;
            if (T_01 < Str_Taille)
            {
                Max = T_01;
            }

            Result = Data1.Substring(0, Max);

        }
        return Result;

    }

    /// <summary>
    /// 
    /// </summary>
    public static int ControlInt(string Data1)
    {
        int Result = 0;
        int value;
        if (int.TryParse(Data1, out value))
        {
            Result = value;
        }
        return Result;

    }

    /// <summary>
    /// 
    /// </summary>
    public static double ControlDouble(string Data1)
    {
        double Result = 0;
        double value;
        if (double.TryParse(Data1, out value))
        {
            Result = value;
        }
        return Result;

    }

    /// <summary>
    /// 
    /// </summary>
    public static string Convert_TAWO_Date(string Data1)
    {
        string Result = "";
        DateTime oDate1 = Convert.ToDateTime(Data1);
        string Str_011 = oDate1.ToShortDateString();
        Result = Str_011;
        return Result;

    }

   

    /// <summary>
    /// 
    /// </summary>
    public static string Number_With_Dash(int Str_Num)
    {
        string result = "-";
        if (Str_Num != 0)
        {
            result = Str_Num.ToString();
        }
        return result;
    }

    /// <summary>
    /// 
    /// </summary>
    public static string Write_Data_In_Database(DataTable Dt_File, string Str_Table, string Str_Connection)
    {
        string result = "";
        using (SqlBulkCopy bulkCopy = new SqlBulkCopy(Str_Connection))
        {
            bulkCopy.DestinationTableName = Str_Table;
            try
            {
                //  Write from the source to the destination.
                bulkCopy.WriteToServer(Dt_File);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        return result;
    }

    /// <summary>
    /// 
    /// </summary>
    public static DataTable Convert_ListofObject_To_DataTable<T>(IList<T> data)
    {
        PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));
        DataTable table = new DataTable();
        foreach (PropertyDescriptor prop in properties)
            table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
        foreach (T item in data)
        {
            DataRow row = table.NewRow();
            foreach (PropertyDescriptor prop in properties)
                row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
            table.Rows.Add(row);
        }
        return table;
    }

    /// <summary>
    /// 
    /// </summary>
    public static string WriteToExcel(System.Data.DataTable dt, string location, string Str_Title)
    {
        string result = "";
        Microsoft.Office.Interop.Excel.Application XlObj = new Microsoft.Office.Interop.Excel.Application();
        XlObj.Visible = false;
        Microsoft.Office.Interop.Excel._Workbook WbObj = (Microsoft.Office.Interop.Excel.Workbook)(XlObj.Workbooks.Add(""));
        Microsoft.Office.Interop.Excel._Worksheet WsObj = (Microsoft.Office.Interop.Excel.Worksheet)WbObj.ActiveSheet;
        try
        {
            if (File.Exists(@location))
            {
                File.Delete(@location);
            }
            int row = 0; int col = 1;
            //if (Header == 1)
            //{
            //row = 1;
            //foreach (DataColumn column in dt.Columns)
            //{
            //    //adding columns
            //    WsObj.Cells[row, col] = column.ColumnName;
            //    col++;
            //}
            //}
            //reset column and row variables
            col = 1;
            row = 1;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                foreach (var cell in dt.Rows[i].ItemArray)
                {
                    WsObj.Cells[row, col] = cell;
                    WsObj.Columns[col].NumberFormat = "@";
                    col++;
                }
                col = 1;
                row++;
            }
            // --------------------------------------------
            // WsObj.Columns.AutoFit();
            String HourMinute = DateTime.Now.ToString("HH:mm");
            int Ligne = 3;
            for (int i = 1; i < dt.Columns.Count + 1; i++)
            {
                WsObj.Cells[Ligne, i].Font.Bold = true;
            }
            WsObj.Range[WsObj.Cells[Ligne, 1], WsObj.Cells[Ligne, dt.Columns.Count]].Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            Microsoft.Office.Interop.Excel.Range cell80 = WsObj.Cells[Ligne, 1];
            Microsoft.Office.Interop.Excel.Range cell81 = WsObj.Cells[dt.Rows.Count, dt.Columns.Count];
            Microsoft.Office.Interop.Excel.Range tRange80 = WsObj.get_Range(cell80, cell81);
            tRange80.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            tRange80.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline;
            WsObj.PageSetup.CenterHeader = "Warehouse DEPT - " + DateTime.Now.ToString();
            WsObj.PageSetup.LeftFooter = "Tayse International 501 Richardson Rd Calhoun, GA 30701, USA.";
            int Ligne0 = 1;
            // WsObj.Cells[Ligne0, 1] = Str_Title;
            WsObj.Cells[Ligne0, 1].Font.Bold = true;
            WsObj.Cells[Ligne0, 1].Font.Size = 20;
            //  WsObj.Range[WsObj.Cells[Ligne0, 1], WsObj.Cells[1, dt.Columns.Count]].Merge();

            if (location != null && location != "")
            {
                try
                {
                    WbObj.SaveAs(location);
                    WbObj.Close();
                    // ------------------------------------
                    System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName("Excel");
                    foreach (System.Diagnostics.Process p in process)
                    {
                        if (!string.IsNullOrEmpty(p.ProcessName) && p.StartTime.AddSeconds(+30) > DateTime.Now)
                        {
                            try
                            {
                                p.Kill();
                            }
                            catch { }
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception("ExportToExcel: Excel file could not be saved! Check filepath.\n"
      + ex.Message);
                }
            }
            else    // no filepath is given
            {
            }
        }
        catch (Exception ex)
        {
            throw new Exception("ExportToExcel: \n" + ex.Message);
        }
        return result;
    }


    /// <summary>
    /// 
    /// </summary>
    public static string ExportToExcel(System.Data.DataTable dt, string Location, int firstRow, int firstCol, int lastRow, int lastCol, string SheetName)
    {
        string result = "";
        Microsoft.Office.Interop.Excel.Application oExcel = new Microsoft.Office.Interop.Excel.Application();
        Microsoft.Office.Interop.Excel.Workbooks oBooks;
        Microsoft.Office.Interop.Excel.Sheets oSheets;
        Microsoft.Office.Interop.Excel.Workbook oBook;
        Microsoft.Office.Interop.Excel.Worksheet oSheet;
        //Create New Excel WorkBook 
        oExcel.Visible = false;
        oExcel.DisplayAlerts = false;
        oExcel.Application.SheetsInNewWorkbook = 1;
        oBooks = oExcel.Workbooks;
        oBook = (Microsoft.Office.Interop.Excel.Workbook)(oExcel.Workbooks.Add(Type.Missing));
        oSheets = oBook.Worksheets;
        oSheet = (Microsoft.Office.Interop.Excel.Worksheet)oSheets.get_Item(1);
        string Str_Name = "Sheet1";
        //if (SheetName != "") { Str_Name = SheetName; };
        oSheet.Name = Str_Name;
        // ---------------------------------------------------------------------------------------------------
        Microsoft.Office.Interop.Excel.Range top = oSheet.Cells[firstRow, firstCol];
        Microsoft.Office.Interop.Excel.Range bottom = oSheet.Cells[lastRow, lastCol];
        Microsoft.Office.Interop.Excel.Range all = (Microsoft.Office.Interop.Excel.Range)oSheet.get_Range(top, bottom);
        string[,] arrayDT = new string[dt.Rows.Count, dt.Columns.Count];
        for (int i = 0; i < dt.Rows.Count; i++)
            for (int j = 0; j < dt.Columns.Count; j++)
                arrayDT[i, j] = dt.Rows[i][j].ToString();
        all.Value2 = arrayDT;
        oExcel.Application.ScreenUpdating = true;
        oExcel.Application.EnableEvents = true;
        oExcel.Application.DisplayAlerts = true;
        // oExcel.Application.Calculation = excel.XlCalculation.xlCalculationAutomatic;
        oSheet.Columns.AutoFit();
        Microsoft.Office.Interop.Excel.Range cell80 = oSheet.Cells[1, 1];
        Microsoft.Office.Interop.Excel.Range cell81 = oSheet.Cells[dt.Rows.Count, dt.Columns.Count];
        Microsoft.Office.Interop.Excel.Range tRange80 = oSheet.get_Range(cell80, cell81);
        tRange80.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
        tRange80.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline;
        oSheet.PageSetup.RightFooter = "Tayse Apps - " + DateTime.Now.ToString();
        oSheet.PageSetup.LeftFooter = "Tayse Inc. 501 Richardson Rd Calhoun, GA 30701, USA.";

        ////////////////////////////////////////////////////////////////////////////////////
        oBook.SaveAs(Location);
        oExcel.Quit();
        //releaseObject(WbObj);
        //releaseObject(XlObj);
        //releaseObject(WsObj);
        // -----------------------------------------
        System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName("Excel");
        foreach (System.Diagnostics.Process p in process)
        {
            if (!string.IsNullOrEmpty(p.ProcessName) && p.StartTime.AddSeconds(+30) > DateTime.Now)
            {
                try
                {
                    p.Kill();
                }
                catch { }
            }
        }
        return result;
    }

    /// <summary>
    /// 
    /// </summary>
    public static string GlobalAction;
    public static string Global_Company;
    public static string GlobalPathImage;
    public static string GlobalAction_Data;
    public static string GlobalAction_Warehouse;
    public static string GlobalAction_Date;
    public static int GlobalAction_Acumatica_Transfer;
    public static int GlobalAction_Acumatica_Empty;
    public static int GlobalAction_Acumatica_Adjust;
    public static string GL_File_Name;
    public static string GL_Review;
    public static string GL_Data_Nbr;
    public static int GlobalOrderID;
    public static int Global_ID;
    public static bool Global_DeliveryData;
    public static bool Global_Save;
    public static string Global_Return_Nbr;
    public static string Global_Status_Action;
    public static int Global_User_ID;
    public static int Global_UserProfil_ID;
    public static int Global_Department_ID;
    public static int Global_ProductType_ID;
    public static int Global_Shipment_ID;
    public static int Global_CustomerClass_ID;
    public static int Global_Program_ID;
    public static int Global_User_CycleCount_ID;
    public static int Global_CycleCount_Identification_ID;
    public static string Global_CycleCount_Identification_Name;

    public static int Detail_PickPack_Qty;
    public static string Global_Menu;

    public static string Global_Path_Save = "\\\\server\\Company\\Tayse_App_Desktop";
    public static string GLobal_User_Name;
    public static bool G_UserFurniture;
    public static bool G_Val_Transfer;
    public static bool G_Val_Adjustment;
    public static bool G_Save;
    public static int G_ModuleID;
    public static int G_Identification_ID;
    public static int Global_Pallet_Num;
    public static string GlobalInvoiceNbr;
    public static string GlobalNbr;
    public static string Global_PickList;
    public static string Global_DateOperation;
    public static string G_Identification_Name;
    public static string GL_Apps = " From Apps TAWO by ";
    public static string GL_Apps_Login;
    public static string GL_Apps_PWD;
    public static string GL_Mode = "";
    public static bool G_Manage_User;
    public static string GL_PWD = "Tayse-h@8-maneg9";
    public static string GlobalInvoiceStatus;
    public static bool Global_Retrieve_Old_Data;
    public static bool Global_User_Allow_Change;


    public static string GL_HOST = "SERVER\\TDM_PROJECT";
    //  public static string GL_HOST = "VMHOST2";

    // public static string GL_HOST = GlobalVar.ReadtextFile_ForBase_ConnectionString();


   

}


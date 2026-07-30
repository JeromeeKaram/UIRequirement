using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeOpenXml;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Windows;

namespace UIRequirement.ViewModels;

public partial class MainViewModel : ObservableObject
{
    private readonly string sDir =
        Path.GetDirectoryName(
            System.Reflection.Assembly.GetEntryAssembly()!.Location)!;

    [ObservableProperty]
    private string? findingTicketNo;

    [ObservableProperty]
    private string? eSN;

    [ObservableProperty]
    private string? tSN;

    [ObservableProperty]
    private string? cSN;

    [ObservableProperty]
    private string? damageDescription;

    [ObservableProperty]
    private bool isLoading;

    [ObservableProperty]
    private string inputFolderPath;

    [ObservableProperty]
    private string ticketDetails;

    private string loginurl = string.Empty;
    private string username = string.Empty;
    private string password = string.Empty;

    List<ConfigInfo> lstConfigs = new List<ConfigInfo>();
    List<AsciiInfo> lstAscii = new List<AsciiInfo>();
    
    List<ManualCriteria> lstManualCriteria = new List<ManualCriteria>(0);

    public ObservableCollection<DamageInfo> Damages { get; } = new();

    [ObservableProperty]
    private Dictionary<string, List<string>> dtTypes = new();

    [ObservableProperty]
    private ObservableCollection<KeyValuePair<string, string>> overviewImages = new();

    [ObservableProperty]
    private ObservableCollection<KeyValuePair<string, string>> partsInformation = new();

    [ObservableProperty]
    private ObservableCollection<KeyValuePair<string, string>> repairParts = new();

    [ObservableProperty]
    private ObservableCollection<KeyValuePair<string, string>> zoomedViews = new();

    [ObservableProperty]
    private ObservableCollection<KeyValuePair<string, string>> nonConformanceTypes = new();

    [ObservableProperty]
    private ObservableCollection<KeyValuePair<string, string>> locations = new();

    [ObservableProperty]
    private string? selectedZoomedView;

    [ObservableProperty]
    private string? selectedOverviewImage;

    [ObservableProperty]
    private string? selectedPartInformation;

    [ObservableProperty]
    private string? selectedLocation;

    [ObservableProperty]
    private string? selectedNonConformanceType;

    [ObservableProperty]
    private string? selectedRepairPart;

    [ObservableProperty]
    private string? windowTitle;

    [RelayCommand]
    private async Task Load()
    {
        //MessageBox.Show("Load clicked");
        await InputFolderSelected();

        //var retObject = await GetDataAsync();

    //    ESN = retObject.ESN;
    //    TSN = retObject.TSN;
    //    CSN = retObject.CSN;

    //    TicketDetails = retObject.TicketDetails;

    //    var imagesKeyValueList = retObject.Images
    //.Select(path => new KeyValuePair<string, string>(
    //    Path.GetFileName(path),
    //    path))
    //.ToList();

    //    overviewImages.Clear();

    //    foreach (var item in imagesKeyValueList)
    //    {
    //        overviewImages.Add(item);
    //    }

    //    zoomedViews.Clear();

    //    foreach (var item in imagesKeyValueList)
    //    {
    //        zoomedViews.Add(item);
    //    }

    //    partsInformation.Clear();

    //    foreach (var item in imagesKeyValueList)
    //    {
    //        partsInformation.Add(item);
    //    }
    }

    [RelayCommand]
    private void Add()
    {
        MessageBox.Show("Add Clicked");
    }

    [RelayCommand]
    private void Clear()
    {
        MessageBox.Show("Clear clicked");
    }

    [RelayCommand]
    private void AddMoreImages()
    {
        MessageBox.Show("Add More Images");
    }

    [RelayCommand]
    private void Run()
    {
        MessageBox.Show("Run");
    }

    [RelayCommand]
    private void Close()
    {
        MessageBox.Show("Close clicked");
        //Application.Current.Shutdown();
    }

    private async Task InputFolderSelected()
    {
        try
        {
            IsLoading = true;
            var folder = Path.Combine(sDir, "bin", "images", "Finding");

            //Utility.WriteErrorLog("", "", folder);
            Console.WriteLine($"Finding folder: {folder}");

            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }
            else
            {
                foreach (string file in Directory.GetFiles(folder))
                {
                    //Utility.DeleteFile(file);
                    System.IO.File.Delete(file);
                }
            }

            string filename = string.Empty;

            if (Utility.m_eRUN == RUN.IHI)
            {
                string input = InputFolderPath?.Trim() ?? "";

                string findingTicket = "";
                string encTicket = "";

                if (input.Contains(','))
                {
                    string[] parts = input.Split(',');

                    findingTicket = parts[0].Trim();

                    if (parts.Length > 1)
                        encTicket = parts[1].Trim();
                }
                else
                {
                    findingTicket = input;
                }

                //int imageCount = await getRedmineImages(findingTicket, encTicket);

                //await WaitForImagesAsync(FindingFolder, imageCount);

                filename = Path.Combine(sDir, "bin", "description.txt");

                string encFile = Path.Combine(sDir, "bin", "ENC.txt");

                if (File.Exists(encFile))
                {
                    string[] lines = File.ReadAllLines(encFile);

                    if (lines.Length > 0) ESN = lines[0];
                    if (lines.Length > 1) TSN = lines[1];
                    if (lines.Length > 2) CSN = lines[2];
                }
            }
            else
            {
                string tempFolder = Path.Combine(sDir, "bin", "test");

                filename = Path.Combine(tempFolder, "description_15April2025.txt");

                foreach (string file in Directory.GetFiles(tempFolder, "*.jpg"))
                {
                    File.Copy(file, Path.Combine(folder, Path.GetFileName(file)), true);
                }
            }

            //--Show the Damage information in UI
            var oCTR = new CTRReader();
            oCTR.updateFromRedmine(filename);
            string desc = "";
            foreach (string s in oCTR.m_lstDamageInfo)
                desc = desc + s + "\r\n";
            TicketDetails = desc.Trim();

            //Update
            if (oCTR.m_sESN.Length > 0) ESN = oCTR.m_sESN;
            if (oCTR.m_sTSN.Length > 0) TSN = oCTR.m_sTSN;
            if (oCTR.m_sCSN.Length > 0) CSN = oCTR.m_sCSN;

            //Load Images

            //cbImage1.Items.Clear();
            //cbImage2.Items.Clear();
            //cbImage3.Items.Clear(); 

            ZoomedViews.Clear();
            OverviewImages.Clear();
            PartsInformation.Clear();

            Utility.WriteErrorLog("", "", folder);
            string[] imageFiles = Directory.GetFiles(folder, "*.*")
                                                    .Where(file => file.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase)
                                                    || file.EndsWith(".jpeg", StringComparison.OrdinalIgnoreCase)
                                                    || file.EndsWith(".png", StringComparison.OrdinalIgnoreCase)
                                                    || file.EndsWith(".bmp", StringComparison.OrdinalIgnoreCase)
                                                    || file.EndsWith(".gif", StringComparison.OrdinalIgnoreCase)).ToArray();

            ZoomedViews.Add(new KeyValuePair<string, string>("No Image", ""));
            OverviewImages.Add(new KeyValuePair<string, string>("No Image", ""));
            PartsInformation.Add(new KeyValuePair<string, string>("No Image", ""));

            //cbImage1.Items.Add("No Image");
            //cbImage2.Items.Add("No Image");
            //cbImage3.Items.Add("No Image");

            if (imageFiles != null && imageFiles.Length > 0)
            {
                foreach (string img in imageFiles)
                {
                    ZoomedViews.Add(new KeyValuePair<string, string>(System.IO.Path.GetFileName(img), img));
                    OverviewImages.Add(new KeyValuePair<string, string>(System.IO.Path.GetFileName(img), img));
                    PartsInformation.Add(new KeyValuePair<string, string>(System.IO.Path.GetFileName(img), img));
                }
            }
            //cbImage1.SelectedIndex = 0;
            //cbImage2.SelectedIndex = 0;
            //cbImage3.SelectedIndex = 0;

            SelectedOverviewImage = OverviewImages.FirstOrDefault().Value;
            SelectedZoomedView = ZoomedViews.FirstOrDefault().Value;
            SelectedPartInformation = PartsInformation.FirstOrDefault().Value;

        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
        finally
        {
            IsLoading = false;
        }
    }

    private async Task<Ret> GetDataAsync()
    {
        return new Ret
        {
            ESN = "ESN001",
            TSN = "TSN001",
            CSN = "CSN001",
            TicketDetails = "Lorem Ipsum Lorem Ipsum Lorem Ipsum Lorem Ipsum Lorem Ipsum Lorem Ipsum Lorem Ipsum Lorem Ipsum Lorem Ipsum Lorem Ipsum Lorem Ipsum Lorem Ipsum ",

            Images =
        {
            new(@"D:\iHi\CTR\Finding CTR Form Tool\Finding CTR Form Tool\CTR Form Tool\bin\Debug\bin\images\Finding\1.jpg"),
            new(@"D:\iHi\CTR\Finding CTR Form Tool\Finding CTR Form Tool\CTR Form Tool\bin\Debug\bin\images\Finding\R0018608.JPG"),
            new(@"D:\iHi\CTR\Finding CTR Form Tool\Finding CTR Form Tool\CTR Form Tool\bin\Debug\bin\images\Finding\R0018609.JPG"),
            new(@"D:\iHi\CTR\Finding CTR Form Tool\Finding CTR Form Tool\CTR Form Tool\bin\Debug\bin\images\Finding\R0018611.JPG"),
            new(@"D:\iHi\CTR\Finding CTR Form Tool\Finding CTR Form Tool\CTR Form Tool\bin\Debug\bin\images\Finding\R0018612.JPG"),
        }
        };
    }

    public void Initialise()
    {
        try
        {
            WindowTitle = Utility.m_sToolName + " V" + Utility.m_sVersion;

            ReadTheConfigurationFile();
            ReadAscii();

            //URLS
            //string sDir = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);
            List<string> lstLines = Utility.ReadFile(sDir + "\\bin\\redminedata.txt");
            for (int i = 0; i < lstLines.Count; i++)
            {
                string line = lstLines[i].Trim();

                if (line == "loginurl:")
                    loginurl = lstLines[i + 1].Trim();

                if (line == "username:")
                    username = lstLines[i + 1].Trim();

                if (line == "password:")
                    password = lstLines[i + 1].Trim();
            }
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
    }


    private void ReadTheConfigurationFile()
    {
        try
        {
            lstConfigs.Clear();
            dtTypes.Clear();

            //Part Repair Combo
            RepairParts.Clear();   //cbType.Items.Clear();



            lstManualCriteria.Clear();

            string err = "";

            //string xlFile = Utility.CopyFileToTempPath(Utility.m_sBinPath + "Config.xlsx");

            //string sDir = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);
            string xlFile = Utility.CopyFileToTempPath(sDir + "\\bin\\Config.xlsx");

            ExcelPackage oExcel = new ExcelPackage(new FileInfo(xlFile));
            ExcelWorksheet oXlWorkSheet = oExcel.Workbook.Worksheets[1];

            int nRows = oXlWorkSheet.Dimension.Rows + 10;
            for (int r = 3; r <= nRows; r++)
            {
                string s2 = GetCellValue(oXlWorkSheet, r, 2);
                string s3 = GetCellValue(oXlWorkSheet, r, 3);
                string s4 = GetCellValue(oXlWorkSheet, r, 4);
                if (s2.Length > 0 && s3.Length > 0 && s4.Length > 0)
                {
                    //string templatePath = Path.Combine(sDir, Utility.m_sBinPath, "Templates", s4); Needs to check TODO
                    string templatePath = Path.Combine(sDir, "bin\\Templates", s4);
                    if (System.IO.File.Exists(templatePath) == false)
                    {
                        err = err + s4 + "\n";
                    }
                    //Add to list
                    ConfigInfo config = new ConfigInfo();
                    config.m_sType = s2;
                    config.m_sSubType = s3;
                    config.m_sFile = s4;

                    CoordinatesInfo c1 = new CoordinatesInfo();
                    c1.m_fImgTop = float.Parse(GetCellValue(oXlWorkSheet, r, 5));
                    c1.m_fImageHeight = float.Parse(GetCellValue(oXlWorkSheet, r, 6));
                    config.lstCoordinates.Add(c1);

                    lstConfigs.Add(config);

                    if (dtTypes.ContainsKey(s2))
                    {
                        dtTypes[s2].Add(s3);
                    }
                    else
                    {
                        dtTypes.Add(s2, new List<string>() { s3 });
                    }
                }
            }
            //Read Non-conformance type
            List<string> lstNonconformanceType = new List<string>();
            oXlWorkSheet = oExcel.Workbook.Worksheets[2];
            nRows = oXlWorkSheet.Dimension.Rows;
            for (int r = 1; r <= nRows; r++)
            {
                string s2 = GetCellValue(oXlWorkSheet, r, 1);
                if (s2.Length == 0) break;
                lstNonconformanceType.Add(s2);
            }
            //cbNonConformanceType.Items.AddRange(lstNonconformanceType.ToArray());
            //cbNonConformanceType.SelectedIndex = 0;

            nonConformanceTypes.Clear();

            foreach (string s in lstNonconformanceType)
            {
                nonConformanceTypes.Add(new KeyValuePair<string, string>(s, s));
            }

            //cbType.Items.AddRange(dtTypes.Keys.ToArray());

            foreach(var item in dtTypes)
            {
                RepairParts.Add(new KeyValuePair<string, string>(item.Key, item.Key));
            }

            //---- Read the ManulaCriteria ----
            oXlWorkSheet = oExcel.Workbook.Worksheets[3];
            nRows = oXlWorkSheet.Dimension.Rows;
            int Cols = oXlWorkSheet.Dimension.Columns;
            for (int r = 2; r <= nRows; r++)
            {
                string type1 = GetCellValue(oXlWorkSheet, r, 2);
                string type2 = GetCellValue(oXlWorkSheet, r, 3);
                int startrow = r, endrow = r;
                var cell = oXlWorkSheet.Cells[r, 2];
                if (cell.Merge)
                {
                    string mergedAddress = oXlWorkSheet.MergedCells[r, 2];
                    var mergedRange = oXlWorkSheet.Cells[mergedAddress];
                    int mergedRowCount = mergedRange.End.Row - mergedRange.Start.Row + 1;
                    endrow = startrow + mergedRowCount - 1;
                }
                Dictionary<string, List<string>> dt = new Dictionary<string, List<string>>();
                for (int c = 4; c <= Cols; c++)
                {
                    string nc = GetCellValue(oXlWorkSheet, 1, c);
                    if (nc.Length == 0) break;
                    List<string> lst = new List<string>();
                    for (int r1 = startrow; r1 <= endrow; r1++)
                    {
                        string v = GetCellValue(oXlWorkSheet, r1, c).Replace("\n", " ").Trim();
                        if (v.Length > 0 && lst.Contains(v) == false)
                        {
                            lst.Add(v);
                        }
                    }
                    dt.Add(nc, lst);
                }
                //create object
                var mc = new ManualCriteria();
                mc.m_sType = type1;
                mc.m_sSubType = type2;
                mc.dtVals = dt;
                lstManualCriteria.Add(mc);

                r = endrow;
            }
            //---------------------------------
            oExcel.Dispose();
            Utility.DeleteFile(xlFile);


            if (err.Length > 0)
            {
                Utility.WriteErrorLog(err);
                err = "Some files are missing. Please update the configuration file properly\n\n" + err.Trim();
                Utility.WarnUser(err);
                Application.Current.Shutdown();
            }
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog("", ee.Message, ee.StackTrace);
            MessageBox.Show("Failed to read the configuration file. Please check if the Excel file is under IHI protection.\n\nKindly contact the CYIENT team for assistance.",
            "Configuration Error",
            MessageBoxButton.OK,
            MessageBoxImage.Error);
            Application.Current.Shutdown();
        }
    }

    private void ReadAscii()
    {
        try
        {
            //Read Ascii
            //string sDir = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);
            List<string> lst = Utility.ReadFile(sDir + "\\bin\\ascii.txt");
            foreach (string s in lst)
            {
                List<string> lstSplits = Utility.SplitString(s, "=");
                AsciiInfo oo = new AsciiInfo();
                oo.m_nAsciiNum = int.Parse(lstSplits[0]);
                oo.m_sVal = lstSplits[1];
                lstAscii.Add(oo);
            }
        }
        catch (Exception ee)
        {
            //Utility.WriteErrorLog(ee);
        }
    }

    string GetCellValue(ExcelWorksheet oXlWorkSheet, int nRow, int nCol)
    {
        string sValue = "";
        try
        {
            if (nRow > 0 && nCol > 0)
            {
                object oCell = oXlWorkSheet.Cells[nRow, nCol].Value;
                if (oCell != null)
                {
                    sValue = oCell.ToString().Trim();
                }
            }
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
        return sValue;
    }
}

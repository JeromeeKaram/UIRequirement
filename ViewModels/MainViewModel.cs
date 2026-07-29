using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using CTR_Form_Tool;
using System.Collections.ObjectModel;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using UIRequirement.Models;

namespace UIRequirement.ViewModels;

public partial class MainViewModel : ObservableObject
{
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
    private string ticketDescription;

    public ObservableCollection<DamageInfo> Damages { get; } = new();

    [RelayCommand]
    private async Task Load()
    {
        //MessageBox.Show("Load clicked");
        await InputFolderSelected();
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
            string sDir = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);
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
                //string input = InputFolderPath?.Trim() ?? "";

                //string findingTicket = "";
                //string encTicket = "";

                //if (input.Contains(','))
                //{
                //    string[] parts = input.Split(',');

                //    findingTicket = parts[0].Trim();

                //    if (parts.Length > 1)
                //        encTicket = parts[1].Trim();
                //}
                //else
                //{
                //    findingTicket = input;
                //}

                //int imageCount = await getRedmineImages(findingTicket, encTicket);

                //await WaitForImagesAsync(FindingFolder, imageCount);

                //filename = Path.Combine(sDir, "bin", "description.txt");

                //string encFile = Path.Combine(sDir, "bin", "ENC.txt");

                //if (File.Exists(encFile))
                //{
                //    string[] lines = File.ReadAllLines(encFile);

                //    if (lines.Length > 0) ESN = lines[0];
                //    if (lines.Length > 1) TSN = lines[1];
                //    if (lines.Length > 2) CSN = lines[2];
                //}
            }
            else
            {
                string tempFolder = Path.Combine(sDir, "bin", "test");

                filename = Path.Combine(tempFolder, "description_15April2025.txt");

                foreach (string file in Directory.GetFiles(tempFolder, "*.jpg"))
                {
                    File.Copy(file,Path.Combine(folder, Path.GetFileName(file)), true);
                }
            }

            //-- Show the Damage information in UI
            var oCTR = new CTRReader();
            oCTR.updateFromRedmine(filename);
            string desc = "";
            foreach (string s in oCTR.m_lstDamageInfo)
                desc = desc + s + "\r\n";
            TicketDescription = desc.Trim();

            //Update
            if (oCTR.m_sESN.Length > 0) ESN = oCTR.m_sESN;
            if (oCTR.m_sTSN.Length > 0) TSN = oCTR.m_sTSN;
            if (oCTR.m_sCSN.Length > 0) CSN = oCTR.m_sCSN;

            //Load Images
            //cbImage1.Items.Clear();
            //cbImage2.Items.Clear();


            Utility.WriteErrorLog("", "", folder);
            string[] imageFiles = Directory.GetFiles(folder, "*.*")
                                                    .Where(file => file.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase)
                                                    || file.EndsWith(".jpeg", StringComparison.OrdinalIgnoreCase)
                                                    || file.EndsWith(".png", StringComparison.OrdinalIgnoreCase)
                                                    || file.EndsWith(".bmp", StringComparison.OrdinalIgnoreCase)
                                                    || file.EndsWith(".gif", StringComparison.OrdinalIgnoreCase)).ToArray();


            //cbImage1.Items.Add("No Image");
            //cbImage2.Items.Add("No Image");
            //cbImage3.Items.Add("No Image");

            //if (imageFiles != null && imageFiles.Length > 0)
            //{
            //    foreach (string img in imageFiles)
            //    {
            //        cbImage1.Items.Add(System.IO.Path.GetFileName(img));
            //        cbImage2.Items.Add(System.IO.Path.GetFileName(img));
            //        cbImage3.Items.Add(System.IO.Path.GetFileName(img));
            //    }
            //}
            //cbImage1.SelectedIndex = 0;
            //cbImage2.SelectedIndex = 0;
            //cbImage3.SelectedIndex = 0;

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

    //private async Task<int> getRedmineImages(string issueId, string encticket)
    //{

    //    IWebDriver driver = null;
    //    int count = 0;

    //    try
    //    {

    //        string baseUrl = "http://0000a1000452";
    //        string loginPage = baseUrl + "/ics/my/page";
    //        //  string issueId = tbInputFolderPath.Text.Trim();
    //        string targetPage = $"http://0000a1000452/ics/issues/{issueId}";
    //        string enctargetPage = $"http://0000a1000452/ics/issues/{encticket}";
    //        //string targetPage = "http://0000a1000452/ics/issues/" + textBox4.Text.Trim();
    //        string folderPath = FindingFolder;
    //        //tbInputFolderPath.Text + "//Finding";

    //        Directory.CreateDirectory(folderPath);

    //        ChromeOptions options = new ChromeOptions();
    //        options.AddArgument("--start-maximized");

    //        driver = new ChromeDriver(options);

    //        // ================= LOGIN =================
    //        driver.Navigate().GoToUrl(loginPage);

    //        WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
    //        wait.Until(d => d.FindElements(By.Id("username")).Count > 0);

    //        driver.FindElement(By.Id("username")).SendKeys(username.Trim());
    //        driver.FindElement(By.Id("password")).SendKeys(password.Trim());
    //        driver.FindElement(By.Id("kc-login")).Click();


    //        // ================= OPEN ENC TARGET PAGE =================
    //        File.WriteAllLines(sDir + "\\bin\\description.txt", new string[1] { "" });
    //        File.WriteAllLines(sDir + "\\bin\\ENC.txt", new string[1] { "" });
    //        if (encticket != "")
    //        {
    //            driver.Navigate().GoToUrl(enctargetPage);

    //            WebDriverWait wait1 = new WebDriverWait(driver, TimeSpan.FromSeconds(15));

    //            wait1.Until(d =>
    //            {
    //                var body = d.FindElement(By.TagName("body"));
    //                return !string.IsNullOrWhiteSpace(body.Text);
    //            });

    //            // Get all visible text from page
    //            string pageText = driver.FindElement(By.TagName("body")).Text;

    //            //  driver.Quit();

    //            // Regex patterns
    //            string esn = Regex.Match(pageText, @"ESN:\s*(\S+)").Groups[1].Value;
    //            string tsn = Regex.Match(pageText, @"TSN:\s*(\S+)").Groups[1].Value;
    //            string csn = Regex.Match(pageText, @"CSN:\s*(\S+)").Groups[1].Value;

    //            result.Add(esn);
    //            result.Add(tsn);
    //            result.Add(csn);
    //            File.WriteAllLines(sDir + "\\bin\\ENC.txt", result.ToArray());
    //        }


    //        // Wait for login cookies
    //        //  wait.Until(d => d.Manage().Cookies.AllCookies.Count > 0);


    //        // ================= OPEN TARGET PAGE =================
    //        driver.Navigate().GoToUrl(targetPage);

    //        wait.Until(d =>
    //            d.FindElements(By.CssSelector("span.existing-attachment")).Count > 0
    //        );

    //        // get page source
    //        string pageSource = driver.PageSource;
    //        WebDriverWait wait3 = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
    //        wait3.Until(d => d.FindElement(By.TagName("body")));

    //        string FindingpageInnerText = (string)((IJavaScriptExecutor)driver)
    //            .ExecuteScript("return document.body.innerText;");
    //        Utility.WriteErrorLog("", "", FindingpageInnerText);

    //        // find the description wiki div
    //        var wikiDiv = driver.FindElement(By.CssSelector("div.description div.wiki"));

    //        // get visible text exactly as browser shows it
    //        string extracteddata = FindingpageInnerText;// wikiDiv.Text;

    //        // split by lines
    //        var lines = extracteddata.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);

    //        // skip the line that contains "ODR No."
    //        descriptionText = string.Join(
    //            Environment.NewLine,
    //            lines.SkipWhile(l => !l.StartsWith("ODR"))
    //                 .Skip(1)
    //        ).Trim();

    //        // save it
    //        File.WriteAllLines(sDir + "\\bin\\description.txt", lines.ToArray());

    //        // ================= COPY COOKIES =================
    //        CookieContainer cookieContainer = new CookieContainer();

    //        foreach (var c in driver.Manage().Cookies.AllCookies)
    //        {
    //            cookieContainer.Add(
    //                new System.Net.Cookie(c.Name, c.Value, c.Path, c.Domain)
    //            );
    //        }

    //        HttpClientHandler handler = new HttpClientHandler
    //        {
    //            UseCookies = true,
    //            CookieContainer = cookieContainer
    //        };

    //        // ================= DOWNLOAD ATTACHMENTS =================
    //        using (HttpClient client = new HttpClient(handler))
    //        {
    //            var attachments = driver.FindElements(
    //                By.CssSelector("span.existing-attachment")
    //            );



    //            foreach (var attachment in attachments)
    //            {
    //                // Get filename and attachment ID from the page
    //                string fileName = attachment
    //                    .FindElement(By.CssSelector("input.filename"))
    //                    .GetAttribute("value");

    //                string attachmentId = attachment
    //                    .FindElement(By.CssSelector("input.deleted_attachment"))
    //                    .GetAttribute("value");

    //                // Build the **full correct URL** for the image/pdf
    //                string downloadUrl = $"{baseUrl}/ics/attachments/download/{attachmentId}/{fileName}";
    //                Utility.WriteErrorLog("", "", downloadUrl);

    //                HttpResponseMessage response = await client.GetAsync(downloadUrl);

    //                if (!response.IsSuccessStatusCode)
    //                    continue;

    //                string contentType = response.Content.Headers.ContentType?.MediaType ?? "";

    //                // Only download images or PDFs
    //                //if (!contentType.StartsWith("image") && !contentType.Contains("pdf"))

    //                // Only download images (NO PDFs)
    //                if (!contentType.StartsWith("image", StringComparison.OrdinalIgnoreCase))
    //                    continue;

    //                byte[] data = await response.Content.ReadAsByteArrayAsync();

    //                string savePath = Path.Combine(folderPath, fileName);
    //                File.WriteAllBytes(savePath, data);

    //                count++;
    //            }

    //            //MessageBox.Show($"Downloaded {count} files successfully.");
    //        }

    //    }

    //    catch (Exception ex)
    //    {
    //        MessageBox.Show("Error: " + ex.Message);
    //    }
    //    finally
    //    {
    //        driver.Quit();
    //    }

    //    return count;
    //}
}

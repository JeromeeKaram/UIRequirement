

//Import libraries
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Media.Imaging;


namespace UIRequirement;

class PDFReader
{
    public List<PDFPageInfo> m_lstPdfs = new List<PDFPageInfo>();
    //*******************************************************
    //Function  : ReadPDF
    //Purpose   : Read the PDF
    //*******************************************************
    public void ReadPDF_dummy(string inputFile)
    {
        try
        {
            //copy the pdf to temp
            string sPdfFile = Utility.m_sTempPath + "123.pdf";// System.IO.Path.GetFileName(inputFile);
            System.IO.File.Copy(inputFile, sPdfFile, true);
            //----------------------------- Read PDF -------------------------------
            //----------------------------------------------------------------------
            PdfReader reader = new PdfReader(sPdfFile);
            int nPages = reader.NumberOfPages;
            int pg = 1;
            PDFPageInfo oPage = new PDFPageInfo(1);
            m_lstPdfs.Add(oPage);

            iTextSharp.text.Rectangle rect = reader.GetPageSize(pg);
            oPage.m_fPageHeight = rect.Height;
            oPage.m_fPageWidth = rect.Width;
            oPage.m_nPageNumber = pg;
            //
            string text = PdfTextExtractor.GetTextFromPage(reader, pg);
            System.IO.File.WriteAllLines(Utility.m_sTempPath + "op.txt", new string[1] { text });
            oPage.m_lstPageLinesFromFile = Utility.ReadFile(Utility.m_sTempPath + "op.txt");

            var t = new MyLocationTextExtractionStrategy();
            text = PdfTextExtractor.GetTextFromPage(reader, pg, t);
            for (int i = 0; i < t.myPoints.Count; i++)
            {
                var p = t.myPoints[i];
                int x = 5;
            }

            bool resp = ReadPDFPage(t, pg, oPage, 1, out oPage.m_lstHeights);
            reader.Close();
            Utility.DeleteFile(sPdfFile);
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
    }
    public void ReadPDF(string inputFile)
    {
        try
        {
            //copy the pdf to temp
            string sPdfFile = Utility.m_sTempPath + "123.pdf";// System.IO.Path.GetFileName(inputFile);
            System.IO.File.Copy(inputFile, sPdfFile, true);
            //----------------------------- Read PDF -------------------------------
            //----------------------------------------------------------------------
            PdfReader reader = new PdfReader(sPdfFile);
            int nPages = reader.NumberOfPages;
            //for (int pg = 4; pg <= 4; pg++)
            for (int pg = 1; pg <= nPages; pg++)
            {
                //if (pg != 5) continue;
                PDFPageInfo oPage = new PDFPageInfo(pg);
                m_lstPdfs.Add(oPage);

                iTextSharp.text.Rectangle rect = reader.GetPageSize(pg);
                oPage.m_fPageHeight = rect.Height;
                oPage.m_fPageWidth = rect.Width;
                oPage.m_nPageNumber = pg;
                //
                string text = PdfTextExtractor.GetTextFromPage(reader, pg);
                System.IO.File.WriteAllLines(Utility.m_sTempPath + "op.txt", new string[1] { text });
                oPage.m_lstPageLinesFromFile = Utility.ReadFile(Utility.m_sTempPath + "op.txt");

                //-->Read Page by page, 3 ways to read...Exact match then exit
                int times = 3, index = -1;
                double percentage = 0.0;
                for (int x = 0; x < times; x++)
                {
                    var t = new MyLocationTextExtractionStrategy();
                    text = PdfTextExtractor.GetTextFromPage(reader, pg, t);
                    bool resp = ReadPDFPage(t, pg, oPage, x, out oPage.m_lstHeights);
                    if (resp)
                    {
                        index = -1;
                        break;
                    }
                    //check %
                    double d = (double)oPage.m_lstPageLinesFromFile.Count / (double)oPage.m_lstHeights.Count;
                    if (d > percentage)
                    {
                        index = x;
                        percentage = d;
                    }
                }
                //No exact match, take highest match
                if (index != -1)
                {
                    var t = new MyLocationTextExtractionStrategy();
                    text = PdfTextExtractor.GetTextFromPage(reader, pg, t);
                    bool resp = ReadPDFPage(t, pg, oPage, index, out oPage.m_lstHeights);
                }
            }
            reader.Close();
            Utility.DeleteFile(sPdfFile);


            //---Handle the wrong pdf pages---
            if (System.IO.File.Exists(Utility.m_sBinPath + "wrongPDFpage.txt"))
            {
                List<string> lstwrongdata = Utility.ReadFile(Utility.m_sBinPath + "wrongPDFpage.txt");
                foreach (PDFPageInfo oPage in m_lstPdfs)
                {
                    bool yes = false;
                    foreach (string ss in oPage.m_lstPageLines)
                    {
                        foreach (string sx in lstwrongdata)
                        {
                            if (ss.Contains(sx))
                            {
                                yes = true; break;
                            }
                        }
                        if (yes) break;
                    }
                    if (yes)
                    {
                        oPage.m_lstPageLines.Clear();
                        oPage.m_lstPageLinesFromFile.Clear();
                        oPage.m_lstHeights.Clear();
                        Utility.WriteErrorLog("", "", "Invalid PDF page, File Name : " + inputFile + ", Page Number : " + oPage.m_nPageNumber.ToString());
                    }
                }
            }

            //Update the Ht info
            foreach (PDFPageInfo opage in m_lstPdfs)
            {
                foreach (HeightInfo ht in opage.m_lstHeights)
                {
                    ht.m_oPage = opage;
                }
            }
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
    }
    public void SplitAndSave(string inputPath, string outputPath)
    {
        try
        {
            System.IO.FileInfo file = new System.IO.FileInfo(inputPath);
            //string name = file.Name.Substring(0, file.Name.LastIndexOf("."));

            using (PdfReader reader = new PdfReader(inputPath))
            {

                for (int pagenumber = 1; pagenumber <= reader.NumberOfPages; pagenumber++)
                {
                    string filename = pagenumber.ToString() + ".pdf";

                    iTextSharp.text.Document document = new iTextSharp.text.Document();
                    PdfCopy copy = new PdfCopy(document, new System.IO.FileStream(outputPath + "\\" + filename, System.IO.FileMode.Create));

                    document.Open();

                    copy.AddPage(copy.GetImportedPage(reader, pagenumber));

                    document.Close();
                }
                //return reader.NumberOfPages;
            }
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
    }
    //*******************************************************
    //Function  : UpdateDataByRotation
    //Purpose   : handle rotated pdf's
    //*******************************************************
    public void UpdateDataByRotation(MyLocationTextExtractionStrategy oPdfData, int type)
    {
        try
        {
            foreach (RectAndText rt in oPdfData.myPoints)
            {
                if (type == 0)//No change
                {
                }
                else if (type == 1)
                {
                    float Top = rt.Rect.Top;
                    float Bottom = rt.Rect.Bottom;
                    float Left = rt.Rect.Left;
                    float Right = rt.Rect.Right;
                    rt.Rect.Top = Left;
                    rt.Rect.Bottom = Right;
                    rt.Rect.Left = Top;
                    rt.Rect.Right = Bottom;
                }
                else if (type == 2)
                {
                    float Top = rt.Rect.Top;
                    float Bottom = rt.Rect.Bottom;
                    float Left = rt.Rect.Left;
                    float Right = rt.Rect.Right;
                    rt.Rect.Top = Right;
                    rt.Rect.Bottom = Left;
                    rt.Rect.Left = Bottom;
                    rt.Rect.Right = Top;
                }
            }
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
    }
    //*******************************************************
    //Function  : ReadPDFPage
    //Purpose   : reafd each pdf page data and store in data structures
    //*******************************************************
    private bool ReadPDFPage(MyLocationTextExtractionStrategy oPdfData, int page, PDFPageInfo oPage, int type, out List<HeightInfo> lstHeights)
    {
        bool response = false;
        List<string> lstData = new List<string>();
        lstHeights = new List<HeightInfo>();
        try
        {
            double ytol = 0.001;

            UpdateDataByRotation(oPdfData, type);
            //--> Get all the  Y points
            List<float> lstDbl = new List<float>();
            List<string> lst = new List<string>();
            for (int i = 0; i < oPdfData.myPoints.Count; i++)
            {
                var p = oPdfData.myPoints[i];
                lstDbl.Add(p.Rect.Bottom);
            }
            //--> Delete the duplicate points
            for (int i = 0; i < lstDbl.Count; i++)
            {
                for (int j = i + 1; j < lstDbl.Count; j++)
                {
                    if (Math.Abs(lstDbl[i] - lstDbl[j]) < ytol)
                    {
                        lstDbl.RemoveAt(j);
                        j--;
                    }
                }
            }
            //--> Sort in Y direction
            for (int i = 0; i < lstDbl.Count; i++)
            {
                for (int j = i + 1; j < lstDbl.Count; j++)
                {
                    if (lstDbl[i] < lstDbl[j])
                    {
                        float dummy = lstDbl[j];
                        lstDbl[j] = lstDbl[i];
                        lstDbl[i] = dummy;
                    }
                }
            }
            //-->combine the line text based on Y direction

            for (int j = 0; j < lstDbl.Count; j++)
            {
                HeightInfo oo = new HeightInfo();
                oo.m_fVal = lstDbl[j];
                for (int i = 0; i < oPdfData.myPoints.Count; i++)
                {
                    var p = oPdfData.myPoints[i];
                    if (Math.Abs(lstDbl[j] - p.Rect.Bottom) < ytol)
                    {
                        oo.m_sValue = oo.m_sValue + p.Text;
                        oo.m_lstRectAndText.Add(p);
                        oPdfData.myPoints.RemoveAt(i);
                        i--;
                    }
                }
                lstHeights.Add(oo);
            }
            //------------- Segregate the texts----
            //-->find the text contain rows and non text text contain rows
            List<HeightInfo> lst1 = new List<HeightInfo>();
            List<HeightInfo> lst2 = new List<HeightInfo>();
            for (int j = 0; j < lstHeights.Count; j++)
            {
                lstHeights[j].m_nIndex = j + 1;
                string sValue = "";
                foreach (RectAndText p in lstHeights[j].m_lstRectAndText)
                {
                    sValue = sValue + p.Text;
                }
                //remove eiphen
                string ss1 = sValue.Replace("Ѹ", "").Replace("−", "").Trim();
                string ss2 = sValue.Trim();
                if (ss1.Length > 0)
                {
                    lst1.Add(lstHeights[j]);
                    lstHeights[j].m_sValue = ss1;
                }
                else if (ss2.Length > 0)
                {
                    lst2.Add(lstHeights[j]);
                }
            }
            //--> link items with some tolerance
            double dTol = 3.0;
            for (int j = 0; j < lst1.Count; j++)
            {
                for (int k = j + 1; k < lst1.Count; k++)
                {
                    if (Math.Abs(lst1[j].m_fVal - lst1[k].m_fVal) < dTol)
                    {
                        foreach (RectAndText p in lst1[k].m_lstRectAndText)
                        {
                            lst1[j].m_lstRectAndText.Add(p);
                        }
                        lst1.RemoveAt(k);
                        k--;
                    }
                }
            }
            for (int j = 0; j < lst1.Count; j++)
            {
                for (int k = 0; k < lst2.Count; k++)
                {
                    if (Math.Abs(lst1[j].m_fVal - lst2[k].m_fVal) < dTol)
                    {
                        foreach (RectAndText p in lst2[k].m_lstRectAndText)
                        {
                            lst1[j].m_lstRectAndText.Add(p);
                        }
                        lst2.RemoveAt(k);
                        k--;
                    }
                }
            }
            //Add
            lstHeights.Clear();
            for (int j = 0; j < lst1.Count; j++)
            {
                lstHeights.Add(lst1[j]);
            }




            //Sort the data in X direction
            int sign = 1;
            foreach (HeightInfo oo in lstHeights)
            {
                List<float> ll = new List<float>();
                for (int p = 0; p < oo.m_lstRectAndText.Count; p++)
                {
                    ll.Add(oo.m_lstRectAndText[p].Rect.Left);
                    for (int q = p + 1; q < oo.m_lstRectAndText.Count; q++)
                    {
                        if (oo.m_lstRectAndText[p].Rect.Left * sign > oo.m_lstRectAndText[q].Rect.Left * sign)
                        {
                            RectAndText pp = oo.m_lstRectAndText[p];
                            oo.m_lstRectAndText[p] = oo.m_lstRectAndText[q];
                            oo.m_lstRectAndText[q] = pp;
                        }
                    }
                }
            }
            //Create text
            foreach (HeightInfo oo in lstHeights)
            {
                oo.m_sValue = "";
                for (int x = 0; x < oo.m_lstRectAndText.Count; x++)
                {
                    string sSpace = "";
                    if (x > 0)
                    {
                        float left = oo.m_lstRectAndText[x].Rect.Left;
                        if (oo.m_lstRectAndText[x].Rect.Right < left) left = oo.m_lstRectAndText[x].Rect.Right;
                        float right = oo.m_lstRectAndText[x - 1].Rect.Right;
                        if (oo.m_lstRectAndText[x - 1].Rect.Right < oo.m_lstRectAndText[x - 1].Rect.Left) right = oo.m_lstRectAndText[x - 1].Rect.Left;

                        //  float f = oo.m_lstRectAndText[x].Rect.Left - oo.m_lstRectAndText[x - 1].Rect.Right;
                        float f = left - right;
                        if (f > 1.5f) sSpace = " ";
                    }
                    oo.m_sValue = oo.m_sValue + sSpace + oo.m_lstRectAndText[x].Text;
                }
            }

            //Add the final text
            foreach (HeightInfo ht in lstHeights)
            {
                string ss = ht.m_sValue.Replace("Ѹ", "").Replace("−", "").Trim();
                if (ss.Length > 0)
                {
                    ht.m_sValue = cleanString(ht.m_sValue.Replace("Ѹ", "-").Trim());
                    lstData.Add(ht.m_sValue);
                }
            }


            if (lstData.Count > 0)
            {
                oPage.m_lstPageLines = lstData;

                string s1 = Utility.CombineString(Utility.SplitString(oPage.m_lstPageLinesFromFile[0], " "), " ");
                string s2 = Utility.CombineString(Utility.SplitString(lstData[0], " "), " ");
                string s3 = Utility.CombineString(Utility.SplitString(oPage.m_lstPageLinesFromFile[oPage.m_lstPageLinesFromFile.Count - 1], " "), " ");
                string s4 = Utility.CombineString(Utility.SplitString(lstData[lstData.Count - 1], " "), " ");

                //Straight solution
                if (s1 == s2 || s3 == s4)
                {
                    oPage.m_lstPageLines = lstData;
                    response = true;
                }
                //reverse solution
                else if (s1 == s4 && s3 == s2)
                {
                    oPage.m_lstPageLines = lstData;
                    lstData.Reverse();
                    lstHeights.Reverse();
                    response = true;
                }
                else
                {
                }
            }
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog("", ee.Message, ee.StackTrace + "; Page : " + page.ToString());
        }
        return response;
    }
    //*******************************************************
    //Function  : SplitPage
    //Purpose   : split the pdf into multiple pdfs(each page one pdf)
    //*******************************************************
    public void SplitPage(string sourcePdfPath, string outputPdfPath)
    {
        try
        {
            PdfReader reader = new PdfReader(sourcePdfPath);
            int startPage = 1;
            int endPage = reader.NumberOfPages;


            for (int i = startPage; i <= endPage; i++)
            {
                iTextSharp.text.Document sourceDocument = new iTextSharp.text.Document(reader.GetPageSizeWithRotation(startPage));
                PdfCopy pdfCopyProvider = new PdfCopy(sourceDocument, new System.IO.FileStream(outputPdfPath + i.ToString() + ".pdf", System.IO.FileMode.Create));
                sourceDocument.Open();
                PdfImportedPage importedPage = pdfCopyProvider.GetImportedPage(reader, i); pdfCopyProvider.AddPage(importedPage);
                sourceDocument.Close();
            }
            reader.Close();
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
    }
    //*******************************************************
    //Function  : cleanString
    //Purpose   : clean the string
    //*******************************************************
    private string cleanString(string str)
    {
        try
        {
            string s1 = str.Replace(((char)8211).ToString(), ((char)45).ToString());
            string s2 = s1.Replace("”", "\"");
            return s2.Trim();
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
        return str.Trim();
    }
}




public class RectAndText
{
    public iTextSharp.text.Rectangle Rect;
    public String Text;
    public iTextSharp.text.pdf.parser.LocationTextExtractionStrategy.TextChunk m_oChunk = null;
    public int m_nMode = 0;
    public RectAndText(iTextSharp.text.Rectangle rect, String text)
    {
        this.Rect = rect;
        this.Text = text;
    }
}
public class MyLocationTextExtractionStrategy5 : LocationTextExtractionStrategy
{
    //Hold each coordinate
    public List<RectAndText> myPoints = new List<RectAndText>();
    //The string that we're searching for
    public String TextToSearchFor { get; set; }
    //How to compare strings
    public System.Globalization.CompareOptions CompareOptions { get; set; }
    public MyLocationTextExtractionStrategy5(String textToSearchFor, System.Globalization.CompareOptions compareOptions = System.Globalization.CompareOptions.None)
    {
        this.TextToSearchFor = textToSearchFor;
        this.CompareOptions = compareOptions;
    }
    //Automatically called for each chunk of text in the PDF
    public override void RenderText(TextRenderInfo renderInfo)
    {
        base.RenderText(renderInfo);
        //See if the current chunk contains the text
        var startPosition = System.Globalization.CultureInfo.CurrentCulture.CompareInfo.IndexOf(renderInfo.GetText(), this.TextToSearchFor, this.CompareOptions);
        //If not found bail
        if (startPosition < 0)
        {
            return;
        }
        //Grab the individual characters
        var chars = renderInfo.GetCharacterRenderInfos().Skip(startPosition).Take(this.TextToSearchFor.Length).ToList();
        //Grab the first and last character
        var firstChar = chars.First();
        var lastChar = chars.Last();
        //Get the bounding box for the chunk of text
        var bottomLeft = firstChar.GetDescentLine().GetStartPoint();
        var topRight = lastChar.GetAscentLine().GetEndPoint();
        //Create a rectangle from it
        var rect = new iTextSharp.text.Rectangle(
        bottomLeft[Vector.I1],
        bottomLeft[Vector.I2],
        topRight[Vector.I1],
        topRight[Vector.I2]
        );
        //Add this to our main collection
        this.myPoints.Add(new RectAndText(rect, this.TextToSearchFor));
    }
}
public class MyLocationTextExtractionStrategy : LocationTextExtractionStrategy
{
    //Hold each coordinate
    public List<RectAndText> myPoints = new List<RectAndText>();

    //Automatically called for each chunk of text in the PDF
    public override void RenderText(TextRenderInfo renderInfo)
    {
        base.RenderText(renderInfo);

        //Get the bounding box for the chunk of text
        var bottomLeft = renderInfo.GetDescentLine().GetStartPoint();
        var topRight = renderInfo.GetAscentLine().GetEndPoint();


        //Create a rectangle from it
        var rect = new iTextSharp.text.Rectangle(
                                                bottomLeft[Vector.I1],
                                                bottomLeft[Vector.I2],
                                                topRight[Vector.I1],
                                                topRight[Vector.I2]
                                                );
        //Add this to our main collection
        RectAndText rat = new RectAndText(rect, renderInfo.GetText());
        this.myPoints.Add(rat);
    }
}
public class PDFUtility
{
    //*******************************************************
    //Function  : getAttribute
    //Purpose   : get the pdf attribute value
    //*******************************************************
    public string getAttribute(HeightInfo ht, string attribute)
    {
        string attributeValue = "";
        try
        {
            string content = "";
            for (int x = 0; x < ht.m_lstRectAndText.Count; x++)
            {

                if (x == 0)
                {
                    content = ht.m_lstRectAndText[x].Text;
                }
                else
                {
                    string sSpace = "";
                    float f = ht.m_lstRectAndText[x].Rect.Left - ht.m_lstRectAndText[x - 1].Rect.Right;
                    if (f > 1.5f) sSpace = " ";
                    content = content + sSpace + ht.m_lstRectAndText[x].Text;
                }

                //check it
                if (Utility.MyStringComparison(content, attribute, 2, false) == true)//content.EndsWith(attribute))
                {
                    //now do it
                    for (int y = x + 1; y < ht.m_lstRectAndText.Count; y++)
                    {
                        if (y == x + 1)
                        {
                            attributeValue = ht.m_lstRectAndText[y].Text;
                        }
                        else
                        {
                            string sSpace = "";
                            float f = ht.m_lstRectAndText[y].Rect.Left - ht.m_lstRectAndText[y - 1].Rect.Right;
                            if (f > 10f && attributeValue.Trim().Length > 0) break;
                            else if (f > 1.5f) sSpace = " ";
                            attributeValue = attributeValue + sSpace + ht.m_lstRectAndText[y].Text;
                        }
                    }
                    break;
                }
                else if (Utility.MyStringComparison(content, attribute, 1, false) == true)
                {
                    attributeValue = Utility.SplitString("ww" + content, attribute)[1];
                    break;
                }
            }
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
        return attributeValue.Trim();
    }
    //*******************************************************
    //Function  : GetStringInRange
    //Purpose   : get string from given range in pdf
    //*******************************************************
    public string GetStringInRange(List<WordInfo> lstWords, double min, double max, ref WordInfo ofoundWord)
    {
        string txt = "";
        try
        {
            for (int i = 0; i < lstWords.Count; i++)
            {
                double dratio = (lstWords[i].m_dMax - lstWords[i].m_dMin) / 10.0;
                for (int j = 0; j <= 10; j++)
                {
                    double dv = lstWords[i].m_dMin + dratio * j;
                    if (dv >= min && dv <= max)
                    {
                        ofoundWord = lstWords[i];
                        txt = lstWords[i].m_sTxt;
                        break;
                    }
                }
                if (txt.Length > 0) break;
            }
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
        return txt;
    }

    public bool IsStringInRange(WordInfo oword, double min, double max)
    {
        try
        {
            double dratio = (oword.m_dMax - oword.m_dMin) / 10.0;
            for (int j = 0; j <= 10; j++)
            {
                double dv = oword.m_dMin + dratio * j;
                if (dv >= min && dv <= max)
                {
                    return true;
                }
            }
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
        return false;
    }
    //*******************************************************
    //Function  : GetStringInRange
    //Purpose   : get string from given range in pdf
    //*******************************************************
    public string GetStringInRange(List<WordInfo> lstWords, double min, double max, ref List<WordInfo> lstfounds)
    {
        string txt = "";
        try
        {
            for (int i = 0; i < lstWords.Count; i++)
            {
                double dratio = (lstWords[i].m_dMax - lstWords[i].m_dMin) / 10.0;
                for (int j = 0; j <= 10; j++)
                {
                    double dv = lstWords[i].m_dMin + dratio * j;
                    if (dv >= min && dv <= max)
                    {
                        lstfounds.Add(lstWords[i]);
                        txt = txt + lstWords[i].m_sTxt + " ";
                        break;
                    }
                }
            }
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
        return txt.Trim();
    }
    //*******************************************************
    //Function  : GetStringInRange
    //Purpose   : get string from given range in pdf
    //*******************************************************
    public WordInfo GetStringInRange(List<WordInfo> lstWords, double min, double max)
    {
        WordInfo ofoundWord = new WordInfo();
        try
        {
            List<WordInfo> lstfounds = new List<WordInfo>();
            for (int i = 0; i < lstWords.Count; i++)
            {
                //bool byes = false;
                double dratio = (lstWords[i].m_dMax - lstWords[i].m_dMin) / 10.0;
                for (int j = 0; j <= 10; j++)
                {
                    double dv = lstWords[i].m_dMin + dratio * j;
                    if (dv >= min && dv <= max)
                    {
                        //byes = true;
                        lstfounds.Add(lstWords[i]);
                        break;
                    }
                }
            }
            //create the final one
            if (lstfounds.Count > 0)
            {
                string txt = "";
                List<RectAndText> lstTexts = new List<RectAndText>();
                foreach (WordInfo oo in lstfounds)
                {
                    txt = txt + oo.m_sTxt + " ";
                    foreach (RectAndText rt in oo.m_lstRects)
                        lstTexts.Add(rt);
                }
                ofoundWord.m_dMax = lstfounds[lstfounds.Count - 1].m_dMax;
                ofoundWord.m_dMin = lstfounds[0].m_dMin;
                ofoundWord.m_fVal = lstfounds[0].m_fVal;
                ofoundWord.m_lstRects = lstTexts;
                ofoundWord.m_sTxt = txt.Trim();
            }
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
        return ofoundWord;
    }
    public string GetStringInRangeWithPlusTolerance(List<WordInfo> lstWords, double dMin, double dMax)
    {
        WordInfo ofoundWord = null;
        string str = GetStringInRange(lstWords, dMin, dMax, ref ofoundWord);
        string newstr = str;
        try
        {
            if (str.Length == 0)
            {
                if (dMax > 0.001) //Means title exists in the table
                {
                    double dx = (dMax - dMin) / 10.0;
                    for (int p = 1; p < 10; p++)
                    {
                        newstr = GetStringInRange(lstWords, dMin - p * dx, dMax + p * dx, ref ofoundWord);
                        if (newstr.Length > 0) break;
                    }
                }
            }
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
        return newstr;
    }
    public string GetStringInRangeWithPlusTolerance5(List<WordInfo> lstWords, double dMin, double dMax, int runs)
    {
        List<WordInfo> lstfounds = new List<WordInfo>();
        string str = GetStringInRange(lstWords, dMin, dMax, ref lstfounds);
        string newstr = str;
        try
        {
            if (str.Length == 0)
            {
                if (dMax > 0.001) //Means title exists in the table
                {
                    double dx = (dMax - dMin) / 10.0;
                    for (int p = 1; p < runs; p++)
                    {
                        newstr = GetStringInRange(lstWords, dMin - p * dx, dMax + p * dx, ref lstfounds);
                        if (newstr.Length > 0) break;
                    }
                }
            }
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
        return newstr;
    }
    public string GetStringInRangeWithPlusTolerance6(List<WordInfo> lstWords, double dMin, double dMax, int num)
    {
        WordInfo ofoundWord = null;
        string str = GetStringInRange(lstWords, dMin, dMax, ref ofoundWord);
        string newstr = str;
        try
        {
            if (str.Length == 0)
            {
                if (dMax > 0.001) //Means title exists in the table
                {
                    double dx = (dMax - dMin) / 10.0;
                    for (int p = 1; p < num; p++)
                    {
                        newstr = GetStringInRange(lstWords, dMin - p * dx, dMax + p * dx, ref ofoundWord);
                        if (newstr.Length > 0) break;
                    }
                }
            }
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
        return newstr;
    }
    public string GetStringInRangeWithPlusTolerance2(List<WordInfo> lstWords, double dMin, double dMax)
    {
        List<WordInfo> ofoundWord = new List<WordInfo>();
        string str = GetStringInRange(lstWords, dMin, dMax, ref ofoundWord);
        string newstr = str;
        try
        {
            if (str.Length == 0)
            {
                if (dMax > 0.001) //Means title exists in the table
                {
                    double dx = (dMax - dMin) / 10.0;
                    for (int p = 1; p < 10; p++)
                    {
                        newstr = GetStringInRange(lstWords, dMin - p * dx, dMax + p * dx, ref ofoundWord);
                        if (newstr.Length > 0) break;
                    }
                }
            }
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
        return newstr;
    }
    public string GetStringInRangeWithPlusTolerance(string str, List<WordInfo> lstWords, double dMin, double dMax)
    {
        string newstr = str;
        try
        {
            if (str.Length == 0)
            {
                double dx = (dMax - dMin) / 10.0;
                for (int p = 1; p < 10; p++)
                {
                    WordInfo ofoundWord = null;
                    newstr = GetStringInRange(lstWords, dMin - p * dx, dMax + p * dx, ref ofoundWord);
                    if (newstr.Length > 0) break;
                }
            }
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
        return newstr;
    }
    //*******************************************************
    //Function  : GetMaxMinRangeOfWord
    //Purpose   : get min/max range from the given words in PDF
    //*******************************************************
    public void GetMaxMinRangeOfWord(List<WordInfo> lstWords, string start, string end, ref double min, ref double max, int my_index = 1)
    {
        try
        {
            bool bYesFound = false;
            int cnt = 0;
            for (int i = 0; i < lstWords.Count; i++)
            {
                if (lstWords[i].m_sTxt.ToUpper() == start.ToUpper())
                {
                    for (int j = i; j < lstWords.Count; j++)
                    {
                        string sp = Utility.trimString(lstWords[j].m_sTxt.ToUpper());
                        if (sp == end.ToUpper())
                        {
                            cnt = cnt + 1;
                            if (cnt == my_index)
                            {
                                bYesFound = true;
                                min = lstWords[i].m_dMin;
                                max = lstWords[j].m_dMax;
                                //re found the min value again
                                for (int k = j; k >= 0; k--)
                                {
                                    if (lstWords[k].m_sTxt.ToUpper() == start.ToUpper())
                                    {
                                        min = lstWords[k].m_dMin;
                                        break;
                                    }
                                }
                                break;
                            }
                        }
                    }
                    if (bYesFound) break;
                }
            }
            //Not found
            cnt = 0;
            if (bYesFound == false)
            {
                for (int i = 0; i < lstWords.Count; i++)
                {
                    if (Utility.trimString(lstWords[i].m_sTxt.ToUpper()) == start.ToUpper() + " " + end.ToUpper())
                    {
                        cnt = cnt + 1;
                        if (cnt == my_index)
                        {
                            bYesFound = true;
                            min = lstWords[i].m_dMin;
                            max = lstWords[i].m_dMax;
                            break;
                        }
                    }
                }
            }
            //
            cnt = 0;
            if (bYesFound == false)
            {
                for (int i = 0; i < lstWords.Count; i++)
                {
                    string sx = Utility.trimString(lstWords[i].m_sTxt.ToUpper());
                    if (sx.StartsWith(start.ToUpper()) && sx.EndsWith(end.ToUpper()))
                    {
                        cnt = cnt + 1;
                        if (cnt == my_index)
                        {
                            bYesFound = true;
                            min = lstWords[i].m_dMin;
                            max = lstWords[i].m_dMax;
                            break;
                        }
                    }
                }
            }
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
    }
    //*******************************************************
    //Function  : SplitLineToWordsNew
    //Purpose   : split the line in Pdf to number of words
    //*******************************************************
    public List<WordInfo> SplitLineToWordsNew(HeightInfo ht)
    {
        List<WordInfo> lstWords = new List<WordInfo>();
        try
        {
            List<List<RectAndText>> lstAll = new List<List<RectAndText>>();
            List<RectAndText> lstRts = new List<RectAndText>();
            for (int x = 0; x < ht.m_lstRectAndText.Count; x++)
            {
                string sSpace = "", txt = ht.m_lstRectAndText[x].Text;
                if (x > 0)
                {
                    float f = ht.m_lstRectAndText[x].Rect.Left - ht.m_lstRectAndText[x - 1].Rect.Right;
                    if (f > 1.5f) sSpace = " ";
                }
                //-->>create list
                if (sSpace.Length > 0)
                {
                    //new list starts
                    if (lstRts.Count > 0) lstAll.Add(lstRts);
                    lstRts = new List<RectAndText>() { ht.m_lstRectAndText[x] };
                }
                else
                {
                    lstRts.Add(ht.m_lstRectAndText[x]);
                }
                //last item
                if (x == ht.m_lstRectAndText.Count - 1 && lstRts.Count > 0) lstAll.Add(lstRts);
            }
            //----Create list objects ---
            foreach (List<RectAndText> lstRects in lstAll)
            {
                string str = "";
                foreach (RectAndText rt in lstRects)
                    str = str + rt.Text;
                //create word
                WordInfo wrd = new WordInfo();
                wrd.m_sTxt = str.Trim();
                wrd.m_lstRects = lstRects;
                wrd.m_dMin = lstRects[0].Rect.Left;
                wrd.m_dMax = lstRects[lstRects.Count - 1].Rect.Right;
                wrd.m_fVal = ht.m_fVal;
                lstWords.Add(wrd);
            }
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
        return lstWords;
    }
    //*******************************************************
    //Function  : SplitLineToWords
    //Purpose   : split the line in Pdf to number of words
    //*******************************************************
    //public List<WordInfo> SplitLineToWords(HeightInfo ht)
    //{
    //    List<WordInfo> lstWords = new List<WordInfo>();
    //    try
    //    {
    //        List<List<RectAndText>> lstAll = new List<List<RectAndText>>();
    //        List<RectAndText> lstRts = new List<RectAndText>();
    //        for (int x = 0; x < ht.m_lstRectAndText.Count; x++)
    //        {
    //            string sSpace = "", txt = ht.m_lstRectAndText[x].Text;
    //            if (x > 0)
    //            {
    //                float f = ht.m_lstRectAndText[x].Rect.Left - ht.m_lstRectAndText[x - 1].Rect.Right;
    //                if (f > 1.5f) sSpace = " ";
    //            }
    //            //-->>create list
    //            if (sSpace.Length > 0 || txt.Trim().Length == 0)
    //            {
    //                //new list starts
    //                if (lstRts.Count > 0) lstAll.Add(lstRts);
    //                lstRts = new List<RectAndText>() { ht.m_lstRectAndText[x] };
    //            }
    //            else if (txt.Trim().Length > 0 && txt[txt.Length - 1] == ' ')
    //            {
    //                lstRts.Add(ht.m_lstRectAndText[x]);
    //                //new list
    //                if (lstRts.Count > 0) lstAll.Add(lstRts);
    //                lstRts = new List<RectAndText>() { };
    //            }
    //            else
    //            {
    //                lstRts.Add(ht.m_lstRectAndText[x]);
    //            }
    //            //last item
    //            if (x == ht.m_lstRectAndText.Count - 1 && lstRts.Count > 0) lstAll.Add(lstRts);
    //        }
    //        //----Create list objects ---
    //        foreach (List<RectAndText> lstRects in lstAll)
    //        {
    //            string str = "";
    //            foreach (RectAndText rt in lstRects)
    //                str = str + rt.Text;
    //            //create word
    //            WordInfo wrd = new WordInfo();
    //            wrd.m_sTxt = str.Trim();
    //            wrd.m_lstRects = lstRects;
    //            wrd.m_dMin = lstRects[0].Rect.Left;
    //            wrd.m_dMax = lstRects[lstRects.Count - 1].Rect.Right;
    //            wrd.m_fVal = ht.m_fVal;
    //            lstWords.Add(wrd);
    //        }
    //    }
    //    catch (Exception ee)
    //    {
    //        Utility.WriteErrorLog(ee);
    //    }
    //    return lstWords;
    //}

}
//---------------------- WordInfo --------------------
public class WordInfo
{
    public string m_sTxt = "";
    public List<RectAndText> m_lstRects = null;
    public double m_dMin = 0.0;
    public double m_dMax = 0.0;
    public double m_fVal = 0.0;
}
internal class ImageRenderListener : IRenderListener
{
    #region Fields
    private Dictionary<BitmapImage, string> images =
    new Dictionary<BitmapImage, string>();
    #endregion Fields
    #region Properties
    public Dictionary<BitmapImage, string> Images
    {
        get { return images; }
    }
    #endregion Properties
    #region Methods
    #region Public Methods
    public void BeginTextBlock() { }
    public void EndTextBlock() { }
    public void RenderImage(ImageRenderInfo renderInfo)
    {
        PdfImageObject image = renderInfo.GetImage();
        PdfName filter = (PdfName)image.Get(PdfName.FILTER);

        //int width = Convert.ToInt32(image.Get(PdfName.WIDTH).ToString());
        //int bitsPerComponent = Convert.ToInt32(image.Get(PdfName.BITSPERCOMPONENT).ToString());
        //string subtype = image.Get(PdfName.SUBTYPE).ToString();
        //int height = Convert.ToInt32(image.Get(PdfName.HEIGHT).ToString());
        //int length = Convert.ToInt32(image.Get(PdfName.LENGTH).ToString());
        //string colorSpace = image.Get(PdfName.COLORSPACE).ToString();
        /* It appears to be safe to assume that when filter == null, PdfImageObject
         * does not know how to decode the image to a System.Drawing.Image.
         *
         * Uncomment the code above to verify, but when I’ve seen this happen,
         * width, height and bits per component all equal zero as well. */
        if (filter != null)
        {
            var drawingImage = new BitmapImage();// image.GetDrawingImage();
            string extension = ".";
            if (filter == PdfName.DCTDECODE)
            {
                extension += PdfImageObject.ImageBytesType.JPG.FileExtension;
            }
            else if (filter == PdfName.JPXDECODE)
            {
                extension += PdfImageObject.ImageBytesType.JP2.FileExtension;
            }
            else if (filter == PdfName.FLATEDECODE)
            {
                extension += PdfImageObject.ImageBytesType.PNG.FileExtension;
            }
            else if (filter == PdfName.LZWDECODE)
            {
                extension += PdfImageObject.ImageBytesType.CCITT.FileExtension;
            }
            /* Rather than struggle with the image stream and try to figure out how to handle
             * BitMapData scan lines in various formats (like virtually every sample I’ve found
             * online), use the PdfImageObject.GetDrawingImage() method, which does the work for us. */
            this.Images.Add(drawingImage, extension);
        }
    }
    public void RenderText(TextRenderInfo renderInfo) { }
    #endregion Public Methods
    #endregion Methods
}

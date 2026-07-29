//using CTR_Form_Tool.CS_FILES;
using iTextSharp.text.pdf.parser;
using Microsoft.Office.Interop.Word;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Numeric;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
//using System.Windows.Forms;
//using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Word = Microsoft.Office.Interop.Word;

namespace CTR_Form_Tool
{
    class CTRReader
    {
        Dictionary<string, Word.ContentControl> dtcontrols = new Dictionary<string, Word.ContentControl>();
        //  Word.Application wordApp = null;
        bool bIsSingleRepair = false;
        string m_sATA = "";

        public string m_sESN = "";
        public string m_sTSN = "";
        public string m_sCSN = "";

        string m_sPN = "";
        string m_sSN = "";
        string m_sNOMEN = "";

        string m_sConstructionNum = "";
        public List<string> m_lstDamageInfo = new List<string>();
        List<AsciiInfo> m_lstAscii = new List<AsciiInfo>();
        string CleanText(string input)
        {
            if (string.IsNullOrEmpty(input))
                return input;

            input = input.Normalize(NormalizationForm.FormKC);

            return input
                .Replace('’', '\'')
                .Replace('“', '"')
                .Replace('”', '"')
                .Replace('–', '-')
                .Replace('—', '-')
                .Replace('：', ':'); // safety
        }
        string SplitIt(string value2, string type1)
        {
            try
            {
                string s1 = CleanText(value2.Replace(type1, "").Trim());
                s1 = s1.TrimStart(new char[1] { ':' }).Trim();
                return s1;
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return value2;
        }
        public void updateFromRedmine(string path)
        {
            try
            {
                List<string> lines = Utility.ReadFile(path);

                //
                int DescStart = -1;
                for (int i = 0; i < lines.Count; i++)
                {
                    string value = updateAscii(lines[i]);
                    string value2 = CleanText(value.Replace(" ", "").Trim());
                    if (value2.StartsWith("P/N") && m_sPN.Length == 0)
                    {
                        //Find the Single Repair
                        for (int p = i - 1; p >= 0; p--)
                        {
                            if (lines[i].ToUpper().Contains("Single Repair".ToUpper()))
                            {
                                bIsSingleRepair = true;
                                break;
                            }
                        }
                        //--
                        m_sPN = value2.Replace("P/N", "").Trim().TrimStart(new char[1] { ':' }).Trim();
                        DescStart = i + 1;
                    }
                    else if (value2.StartsWith("S/N") && m_sSN.Length == 0)
                    {
                        m_sSN = value2.Replace("S/N", "").Trim().TrimStart(new char[1] { ':' }).Trim();
                        DescStart = i + 1;
                    }
                    else if (value.StartsWith("NOMEN") && m_sNOMEN.Length == 0)
                    {
                        m_sNOMEN = value.Replace("NOMEN", "").Trim().TrimStart(new char[1] { ':' }).Trim();
                        DescStart = i + 1;
                    }
                    else if (value2.StartsWith("CSN") && m_sCSN.Length == 0)
                    {
                        m_sCSN = value2.Replace("CSN", "").Trim().TrimStart(new char[1] { ':' }).Trim();
                    }
                    else if (value2.StartsWith("TSN") && m_sTSN.Length == 0)
                    {
                        m_sTSN = value2.Replace("TSN", "").Trim().TrimStart(new char[1] { ':' }).Trim();
                    }
                    else if (value2.StartsWith("ESN") && m_sESN.Length == 0)
                    {
                        m_sESN = value2.Replace("ESN", "").Trim().TrimStart(new char[1] { ':' }).Trim();
                    }
                    else if (value2.ToUpper().StartsWith("ATANo".ToUpper()) && m_sATA.Length == 0)
                    {
                        m_sATA = value2.ToUpper().Replace("ATANo".ToUpper(), "").Trim().TrimStart(new char[1] { ':' }).TrimStart(new char[1] { '.' }).Trim();
                    }
                }
                if (DescStart != -1)
                {
                    if (m_lstDamageInfo.Count == 0)
                    {
                        for (int j = DescStart; j < lines.Count; j++)
                        {
                            string value = updateAscii(lines[j]);
                            if (value.Length > 0)
                            {
                                if (bIsContinue(value)) continue;
                                if (bIsEnd(value) && m_lstDamageInfo.Count > 0) break;
                                if (value.StartsWith("CSN") && m_lstDamageInfo.Count == 0) continue;
                                if (bIsRemove(value) ) continue;
                                m_lstDamageInfo.Add(value);
                            }
                        }
                    }
                }

                if (m_sNOMEN.Length == 0)
                {
                    for (int i = 0; i < lines.Count; i++)
                    {
                        string v = FindNomenclature(lines[i]);
                        if (v.Length > 0)
                        {
                            m_sNOMEN = v;
                            break;
                        }
                    }
                }

                //Find the construction number
                if (m_sPN.Length > 0 && m_sConstructionNum.Length == 0)
                {
                    //find from Text
                    foreach (string sx in lines)
                    {
                        List<string> lstp = Utility.SplitString(sx, " ");
                        foreach (string s in lstp)
                        {
                            if (s.StartsWith("#"))
                            {
                                m_sConstructionNum = s;
                                break;
                            }
                        }
                        if (m_sConstructionNum.Length > 0) break;
                    }
                }

                //Update nomenclature
                if (m_sNOMEN.Length > 0)
                {
                    try
                    {
                        m_sNOMEN = Utility.SplitString(m_sNOMEN, "Area")[0].Trim();
                    }
                    catch { }
                }


                //---
                foreach (string lne in lines)
                {
                    if (m_sESN.Length == 0 && lne.StartsWith("ESN:"))
                    {
                        m_sESN = lne.Replace("ESN:", "").Trim();
                    }
                    else if (m_sTSN.Length == 0 && lne.StartsWith("TSN:"))
                    {
                        m_sTSN = lne.Replace("TSN:", "").Trim();
                    }
                    else if (m_sCSN.Length == 0 && lne.StartsWith("CSN:"))
                    {
                        m_sCSN = lne.Replace("CSN:", "").Trim();
                    }
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
        }

        public string Update(string folderPath, List<AsciiInfo> lstAscii)
        {
            string givenpdffile = "";
            try
            {
                m_lstAscii = lstAscii;
                //Get the PDFs & sort it
                List<string> sArrPDFNew = new List<string>();
                List<string> sArrPDFs = System.IO.Directory.GetFiles(folderPath, "*.pdf").ToList();
                for (int i = 0; i < sArrPDFs.Count; i++)
                {
                    if (System.IO.Path.GetFileNameWithoutExtension(sArrPDFs[i]).Contains("Finding"))
                    {
                        sArrPDFNew.Add(sArrPDFs[i]);
                        sArrPDFs.RemoveAt(i);
                        i--;
                    }
                }
                foreach (string s in sArrPDFs) sArrPDFNew.Add(s);
                //Read the PDFs
                foreach (string pdffile in sArrPDFNew)
                {
                    //Read PDF using pdf reader
                    read_PDF(pdffile);
                    read_Word(pdffile);
                    givenpdffile = pdffile;
                }

                //read ESN/TSN/CSN
                Read_ESN_TSN_CSN(folderPath);
                //Write
                //write(folderPath);
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return givenpdffile;
        }
        public string updateAscii(string val)
        {
            string updatedval = val;
            foreach (AsciiInfo oo in m_lstAscii)
            {
                updatedval = updatedval.Replace(((char)oo.m_nAsciiNum).ToString(), oo.m_sVal);
            }
            return updatedval.Trim();
        }
        void InsertBlankPage(Word.Application wordApp, Word.Document docMain, int pagenum)
        {
            Word.Range insertRange = GetPageRange(wordApp, docMain, pagenum)?.Duplicate;
            if (insertRange != null)
            {
                insertRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                insertRange.InsertBreak(Word.WdBreakType.wdPageBreak); // Insert new page break
            }
            else
            {
                // If target doc has less than 4 pages, just paste at the end
                docMain.Content.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                docMain.Content.InsertBreak(Word.WdBreakType.wdPageBreak);
            }
        }
        void SetCheckBoxStatus(string id, bool status)
        {
            if (dtcontrols.ContainsKey(id))
            {
                dtcontrols[id].Checked = status;
            }
        }
        public DialogResult write(string outputDir, List<DamageInfo> lstDamages, string esn, string tsn, string csn, string finding)
        {

            try
            {
                //Update ESN/TSN
                m_sESN = esn;
                m_sTSN = tsn;
                m_sCSN = csn;
                if (m_sConstructionNum.Length == 0) m_sConstructionNum = finding;

                if (m_sPN.Length > 0 && m_sSN.Length > 0)
                {
                    string filename = "CTR_MRO_Form_PN_" + m_sPN + "_SN_" + m_sSN + ".docx";
                    string pdffile = outputDir + "\\" + filename;

                    SaveFileDialog saveFileDialog = new SaveFileDialog();
                    saveFileDialog.Title = "Save Word Document";
                    saveFileDialog.Filter = "Word Document (*.docx)|*.docx";
                    saveFileDialog.FileName = filename;// "Report.docx";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        pdffile = saveFileDialog.FileName;

                        //--- Open the 1st template and copy to required folder
                        string templatePath = Utility.m_sBinPath + "Templates\\" + lstDamages[0].m_oConfig.m_sFile;
                        Utility.WriteErrorLog("Type : " + lstDamages[0].m_sType + " ; " + "Sub Type : " + lstDamages[0].m_sSubType + " ; Template 1 : " + templatePath);
                        System.IO.File.Copy(templatePath, pdffile, true);
                        //Open word && Change the title text
                        object fileName = pdffile;
                        object readOnly = false;
                        object isVisible = true;
                        object missing = System.Reflection.Missing.Value;
                        Word.Application wordApp = new Word.Application();
                        wordApp.Visible = true;
                        wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                        Word.Document docMain = wordApp.Documents.Open(ref fileName, ref missing, ref readOnly, ref missing, ref missing, ref missing,
                                                                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                        for (int i = 1; i <= docMain.Shapes.Count; i++)
                        {
                            Word.Shape oshape = docMain.Shapes[i];
                            string text = "";
                            try
                            {
                                text = oshape.TextFrame.TextRange.Text;
                                text = text.Replace("\r", "").Replace("\a", "").Replace("\n", "").Trim();
                            }
                            catch { }
                            if (text == "Part name:")
                            {
                                oshape.TextFrame.TextRange.Text = "Partname1xx";
                            }
                            else if (text == "Description")
                            {
                                oshape.TextFrame.TextRange.Text = "Description1xx";
                            }
                        }
                        //open new templates and insert into the page
                        int pagenum = 4;
                        //Add New page...Images are more
                        if (lstDamages[0].m_nNoOfPages > 1)
                        {
                            InsertBlankPage(wordApp, docMain, pagenum);
                            pagenum = pagenum + 1;
                        }
                        for (int p = 1; p < lstDamages.Count; p++)
                        {
                            string newTemplate = Utility.m_sBinPath + "Templates\\" + lstDamages[p].m_oConfig.m_sFile;
                            Utility.WriteErrorLog("Type : " + lstDamages[p].m_sType + " ; " + "Sub Type : " + lstDamages[0].m_sSubType + " ; Template " + (p + 1).ToString() + " : " + templatePath);
                            if (newTemplate.Length > 0)
                            {
                                fileName = newTemplate;
                                Word.Document sourceDoc = wordApp.Documents.Open(ref fileName, ref missing, ref readOnly, ref missing, ref missing, ref missing,
                                                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);


                                // Get range of page 1 from source
                                Word.Range pageRange = GetPageRange(wordApp, sourceDoc, 4);

                                if (pageRange != null)
                                {
                                    // Copy the content of the 4th page
                                    pageRange.Copy();

                                    // Go to the end of the 4th page in targetDoc (or create page 5)
                                    Word.Range insertRange = GetPageRange(wordApp, docMain, pagenum)?.Duplicate;

                                    if (insertRange != null)
                                    {
                                        insertRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                                        insertRange.InsertBreak(Word.WdBreakType.wdPageBreak); // Insert new page break
                                        insertRange.Paste();
                                    }
                                    else
                                    {
                                        // If target doc has less than 4 pages, just paste at the end
                                        docMain.Content.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                                        docMain.Content.InsertBreak(Word.WdBreakType.wdPageBreak);
                                        docMain.Content.Paste();
                                    }

                                    //No of images
                                    pagenum = pagenum + 1;
                                    //Update page
                                    for (int i = 1; i <= docMain.Shapes.Count; i++)
                                    {
                                        Word.Shape oshape = docMain.Shapes[i];
                                        string text = "";
                                        try
                                        {
                                            text = oshape.TextFrame.TextRange.Text;
                                            text = text.Replace("\r", "").Replace("\a", "").Replace("\n", "").Trim();
                                        }
                                        catch { }
                                        if (text == "Part name:")
                                        {
                                            oshape.TextFrame.TextRange.Text = "Partname" + (p + 1).ToString() + "xx";
                                        }
                                        else if (text == "Description")
                                        {
                                            oshape.TextFrame.TextRange.Text = "Description" + (p + 1).ToString() + "xx";
                                        }
                                    }
                                    //Add New page...Images are more
                                    if (lstDamages[p].m_nNoOfPages > 1)
                                    {
                                        InsertBlankPage(wordApp, docMain, pagenum);
                                        pagenum = pagenum + 1;
                                    }
                                }
                                sourceDoc.Close();
                            }
                        }
                        docMain.Save();

                        //-------------- UPDATE CHECK BOXES & DATA ---------------
                        dtcontrols.Clear();
                        Word.ContentControls contentControls = docMain.ContentControls;
                        foreach (Word.ContentControl contentControl in contentControls)
                        {
                            if (contentControl.Type == Word.WdContentControlType.wdContentControlCheckBox)
                            {
                                if (contentControl.ID != null)
                                {
                                    dtcontrols.Add(contentControl.ID, contentControl);
                                }
                            }
                        }
                        //20) DAT use approved process when part was damaged:
                        if (bIsSingleRepair)
                        {
                            SetCheckBoxStatus("1767575576", true);
                            SetCheckBoxStatus("1954678652", false);
                        }
                        //--------------------------------------------------------

                        //data
                        string pnNum = "Part name: " + m_sNOMEN + "\r\n";
                        pnNum = pnNum + "Part number: " + m_sPN + "\r\n";
                        pnNum = pnNum + "Serial number: " + m_sSN + "\r\n";
                        pnNum = pnNum + "CSN: " + m_sCSN;
                        string desc = "";
                        foreach (string s in m_lstDamageInfo)
                            desc = desc + s + "\r\n";

                        //----
                        int extrapages = 0;
                        string Nonconformancetype = "";
                        foreach (DamageInfo dmg in lstDamages)
                        {
                            if (Nonconformancetype.Contains(dmg.m_sNonConformanceType) == false)
                                Nonconformancetype = Nonconformancetype + dmg.m_sNonConformanceType + ",";
                            extrapages = extrapages + dmg.m_nNoOfPages - 1;
                        }
                        int endPage = 3 + lstDamages.Count + extrapages;
                        Nonconformancetype = Nonconformancetype.TrimEnd(',').Trim();
                        string str26 = "Refer to Page no 4";
                        if (endPage < 5)
                        {
                            List<string> lx = Utility.SplitString(lstDamages[0].m_sDamageInfo, "\r\n");
                            if (lx.Count > 0)
                            {
                                if (lx[lx.Count - 1].ToUpper().StartsWith("(Limit".ToUpper()))
                                    lx.RemoveAt(lx.Count - 1);
                                else if (lx[lx.Count - 1].ToUpper().StartsWith("Limit".ToUpper()))
                                    lx.RemoveAt(lx.Count - 1);
                            }
                            str26 = Utility.CombineString(lx, "\r").Trim();
                            // str26 = str26.Replace("\n", " ").Trim();
                        }
                        else if (endPage == 5) str26 = "Refer to Page no 4 and 5";
                        else if (endPage > 5) str26 = "Refer to Page no 4 to " + endPage.ToString();
                        //--
                        //Update ATA for Wear
                        if (Nonconformancetype.ToUpper().Contains("WEAR") && m_sATA.Length > 0)
                        {
                            m_sATA = m_sATA + ", EA23CJ182";
                        }
                        //----
                        List<WordDate> lstParas = new List<WordDate>();
                        string Manualcriteria = "";
                        List<string> lsx = new List<string>();
                        foreach(DamageInfo odmg in lstDamages)
                        {
                            if (odmg.m_sManualCriteria.Length > 0 && lsx.Contains(odmg.m_sManualCriteria) == false) lsx.Add(odmg.m_sManualCriteria);
                        }
                        if (lsx.Count > 0)
                        {
                            Manualcriteria = Utility.CombineString(lsx, Environment.NewLine).Trim();
                            Manualcriteria = Manualcriteria.TrimEnd();
                        }
                        //if (Nonconformancetype.ToUpper().Contains("WEAR")) Manualcriteria = "[15] GROOVE - Wear - Not permitted";


                        //Update the description and part number
                        for (int x = 0; x < lstDamages.Count; x++)
                        {
                            string pn1 = "Partname" + (x + 1).ToString() + "xx";
                            string dc1 = "Description" + (x + 1).ToString() + "xx";

                            for (int i = 1; i <= docMain.Shapes.Count; i++)
                            {
                                Word.Shape oshape = docMain.Shapes[i];
                                string text = "";
                                try
                                {
                                    text = oshape.TextFrame.TextRange.Text;
                                    text = text.Replace("\r", "").Replace("\a", "").Replace("\n", "").Trim();
                                }
                                catch { }
                                if (text == pn1)
                                {
                                    oshape.TextFrame.TextRange.Text = pnNum.Trim();
                                    oshape.TextFrame.TextRange.Font.Size = 11f;
                                    oshape.TextFrame.TextRange.Font.Name = "Arial";
                                }
                                else if (text == dc1)
                                {
                                    oshape.TextFrame.TextRange.Text = lstDamages[x].m_sDamageInfo;
                                    oshape.TextFrame.TextRange.Font.Size = 11f;
                                    oshape.TextFrame.TextRange.Font.Name = "Arial";
                                }
                            }
                        }

                        //------------Add the Zoomed Images ---------
                        int pageNum = 4;
                        for (int p = 0; p < lstDamages.Count; p++)
                        {
                            string txt = "";
                            try
                            {
                                txt = Utility.SplitString(lstDamages[p].m_sDamageInfo, "\r\n")[0].Trim();
                                //Divide into Area
                                int n = txt.ToUpper().IndexOf("AREA");
                                if (n != -1)
                                {
                                    string s1 = txt.Substring(0, n).Trim();
                                    string s2 = txt.Substring(s1.Length, txt.Length - s1.Length).Trim();
                                    s2 = s2.Substring(4, s2.Length - 4);
                                    List<string> lx2 = Utility.SplitString(s2, new string[2] { " ", "(" });
                                    txt = s1 + " " + "Area " + lx2[0];
                                    txt = Regex.Replace(txt, @"^\d+\.\s*", "");
                                }
                            }
                            catch { }
                            //find page num
                            if (p > 0) pageNum = pageNum + lstDamages[p - 1].m_nNoOfPages;
                            //Find shapes
                            Dictionary<int, ShapeInfo> dtShapesZoomed = new Dictionary<int, ShapeInfo>();

                            //Add images
                            ConfigInfo cfg = lstDamages[p].m_oConfig;
                            int imagecnt = 1;
                            List<string> ArrFilesZoomed = new List<string>();
                            if (lstDamages[p].m_sImage1.Length > 0 && System.IO.File.Exists(lstDamages[p].m_sImage1))
                            {
                                ArrFilesZoomed.Add(lstDamages[p].m_sImage1);
                                dtShapesZoomed.Add(imagecnt, getImageLocation(pageNum, imagecnt, cfg, "Overview of " + txt));
                                imagecnt = imagecnt + 1;
                            }
                            if (lstDamages[p].m_sImage2.Length > 0 && System.IO.File.Exists(lstDamages[p].m_sImage2))
                            {
                                ArrFilesZoomed.Add(lstDamages[p].m_sImage2);
                                dtShapesZoomed.Add(imagecnt, getImageLocation(pageNum, imagecnt, cfg, "Zoomed view of " + txt));
                                imagecnt = imagecnt + 1;
                            }
                            if (lstDamages[p].m_sImage3.Length > 0 && System.IO.File.Exists(lstDamages[p].m_sImage3))
                            {
                                ArrFilesZoomed.Add(lstDamages[p].m_sImage3);
                                dtShapesZoomed.Add(imagecnt, getImageLocation(pageNum, imagecnt, cfg, "Zoomed view of " + txt));
                                imagecnt = imagecnt + 1;
                            }
                            //Add other images
                            foreach (MoreImages mi in lstDamages[p].addmoreimages)
                            {
                                ArrFilesZoomed.Add(mi.m_sFilename);
                                dtShapesZoomed.Add(imagecnt, getImageLocation(pageNum, imagecnt, cfg, mi.m_sName));
                                imagecnt = imagecnt + 1;
                            }


                            ////Only one image no need of over view or zoomed text
                            if (ArrFilesZoomed.Count % 2 == 1)
                            {
                                dtShapesZoomed[ArrFilesZoomed.Count].m_sText = txt;
                                dtShapesZoomed[ArrFilesZoomed.Count].m_fImgLeft = dtShapesZoomed[ArrFilesZoomed.Count].m_fImgLeft + 120;
                                dtShapesZoomed[ArrFilesZoomed.Count].m_fTitleLeft = dtShapesZoomed[ArrFilesZoomed.Count].m_fTitleLeft + 120;
                                dtShapesZoomed[ArrFilesZoomed.Count].m_fNumLeft = dtShapesZoomed[ArrFilesZoomed.Count].m_fNumLeft + 120;
                            }

                            //if (ArrFilesZoomed.Count == 3)
                            //{
                            //    dtShapesZoomed[3].m_sText = txt;
                            //    dtShapesZoomed[3].m_fImgLeft = 0.5f * (dtShapesZoomed[3].m_fImgLeft + dtShapesZoomed[4].m_fImgLeft);
                            //    dtShapesZoomed[3].m_fTitleLeft = 0.5f * (dtShapesZoomed[3].m_fTitleLeft + dtShapesZoomed[4].m_fTitleLeft);
                            //    dtShapesZoomed[3].m_fNumLeft = 0.5f * (dtShapesZoomed[3].m_fNumLeft + dtShapesZoomed[4].m_fNumLeft);
                            //}
                            //else if (ArrFilesZoomed.Count > 3)
                            //{

                            //}

                            for (int x = 0; x < ArrFilesZoomed.Count; x++)
                            {
                                AddImage(wordApp, docMain, dtShapesZoomed, ArrFilesZoomed[x], x + 1, ArrFilesZoomed.Count);
                            }


                            //Add images
                            /*int cnt = 0;
                            foreach (string file in ArrFilesZoomed)
                            {
                                cnt = cnt + 1;

                                if (dtShapesZoomed.ContainsKey(cnt) == false) continue;
                                if (System.IO.File.Exists(file) == false) continue;

                                ShapeInfo oshObj = dtShapesZoomed[cnt];
                                // Ensure pages exist
                                //for (int i = 1; i < oshObj.m_nSheetNo; i++)
                                //{
                                //    wordApp.Selection.InsertBreak(Word.WdBreakType.wdPageBreak);
                                //}

                                // Force Word to update layout
                                docMain.Repaginate();
                                wordApp.ScreenRefresh();
                                System.Threading.Thread.Sleep(100);
                                //

                                Word.Range rng = wordApp.Selection.GoTo(Word.WdGoToItem.wdGoToPage, Word.WdGoToDirection.wdGoToAbsolute, oshObj.m_nSheetNo);
                                System.Threading.Thread.Sleep(1000);
                                rng.Collapse(Word.WdCollapseDirection.wdCollapseStart);

                                // Define where to insert the image
                                Word.Range range = docMain.Range(0, 0);
                                range = docMain.Range(range.End, range.End);  // This moves the range to the end of the current text
                                                                              // Insert the picture (using Shape, not InlineShape, for positioning with left and top)


                                Word.Shape picture = docMain.Shapes.AddPicture(file,
                                    LinkToFile: false, SaveWithDocument: true,
                                    Left: oshObj.m_fImgLeft, Top: oshObj.m_fImgTop, Width: oshObj.m_fWidth, Height: oshObj.m_fHeight);
                                // Important settings
                                picture.WrapFormat.Type = Word.WdWrapType.wdWrapNone;
                                picture.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
                                picture.Width = oshObj.m_fWidth; 
                                picture.Left = oshObj.m_fImgLeft;
                                picture.Top = oshObj.m_fImgTop;


                                picture.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoSendToBack);
                                range = docMain.Range(picture.Anchor.End, picture.Anchor.End);

                                if (cnt == 2)
                                {
                                    picture.Line.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                                    // Set border color to Red (use System.Drawing.Color to get RGB OLE int)
                                    picture.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Red);
                                    // Set border width/weight to 1.5 points
                                    picture.Line.Weight = 1.5f;
                                    // Set dash style to long dash
                                    picture.Line.DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineLongDash;
                                }
                                //if (cnt == 2 || cnt == 4) pos = 2;

                                // Add a TextBox shape that can hold text at a specific location
                                Word.Shape textBox1 = docMain.Shapes.AddTextbox(
                                    Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                                    oshObj.m_fNumLeft, oshObj.m_fNumTop, 210, 30); // Width and height for the TextBox
                                textBox1.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                                textBox1.TextFrame.TextRange.Text = m_sConstructionNum.TrimEnd(new char[1] { ':' }) + " " + System.IO.Path.GetFileName(file);
                                textBox1.TextFrame.TextRange.Font.Size = 8f;
                                textBox1.TextFrame.TextRange.Font.Name = "Arial";
                                textBox1.TextFrame.TextRange.Font.Color = Word.WdColor.wdColorWhite;
                                textBox1.TextFrame.TextRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                                // Add a TextBox shape that can hold text at a specific location
                                Word.Shape textBox2 = docMain.Shapes.AddTextbox(
                                    Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                                    oshObj.m_fTitleLeft, oshObj.m_fTitleTop, 230, 30); // Width and height for the TextBox
                                textBox2.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                                textBox2.TextFrame.TextRange.Text = oshObj.m_sText;
                                textBox2.TextFrame.TextRange.Font.Name = "Arial";
                                textBox2.TextFrame.TextRange.Font.Size = 9f;
                                textBox2.TextFrame.TextRange.Font.Bold = 1;
                                textBox2.TextFrame.TextRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                                //Handle special case, image resize
                                if (lstDamages[p].m_oConfig.m_sType == "HUB - LPC ROTOR, 1ST STAGE" && lstDamages[p].m_oConfig.m_sSubType == "Loc[15]" && ArrFilesZoomed.Count == 1)
                                {
                                    picture.Height = 288;
                                    picture.Width = 360;
                                    picture.Left = 94;
                                    picture.Top = 320;

                                    textBox1.Left = 240;
                                    textBox1.Top = 580;


                                    textBox2.Left = 170;
                                    textBox2.Top = 610;
                                }
                            }*/
                        }
                        //---
                        System.Windows.Forms.Clipboard.Clear();


                        //--------------------------------------------------------------
                        //Update paragraphs
                        // lstParas.Add(new WordDate("ESN:", m_sESN));
                        lstParas.Add(new WordDate("Hardware PN:", m_sPN));
                        lstParas.Add(new WordDate("Hardware SN:", m_sSN));
                        lstParas.Add(new WordDate("Hardware Time:", m_sTSN));
                        lstParas.Add(new WordDate("Hardware Cycle:", m_sCSN));
                        lstParas.Add(new WordDate("Manual Reference(s) (SB/EA#/DMC/ATA):", m_sATA));
                        lstParas.Add(new WordDate("NOTE: please provide picture of the damage", Nonconformancetype));
                        lstParas.Add(new WordDate("NOTE: If BSI please ref. QRG form (attached below)", str26, Word.WdColor.wdColorRed));
                        lstParas.Add(new WordDate("Manual criteria & inspector disposition per criteria:", Manualcriteria));
                        //Read/Write the paragraphs
                        foreach (WordDate oo in lstParas)
                        {
                            if (oo.m_sValue.Length == 0) continue;
                            foreach (Word.Paragraph paragraph in docMain.Paragraphs)
                            {
                                string paragraphText = paragraph.Range.Text;
                                if (paragraphText.Contains(oo.m_sTitle))
                                {
                                    Word.Range range = paragraph.Range;
                                    range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                                    range.InsertAfter(oo.m_sValue);
                                    //if (oo.m_oclr != Word.WdColor.wdColorBlack)
                                    {
                                        Word.Range coloredRange = docMain.Range(range.End - (oo.m_sValue.Length + 1), range.End);
                                        coloredRange.Font.Color = oo.m_oclr;
                                        coloredRange.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight;
                                        coloredRange.Font.Bold = 0;// Word.WdConstants.wdFalse;
                                        if (oo.m_sTitle == "Hardware PN:" || oo.m_sTitle == "Hardware SN:" || oo.m_sTitle == "Hardware Time:" || oo.m_sTitle == "Hardware Cycle:")
                                        {
                                            coloredRange.Font.Bold = 1;
                                        }
                                        //
                                        if (oo.m_sTitle == "Manual criteria & inspector disposition per criteria:")
                                        {
                                            coloredRange.Font.Color = ConvertToWdColor(Color.FromArgb(0, 176, 240));
                                        }
                                    }
                                    break;
                                }
                            }
                        }
                        lstParas.Clear();
                        lstParas.Add(new WordDate("ESN:", m_sESN));
                        lstParas.Add(new WordDate("Engine Operator:", "UNK"));
                        lstParas.Add(new WordDate("Engine Time On-wing target:", "UNK"));
                        lstParas.Add(new WordDate("Engine Cycles On-wing target:", "UNK"));
                        lstParas.Add(new WordDate("Engine TSN:", "UNK"));
                        lstParas.Add(new WordDate("Engine CSN:", "UNK"));
                        foreach (WordDate oo in lstParas)
                        {
                            foreach (Word.Paragraph paragraph in docMain.Paragraphs)
                            {
                                string paragraphText = paragraph.Range.Text;
                                if (paragraphText.Contains(oo.m_sTitle))
                                {
                                    Word.Range range = paragraph.Range;
                                    range.InsertAfter(" " + oo.m_sValue);
                                    try
                                    {
                                        Word.Range coloredRange = docMain.Range(range.End - (oo.m_sValue.Length + 1), range.End);
                                        if (oo.m_sTitle == "ESN:")
                                        {
                                            coloredRange.Font.Bold = 1;
                                        }
                                    }
                                    catch { }
                                    break;
                                }
                            }
                        }
                        //--------------------------------------------------------------

                        System.Windows.Forms.Clipboard.Clear();
                        //wordApp.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                        docMain.Save();
                        docMain.Close();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(docMain);
                        wordApp.Quit();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);


                        return DialogResult.Yes;
                    }
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return DialogResult.No;
        }
        private int GetPageNumber(Word.Range range)
        {
            return (int)range.Information[Word.WdInformation.wdActiveEndPageNumber];
        }
        public void doit(List<ConfigInfo> lstConfigs)
        {
            //string type = "LOW PRESSURE TURBINE DRIVE SHAFT", sSubType = "LOC [24]";
            //ConfigInfo cfg = null;
            //foreach (ConfigInfo config in lstConfigs)
            //{
            //    if (config.m_sType == type && config.m_sSubType == sSubType)
            //    {
            //        cfg = config;
            //        break;
            //    }
            //}
            for (int i = 0; i < 1; i++)
            {
                ConfigInfo cfg = lstConfigs[i];
                List<string> ArrFilesZoomed = new List<string>();
                string s1 = @"D:\veeru\IHI\Source\IHI - Tooling\Finding CTR Form Tool\CTR Form Tool\bin\Debug\bin\test\R0018608.JPG";
                string s2 = @"D:\veeru\IHI\Source\IHI - Tooling\Finding CTR Form Tool\CTR Form Tool\bin\Debug\bin\test\1.jpg";
                ArrFilesZoomed.Add(s1);
                ArrFilesZoomed.Add(s2);
                ArrFilesZoomed.Add(s1);
                ArrFilesZoomed.Add(s2);

                object fileName = @"C:\Users\vk11549.CORP\Desktop\Templates\" + cfg.m_sFile;
                Dictionary<int, ShapeInfo> dtShapesZoomed = new Dictionary<int, ShapeInfo>();
                dtShapesZoomed.Add(1, getImageLocation(4, 1, cfg, "Overview of Aarya"));
                dtShapesZoomed.Add(2, getImageLocation(4, 2, cfg, "Overview of Aarya"));
                dtShapesZoomed.Add(3, getImageLocation(4, 3, cfg, "Overview of Aarya"));
                dtShapesZoomed.Add(4, getImageLocation(4, 4, cfg, "Overview of Aarya"));


                object readOnly = false;
                object isVisible = true;
                object missing = System.Reflection.Missing.Value;
                Word.Application wordApp = new Word.Application();
                wordApp.Visible = true;
                wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                Word.Document docMain = wordApp.Documents.Open(ref fileName, ref missing, ref readOnly, ref missing, ref missing, ref missing,
                                                            ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);



                //foreach (Word.Shape shape in docMain.Shapes)
                //{
                //    try
                //    {
                //        if (shape.Anchor != null)
                //        {
                //            int pageNum = GetPageNumber(shape.Anchor);
                //            if (pageNum == 4)
                //            {
                //                if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoPicture)
                //                {
                //                    shape.RelativeHorizontalPosition = Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
                //                    shape.RelativeVerticalPosition = Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
                //                }
                //            }
                //        }
                //    }
                //    catch
                //    {
                //        // ignore shapes that can't report anchor/page
                //    }
                //}


                for (int x = 0; x < ArrFilesZoomed.Count; x++)
                {
                    AddImage(wordApp, docMain, dtShapesZoomed, ArrFilesZoomed[x], x + 1, ArrFilesZoomed.Count);
                }


            }

            return;




            //    int cnt = 0;
            //foreach (string file in ArrFilesZoomed)
            //{
            //    cnt = cnt + 1;

            //    if (dtShapesZoomed.ContainsKey(cnt) == false) continue;
            //    if (System.IO.File.Exists(file) == false) continue;

            //    ShapeInfo oshObj = dtShapesZoomed[cnt];

            //    // Force Word to update layout
            //    docMain.Repaginate();
            //    wordApp.ScreenRefresh();
            //    System.Threading.Thread.Sleep(100);

            //    //Go to the page
            //    Word.Range rng = wordApp.Selection.GoTo(Word.WdGoToItem.wdGoToPage, Word.WdGoToDirection.wdGoToAbsolute, oshObj.m_nSheetNo);
            //    System.Threading.Thread.Sleep(1000);
            //    rng.Collapse(Word.WdCollapseDirection.wdCollapseStart);

            //    // Define where to insert the image
            //    Word.Range range = docMain.Range(0, 0);
            //    range = docMain.Range(range.End, range.End);  // This moves the range to the end of the current text


            //    // -------Insert the picture -----------
            //    Word.Shape picture = docMain.Shapes.AddPicture(file);//
            //    //Image rotated...Portrait mode
            //    float diff = 0f;
            //    if (picture.Rotation == 90 || picture.Rotation == 270)
            //    {
            //        Image img = Image.FromFile(file);
            //        float before = oshObj.m_fHeight;
            //        oshObj.m_fHeight = oshObj.m_fHeight * img.Height / img.Width;
            //        diff = before - oshObj.m_fHeight;
            //        oshObj.m_fImgTop = oshObj.m_fImgTop + diff;
            //    }
            //    // Important settings
            //    picture.WrapFormat.Type = Word.WdWrapType.wdWrapNone;
            //    picture.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
            //    picture.Height = oshObj.m_fHeight;
            //    picture.Left = oshObj.m_fImgLeft;
            //    picture.Top = oshObj.m_fImgTop;
            //    picture.RelativeVerticalPosition = Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
            //    float ImageBottom = picture.Top + picture.Height;
            //    float imageWidth = picture.Width;
            //    float imageLeft = picture.Left;
            //    if (picture.Rotation == 90 || picture.Rotation == 270)
            //    {
            //        imageWidth = picture.Height;
            //        float visualTop = picture.Top;
            //        visualTop += (picture.Width - picture.Height) / 2;
            //        ImageBottom = visualTop + Math.Min(picture.Width, picture.Height);
            //        float leftOffset = (picture.Width - picture.Height) / 2;
            //        imageLeft = picture.Left + leftOffset;
            //    }
            //    //Single image special case for HUB - LPC ROTOR, 1ST STAGE, Loc[15], as per request, need to confirm with customer if this is only for this case or we need to handle in other cases as well
            //    string stype = "", ssubtype = "";
            //    if (stype == "HUB - LPC ROTOR, 1ST STAGE" && ssubtype == "Loc[15]" && ArrFilesZoomed.Count == 1)
            //    {
            //        picture.Height = 288;
            //        picture.Width = 360;
            //        picture.Left = 94;
            //        picture.Top = 320;
            //    }

            //    //Z order
            //    picture.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoSendToBack);
            //    range = docMain.Range(picture.Anchor.End, picture.Anchor.End);


            //    //2nd image, make it red border and dash line
            //    if (cnt == 2)
            //    {
            //        picture.Line.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
            //        // Set border color to Red (use System.Drawing.Color to get RGB OLE int)
            //        picture.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Red);
            //        // Set border width/weight to 1.5 points
            //        picture.Line.Weight = 1.5f;
            //        // Set dash style to long dash
            //        picture.Line.DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineLongDash;
            //    }

            //    //------- Add the Image number text box ------
            //    Word.Shape textBox1 = docMain.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, oshObj.m_fNumLeft, oshObj.m_fNumTop, 200, 20, picture.Anchor);
            //    // Important settings
            //    textBox1.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            //    textBox1.TextFrame.TextRange.Text = m_sConstructionNum.TrimEnd(new char[1] { ':' }) + " " + System.IO.Path.GetFileName(file);
            //    textBox1.TextFrame.TextRange.Font.Size = 8f;
            //    textBox1.TextFrame.TextRange.Font.Name = "Arial";
            //    textBox1.TextFrame.TextRange.Font.Color = Word.WdColor.wdColorWhite;
            //    textBox1.TextFrame.TextRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            //    textBox1.Left = imageLeft;
            //    textBox1.Top = ImageBottom;
            //    textBox1.Width = imageWidth;
            //    textBox1.RelativeHorizontalPosition = picture.RelativeHorizontalPosition;
            //    textBox1.WrapFormat.Type = Word.WdWrapType.wdWrapNone;
            //    textBox1.RelativeVerticalPosition = Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
            //    textBox1.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
            //    textBox1.Top = ImageBottom-20;

            //    //------- Add the Title text box ------
            //    Word.Shape textBox2 = docMain.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, oshObj.m_fNumLeft, oshObj.m_fNumTop, 200, 20, picture.Anchor);
            //    // Important settings
            //    textBox2.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            //    textBox2.TextFrame.TextRange.Text = oshObj.m_sText;
            //    textBox2.TextFrame.TextRange.Font.Name = "Arial";
            //    textBox2.TextFrame.TextRange.Font.Size = 9f;
            //    textBox2.TextFrame.TextRange.Font.Bold = 1;
            //    textBox2.TextFrame.TextRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //    textBox2.Left = imageLeft;
            //    textBox2.Top = ImageBottom;
            //    textBox2.Width = imageWidth;
            //    textBox2.RelativeHorizontalPosition = picture.RelativeHorizontalPosition;
            //    textBox2.WrapFormat.Type = Word.WdWrapType.wdWrapNone;
            //    textBox2.RelativeVerticalPosition = Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
            //    textBox2.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
            //    textBox2.Top = ImageBottom;
            //}
        }
        void AddImage(Word.Application wordApp, Word.Document docMain, Dictionary<int, ShapeInfo> dtShapesZoomed, string file, int imgNumCount, int NumOfImages)
        {
            try
            {

                    if (dtShapesZoomed.ContainsKey(imgNumCount) == false) return;
                    if (System.IO.File.Exists(file) == false) return;

                    ShapeInfo oshObj = dtShapesZoomed[imgNumCount];
                    
                     

                    // Force Word to update layout
                    docMain.Repaginate();
                    wordApp.ScreenRefresh();
                    System.Threading.Thread.Sleep(100);

                    //Go to the page
                    Word.Range rng = wordApp.Selection.GoTo(Word.WdGoToItem.wdGoToPage, Word.WdGoToDirection.wdGoToAbsolute, oshObj.m_nSheetNo);
                    System.Threading.Thread.Sleep(1000);
                    rng.Collapse(Word.WdCollapseDirection.wdCollapseStart);

                    // Define where to insert the image
                    Word.Range range = docMain.Range(0, 0);
                    range = docMain.Range(range.End, range.End);  // This moves the range to the end of the current text


                    // -------Insert the picture -----------
                    Word.Shape picture = docMain.Shapes.AddPicture(file);//
                                                                         //Image rotated...Portrait mode
                    float diff = 0f;
                    if (picture.Rotation == 90 || picture.Rotation == 270)
                    {
                        Image img = Image.FromFile(file);
                        float before = oshObj.m_fHeight;
                        oshObj.m_fHeight = oshObj.m_fHeight * img.Height / img.Width;
                        diff = before - oshObj.m_fHeight;
                        oshObj.m_fImgTop = oshObj.m_fImgTop + diff;
                    }
                    // Important settings
                    picture.WrapFormat.Type = Word.WdWrapType.wdWrapNone;
                    picture.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
                    picture.Height = oshObj.m_fHeight;
                    picture.Left = oshObj.m_fImgLeft;
                    if (imgNumCount % 2 == 1 && imgNumCount == NumOfImages) picture.Left = 170;
                    picture.Top = oshObj.m_fImgTop;
                    picture.RelativeVerticalPosition = Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
                    float ImageBottom = picture.Top + picture.Height;
                    float imageWidth = picture.Width;
                    float imageLeft = picture.Left;
                    if (picture.Rotation == 90 || picture.Rotation == 270)
                    {
                        imageWidth = picture.Height;
                        float visualTop = picture.Top;
                        visualTop += (picture.Width - picture.Height) / 2;
                        ImageBottom = visualTop + Math.Min(picture.Width, picture.Height);
                        float leftOffset = (picture.Width - picture.Height) / 2;
                        imageLeft = picture.Left + leftOffset;
                    }
                    //Single image special case for HUB - LPC ROTOR, 1ST STAGE, Loc[15], as per request, need to confirm with customer if this is only for this case or we need to handle in other cases as well
                    string stype = "", ssubtype = "";
                    if (stype == "HUB - LPC ROTOR, 1ST STAGE" && ssubtype == "Loc[15]" && NumOfImages == 1)
                    {
                        picture.Height = 288;
                        picture.Width = 360;
                        picture.Left = 94;
                        picture.Top = 320;
                    }

                    //Z order
                    picture.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoSendToBack);
                    range = docMain.Range(picture.Anchor.End, picture.Anchor.End);


                    //2nd image, make it red border and dash line
                    if (imgNumCount == 2)
                    {
                        picture.Line.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                        // Set border color to Red (use System.Drawing.Color to get RGB OLE int)
                        picture.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Red);
                        // Set border width/weight to 1.5 points
                        picture.Line.Weight = 1.5f;
                        // Set dash style to long dash
                        picture.Line.DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineLongDash;
                    }

                    //------- Add the Image number text box ------
                    Word.Shape textBox1 = docMain.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, oshObj.m_fNumLeft, oshObj.m_fNumTop, 200, 20, picture.Anchor);
                    // Important settings
                    textBox1.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                    textBox1.TextFrame.TextRange.Text = m_sConstructionNum.TrimEnd(new char[1] { ':' }) + " " + System.IO.Path.GetFileName(file);
                    textBox1.TextFrame.TextRange.Font.Size = 8f;
                    textBox1.TextFrame.TextRange.Font.Name = "Arial";
                    textBox1.TextFrame.TextRange.Font.Color = Word.WdColor.wdColorWhite;
                    textBox1.TextFrame.TextRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                    textBox1.Left = imageLeft-40;
                    textBox1.Top = ImageBottom;
                    textBox1.Width = imageWidth+40;
                    textBox1.RelativeHorizontalPosition = picture.RelativeHorizontalPosition;
                    textBox1.WrapFormat.Type = Word.WdWrapType.wdWrapNone;
                    textBox1.RelativeVerticalPosition = Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
                    textBox1.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                    textBox1.Top = ImageBottom - 20;

                    //------- Add the Title text box ------
                    Word.Shape textBox2 = docMain.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, oshObj.m_fNumLeft, oshObj.m_fNumTop, 200, 20, picture.Anchor);
                    // Important settings
                    textBox2.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                    textBox2.TextFrame.TextRange.Text = oshObj.m_sText;
                    textBox2.TextFrame.TextRange.Font.Name = "Arial";
                    textBox2.TextFrame.TextRange.Font.Size = 9f;
                    textBox2.TextFrame.TextRange.Font.Bold = 1;
                    textBox2.TextFrame.TextRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    textBox2.Left = imageLeft-40;
                    textBox2.Top = ImageBottom;
                    textBox2.Width = imageWidth+80;
                    textBox2.RelativeHorizontalPosition = picture.RelativeHorizontalPosition;
                    textBox2.WrapFormat.Type = Word.WdWrapType.wdWrapNone;
                    textBox2.RelativeVerticalPosition = Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
                    textBox2.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                    textBox2.Top = ImageBottom;
                
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
        }
        Word.WdColor ConvertToWdColor(System.Drawing.Color color)
        {
            return (Word.WdColor)(color.R + (color.G << 8) + (color.B << 16));
        }
        ShapeInfo getImageLocation(int pageNum, int imagecnt, ConfigInfo cfg, string text)
        {
            //if(imagecnt > 8)
            //{
            //    if (imagecnt % 2 == 1) imagecnt = 7;
            //    imagecnt = 8;
            //}
            //if (imagecnt <= 4)
            {
                int pgnum = pageNum + getPageNumber(imagecnt);
                //Left
                float imgLeft = 60;
                if (imagecnt % 2 == 0) imgLeft = 280;

                //Top
                float imgTop = cfg.lstCoordinates[0].m_fImgTop;

                //Height
                float height = cfg.lstCoordinates[0].m_fImageHeight;// 165;// Utility.m_fLast - imgTop;
                if (imagecnt > 2) height = 165;

                //Next page images
                if (imagecnt > 2)
                {
                    imgTop = ((imagecnt - 3) / 2) * 175 + 35;
                }


                ShapeInfo shObj = new ShapeInfo(pageNum + getPageNumber(imagecnt), imgLeft, imgTop, 0, 0, 0, 0, text);
                shObj.m_fHeight = height;
                return shObj;
            }




            if (imagecnt <= 4)
            {
                return new ShapeInfo(pageNum + getPageNumber(imagecnt), cfg.lstCoordinates[imagecnt - 1].m_fImgLeft, cfg.lstCoordinates[imagecnt - 1].m_fImgTop, cfg.lstCoordinates[imagecnt - 1].m_fNumLeft, cfg.lstCoordinates[imagecnt - 1].m_fNumTop, cfg.lstCoordinates[imagecnt - 1].m_fTitleLeft, cfg.lstCoordinates[imagecnt - 1].m_fTitleTop, text);
            }
            else
            {
                int loc = (imagecnt - 5) / 2 + 1;
                float dist = loc * 200;
                int index = 3;
                if (imagecnt % 2 == 1) index = 2;//Odd image
                return new ShapeInfo(pageNum + getPageNumber(imagecnt), cfg.lstCoordinates[index].m_fImgLeft, cfg.lstCoordinates[index].m_fImgTop + dist, cfg.lstCoordinates[index].m_fNumLeft, cfg.lstCoordinates[index].m_fNumTop + dist, cfg.lstCoordinates[index].m_fTitleLeft, cfg.lstCoordinates[index].m_fTitleTop + dist, text);
            }
            return null;
        }
        int getPageNumber(int imgcnt)
        {
            if (imgcnt > 2) return 1;
            return 0;
        }
        Word.Range GetPageRange(Word.Application wordApp, Word.Document doc, int pageNumber)
        {
            object what = Word.WdGoToItem.wdGoToPage;
            object which = Word.WdGoToDirection.wdGoToAbsolute;
            object page = pageNumber;

            try
            {
                // Force Word to update layout veeru
                doc.Repaginate();
                wordApp.ScreenRefresh();
                System.Threading.Thread.Sleep(100);
                //veeru

                object obj = Type.Missing;
                Word.Range startRange = doc.GoTo(ref what, ref which, ref page, ref obj);

                //--veeru
                System.Threading.Thread.Sleep(1000);
                startRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                //--veeru


                // Try to get start of next page
                object nextPage = pageNumber + 1;
                Word.Range endRange;

                try
                {
                    endRange = doc.GoTo(ref what, ref which, ref nextPage, ref obj);
                }
                catch
                {
                    // If next page doesn't exist, use document end
                    endRange = doc.Content;
                    endRange.Start = doc.Content.End;
                }

                int start = startRange.Start;
                int end = (endRange.Start > start) ? endRange.Start - 1 : doc.Content.End;

                Word.Range pageRange = doc.Range(start, end);

                // Check if range has text or content
                if (pageRange != null && pageRange.End > pageRange.Start)
                    return pageRange;
                else
                    return null;
            }
            catch
            {
                // Page doesn't exist or error
                return null;
            }
        }

        private void Read_ESN_TSN_CSN(string folderPath)
        {
            try
            {
                string upfolder = System.IO.Path.GetDirectoryName(folderPath);
                //Find all PDFs
                string[] ArrFiles = System.IO.Directory.GetFiles(upfolder, "*.pdf");
                foreach (string file in ArrFiles)
                {
                    PDFReader oreader = new PDFReader();
                    oreader.ReadPDF(file);
                    List<PDFPageInfo> lstPdfPages = oreader.m_lstPdfs;
                    foreach (PDFPageInfo opage in lstPdfPages)
                    {
                        foreach (string lne in opage.m_lstPageLines)
                        {
                            if (m_sESN.Length == 0 && lne.StartsWith("ESN:"))
                            {
                                m_sESN = lne.Replace("ESN:", "").Trim();
                            }
                            else if (m_sTSN.Length == 0 && lne.StartsWith("TSN:"))
                            {
                                m_sTSN = lne.Replace("TSN:", "").Trim();
                            }
                            else if (m_sCSN.Length == 0 && lne.StartsWith("CSN:"))
                            {
                                m_sCSN = lne.Replace("CSN:", "").Trim();
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
        private void read_PDF(string pdffile)
        {

            try
            {
                PDFReader oreader = new PDFReader();
                oreader.ReadPDF(pdffile);
                List<PDFPageInfo> lstPdfPages = oreader.m_lstPdfs;

                foreach (PDFPageInfo opage in lstPdfPages)
                {
                    int DescStart = -1;
                    for (int i = 0; i < opage.m_lstHeights.Count; i++)
                    {
                        HeightInfo oht = opage.m_lstHeights[i];
                        string value = updateAscii(oht.m_sValue);
                        string value2 = value.Replace(" ", "").Trim();
                        if (value2.StartsWith("P/N") && m_sPN.Length == 0)
                        {
                            //Find the Single Repair
                            for (int p = i - 1; p >= 0; p--)
                            {
                                if (opage.m_lstHeights[p].m_sValue.ToUpper().Contains("Single Repair".ToUpper()))
                                {
                                    bIsSingleRepair = true;
                                    break;
                                }
                            }
                            //--
                            m_sPN = value2.Replace("P/N", "").TrimStart(new char[1] { ':' }).Trim();
                            DescStart = i + 1;
                        }
                        else if (value2.StartsWith("S/N") && m_sSN.Length == 0)
                        {
                            m_sSN = value2.Replace("S/N", "").TrimStart(new char[1] { ':' }).Trim();
                            DescStart = i + 1;
                        }
                        else if (value.StartsWith("NOMEN") && m_sNOMEN.Length == 0)
                        {
                            m_sNOMEN = value.Replace("NOMEN", "").TrimStart(new char[1] { ':' }).Trim();
                            DescStart = i + 1;
                        }
                        else if (value2.StartsWith("CSN") && m_sCSN.Length == 0)
                        {
                            m_sCSN = value2.Replace("CSN", "").TrimStart(new char[1] { ':' }).Trim();
                        }
                        else if (value2.StartsWith("TSN") && m_sTSN.Length == 0)
                        {
                            m_sTSN = value2.Replace("TSN", "").TrimStart(new char[1] { ':' }).Trim();
                        }
                        else if (value2.StartsWith("ESN") && m_sESN.Length == 0)
                        {
                            m_sESN = value2.Replace("ESN", "").TrimStart(new char[1] { ':' }).Trim();
                        }
                        else if (value2.ToUpper().StartsWith("ATANo".ToUpper()) && m_sATA.Length == 0)
                        {
                            m_sATA = value2.ToUpper().Replace("ATANo".ToUpper(), "").TrimStart(new char[1] { ':' }).TrimStart(new char[1] { '.' }).Trim();
                        }
                    }
                    if (DescStart != -1)
                    {
                        if (m_lstDamageInfo.Count == 0)
                        {
                            for (int j = DescStart; j < opage.m_lstHeights.Count; j++)
                            {
                                HeightInfo oht = opage.m_lstHeights[j];
                                string value = updateAscii(oht.m_sValue);
                                if (value.Length > 0)
                                {
                                    if (bIsContinue(value)) continue;
                                    if (bIsEnd(value) && m_lstDamageInfo.Count > 0) break;
                                    if (value.StartsWith("CSN") && m_lstDamageInfo.Count == 0) continue;
                                    m_lstDamageInfo.Add(value);
                                }
                            }
                        }
                    }
                }

                if (m_sNOMEN.Length == 0)
                {
                    foreach (PDFPageInfo opage in lstPdfPages)
                    {
                        //foreach (string s in opage.m_lstPageLines)
                        for (int i = 1; i < opage.m_lstPageLines.Count; i++)
                        {
                            string v = FindNomenclature(opage.m_lstPageLines[i]);
                            if (v.Length > 0)
                            {
                                m_sNOMEN = v;
                                break;
                            }
                        }
                        if (m_sNOMEN.Length > 0) break;
                    }
                }

                //Find the construction number
                if (m_sPN.Length > 0 && m_sConstructionNum.Length == 0)
                {
                    //find from Text
                    foreach (PDFPageInfo opage in lstPdfPages)
                    {
                        foreach (string sx in opage.m_lstPageLines)
                        {
                            List<string> lstp = Utility.SplitString(sx, " ");
                            foreach (string s in lstp)
                            {
                                if (s.StartsWith("#"))
                                {
                                    m_sConstructionNum = s;
                                    break;
                                }
                            }
                            if (m_sConstructionNum.Length > 0) break;
                        }
                        if (m_sConstructionNum.Length > 0) break;
                    }
                    //find from file name
                    if (m_sConstructionNum.Length == 0)
                    {
                        List<string> lstp = Utility.SplitString(System.IO.Path.GetFileNameWithoutExtension(pdffile), "_");
                        m_sConstructionNum = lstp[0];
                    }
                }

                //Update nomenclature
                if (m_sNOMEN.Length > 0)
                {
                    try
                    {
                        m_sNOMEN = Utility.SplitString(m_sNOMEN, "Area")[0].Trim();
                    }
                    catch { }
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
        }
        private string FindNomenclature(string value)
        {
            try
            {
                if (value.Contains("Finding"))
                {
                    string final = "";
                    List<string> lst = Utility.SplitString(value, " ");
                    foreach (string s in lst)
                    {
                        if (s.Contains("Finding")) continue;
                        if (s.Contains("_")) continue;
                        if (s.Contains("(")) continue;
                        final = final + s + " ";
                    }
                    return final.Trim();
                }
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return "";
        }
        private void read_Word(string pdffile)
        {

            try
            {
                object fileName = pdffile;
                object readOnly = true;
                object isVisible = true;
                object missing = System.Reflection.Missing.Value;
                //Create word object
                Word.Application wordApp = new Word.Application();
                wordApp.Visible = false;
                wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                Word.Document doc = wordApp.Documents.Open(ref fileName, ref missing, ref readOnly, ref missing, ref missing, ref missing,
                                                            ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

                //read it here
                int DescStart = -1;
                for (int i = 1; i <= doc.Paragraphs.Count; i++)
                {
                    Word.Paragraph paragraph = doc.Paragraphs[i];
                    string value = paragraph.Range.Text;
                    value = Regex.Replace(value, @"[\uFF00-\uFFFF]", "");
                    value = value.Replace("\r", "").Replace("\a", "").Replace("\n", "").Trim();
                    //Utility.WriteErrorLog("", "", value);
                    if (value.Length == 0) continue;
                    string value2 = value.Replace(" ", "").Trim();
                    if (value2.StartsWith("P/N") && m_sPN.Length == 0)
                    {
                        //Find the Single Repair
                        for (int p = i - 1; p >= 1; p--)
                        {
                            Word.Paragraph paragraph1 = doc.Paragraphs[p];
                            string value1 = paragraph1.Range.Text;
                            if (value1.ToUpper().Contains("Single Repair".ToUpper()))
                            {
                                bIsSingleRepair = true;
                                break;
                            }
                        }
                        //--
                        m_sPN = value2.Replace("P/N", "").TrimStart(new char[1] { ':' }).Trim();
                        DescStart = i + 1;
                    }
                    else if (value2.StartsWith("S/N") && m_sSN.Length == 0)
                    {
                        m_sSN = value2.Replace("S/N", "").TrimStart(new char[1] { ':' }).Trim();
                        DescStart = i + 1;
                    }
                    else if (value2.StartsWith("CSN") && m_sCSN.Length == 0)
                    {
                        m_sCSN = value2.Replace("CSN", "").TrimStart(new char[1] { ':' }).Trim();
                    }
                    else if (value2.StartsWith("TSN") && m_sTSN.Length == 0)
                    {
                        m_sTSN = value2.Replace("TSN", "").TrimStart(new char[1] { ':' }).Trim();
                    }
                    else if (value2.StartsWith("ESN") && m_sESN.Length == 0)
                    {
                        m_sESN = value2.Replace("ESN", "").TrimStart(new char[1] { ':' }).Trim();
                    }
                    else if (value2.ToUpper().StartsWith("ATANo".ToUpper()) && m_sATA.Length == 0)
                    {
                        m_sATA = value2.ToUpper().Replace("ATANo".ToUpper(), "").TrimStart(new char[1] { ':' }).TrimStart(new char[1] { '.' }).Trim();
                    }
                    else if (value.StartsWith("NOMEN") && m_sNOMEN.Length == 0)
                    {
                        m_sNOMEN = value.Replace("NOMEN", "").TrimStart(new char[1] { ':' }).Trim();
                        if (m_sNOMEN.Contains("Damage"))
                        {
                            int n = m_sNOMEN.IndexOf("Damage");
                            m_lstDamageInfo.Add(m_sNOMEN.Substring(n, m_sNOMEN.Length - n).Trim());
                            m_sNOMEN = m_sNOMEN.Substring(0, n).Trim();
                        }
                        DescStart = i + 1;
                        break;
                    }
                }
                if (DescStart != -1)
                {
                    if (m_lstDamageInfo.Count == 0)
                    {
                        for (int j = DescStart; j <= doc.Paragraphs.Count; j++)
                        {
                            Word.Paragraph paragraph = doc.Paragraphs[j];
                            string value = paragraph.Range.Text;
                            value = Regex.Replace(value, @"[\uFF00-\uFFFF]", "");
                            value = value.Replace("\r", "").Replace("\a", "").Replace("\n", "").Trim();
                            if (value.Length > 0)
                            {
                                if (bIsContinue(value)) continue;
                                if (bIsEnd(value) && m_lstDamageInfo.Count > 0) break;
                                if (value.StartsWith("CSN") && m_lstDamageInfo.Count == 0) continue;

                                if (value.Contains("Insp Task"))
                                {
                                    int n = value.IndexOf("Insp Task");
                                    m_lstDamageInfo.Add(value.Substring(0, n).Trim());
                                    break;
                                }
                                m_lstDamageInfo.Add(value);
                            }
                        }
                    }
                }
                if (m_sNOMEN.Length == 0)
                {
                    for (int i = 1; i <= doc.Paragraphs.Count; i++)
                    {
                        Word.Paragraph paragraph = doc.Paragraphs[i];
                        string value = paragraph.Range.Text;
                        value = Regex.Replace(value, @"[\uFF00-\uFFFF]", "");
                        value = value.Replace("\r", "").Replace("\a", "").Replace("\n", "").Trim();
                        string v = FindNomenclature(value);
                        if (v.Length > 0)
                        {
                            m_sNOMEN = v;
                            break;
                        }
                    }
                }
                if (m_sPN.Length > 0 && m_sConstructionNum.Length == 0)
                {
                    //find from Text
                    for (int i = 1; i <= doc.Paragraphs.Count; i++)
                    {
                        Word.Paragraph paragraph = doc.Paragraphs[i];
                        string value = paragraph.Range.Text;
                        value = Regex.Replace(value, @"[\uFF00-\uFFFF]", "");
                        value = value.Replace("\r", "").Replace("\a", "").Replace("\n", "").Trim();
                        List<string> lstp = Utility.SplitString(value, " ");
                        foreach (string s in lstp)
                        {
                            if (s.StartsWith("#"))
                            {
                                m_sConstructionNum = s;
                                break;
                            }
                        }
                        if (m_sConstructionNum.Length > 0) break;
                    }
                    //find from file name
                    if (m_sConstructionNum.Length == 0)
                    {
                        List<string> lstp = Utility.SplitString(System.IO.Path.GetFileNameWithoutExtension(pdffile), "_");
                        m_sConstructionNum = lstp[0];
                    }
                }
                //---
                doc.Close();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                wordApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
        }
        bool bIsEnd(string value)
        {

            try
            {
                value = value.ToLower();
                if(value == "file") return true;
                if (value == "files") return true;
                if (value == "others") return true;
                if (value == "参照") return true;
                if (value == "その他") return true;
                if (value == "ファイル") return true;
                if (value == "追加スキム無し") return true;
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return false;
        }
        bool bIsRemove(string value)
        {

            try
            {
                if (value.Replace(" ", "").ToUpper().StartsWith("EA:".Replace(" ", "").ToUpper())) return true;
                if (value.Replace(" ", "").ToUpper().StartsWith("Price:".Replace(" ", "").ToUpper())) return true;
                if (value.Replace(" ", "").ToUpper().StartsWith("Insp Task".Replace(" ", "").ToUpper())) return true;
                if (value.Replace(" ", "").ToUpper().StartsWith("CSN".Replace(" ", "").ToUpper())) return true;
                if (value.Replace(" ", "").ToUpper().StartsWith("TSN".Replace(" ", "").ToUpper())) return true;
                if (value.Replace(" ", "").ToUpper().StartsWith("ESN".Replace(" ", "").ToUpper())) return true;
                if (value.Replace(" ", "").ToUpper().StartsWith("Engine Manual No".Replace(" ", "").ToUpper())) return true;
                if (value.Replace(" ", "").ToUpper().StartsWith("ATA".Replace(" ", "").ToUpper())) return true;
                if (value.Replace(" ", "").ToUpper().StartsWith("CIR MANUAL No".Replace(" ", "").ToUpper())) return true;
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return false;
        }
        bool bIsContinue(string value)
        {
            try
            {
                if (value.ToUpper().StartsWith("ODR")) return true;
                if (value.ToUpper().StartsWith("CUST")) return true;
            }
            catch (Exception ee)
            {
                Utility.WriteErrorLog(ee);
            }
            return false;
        }
    }
    public class AsciiInfo
    {
        public int m_nAsciiNum = 0;
        public string m_sVal = "";
    }
    public class ShapeInfo
    {
        public int m_nSheetNo = 1;
        public float m_fWidth = 0f;
        public float m_fHeight = 0f;
        public float m_fImgLeft = 0f;
        public float m_fImgTop = 0f;
        public float m_fNumLeft = 0f;
        public float m_fNumTop = 0f;
        public float m_fTitleLeft = 0f;
        public float m_fTitleTop = 0f;

        public string m_sText = "";
        public ShapeInfo(int shno, float left, float top, float numleft, float numtop, float Titleleft, float Titletop, string txt)
        {
            m_nSheetNo = shno;
            m_fImgLeft = left;
            m_fImgTop = top;

            m_fNumLeft = numleft;
            m_fNumTop = numtop;
            m_fTitleLeft = Titleleft;
            m_fTitleTop = Titletop;


            m_fWidth = (float)(2.98 * 72.0);
            m_fHeight = (float)(2.24 * 72.0);
            m_sText = txt;
        }
        //public ShapeInfo(int shno, float left, float top, float wd, float ht)
        //{
        //    m_nSheetNo = shno;
        //    m_fLeft = left;
        //    m_fTop = top;
        //    m_fWidth = (float)(wd * 72.0);
        //    m_fHeight = (float)(ht * 72.0);

        //}
    }
    public class WordDate
    {
        public Word.WdColor m_oclr = Word.WdColor.wdColorBlack;
        public string m_sTitle = "";
        public string m_sValue = "";
        public WordDate(string title, string value, Word.WdColor clr = Word.WdColor.wdColorBlack)
        {
            m_sTitle = title;
            m_sValue = value;
            m_oclr = clr;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UIRequirement;

public class DamageInfo
{
    public int m_nNoOfPages = 1;
    public string m_sDamageInfo = "";
    public string m_sType = "";
    public string m_sSubType = "";
    public string m_sImage1 = "";
    public string m_sImage2 = "";
    public string m_sImage3 = "";
    // public string m_sImage4 = "";
    public List<MoreImages> addmoreimages = new List<MoreImages>();
    //public List<string> addmoreimages { get; set; } = new List<string>();
    public string m_sNonConformanceType = "";
    public ConfigInfo m_oConfig = null;
    public string m_sManualCriteria = "";
}

public class ConfigInfo
{
    public string m_sType = "";
    public string m_sSubType = "";
    public string m_sFile = "";
    public List<CoordinatesInfo> lstCoordinates = new List<CoordinatesInfo>();
}

public class CoordinatesInfo
{
    public float m_fImgLeft = 0f;
    public float m_fImgTop = 0f;
    public float m_fImgLeft2 = 0f;
    public float m_fNumLeft = 0f;
    public float m_fNumTop = 0f;
    public float m_fTitleLeft = 0f;
    public float m_fTitleTop = 0f;
    public float m_fImageHeight = 0f;
}

public class MoreImages
{
    public string m_sFilename = "";
    public string m_sName = "";
}

public class PDFPageInfo
{
    public int m_nPageNumber = 1;
    //public string m_sPageNum = "";
    public float m_fPageHeight = 0f;
    public float m_fPageWidth = 0f;
    public List<HeightInfo> m_lstHeights = null;
    public List<string> m_lstPageLines = new List<string>();
    public List<string> m_lstPageLinesFromFile = new List<string>();
    public PDFPageInfo() { }
    public PDFPageInfo(int pg)
    {
        m_nPageNumber = pg;
    }
}

public class HeightInfo
{
    public PDFPageInfo m_oPage = null;
    public int m_nIndex = 0;
    public float m_fVal = 0f;
    public string m_sValue = "";
    public List<RectAndText> m_lstRectAndText = new List<RectAndText>();
    //
    public List<string> m_lstTexts = new List<string>();
    //
    public List<HeightInfo> m_lstHts = new List<HeightInfo>();
}


public class Ret
{
    public string ESN { get; set; }
    public string TSN { get; set; }
    public string CSN { get; set; }
    public string TicketDetails{ get; set; }


    public List<string> Images { get; set; }
        = new();

}
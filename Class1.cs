using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UIRequirement
{
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
}

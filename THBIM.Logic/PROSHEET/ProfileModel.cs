using Autodesk.Revit.DB.ExtensibleStorage;
using System;
using System.Collections.Generic;
using System.Xml.Serialization;
using Autodesk.Revit.DB;

namespace THBIM
{
    

    [XmlRoot("Profiles")]

    public class ProfilesRoot
    {
        [XmlArray("List")]
        [XmlArrayItem("Profile")]
        public List<ProfileXml> List { get; set; } = new List<ProfileXml>();
    }

    public class ProfileXml
    {
        public string Name { get; set; }
        public bool IsCurrent { get; set; }
        public string FilePath { get; set; }
        public int Version { get; set; } = 1;
        public TemplateInfoXml TemplateInfo { get; set; } = new TemplateInfoXml();
    }

    public class TemplateInfoXml
    {
        // Các thông số Tab Create
        public string CreateExportFolderPath { get; set; }
        public bool CreateSplitFolder { get; set; }

        // Các thông số Tab Format (General)
        public string DWGSettingName { get; set; }
        public bool IsCenter { get; set; }
        public string OffsetX { get; set; }
        public string OffsetY { get; set; }
        public bool IsFitToPage { get; set; }
        public int Zoom { get; set; }
        public string RasterQuality { get; set; }
        public string Color { get; set; }
        public bool HidePlanes { get; set; }
        public bool HideScopeBox { get; set; }
        public bool HideUnreferencedTags { get; set; }
        public bool HideCropBoundaries { get; set; }
        public bool IsSeparateFile { get; set; }
        public bool IsPDFChecked { get; set; }
        public bool IsDWGChecked { get; set; }

        // Các thông số DWG
        public bool DWG_MergedViews { get; set; }
        public bool DWG_BindImages { get; set; }
        public bool CleanPcp { get; set; }

        // Phần đặt tên (Naming Rules)
        public SelectSheetParametersXml SelectSheetParameters { get; set; } = new SelectSheetParametersXml();
    }

    public class SelectSheetParametersXml
    {
        public string CombineParameterName { get; set; }
        [XmlArray("CombineParameters")]
        [XmlArrayItem("ParameterModel")]
        public List<ParameterModelXml> CombineParameters { get; set; } = new List<ParameterModelXml>();
    }

    public class ParameterModelXml
    {
        public string ParameterName { get; set; }
        public string Prefix { get; set; }
        public string Suffix { get; set; }
    }
}
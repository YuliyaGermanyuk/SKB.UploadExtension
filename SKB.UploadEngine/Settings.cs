using SKB.Base;
using System;
using System.Xml.Linq;

namespace SKB.UploadExtension
{
    static class Settings
    {
        #region Properties
        public static String Domain { get; private set; }
        public static String ServerName { get; private set; }
        public static String ArchiveName { get; private set; }
        public static String ArchiveTempName { get; private set; }
        public static String ArchiveDeleteName { get; private set; }
        public static String MatrixPath { get; private set; }
        public static String CardSheetName { get; private set; }
        public static String FolderSheetName { get; private set; }
        #endregion

        #region Methods

        public static void Load (String xmlConfigPath)
        {
            XElement Element = XElement.Load(xmlConfigPath);
            XElement SupElement = Element.Element("domain");
            Domain = SupElement.IsNull() ? String.Empty : SupElement.Value;
            SupElement = Element.Element("serverName");
            ServerName = SupElement.IsNull() ? String.Empty : SupElement.Value;
            SupElement = Element.Element("archiveName");
            ArchiveName = SupElement.IsNull() ? String.Empty : SupElement.Value;
            SupElement = Element.Element("archiveTempName");
            ArchiveTempName = SupElement.IsNull() ? String.Empty : SupElement.Value;
            SupElement = Element.Element("archiveDeleteName");
            ArchiveDeleteName = SupElement.IsNull() ? String.Empty : SupElement.Value;
            SupElement = Element.Element("matrixPath");
            MatrixPath = SupElement.IsNull() ? String.Empty : SupElement.Value;
            CardSheetName = SupElement.IsNull() || SupElement.Attribute("cardSheetName").IsNull() ? "CardRights" : SupElement.Attribute("cardSheetName").Value;
            FolderSheetName = SupElement.IsNull() || SupElement.Attribute("folderSheetName").IsNull() ? "FolderRights" : SupElement.Attribute("folderSheetName").Value;
        }
        #endregion
    }
}
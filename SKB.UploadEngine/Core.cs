using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;

using DocsVision.Platform.Cards.Constants;
using DocsVision.Platform.ObjectManager;
using DocsVision.Platform.ObjectManager.Metadata;
using DocsVision.Platform.ObjectManager.SearchModel;
using DocsVision.Platform.ObjectManager.SystemCards;
using DocsVision.Platform.Security.AccessControl;
using DocsVision.Platform.StorageServer;

using DocsVision.TakeOffice.Cards.Constants;
using DocsVision.TakeOffice.ObjectModel.Resolution;

using SKB.Base;
using SKB.Base.AssignRights;
using SKB.Base.Enums;
using SKB.Base.Ref;

using AssignGroup = SKB.Base.AssignRights.Group;
using CardResolution = DocsVision.TakeOffice.ObjectModel.Resolution.CardResolution;
using VersionedFileCard = DocsVision.Platform.ObjectManager.SystemCards.VersionedFileCard;

namespace SKB.UploadExtension
{
    public static class Core
    {
        #region Fields

        private static readonly NLog.Logger logger;
        public static readonly String ArchivePath;
        public static readonly String ArchiveTempPath;
        public static readonly String ArchiveDeletePath;

        #endregion

        static Core()
        {
            logger = NLog.LogManager.GetCurrentClassLogger();
            String settingsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "settings.xml");
            Settings.Load(settingsPath);

            ArchivePath = String.Format(@"\\{0}\{1}", Settings.ServerName, Settings.ArchiveName);
            ArchiveTempPath = String.Format(@"\\{0}\{1}", Settings.ServerName, Settings.ArchiveTempName);
            ArchiveDeletePath = String.Format(@"\\{0}\{1}", Settings.ServerName, Settings.ArchiveDeleteName);
        }

        public static Boolean CheckDuplication(UserSession Session, String CheckName)
        {
            logger.Info("checkName='{0}'", CheckName);

            String[] CheckNameParts = CheckName.Split('\t');
            if (CheckNameParts.Length == 2)
                return Directory.Exists(Path.Combine(ArchivePath, MyHelper.RemoveInvalidFileNameChars(CheckNameParts[0]), MyHelper.RemoveInvalidFileNameChars(CheckNameParts[1])));
            else if (CheckNameParts.Length == 3)
                return Directory.Exists(Path.Combine(ArchivePath, MyHelper.RemoveInvalidFileNameChars(CheckNameParts[0]), MyHelper.RemoveInvalidFileNameChars(CheckNameParts[1]), MyHelper.RemoveInvalidFileNameChars(CheckNameParts[2])));
            else if (CheckName.Contains("Админ"))
            {
                logger.Info("Big files permission");
                return false;
            }
            else
            {

                CardData UniversalCard = Session.CardManager.GetDictionaryData(RefUniversal.ID);
                SectionQuery Query_Section = UniversalCard.Session.CreateSectionQuery();
                Query_Section.ConditionGroup.Conditions.AddNew(RefUniversal.Item.Name, FieldType.Unistring, ConditionOperation.StartsWith, CheckName);
                RowDataCollection FoundItems = UniversalCard.Sections[RefUniversal.Item.ID].FindRows(Query_Section.GetXml(null, true));

                if (FoundItems.Count > 0)
                {
                    logger.Info("Запись найдена");
                    logger.Info("Идентификатор: " + FoundItems[0].Id);
                    
                    return true;
                }

                logger.Info("Запись не найдена");
                return false;
            }
        }

        public static String PlaceFiles(String TempName, String Label, String Code, String Type, String Version, String Name, String Number, String UnitName)
        {
            using (new Impersonator(ServerExtension.Domain, ServerExtension.User, ServerExtension.SecurePassword))
            {
                logger.Info(String.Format("tempName='{0}'; label='{7}'; code='{1}'; type='{2}'; version='{3}'; name='{4}'; number='{5}'; unitName='{6}'",
                    TempName, Code, Type, Version, Name, Number, UnitName, Label));

                DirectoryInfo TempFolder = new DirectoryInfo(Path.Combine(ArchiveTempPath, TempName));
                DirectoryInfo ArchiveFolder = new DirectoryInfo(ArchivePath);

                if (TempFolder.Exists)
                {
                    Name = MyHelper.RemoveInvalidFileNameChars(Name);
                    UnitName = MyHelper.RemoveInvalidFileNameChars(UnitName);

                    /* Создание папки в архиве */
                    DirectoryInfo UnitDir = ArchiveFolder.CreateSubdirectory(UnitName);

                    /* Формирование названия подпапки */
                    String DocDirName;
                    switch (UnitName)
                    {
                        case RefMarketingFilesCard.Name: DocDirName = Name; break;
                        default: DocDirName = String.Format("{5}{0} {1} [Версия {2}] {3} [ИН {4}]", Code, Type, Version, Name, Number, String.IsNullOrWhiteSpace(Label) || Label.Contains("нет") ? String.Empty : Label.First() + " "); break;
                    }
                    
                    DocDirName = MyHelper.RemoveInvalidFileNameChars(DocDirName);

                    logger.Info("Папка архива: " + UnitName);
                    logger.Info("Подпапка архива: " + DocDirName);

                    DirectoryInfo TargetFolder = UnitDir.CreateSubdirectory(DocDirName);

                    FileInfo[] TempDirFiles = TempFolder.GetFiles();
                    foreach (FileInfo TempFile in TempDirFiles)
                    {
                        String TargetFilePath = Path.Combine(TargetFolder.FullName, TempFile.Name);

                        if (File.Exists(TargetFilePath))
                            File.Delete(TargetFilePath);

                        TempFile.MoveTo(TargetFilePath);
                    }

                    logger.Info("В папку " + TargetFolder.FullName + "загружено файлов: " + TempDirFiles.Length);
                    TempFolder.Delete(true);
                    return TargetFolder.FullName.Last() == '\\' ? TargetFolder.FullName : (TargetFolder.FullName + @"\");
                }
                logger.Warn("Папка " + TempFolder.FullName + " не существует");

                return Boolean.FalseString;
            }
        }

        public static String PlaceFiles(String TempName, String NewFolder)
        {
            using (new Impersonator(ServerExtension.Domain, ServerExtension.User, ServerExtension.SecurePassword))
            {
                logger.Info(String.Format("TempName='{0}'; NewFolder='{1}'", TempName, NewFolder));

                DirectoryInfo TempFolder = new DirectoryInfo(TempName);
                DirectoryInfo ArchiveFolder = new DirectoryInfo(NewFolder);

                if (TempFolder.Exists)
                {
                    if (!ArchiveFolder.Exists)
                        ArchiveFolder.Create();

                    logger.Info("Папка архива: " + ArchiveFolder.FullName);

                    FileInfo[] TempDirFiles = TempFolder.GetFiles();
                    foreach (FileInfo TempFile in TempDirFiles)
                    {
                        String TargetFilePath = Path.Combine(ArchiveFolder.FullName, TempFile.Name);

                        if (File.Exists(TargetFilePath))
                            File.Delete(TargetFilePath);

                        TempFile.MoveTo(TargetFilePath);
                    }

                    logger.Info("В папку " + ArchiveFolder.FullName + "загружено файлов: " + TempDirFiles.Length);
                    TempFolder.Delete(true);
                    return ArchiveFolder.FullName.Last() == '\\' ? ArchiveFolder.FullName : (ArchiveFolder.FullName + @"\");
                }
                logger.Warn("Папка " + TempFolder.FullName + " не существует");

                return Boolean.FalseString;
            }
        }

        public static String AddFiles(String TempName, String TargetPath)
        {
            using (new Impersonator(ServerExtension.Domain, ServerExtension.User, ServerExtension.SecurePassword))
            {
                logger.Info("tempName='{0}'; docDirPath='{1}'", TempName, TargetPath);
                DirectoryInfo TargetFolder = new DirectoryInfo(TargetPath);
                if (TargetFolder == null)
                    logger.Info("docDir=null");
                if (TargetFolder.Exists)
                {
                    DirectoryInfo TempFolder = new DirectoryInfo(Path.Combine(ArchiveTempPath, TempName));

                    if (TempFolder.Exists)
                    {
                        FileInfo[] TempDirFiles = TempFolder.GetFiles();
                        foreach (FileInfo TempFile in TempDirFiles)
                        {
                            String TargetFilePath = Path.Combine(TargetFolder.FullName, TempFile.Name);

                            if (File.Exists(TargetFilePath))
                            {
                                File.SetAttributes(TargetFilePath, FileAttributes.Normal);
                                File.Delete(TargetFilePath);
                            }

                            TempFile.MoveTo(TargetFilePath);
                        }

                        logger.Info("В папку " + TargetFolder.FullName + "добавлено файлов: " + TempDirFiles.Length);
                        TempFolder.Delete(true);
                        return TargetFolder.FullName.Last() == '\\' ? TargetFolder.FullName : (TargetFolder.FullName + @"\");
                    }
                    logger.Warn("Папка " + TempFolder.FullName + " не существует");
                }
                else
                    logger.Warn("Папка " + TargetFolder.FullName + " не существует");

                return Boolean.FalseString;
            }
        }

        public static Boolean Synchronize(UserSession Session, String CardID)
        {
            logger.Info("cardId='{0}'", CardID);
            using (new Impersonator(ServerExtension.Domain, ServerExtension.User, ServerExtension.SecurePassword))
            {
                Guid CardId = new Guid(CardID);
                switch (Session.CardManager.GetCardState(CardId))
                {
                    case ObjectState.Existing:
                        CardData Card = Session.CardManager.GetCardData(CardId);

                        Card.UnlockCard();

                        ExtraCard Extra = Card.GetExtraCard();
                        if (Extra.IsNull())
                        {
                            logger.Info("Синхронизация не производилась, неверный тип карточки.");
                            return false;
                        }

                        if (String.IsNullOrEmpty(Extra.Name) || String.IsNullOrEmpty(Extra.ShortCategory) || (String.IsNullOrEmpty(Extra.Code) && !(Extra is ExtraCardMarketingFiles)))
                        {
                            logger.Info("Синхронизация не производилась, не заполнены обязательные поля.");
                            return false;
                        }

                        
                        String Digest = Extra.ToString();
                        Boolean FileSynch = false;

                        /*Синхронизация с реестром только для КД и ТД*/
                        if (Extra is ExtraCardCD || Extra is ExtraCardTD)
                        {
                            Guid RegItemId = Extra.RegID ?? Guid.Empty;
                            if (RegItemId != Guid.Empty)
                            {
                                try
                                {
                                    RowData RegItem = Session.CardManager.GetDictionaryData(RefUniversal.ID).Sections[RefUniversal.Item.ID].GetRow(RegItemId);
                                    if (RegItem.GetString(RefUniversal.Item.Name) != Digest)
                                    {
                                        logger.Info("Требуется синхронизация...");
                                        RegItem.SetString(RefUniversal.Item.Name, Digest);
                                        foreach (RowData rowData in RegItem.ChildSections[0].Rows)
                                        {
                                            String PropertyName = rowData.GetString(RefUniversal.TypeProperties.Name);
                                            if ((PropertyName == "Код СКБ" && Extra is ExtraCardCD) || (PropertyName == "Обозначение документа" && Extra is ExtraCardTD))
                                            {
                                                rowData.SetString(RefUniversal.TypeProperties.Value, Extra.CodeId);
                                                rowData.SetString(RefUniversal.TypeProperties.DisplayValue, Extra.Code);
                                            }
                                            else if (PropertyName == "Тип документа")
                                            {
                                                rowData.SetString(RefUniversal.TypeProperties.Value, Extra.Category);
                                                rowData.SetString(RefUniversal.TypeProperties.DisplayValue, Extra.Category);
                                            }
                                            else if (PropertyName == "Версия")
                                            {
                                                rowData.SetString(RefUniversal.TypeProperties.Value, Extra.Version);
                                                rowData.SetString(RefUniversal.TypeProperties.DisplayValue, Extra.Version);
                                            }
                                        }
                                        logger.Info("Синхронизация с реестром выполнена.");
                                        FileSynch = true;
                                    }
                                    else
                                        logger.Info("Синхронизация не требуется.");

                                }
                                catch (Exception Ex) { logger.WarnException("Не удалось синхронизировать с записью реестра: " + RegItemId, Ex); }
                            }
                            else
                            {
                                logger.Info("Регистрация в реестре");
                                RowData RegisterRow = Session.CardManager.GetDictionaryData(RefUniversal.ID).GetItemTypeRow(MyHelper.RefItem_TechDocRegister);

                                SubSectionData RegisterTypeProperties = RegisterRow.ChildSections[RefUniversal.TypeProperties.ID];

                                RowData RegistrarUnit = Card.Sections[CardOrd.Properties.ID].GetProperty("Подразделение регистратора");
                                RowData Device = Card.Sections[CardOrd.Properties.ID].GetProperty("Прибор");

                                RowData RegItem = RegisterRow.ChildSections[RefUniversal.Item.ID].Rows.AddNew();
                                RegItem.SetString(RefUniversal.Item.Name, Card.Description);
                                RegItem.SetString(RefUniversal.Item.Comments, Extra.Number);

                                RowDataCollection PropertiesSubSectionRows = RegItem.ChildSections[RefUniversal.Properties.ID].Rows;
                                RowData ItemPropertyRow = PropertiesSubSectionRows.AddNew();
                                ItemPropertyRow.SetGuid(RefUniversal.Properties.Property, RegisterTypeProperties.FindRow("@Name='" + (Extra is ExtraCardCD ? "Код СКБ" : "Обозначение документа") + "'").Id);
                                ItemPropertyRow.SetString(RefUniversal.Properties.DisplayValue, Extra.Code);
                                ItemPropertyRow.SetValue(RefUniversal.Properties.Value, Extra.CodeId, FieldType.Variant);
                                ItemPropertyRow = PropertiesSubSectionRows.AddNew();
                                ItemPropertyRow.SetGuid(RefUniversal.Properties.Property, RegisterTypeProperties.FindRow("@Name='Тип документа'").Id);
                                ItemPropertyRow.SetString(RefUniversal.Properties.DisplayValue, Extra.Category);
                                ItemPropertyRow.SetValue(RefUniversal.Properties.Value, Extra.Category, FieldType.Variant);
                                ItemPropertyRow = PropertiesSubSectionRows.AddNew();
                                ItemPropertyRow.SetGuid(RefUniversal.Properties.Property, RegisterTypeProperties.FindRow("@Name='Версия'").Id);
                                ItemPropertyRow.SetString(RefUniversal.Properties.DisplayValue, Extra.Version);
                                ItemPropertyRow.SetValue(RefUniversal.Properties.Value, Extra.Version, FieldType.Variant);
                                ItemPropertyRow = PropertiesSubSectionRows.AddNew();
                                ItemPropertyRow.SetGuid(RefUniversal.Properties.Property, RegisterTypeProperties.FindRow("@Name='Статус'").Id);
                                ItemPropertyRow.SetString(RefUniversal.Properties.DisplayValue, ((DocumentState)Extra.Status).GetDescription());
                                ItemPropertyRow.SetValue(RefUniversal.Properties.Value, (Int32)(DocumentState)Extra.Status, FieldType.Variant);
                                ItemPropertyRow = PropertiesSubSectionRows.AddNew();
                                ItemPropertyRow.SetGuid(RefUniversal.Properties.Property, RegisterTypeProperties.FindRow("@Name='Подразделение регистратора'").Id);
                                ItemPropertyRow.SetString(RefUniversal.Properties.DisplayValue, RegistrarUnit.GetString(CardOrd.Properties.DisplayValue));
                                ItemPropertyRow.SetValue(RefUniversal.Properties.Value, RegistrarUnit.GetString(CardOrd.Properties.Value), FieldType.Variant);
                                ItemPropertyRow = PropertiesSubSectionRows.AddNew();
                                ItemPropertyRow.SetGuid(RefUniversal.Properties.Property, RegisterTypeProperties.FindRow("@Name='Прибор'").Id);
                                ItemPropertyRow.SetString(RefUniversal.Properties.DisplayValue, Device.GetString(CardOrd.Properties.DisplayValue));
                                ItemPropertyRow.SetValue(RefUniversal.Properties.Value, Device.ChildSections[CardOrd.SelectedValues.ID].FirstRow.GetString(CardOrd.SelectedValues.SelectedValue), FieldType.Variant);
                                ItemPropertyRow = PropertiesSubSectionRows.AddNew();
                                ItemPropertyRow.SetGuid(RefUniversal.Properties.Property, RegisterTypeProperties.FindRow("@Name='Ссылка на карточку документа'").Id);
                                ItemPropertyRow.SetString(RefUniversal.Properties.DisplayValue, Extra.Card.Description);
                                ItemPropertyRow.SetValue(RefUniversal.Properties.Value, Extra.Card.Id, FieldType.Variant);
                                ItemPropertyRow = PropertiesSubSectionRows.AddNew();
                                ItemPropertyRow.SetGuid(RefUniversal.Properties.Property, RegisterTypeProperties.FindRow("@Name='Файлы документа'").Id);
                                ItemPropertyRow.SetString(RefUniversal.Properties.DisplayValue, Extra.Path);
                                ItemPropertyRow.SetValue(RefUniversal.Properties.Value, Extra.Card.Id, FieldType.Variant);

                                Extra.RegID = RegItem.Id;
                                logger.Info("В реестр добавлена запись: " + Card.Description);
                            }
                        }
                        else
                            FileSynch = true;

                        if (FileSynch)
                        {
                            logger.Info("Синхронизация пути...");
                            String OldPath = Extra.Path;
                            logger.Info("Старый путь: " + OldPath);
                            String NewPath = String.Empty;
                            logger.Info("Новый путь: " + NewPath);
                            if (Directory.Exists(OldPath))
                            {
                                DirectoryInfo OldFolder = new DirectoryInfo(OldPath);
                                DirectoryInfo ParentFolder = OldFolder.Parent;
                                if (ParentFolder != null)
                                {
                                    logger.Info("Старый путь:" + OldPath);
                                    NewPath = MyHelper.RemoveInvalidFileNameChars(Digest);
                                    if (NewPath != Digest)
                                        logger.Info("Удалены недопустимые символы: " + Digest);
                                    NewPath = Path.Combine(ParentFolder.FullName, NewPath);
                                    String TempPath = Path.Combine(ParentFolder.FullName, Path.GetRandomFileName());
                                    logger.Info("Временный путь:" + TempPath);
                                    logger.Info("Новый путь:" + NewPath + @"\");
                                    try
                                    {
                                        logger.Info("Перемещение во временную папку");
                                        Directory.Move(OldPath, TempPath);
                                        try
                                        {
                                            logger.Info("Перемещение в новую папку");
                                            Directory.Move(TempPath, NewPath);
                                        }
                                        catch (Exception Ex)
                                        {
                                            logger.WarnException("Перемещение не удалось!", Ex);
                                            try
                                            {
                                                logger.Info("Перемещение в старую папку");
                                                Directory.Move(TempPath, OldPath);
                                            }
                                            catch (Exception SubEx)
                                            {
                                                logger.Warn("Перемещение не удалось!");
                                                logger.ErrorException("Перемещение не удалось!", SubEx);
                                                NewPath = String.Empty;
                                            }
                                        }
                                    }
                                    catch (Exception Ex)
                                    {
                                        logger.WarnException("Перемещение не удалось!", Ex);
                                        NewPath = String.Empty;
                                    }

                                    if (!String.IsNullOrWhiteSpace(NewPath))
                                    {
                                        if (!(Extra is ExtraCardMarketingFiles))
                                            Extra.Path = NewPath + @"\";
                                        logger.Info("Синхронизация пути выполнена");
                                    }
                                    else
                                        logger.Warn("Синхронизация пути не выполнена");
                                }
                                else
                                    logger.Warn("Отсутствует родительская папка!");
                            }
                            else
                                logger.Info("Отсутствует путь для синхронизации!");

                        }
                        return true;
                    default:
                        return false;
                }
            }
        }

        public static Boolean AssignRights (UserSession Session, String CardID)
        {
            logger.Info("cardId='{0}'", CardID);
            using (new Impersonator(ServerExtension.Domain, ServerExtension.User, ServerExtension.SecurePassword))
            {
                Guid CardId = new Guid(CardID);
                switch (Session.CardManager.GetCardState(CardId))
                {
                    case ObjectState.Existing:
                        CardData Card = Session.CardManager.GetCardData(CardId);

                        List<TargetObject> TObjects = AssignHelper.LoadMatrix(Settings.MatrixPath, Settings.CardSheetName, Settings.Domain);
                        Card.UnlockCard();
                        logger.Debug("Card unlocked.'");
                        ExtraCard Extra = Card.GetExtraCard();
                        if (Extra.IsNull())
                            Extra = new ExtraCard(Card);
                        logger.Debug("ExtraCard received.'");

                        if (!Extra.IsNull())
                        {
                            String CategoryName = Extra.CategoryName;
                            logger.Info("Категория = '{0}'", CategoryName);

                            if (!String.IsNullOrEmpty(CategoryName))
                            {
                                TargetObject TObject = TObjects.FirstOrDefault(to => to.Name == CategoryName);
                                if (!TObject.IsNull())
                                {
                                    CardDataSecurity Rights = Card.GetAccessControl();
                                    Rights.SetOwner(new NTAccount(Settings.Domain, "DV Admins"));
                                    Rights = Rights.RemoveExplicitRights();


                                    foreach (AssignGroup Group in TObject.Groups)
                                        try { Rights.SetAssignGroup(TObject, Group); }
                                        catch (Exception ex) { logger.ErrorException(String.Format("Группа '{0}'", Group.Name), ex); }

                                    AssignGroup[] FolderGroups = TObject.Groups.Where(g => g.DirectoryRights != 0).ToArray();
                                    if(FolderGroups.Length > 0)
                                    {
                                        String FolderPath = Extra.Path;
                                        if (!String.IsNullOrEmpty(FolderPath))
                                        {
                                            DirectoryInfo Folder = new DirectoryInfo(FolderPath);
                                            if (Folder.Exists)
                                                foreach (AssignGroup Group in FolderGroups)
                                                    try { Folder.SetAssignGroup(TObject, Group); }
                                                    catch (Exception ex) { logger.ErrorException(String.Format("Группа '{0}'", Group.Name), ex); }

                                            else
                                                logger.Warn(String.Format("Карточка '{0}' ссылается на несуществующую папку'{1}'", Card.Id, Folder.FullName));
                                        }
                                    }

                                    Card.SetAccessControl(Rights);
                                    logger.Info("Права назначены!");
                                    return true;
                                }
                                logger.Warn("Категория '{0}' не найдена в матрице.", CategoryName);
                            }
                        }

                        logger.Info("Права не назначены.");
                        return false;
                    default:
                        return false;
                }
            }
        }

        public static Boolean CutRights (UserSession Session, String CardID)
        {
            logger.Info("cardId='{0}'", CardID);
            using (new Impersonator(ServerExtension.Domain, ServerExtension.User, ServerExtension.SecurePassword))
            {
                Guid CardId = new Guid(CardID);
                switch (Session.CardManager.GetCardState(CardId))
                {
                    case ObjectState.Existing:
                        CardData Card = Session.CardManager.GetCardData(CardId);
                        Card.UnlockCard();
                        Card.SetAccessControl(Card.GetAccessControl().Restrict());
                        return true;
                    default:
                        return false;
                }
            }
        }

        public static Boolean RegisterProtocol(UserSession Session, String CardID, String TempFolder, Guid EmployeeId)
        {
            logger.Info("cardId='{0}'", CardID);
            logger.Info("tempName='{0}'", TempFolder);
            logger.Info("EmployeeId='{0}'", EmployeeId);
            using (new Impersonator(ServerExtension.Domain, ServerExtension.User, ServerExtension.SecurePassword))
            {
                Guid CardId = new Guid(CardID);
                switch (Session.CardManager.GetCardState(CardId))
                {
                    case ObjectState.Existing:
                        DirectoryInfo TempDirectory = new DirectoryInfo(Path.Combine(ArchiveTempPath, TempFolder));
                        FileInfo[] TempFiles = TempDirectory.GetFiles();
                        logger.Info("В папке файлов: " + TempFiles.Length);
                        if (TempFiles.Length > 0)
                        {
                            logger.Info("Файлы: " + TempFiles.Select(file => file.Name).Aggregate((a, b) => a + "; " + b));
                            CardData Card = Session.CardManager.GetCardData(CardId);
                            Card.UnlockCard();
                            if (Card.InUpdate)
                                Card.CancelUpdate();
                            Card.PlaceLock();
                            Card.BeginUpdate();
                            foreach (Protocol Protocol in TempFiles.Select(fi => new Protocol(fi)))
                            {
                                if (Protocol.IsParsed)
                                {
                                    RowData MainInfoRow = Card.Sections[CardOrd.MainInfo.ID].FirstRow;
                                    Guid FilesID = MainInfoRow.GetGuid(CardOrd.MainInfo.FilesID) ?? Guid.Empty;
                                    Boolean FileListExist = !(FilesID.IsEmpty() && Card.Session.CardManager.GetCardState(FilesID) == ObjectState.Existing);
                                    logger.Info("FileListExist = " + FileListExist);
                                    CardData FileListCard;
                                    if (FileListExist)
                                    {
                                        FileListCard = Card.Session.CardManager.GetCardData(FilesID);
                                    }
                                    else
                                    {
                                        FileListCard = Card.Session.CardManager.CreateCardData(FileList.ID);
                                        MainInfoRow.SetGuid(CardOrd.MainInfo.FilesID, FileListCard.Id);
                                    }

                                    SectionData FileReferencesSection = FileListCard.Sections[FileList.FileReferences.ID];
                                    /* Проверка существования файла протокола в карточке */
                                    if (!FileReferencesSection.Rows.Any(file => Protocol.PhysicalFile.Name.Contains(file.GetString(CardFile.MainInfo.FileName))))
                                    {
                                        FileListCard.UnlockCard();

                                        VersionedFileCard FileCard = (VersionedFileCard)Card.Session.CardManager.CreateCard(DocsVision.Platform.Cards.Constants.VersionedFileCard.ID);
                                        FileVersion FileCardVersion = FileCard.Initialize(Protocol.PhysicalFile.FullName, Guid.Empty, false, true);
                                        CardData FileData = Card.Session.CardManager.CreateCardData(CardFile.ID);

                                        FileData.BeginUpdate();
                                        FileData.Description = "Файл: " + FileCard.Name;

                                        RowData FileMainInfoRow = FileData.Sections[CardFile.MainInfo.ID].Rows.AddNew();
                                        FileMainInfoRow.SetGuid(CardFile.MainInfo.FileID, FileCard.Id);
                                        FileMainInfoRow.SetString(CardFile.MainInfo.FileName, FileCardVersion.Name);
                                        FileMainInfoRow.SetGuid(CardFile.MainInfo.Author, FileCardVersion.AuthorId);
                                        FileMainInfoRow.SetInt32(CardFile.MainInfo.FileSize, FileCardVersion.Size);
                                        FileMainInfoRow.SetInt32(CardFile.MainInfo.VersioningType, 0);

                                        RowData FilePropertiesRow = FileData.Sections[CardFile.Properties.ID].Rows.AddNew();
                                        FilePropertiesRow.SetString(CardFile.Properties.Name, "Дата начала испытаний");
                                        FilePropertiesRow.SetInt32(CardFile.Properties.ParamType, (Int32)PropertieParamType.Date);
                                        FilePropertiesRow.SetString(CardFile.Properties.Value, Protocol.StringDate);
                                        FilePropertiesRow.SetString(CardFile.Properties.DisplayValue, Protocol.StringDate);

                                        FileData.Sections[CardFile.Categories.ID].Rows.AddNew().SetGuid(CardFile.Categories.CategoryID, MyHelper.RefCategory_CalibrationProtocol);

                                        FileData.EndUpdate();

                                        Int32 FilesCount = FileReferencesSection.Rows.Count;
                                        FileReferencesSection.Rows.AddNew().SetGuid(FileList.FileReferences.CardFileID, FileData.Id);
                                        FileListCard.Sections[FileList.MainInfo.ID].FirstRow.SetInt32(FileList.MainInfo.Count, FilesCount + 1);

                                        SectionData PropertiesSection = Card.Sections[CardOrd.Properties.ID];
                                        RowData PropertyRow = PropertiesSection.GetProperty("Дата");
                                        if (!PropertyRow.IsNull())
                                        {
                                            RowDataCollection SelectedValuesRows = PropertyRow.ChildSections[CardOrd.SelectedValues.ID].Rows;
                                            RowData SelectedValuesRow = SelectedValuesRows.AddNew();
                                            SelectedValuesRow.SetInt32(CardOrd.SelectedValues.Order, SelectedValuesRows.Count);
                                            SelectedValuesRow.SetDateTime(CardOrd.SelectedValues.SelectedValue, DateTime.Now);
                                        }
                                        PropertyRow = PropertiesSection.GetProperty("Действие");
                                        if (!PropertyRow.IsNull())
                                        {
                                            RowDataCollection SelectedValuesRows = PropertyRow.ChildSections[CardOrd.SelectedValues.ID].Rows;
                                            RowData SelectedValuesRow = SelectedValuesRows.AddNew();
                                            SelectedValuesRow.SetInt32(CardOrd.SelectedValues.Order, SelectedValuesRows.Count);
                                            SelectedValuesRow.SetString(CardOrd.SelectedValues.SelectedValue, "Прикреплен протокол калибровки");
                                        }
                                        PropertyRow = PropertiesSection.GetProperty("Участник");
                                        if (!PropertyRow.IsNull())
                                        {
                                            RowDataCollection SelectedValuesRows = PropertyRow.ChildSections[CardOrd.SelectedValues.ID].Rows;
                                            RowData SelectedValuesRow = SelectedValuesRows.AddNew();
                                            SelectedValuesRow.SetInt32(CardOrd.SelectedValues.Order, SelectedValuesRows.Count);
                                            SelectedValuesRow.SetGuid(CardOrd.SelectedValues.SelectedValue, EmployeeId);
                                        }
                                        PropertyRow = PropertiesSection.GetProperty("Комментарий");
                                        if (!PropertyRow.IsNull())
                                        {
                                            RowDataCollection SelectedValuesRows = PropertyRow.ChildSections[CardOrd.SelectedValues.ID].Rows;
                                            RowData SelectedValuesRow = SelectedValuesRows.AddNew();
                                            SelectedValuesRow.SetInt32(CardOrd.SelectedValues.Order, SelectedValuesRows.Count);
                                            SelectedValuesRow.SetString(CardOrd.SelectedValues.SelectedValue, "Автоматическое прикрепление протокола калибровки " + Protocol.PhysicalFile.Name);
                                        }
                                        PropertyRow = PropertiesSection.GetProperty("Ссылки");
                                        if (!PropertyRow.IsNull())
                                        {
                                            RowDataCollection SelectedValuesRows = PropertyRow.ChildSections[CardOrd.SelectedValues.ID].Rows;
                                            RowData SelectedValuesRow = SelectedValuesRows.AddNew();
                                            SelectedValuesRow.SetInt32(CardOrd.SelectedValues.Order, SelectedValuesRows.Count);
                                            SelectedValuesRow.SetString(CardOrd.SelectedValues.SelectedValue, null);
                                        }

                                        if (!FileListExist)
                                            Card.Sections[CardOrd.MainInfo.ID].FirstRow.SetGuid(CardOrd.MainInfo.FilesID, FileListCard.Id);
                                    }
                                }
                                else
                                    logger.Warn("Нераспознаный файл: " + Protocol.PhysicalFile.Name);
                            }

                            Card.EndUpdate();
                            Card.RemoveLock();

                            TempDirectory.Delete(true);
                            logger.Info("RegisterProtocol - выполнено.");
                        }
                        else
                            logger.Info("RegisterProtocol - не выполнено.");
                        return true;
                    default:
                        logger.Info("RegisterProtocol - не выполнено. Паспорт прибора не существует.");
                        return false;
                }
            }
        }

        public static Boolean DeleteCard(UserSession Session, String CardID)
        {
            logger.Info("cardId='{0}'", CardID);
            using (new Impersonator(ServerExtension.Domain, ServerExtension.User, ServerExtension.SecurePassword))
            {
                FolderCard FolderCard = (FolderCard)Session.CardManager.GetDictionary(FoldersCard.ID);
                CardData UniversalCard = Session.CardManager.GetDictionaryData(RefUniversal.ID);

                CardData Card = Session.CardManager.GetCardData(new Guid(CardID));

                Card.UnlockCard();
                ExtraCard Extra = Card.GetExtraCard();

                Guid? RegItemId  = null;
                String DirectoryPath = null;

                if (Extra != null)
                {
                    RegItemId = Extra.RegID;
                    DirectoryPath = Extra.Path;
                    logger.Info("Запись в реестре='{0}'", RegItemId.HasValue ? RegItemId.Value.ToString() : "NULL");
                    logger.Info("Путь в архиве='{0}'", DirectoryPath);

                    /* Удаление карточки */
                    FolderCard.DeleteCard(Card.Id, true);
                    logger.Info("Карточка удалена!");

                    /* Удаление записи в реестре */
                    if (RegItemId.HasValue)
                    {
                        RowData ItemRow = UniversalCard.GetItemRow(RegItemId.Value);
                        if (ItemRow != null)
                        {
                            UniversalCard.Sections[RefUniversal.Item.ID].DeleteRow(ItemRow.Id);
                            logger.Info("Запись в реестре тех.док. удалена");
                        }
                        else
                            logger.Info("Запись в реестре тех.док. не найдена");
                    }

                    /* Удаление папки в архиве */
                    if (!String.IsNullOrEmpty(DirectoryPath))
                    {
                        var di = new DirectoryInfo(DirectoryPath);
                        if (di.Exists)
                        {
                            if (di.FullName.Length > 20)
                            {
                                di.Delete(true);
                                logger.Info("Папка удалена!");
                            }
                            else
                                logger.Warn("В карточке '{0}' некорректный путь '{1}'", Card.Id, DirectoryPath);
                        }
                        else
                            logger.Info("Папка не существует!");
                    }

                    return true;
                }
                else
                {
                    logger.Warn("Не верный тип карточки!");
                    return false;
                }
            }
        }

        public static void DeleteDvCard(UserSession Session, String CardID, Boolean Permanent)
        {
            logger.Info("cardId='{0}'", CardID);
            CardData Card = Session.CardManager.GetCardData(new Guid(CardID));
            Card.UnlockCard();
            Session.CardManager.DeleteCard(Card.Id, Permanent);                     
        }

        public static String GetExecutionInfo ()
        {
            logger.Info("GetExecutionInfo");

            Process process = Process.GetCurrentProcess();
            AppDomain domain = AppDomain.CurrentDomain;
            Assembly assembly = Assembly.GetExecutingAssembly();
            WindowsIdentity identity = WindowsIdentity.GetCurrent();

            Int32 id = process.Id;
            String machineName = process.MachineName;
            String processName = process.ProcessName;
            String startDate = process.StartTime.ToShortDateString();
            String startTime = process.StartTime.ToShortTimeString();
            String arguments = process.StartInfo.Arguments;
            String fileName = process.StartInfo.FileName;
            String directory = process.StartInfo.WorkingDirectory;

            String friendlyName = domain.FriendlyName;
            String baseDirectory = domain.BaseDirectory;
            String dynamicDirectory = domain.DynamicDirectory;
            String relativeSearchPath = domain.RelativeSearchPath;

            String codeBase = assembly.CodeBase;
            String fullName = assembly.FullName;
            String location = assembly.Location;

            StringBuilder executionInfo = new StringBuilder(10000);

            if (identity != null)
            {
                String name = identity.Name;
                String authenticationType = identity.AuthenticationType;

                executionInfo.AppendLine("User information:");
                executionInfo.AppendFormat("{0}name: {2}{1}{0}authenticationType: {3}{1}", "\t", "\n", name, authenticationType);
                executionInfo.AppendLine();
            }

            executionInfo.AppendLine("Process information:");
            executionInfo.AppendFormat("\tid: {0}\n", id);
            executionInfo.AppendFormat("\tmachineName: {0}\n", machineName);
            executionInfo.AppendFormat("\tprocessName: {0}\n", processName);
            executionInfo.AppendFormat("\tstartDateTime: {0} {1}\n", startDate, startTime);
            executionInfo.AppendFormat("\targuments: {0}\n", arguments);
            executionInfo.AppendFormat("\tfileName: {0}\n", fileName);
            executionInfo.AppendFormat("\tdirectory: {0}\n", directory);
            executionInfo.AppendFormat("\tworkDirectory: {0}\n", Directory.GetCurrentDirectory());
            executionInfo.AppendLine();

            executionInfo.AppendLine("AppDomain information:");
            executionInfo.AppendFormat("\tfriendlyName: {0}\n", friendlyName);
            executionInfo.AppendFormat("\tbaseDirectory: {0}\n", baseDirectory);
            executionInfo.AppendFormat("\tdynamicDirectory: {0}\n", dynamicDirectory);
            executionInfo.AppendFormat("\trelativeSearchPath: {0}\n", relativeSearchPath);
            executionInfo.AppendLine("Loaded assemblies:");
            foreach (var da in domain.GetAssemblies())
            {
                executionInfo.AppendLine("\t" + da.FullName);
                executionInfo.AppendLine("\t" + da.Location);
            }
            executionInfo.AppendLine();

            executionInfo.AppendLine("Assembly information:");
            executionInfo.AppendFormat("\tcodeBase: {0}\n", codeBase);
            executionInfo.AppendFormat("\tfullName: {0}\n", fullName);
            executionInfo.AppendFormat("\tlocation: {0}\n", location);

            logger.Info("Return ExecutionInfo");
            return executionInfo.ToString();
        }

        public static Boolean SendNoticeOfExecution (SessionParameters SessionParams, UserSession Session, String Subject, String Body, String CardID, String FolderId)
        {
            logger.Info(String.Format("SendNoticeOfExecution: '{0}' on '{1}'", SessionParams.UserAccount, SessionParams.UserComputer));

            var folderCard = (FolderCard)Session.CardManager.GetDictionary(FoldersCard.ID);
            var folder = folderCard.GetFolder(new Guid(FolderId));
            CardData tt = Session.CardManager.GetCardData(new Guid("8F26CB44-BCDC-4040-AE74-3D3F6F6E6425"));
            var pt = tt.Copy();
            pt.IsTemplate = false;
            folder.Shortcuts.AddNew(pt.Id, true);

            var rc = new CardResolution(pt);

            rc.Info.Name = Subject;
            rc.Info.Comments = Body;
            rc.Info.RegisteredBy = SessionParams.EmployeeId;

            var perf = rc.Performers.CreatePerformer();

            var pg = new Guid("A6640FFB-8583-4F73-BBE4-E2FF7F30E119");
            perf.EmployeeId = pg;
            perf.PerformerId = pg;
            perf.RoutingType = PerformerRoutingType.Default;
            perf.Order = 1;

            foreach (var row in pt.Sections[DocsVision.TakeOffice.Cards.Constants.CardResolution.Performers.ID].Rows)
                row.SetInt32("PerformerType", 0);

            var cg = new Guid(CardID);
            Reference rfr = rc.References.CreateReference(cg);
            rfr.Comment = "Ссылка на запрос";
            rfr.RefCardId = cg;
            rfr.RefType = ReferenceRefType.DVDocument;

            rc.Info.ResolutionState = MainInfoResolutionState.ToPerform;

            logger.Info(String.Format("SendNoticeOfExecution - выполнено"));
            return true;
        }

        /// <summary>
        /// Освобождает заблокированную карточку.
        /// </summary>
        /// <param name="Card"></param>
        internal static void UnlockCard (this CardData Card)
        {
            if (Card.LockStatus != LockStatus.Free)
                Card.ForceUnlock();
        }
        /// <summary>
        /// Возвращает надстройку для карточки.
        /// </summary>
        /// <param name="Card">Данные карточки.</param>
        internal static ExtraCard GetExtraCard (this CardData Card)
        {
            ExtraCard Extra = ExtraCardCD.GetExtraCard(Card);
            if (Extra == null)
                Extra = ExtraCardTD.GetExtraCard(Card);
            if (Extra == null)
                Extra = ExtraCardEA.GetExtraCard(Card);
            if(Extra == null)
                Extra = ExtraCardMarketingFiles.GetExtraCard(Card);
            if (Extra == null)
                Extra = ExtraCardDoc.GetExtraCard(Card);
            return Extra;
        }
        /// <summary>
        /// Функция для преобразования пароля для Impersonator.
        /// </summary>
        /// <param name="securePassword"></param>
        /// <returns></returns>
        internal static String ConvertToUnsecureString (SecureString securePassword)
        {
            if (securePassword == null)
                throw new ArgumentNullException("securePassword");

            IntPtr unmanagedString = IntPtr.Zero;
            try
            {
                unmanagedString = Marshal.SecureStringToGlobalAllocUnicode(securePassword);
                return Marshal.PtrToStringUni(unmanagedString);
            }
            finally
            {
                Marshal.ZeroFreeGlobalAllocUnicode(unmanagedString);
            }
        }
    }
}
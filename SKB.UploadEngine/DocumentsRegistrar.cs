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
using DocsVision.Platform.ObjectManager.ViewModel;
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
    public static class DocumentsRegistrar
    {
        #region Fields

        private static readonly NLog.Logger logger;
        public static readonly String ArchivePath;
        public static readonly String ArchiveTempPath;
        public static readonly String ArchiveDeletePath;

        #endregion

        static DocumentsRegistrar()
        {
            logger = NLog.LogManager.GetCurrentClassLogger();
            String settingsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "settings.xml");
            Settings.Load(settingsPath);

            ArchivePath = String.Format(@"\\{0}\{1}", Settings.ServerName, Settings.ArchiveName);
            ArchiveTempPath = String.Format(@"\\{0}\{1}", Settings.ServerName, Settings.ArchiveTempName);
            ArchiveDeletePath = String.Format(@"\\{0}\{1}", Settings.ServerName, Settings.ArchiveDeleteName);
        }
        public static string LoadDocuments(string LoadFolderPath, string ArchiveFolderPath, string PassportFolderID, string TemplateCardID, UserSession Session)
        {
            string TextResult = "";
            CardData TemplateCard = Session.CardManager.GetCardData(new Guid(TemplateCardID));
            Folder PassportFolder = ((FolderCard)Session.CardManager.GetDictionary(FoldersCard.ID)).GetFolder(new Guid(PassportFolderID));
            CardData refUniversal = Session.CardManager.GetDictionaryData(RefUniversal.ID);

            List<RawView> passportsRawView = Session.CardManager.GetViewData(ViewSource.FromFolder(new Guid(PassportFolderID)))
                .Select(ir => new RawView
                {
                    Description = ir.GetString(DigestViewColumns.Description),
                    InstanceId = ir.GetGuid(DigestViewColumns.InstanceId)
                }).ToList();


            IEnumerable<Protocol> allProtocols = GetProtocols(LoadFolderPath);
            logger.Info("Найдено файлов:" + allProtocols.Count());

            IEnumerable<Protocol> parsedProtocols = allProtocols.Where(p => p.IsParsed).ToList();
            IEnumerable<Protocol> incorrectProtocols = allProtocols.Where(p => !p.IsParsed).ToList();

            string incorrectProtocolsList = "";
            foreach (Protocol pr in incorrectProtocols)
            {
                incorrectProtocolsList += "\n" + pr.PhysicalFile.FullName;
            }
            logger.Info("Не распознаны следующие файлы: {0}", incorrectProtocolsList);
            TextResult = TextResult + "Не распознаны следующие файлы: {0}" + incorrectProtocolsList + "\n";

            foreach (Protocol pp in parsedProtocols)
            {
                try
                {
                    string RegisterProtocolResult = "";
                    if (RegisterProtocol(pp, passportsRawView, Session, TemplateCard, PassportFolder, refUniversal, out RegisterProtocolResult))
                    { File.Move(pp.PhysicalFile.FullName, pp.PhysicalFile.FullName.Replace(LoadFolderPath, ArchiveFolderPath)); }
                    TextResult = TextResult + RegisterProtocolResult;
                }
                catch (Exception ex)
                {
                    logger.ErrorException(string.Format("Register error: {0}", pp.PhysicalFile.FullName), ex);
                    TextResult = TextResult + "Ошибка регистрации:\n" + pp.PhysicalFile.FullName + ". " + ex.ToString() + "\n";
                }
            }

            foreach (var ic in incorrectProtocols)
            {
                logger.Info("Файл не соответсвует формату:\n{0}", ic.PhysicalFile.FullName);
                TextResult = TextResult + "Файл не соответсвует формату:\n" + ic.PhysicalFile.FullName + "\n";
            }
            return TextResult;
        }
        internal static IEnumerable<Protocol> GetProtocols(string FolderPath)
        {
            DirectoryInfo protocolDir = new DirectoryInfo(FolderPath);
            return protocolDir.GetFiles("??-*-*-*-*.*", SearchOption.AllDirectories).Select(file => new Protocol(file));
        }
        internal static bool RegisterProtocol(Protocol pd, List<RawView> passportsRawView, UserSession Session, CardData TemplateCard, Folder PassportFolder, CardData refUniversal, out string TextResult)
        {
            TextResult = "";
            logger.Debug("Search passport...");

            var rw = passportsRawView.Find(ir => Regex.IsMatch(ir.Description,
                pd.RegexPatternUnit, RegexOptions.IgnoreCase));

            if (rw != null)
            {
                logger.Debug("Паспорт найден для: {0}", pd.PhysicalFile.Name);
                TextResult = TextResult + "Паспорт найден для: " + pd.PhysicalFile.Name + "\n";
                var value = rw.InstanceId.Value;
                CardData card = Session.CardManager.GetCardData(value);
                card.AttachDocumentToCard(Session, pd);
                logger.Info("Загружен {0}", pd.PhysicalFile.Name);
                TextResult = TextResult + "Загружен " + pd.PhysicalFile.Name + "\n";
                return true;
            }
            else
            {
                logger.Info("Паспорт не найден для '{0}'", pd.PhysicalFile.Name);
                TextResult = TextResult + "Паспорт не найден для " + pd.PhysicalFile.Name + "\n";
                if (SearchParty(Session, pd, refUniversal) != null)
                {
                    //CreateUnitCard(Session, pd, passportsRawView, TemplateCard, PassportFolder, refUniversal);
                    //logger.Info("Загружен {0}", pd.PhysicalFile.Name);
                    //TextResult = TextResult + "Загружен " + pd.PhysicalFile.Name + "\n";
                    //return true;
                    return false;
                }
                else
                {
                    logger.Info("Партия не найдена для прибора: {0}", pd.PhysicalFile.Name);
                    TextResult = TextResult + "Партия не найдена для прибора: " + pd.PhysicalFile.Name + "\n";
                    return false;
                }
            }
        }
        private static void CreateUnitCard(UserSession Session, Protocol protocol, List<RawView> passportsRawView, CardData TemplateCard, Folder PassportFolder, CardData refUniversal)
        {
            var unitCard = TemplateCard.Copy();
            unitCard.IsTemplate = false;
            PassportFolder.Shortcuts.AddNew(unitCard.Id, true);

            var sdp = unitCard.Sections[CardOrd.Properties.ID];
            var rdm = unitCard.Sections[CardOrd.MainInfo.ID].FirstRow;

            #region Filling card fields

            sdp.BeginUpdate();
            sdp.SetRowDoubleValue("Заводской номер прибора",
                protocol.Number.StartsWith("0") ? protocol.Number.Substring(1) : protocol.Number);
            sdp.SetRowDoubleValue("/Год прибора", protocol.Year);
            sdp.SetRowDoubleValue("Год партии", protocol.Year);
            sdp.SetRowEnumValue("Состояние прибора", "В эксплуатации");

            // поиск партии по прибору
            string partyNumber = string.Empty;
            var foundParty = SearchParty(Session, protocol, refUniversal);

            if (foundParty != null)
            {
                RowData rdfp = foundParty.GetRowData();
                partyNumber = rdfp.GetString("Name");
                sdp.SetRowValue("№ партии", "DisplayValue", partyNumber);
                sdp.SetRowValue("№ партии", "Value", rdfp.Id, fieldType: FieldType.RefId);

                foreach (var row in rdfp.ChildSections[0].Rows)
                {
                    if (row.GetString("Name") == "Наименование прибора")
                    {
                        var unitName = row.GetString("DisplayValue");
                        sdp.SetRowValue("Прибор", "DisplayValue", unitName);
                        sdp.SetRowValue("Прибор", "Value", row.GetGuid("Value"), fieldType: FieldType.RefId);

                        // поиск прибора в справочнике
                        var si = refUniversal.Sections[RefUniversal.ItemType.ID].FindRow("@Name = 'Приборы и комплектующие'").ChildSections[RefUniversal.Item.ID].FindRow("@Name = '" + unitName + "'");
                        //var si = SearchItemInReference(refUniversal,RefUniversal.Item.ID, "Name", unitName);
                        if (si != null)
                        {
                            var code = si.ChildSections[0].FindRow("@Name = 'Код СКБ'");
                            sdp.SetRowValue("Код СКБ", "DisplayValue", code.GetString("DisplayValue"));
                            sdp.SetRowValue("Код СКБ", "Value", code.GetGuid("Value"), fieldType: FieldType.RefId);
                        }
                        else
                        {
                            logger.Info("Прибор не найден в справочнике: {0}", protocol.PhysicalFile.Name);
                        }
                    }

                    if (row.GetString("Name") == "Карточка плана")
                    {
                        var guid = row.GetGuid("Value");
                        if (guid.HasValue)
                        {
                            var rdcr = unitCard.Sections[CardOrd.CardReferences.ID].Rows.AddNew();
                            rdcr.SetGuid("Link", guid.Value);
                            rdcr.SetGuid("Type", new Guid("508F52C7-ACE1-48F0-B978-B55FC0649776"));

                            var cdp = Session.CardManager.GetCardData(guid.Value);
                            if (cdp.LockStatus == LockStatus.Free)
                            {
                                rdcr = cdp.Sections[CardOrd.CardReferences.ID].Rows.AddNew();
                                rdcr.SetGuid("Link", unitCard.Id);
                                rdcr.SetGuid("Type", new Guid("C8DE0604-03CD-4806-B0E1-304A92A6E277"));
                            }
                            else
                                logger.Warn("Card of plan is locked, ref not added {0}", protocol.PhysicalFile.Name);
                        }
                    }

                    if (row.GetString("Name") == "Комментарий")
                        sdp.SetRowDoubleValue("Комментарии к партии", row.GetString("Value"));
                }
            }
            else
            {
                logger.Warn("Партия не найдена для прибора: {0}", protocol.PhysicalFile.Name);
            }
            sdp.EndUpdate();

            rdm.BeginUpdate();

            RowData CurrentUser = SearchItemInReference(Session, Session.CardManager
                .GetDictionaryData(DocsVision.Platform.Cards.Constants.RefStaff.ID), DocsVision.Platform.Cards.Constants.RefStaff.Employees.ID, "AccountName",
                (string)Session.Properties["AccountName"].Value);

            var name = string.Format("{0} № {1}/{2} из партии {3}",
                protocol.DVUnitName, protocol.Number, protocol.Year, partyNumber);
            rdm.SetString("Name", name);
            rdm.SetString("Digest", name);
            rdm.SetGuid("RegisteredBy", CurrentUser.Id);
            rdm.EndUpdate();

            #endregion

            unitCard.Description = name;
            passportsRawView.Add(new RawView { Description = unitCard.Description, InstanceId = unitCard.Id });

            sdp.SetRowValue("Запись в справочнике", "Value", unitCard.AddItemToRegestry(refUniversal), fieldType: FieldType.RefId);

            sdp.SetTableProperty("Дата", DateTime.Now, FieldType.DateTime);
            sdp.SetTableProperty("Действие", "Паспорт создан на основании протокола калибровки");
            sdp.SetTableProperty("Участник", CurrentUser.Id, FieldType.UniqueId);
            sdp.SetTableProperty("Комментарий", "Паспорт создан автоматически в процессе работы утилиты \"ProtocolsRegistrar\"");
            sdp.SetTableProperty("Ссылки", null);

            unitCard.AttachDocumentToCard(Session, protocol, false);

            var em = Session.ExtensionManager.GetExtensionMethod("UploadExtension", "AssignRights");
            em.Parameters.AddNew("cardId", 0).Value = unitCard.Id.ToString();
            em.Execute();
        }
        private static void AttachDocumentToCard(this CardData self, UserSession Session,
            Protocol document, bool fileListExist = true)
        {
            var rdm = self.Sections[CardOrd.MainInfo.ID].FirstRow;
            // ReSharper disable PossibleInvalidOperationException
            var flc = Session.CardManager.GetCardData(rdm.GetGuid("FilesID").Value);
            // ReSharper restore PossibleInvalidOperationException
            var fileList = flc.Sections[flc.Type.AllSections["FileReferences"].Id];

            // Проверка существования файла протокола в карточке
            if (fileList.Rows.Any(file => document.PhysicalFile.Name.Contains(file.GetString("FileName"))))
                return;

            RowData si = SearchItemInReference(Session, Session.CardManager
                .GetDictionaryData(DocsVision.Platform.Cards.Constants.RefStaff.ID), DocsVision.Platform.Cards.Constants.RefStaff.Employees.ID, "AccountName",
                (string)Session.Properties["AccountName"].Value);

            var sdp = self.Sections[CardOrd.Properties.ID];
            sdp.SetTableProperty("Дата", DateTime.Now, FieldType.DateTime);
            sdp.SetTableProperty("Действие", "Прикреплен протокол калибровки");
            sdp.SetTableProperty("Участник", si.Id, FieldType.UniqueId);
            sdp.SetTableProperty("Комментарий", "Автоматическое прикрепление протокола калибровки " + document.PhysicalFile.Name);
            sdp.SetTableProperty("Ссылки", null);

            self.AttachFileToCard(Session, document, fileListExist);
        }
        private static void AttachFileToCard(this CardData self, UserSession Session,
            Protocol protocol, bool fileListExist)
        {
            var rdm = self.Sections[CardOrd.MainInfo.ID].FirstRow;
            var fileCard = (VersionedFileCard)Session.CardManager
                .CreateCard(DocsVision.Platform.Cards.Constants.VersionedFileCard.ID);

            fileCard.Initialize(protocol.PhysicalFile.FullName, Guid.Empty, false, true);

            var fileData = Session.CardManager.CreateCardData(CardFile.ID);
            var msr = fileData.Sections[fileData.Type.AllSections["MainInfo"].Id].Rows.AddNew();
            var ps = fileData.Sections[fileData.Type.AllSections["Properties"].Id];
            var cs = fileData.Sections[fileData.Type.AllSections["Categories"].Id];

            msr.BeginUpdate();
            msr.SetGuid("FileID", fileCard.Id);
            msr.SetString("FileName", fileCard.CurrentVersion.Name);
            msr.SetGuid("Author", fileCard.CurrentVersion.AuthorId);
            msr.SetInt32("FileSize", fileCard.CurrentVersion.Size);
            msr.SetInt32("VersioningType", 0);
            msr.EndUpdate();

            fileData.Description = "Файл: " + fileCard.Name;

            var psr = ps.Rows.AddNew();

            psr.BeginUpdate();
            psr.SetString("Name", "Дата начала испытаний");
            psr.SetInt32("ParamType", 17);
            psr.SetDateTime("Value", protocol.Date.Date);
            psr.SetString("DisplayValue", protocol.StringDate);
            psr.EndUpdate();

            var csr = cs.Rows.AddNew();
            csr.SetGuid("CategoryID", new Guid(protocol.DocumentTypeID));

            if (!fileListExist)
            {
                var fileList = Session.CardManager.CreateCardData(FileList.ID);
                var newRef = fileList.Sections[fileList.Type.AllSections["FileReferences"].Id].CreateRow();

                rdm.SetGuid("FilesID", fileList.Id);
                newRef.SetGuid("CardFileID", fileData.Id);
                fileList.Sections[fileList.Type.AllSections["MainInfo"].Id].FirstRow["Count"] = 1;
            }
            else
            {
                var guid = rdm.GetGuid("FilesID");
                if (guid.HasValue)
                {
                    var flc = Session.CardManager.GetCardData(guid.Value);
                    var newFileRow = flc.Sections[flc.Type.AllSections["FileReferences"].Id].Rows.AddNew();
                    newFileRow.SetGuid("CardFileID", fileData.Id);
                }
            }
        }
        private static UnitParty SearchParty(UserSession Session, Protocol protocol, CardData refUniversal)
        {
            var rdi = SearchItemInReference(Session, refUniversal, RefUniversal.ItemType.ID, "Name", "Справочник партий");
            if (rdi != null)
            {
                var rdc = SearchItemsInReference(Session, refUniversal, RefUniversal.Item.ID,
                    "ParentRowID", rdi.Id.ToString(), FieldType.UniqueId);

                // партии для прибора по году и отсортированные по месяцу
                var ups = rdc.Where(rd => Regex.IsMatch(rd.GetString("Name"),
                    string.Format("{0}.*{1}", protocol.DVUnitName, protocol.Year), RegexOptions.IgnoreCase))
                    .Select(rd => new UnitParty(rd)).OrderBy(up => up.Month).ToList();

                return ups.LastOrDefault();
            }

            return null;
        }

        private static Guid AddItemToRegestry(this CardData self, CardData refUniversal)
        {
            var sdp = self.Sections[CardOrd.Properties.ID];

            RowData rtd = refUniversal.Sections[RefUniversal.ItemType.ID]
                .FindRow("@Name='Справочник готовых приборов'");

            SubSectionData items = rtd.ChildSections[refUniversal.Type.AllSections["Item"].Id];
            SubSectionData propertyItem = rtd.ChildSections[refUniversal.Type.AllSections["TypeProperties"].Id];

            // Свойства записи
            RowData partyProperty = propertyItem.FindRow("@Name='Партия'");
            RowData statusProperty = propertyItem.FindRow("@Name='Статус'");
            RowData unitYearProperty = propertyItem.FindRow("@Name='Год прибора'");
            RowData unitNumberProperty = propertyItem.FindRow("@Name='Номер прибора'");
            RowData unitPassportProperty = propertyItem.FindRow("@Name='Паспорт прибора'");
            RowData warrantyTalonProperty = propertyItem.FindRow("@Name='Гарантийный талон'");
            RowData unitNameProperty = propertyItem.FindRow("@Name='Наименование прибора'");

            // Свойства карточки
            RowData unit = sdp.FindRow("@Name='Прибор'");
            RowData party = sdp.FindRow("@Name='№ партии'");
            RowData unitYear = sdp.FindRow(string.Format("@Name = '{0}'", "/Год прибора"));
            RowData unitNumber = sdp.FindRow("@Name='Заводской номер прибора'");

            // Создаем запись в справочнике готовых приборов
            RowData newItem = items.Rows.AddNew();
            newItem.SetString("Name", self.Description);
            SubSectionData subrow = newItem.ChildSections[refUniversal.Type.AllSections["Properties"].Id];

            // Прописываем свойства
            AddItem(subrow.Rows.AddNew(),
                unitNumberProperty.GetString("RowID"), unitNumber.GetString("DisplayValue"), unitNumber.GetString("Value"));
            AddItem(subrow.Rows.AddNew(),
                unitYearProperty.GetString("RowID"), unitYear.GetString("DisplayValue"), unitYear.GetString("Value"));
            AddItem(subrow.Rows.AddNew(),
                partyProperty.GetString("RowID"), party.GetString("DisplayValue"), party.GetString("Value"));
            AddItem(subrow.Rows.AddNew(),
                statusProperty.GetString("RowID"), "В эксплуатации", "6");
            AddItem(subrow.Rows.AddNew(),
                unitNameProperty.GetString("RowID"), unit.GetString("DisplayValue"), unit.GetString("Value"));
            AddItem(subrow.Rows.AddNew(),
                    unitPassportProperty.GetString("RowID"), self.Description, self.Id.ToString());
            AddItem(subrow.Rows.AddNew(),
                unitNumber.GetString("RowID"), self.Description, self.Id.ToString());
            AddItem(subrow.Rows.AddNew(),
                warrantyTalonProperty.GetString("RowID"), "", "");
            return newItem.Id;
        }

        private static void AddItem(RowData newSubRow, params string[] list)
        {
            newSubRow.BeginUpdate();
            newSubRow["Property"] = list[0];
            newSubRow["DisplayValue"] = list[1];
            newSubRow["Value"] = list[2];
            newSubRow.EndUpdate();
        }

        private static void SetTableProperty(this SectionData self,
            string fieldName, object value, FieldType fieldType = FieldType.String)
        {
            self.BeginUpdate();
            var rd = self.FindRow(string.Format("@Name = '{0}'", fieldName));
            var ssd = rd.ChildSections[CardOrd.SelectedValues.ID];
            var nrd = ssd.Rows.AddNew();
            nrd.SetInt32("Order", ssd.Rows.Count);
            nrd.SetValue("SelectedValue", value, fieldType);
            self.EndUpdate();
        }

        private static RowData SearchItemInReference(UserSession Session, CardData reference, Guid sectionId,
            string fieldName, string fieldValue, FieldType fieldType = FieldType.Unistring,
            ConditionOperation operation = ConditionOperation.Equals)
        {
            RowDataCollection rdc = reference.Sections[sectionId]
                .FindRows(GetSectionSearchQuery(Session, fieldName, fieldValue, fieldType, operation));
            return rdc.Count != 0 ? rdc[0] : null;
        }

        private static IEnumerable<RowData> SearchItemsInReference(UserSession Session, CardData reference,
            Guid sectionId, string fieldName, string fieldValue,
            FieldType fieldType = FieldType.Unistring,
            ConditionOperation operation = ConditionOperation.Equals)
        {
            RowDataCollection rdc = reference.Sections[sectionId]
                .FindRows(GetSectionSearchQuery(Session, fieldName, fieldValue, fieldType, operation));
            return rdc;
        }

        private static string GetSectionSearchQuery(UserSession Session, string fieldName, string fieldValue,
            FieldType fieldType = FieldType.Unistring,
            ConditionOperation conditionOperation = ConditionOperation.Equals)
        {
            SectionQuery query = Session.CreateSectionQuery();
            query.ConditionGroup.Conditions.AddNew(fieldName, fieldType, conditionOperation, fieldValue);
            return query.GetXml(null, true);
        }

        private static void SetRowValue(this SectionData self, string fieldValue,
            string alias, object value, string fieldName = "Name",
            FieldType fieldType = FieldType.Unistring)
        {
            RowData rd = self.FindRow(string.Format("@{0}='{1}'", fieldName, fieldValue));
            rd.SetValue(alias, value, fieldType);
        }

        private static void SetRowDoubleValue(this SectionData self, string fieldValue,
            object value, string fieldName = "Name",
            FieldType fieldType = FieldType.Unistring)
        {
            RowData rd = self.FindRow(string.Format("@{0}='{1}'", fieldName, fieldValue));
            rd.SetValue("Value", value, fieldType);
            rd.SetString("DisplayValue", value);
            rd.SetString("TextValue", value);
        }

        private static void SetRowRefValue(this SectionData self, UserSession Session, string fieldValue,
            string value, CardData reference, Guid sectionId,
            string searchFieldName = "Name", FieldType fieldType = FieldType.Unistring,
            ConditionOperation operation = ConditionOperation.Equals)
        {
            if (!String.IsNullOrEmpty(value) && !RegexEngine.IsMatch(value, @"^\s+$"))
            {
                RowData si = SearchItemInReference(Session, reference, sectionId,
                    searchFieldName, value, fieldType, operation);
                if (si != null)
                {
                    self.SetRowValue(fieldValue, "Value", si.Id, fieldType: FieldType.RefId);
                    self.SetRowValue(fieldValue, "DisplayValue", si.DisplayString);
                }
                else
                {
                    if (sectionId == DocsVision.Platform.Cards.Constants.RefStaff.Employees.ID)
                    {
                        RowData unit = SearchItemInReference(Session, reference, DocsVision.Platform.Cards.Constants.RefStaff.Units.ID, "Name", "Временные");
                        RowData ne = unit.ChildSections[DocsVision.Platform.Cards.Constants.RefStaff.Employees.ID].Rows.AddNew();
                        ne.SetString("LastName", value);
                        ne.SetString("DisplayString", value);

                        self.SetRowValue(fieldValue, "Value", ne.Id, fieldType: FieldType.RefId);
                        self.SetRowValue(fieldValue, "DisplayValue", ne.DisplayString);
                    }
                    else
                        logger.Warn("cardId='{0}'; {2}='{1}' value not found", self.Card.Id, value, fieldValue);
                }
            }
        }

        private static void SetRowEnumValue(this SectionData self,
            string fieldValue, string value)
        {
            if (!String.IsNullOrEmpty(value) && !RegexEngine.IsMatch(value, @"^\s+$"))
            {
                RowData rd = self.FindRow(string.Format("@Name='{0}'", fieldValue));
                SubSectionData ssd = rd.ChildSections[self.Card.Type.AllSections["EnumValues"].Id];
                RowData srd = ssd.GetAllRows().ToList().Find(r => RegexEngine.IsMatch(r.GetString("ValueName"), value));
                srd = srd ?? ssd.GetAllRows().ToList().Find(r => RegexEngine.IsMatch(value, r.GetString("ValueName")));
                if (srd != null)
                {
                    rd.SetInt32("Value", srd.GetInt32("ValueID"));
                    rd.SetString("DisplayValue", srd.GetString("ValueName"));
                }
                else
                    logger.Warn("cardId='{0}'; {2}='{1}' value not found", self.Card.Id, value, fieldValue);
            }
        }
    }
}
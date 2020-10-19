using System;
using System.Security;
using System.Security.Cryptography;
using System.Text;
using DocsVision.Platform.ObjectManager;
using DocsVision.Platform.StorageServer.Extensibility;
using Microsoft.Win32;
using DocsVision.BackOffice.ObjectModel.Services;
using SKB.Base;
using DocsVision.Platform.ObjectModel;
using DocsVision.BackOffice.ObjectModel;
using SKB.Base.Task;
using DocsVision.Platform.ObjectManager.SystemCards;
using System.IO;

namespace SKB.UploadExtension
{
    /// <summary>
    /// Представлет набор методов серверного расширения.
    /// </summary>
    public sealed class ServerExtension : DocsVision.Platform.StorageServer.Extensibility.StorageServerExtension
    {
        public static String Server { get; private set; }
        public static String Domain { get; private set; }
        public static String User { get; private set; }
        public static String Password { get; private set; }
        public static SecureString SecurePassword
        {
            get
            {
                SecureString s = new SecureString();
                foreach (Char c in Password)
                    s.AppendChar(c);
                return s;
            }
        }
        private static readonly NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();

        static ServerExtension ()
        {
            /*Получение данных подключения*/
            RegistryKey
                Key = Registry.LocalMachine.OpenSubKey("SOFTWARE\\DocsVision\\BackOffice\\5.0\\Server\\Extension", false);
            if (Key != null)
            {
                String Account = ((String)Key.GetValue("userName", (Object)String.Empty)).Trim();

                if (!String.IsNullOrEmpty(Account))
                {
                    Domain = Account.Split('\\')[0];
                    User = Account.Split('\\')[1];
                }

                String PasswordData = (String)Key.GetValue("password", (Object)String.Empty);

                if (!String.IsNullOrEmpty(PasswordData))
                {
                    if (PasswordData.StartsWith("DVMP", StringComparison.OrdinalIgnoreCase))
                    {
                        Byte[] Data = Convert.FromBase64String(PasswordData.Substring("DVMP".Length));
                        Password = Encoding.Unicode.GetString(ProtectedData.Unprotect(Data, null, DataProtectionScope.LocalMachine));
                    }
                }
            }

            Key = Registry.LocalMachine.OpenSubKey("SOFTWARE\\DocsVision\\Platform\\5.0\\Site", false);

            if (Key != null)
                Server = ((String)Key.GetValue("ConnectAddress", (Object)String.Empty)).Trim();
        }

        /// <summary>
        /// Проверка на дублирование карточки.
        /// </summary>
        /// <param name="checkName">Имя для проверки.</param>
        /// <returns></returns>
        [ExtensionMethod]
        public String CheckDuplication(String checkName)
        {
            UserSession Session = null;
            try
            {
                logger.Info(String.Format("CheckDuplication: '{0}' на '{1}'", Request.Session.UserAccount, Request.Session.UserComputer));
                Session = OpenUserSession();
                return Core.CheckDuplication(Session, checkName).ToString();
            }
            catch (Exception Ex)
            {
                logger.Warn("CheckDuplication: " + Ex.Message);
                logger.ErrorException("CheckDuplication", Ex);
                return Boolean.FalseString;
            }
            finally { CloseUserSession(Session); }
        }
        /// <summary>
        /// Разместить файлы в архиве.
        /// </summary>
        /// <param name="tempName">Имя временной папки.</param>
        /// <param name="code">Код документа.</param>
        /// <param name="type">Код категории документа.</param>
        /// <param name="label">Метка документа.</param>
        /// <param name="version">Версия документа (цифра).</param>
        /// <param name="name">Имя документа.</param>
        /// <param name="number">Номер документа (цифра).</param>
        /// <param name="unitName">Имя прибора.</param>
        /// <returns>Полный путь на размещенную директорию.</returns>
        [ExtensionMethod]
        public String PlaceFiles(String tempName, String label, String code, String type, String version, String name, String number, String unitName)
        {

            try
            {
                logger.Info(String.Format("PlaceFiles: '{0}' на '{1}'", Request.Session.UserAccount, Request.Session.UserComputer));
                return Core.PlaceFiles(tempName, label, code, type, version, name, number, unitName);
            }
            catch (Exception Ex)
            {
                logger.Warn("PlaceFiles: " + Ex.Message);
                logger.ErrorException("PlaceFiles", Ex);
                logger.Error("tempName: {0}; label: {1}; code: {2}; type: {3}; version: {4}; name: {5}; number: {6}; unitName: {7};"
                    , tempName, label, code, type, version, name, number, unitName);
                return Boolean.FalseString;
            }
        }
        /// <summary>
        /// Разместить файлы в архиве.
        /// </summary>
        /// <param name="tempName">Имя временной папки.</param>
        /// <param name="newFolder">Новая папка.</param>
        /// <returns>Полный путь на размещенную директорию.</returns>
        [ExtensionMethod]
        public String PlaceArchiveFiles(String tempName, String newFolder)
        {

            try
            {
                logger.Info(String.Format("PlaceFiles: '{0}' на '{1}'", Request.Session.UserAccount, Request.Session.UserComputer));
                return Core.PlaceFiles(tempName, newFolder);
            }
            catch (Exception Ex)
            {
                logger.Warn("PlaceFiles: " + Ex.Message);
                logger.ErrorException("PlaceFiles", Ex);
                logger.Error("tempName: {0}; newFolder: {1}", tempName, newFolder);
                return Boolean.FalseString;
            }
        }
        /// <summary>
        /// Добавляет файлы в архив.
        /// </summary>
        /// <param name="tempName">Имя временной папки.</param>
        /// <param name="docDirPath">Путь на папку с документами.</param>
        /// <returns>Полный путь на размещенную директорию.</returns>
        [ExtensionMethod]
        public String AddFiles(String tempName, String docDirPath)
        {
            try
            {
                logger.Info(String.Format("AddFiles: '{0}' на '{1}'", Request.Session.UserAccount, Request.Session.UserComputer));
                return Core.AddFiles(tempName, docDirPath);
            }
            catch (Exception ex)
            {
                logger.Warn("AddFiles: " + ex.Message);
                logger.ErrorException("AddFiles", ex);
                logger.Error("docDirPath: " + docDirPath);
                return Boolean.FalseString;
            }
        }
        /// <summary>
        /// Синхронизировать карточку с реестром документации и папкой в архиве.
        /// </summary>
        /// <param name="cardId">
        /// Идентификатор карточки.
        /// </param>
        /// <returns>True, если успех.</returns>
        [ExtensionMethod]
        public String Synchronize(String cardId)
        {
            UserSession Session = null;
            try
            {
                logger.Info(String.Format("Synchronize: '{0}' на '{1}'", Request.Session.UserAccount, Request.Session.UserComputer));
                Session = OpenUserSession();
                return Core.Synchronize(Session, cardId).ToString();
            }
            catch (Exception Ex)
            {
                logger.Warn("Synchronize: " + Ex.Message);
                logger.ErrorException("Synchronize", Ex);
                logger.Error("cardId: " + cardId);
                return Boolean.FalseString;
            }
            finally { CloseUserSession(Session); }
        }
        /// <summary>
        /// Назначает права на карточку и файлы.
        /// </summary>
        /// <param name="cardId">
        /// Идентификатор карточки.
        /// </param>
        /// <returns>True, если успех.</returns>
        [ExtensionMethod]
        public String AssignRights(String cardId)
        {
            logger.Info(String.Format("AssignRights: '{0}' на '{1}'", Request.Session.UserAccount, Request.Session.UserComputer));
            UserSession Session = null;
            try
            {
                Session = OpenUserSession();
                return Core.AssignRights(Session, cardId).ToString();
            }
            catch (Exception Ex)
            {
                logger.Warn("AssignRights: " + Ex.Message);
                logger.ErrorException("AssignRights", Ex);
                logger.Error("cardId: " + cardId);
                return Boolean.FalseString;
            }
            finally { CloseUserSession(Session); }
        }
        /// <summary>
        /// Удаляет права на удаление у карточки.
        /// </summary>
        /// <param name="cardId">
        /// Идентификатор карточки.
        /// </param>
        /// <returns></returns>
        [ExtensionMethod]
        public String CutRights(String cardId)
        {
            UserSession Session = null;
            try
            {
                logger.Info(String.Format("CutRights: '{0}' на '{1}'", Request.Session.UserAccount, Request.Session.UserComputer));
                Session = OpenUserSession();
                return Core.CutRights(Session, cardId).ToString();
            }
            catch (Exception ex)
            {
                logger.Warn("CutRights: " + ex.Message);
                logger.ErrorException("CutRights", ex);
                logger.Error("cardId: " + cardId);
                return Boolean.FalseString;
            }
            finally { CloseUserSession(Session); }
        }
        /// <summary>
        /// Удаляет карточку, включая запись в реестре и папку в архиве.
        /// </summary>
        /// <param name="cardId">
        /// Идентификатор карточки.
        /// </param>
        /// <returns>True, если успех.</returns>
        [ExtensionMethod]
        public String DeleteCard(String cardId)
        {
            UserSession Session = null;
            try
            {
                logger.Info(String.Format("DeleteCard: '{0}' на '{1}'", Request.Session.UserAccount, Request.Session.UserComputer));
                Session = OpenUserSession();
                return Core.DeleteCard(Session, cardId).ToString();
            }
            catch (Exception Ex)
            {
                logger.Warn("DeleteCard: " + Ex.Message);
                logger.ErrorException("DeleteCard", Ex);
                logger.Error("cardId: " + cardId);
                return Boolean.FalseString;
            }
            finally { CloseUserSession(Session); }
        }
        /// <summary>
        /// Удаляет карточку.
        /// </summary>
        /// <param name="cardId">
        /// Идентификатор карточки.
        /// </param>
        /// <param name="permanent">
        /// Полное удаление.
        /// </param>
        /// <returns>True, если успех.</returns>
        [ExtensionMethod]
        public void DeleteDvCard(string cardId, bool permanent)
        {
            UserSession Session = null;
            try
            {
                logger.Info(String.Format("DeleteDvCard: '{0}' на '{1}'", Request.Session.UserAccount, Request.Session.UserComputer));
                Session = OpenUserSession();
                Core.DeleteDvCard(Session, cardId, permanent);
            }
            catch (Exception Ex)
            {
                logger.Warn("DeleteDvCard: " + Ex.Message);
                logger.ErrorException("DeleteDvCard", Ex);
                logger.Error("cardId: " + cardId);
            }
            finally { CloseUserSession(Session); }            
        }
        /// <summary>
        /// Прикрепляет протокол калиброки к карточке паспорта прибора.
        /// </summary>
        /// <param name="cardId"></param>
        /// <param name="tempName"></param>
        /// <returns></returns>
        [ExtensionMethod]
        public String RegisterProtocol(String cardId, String tempName)
        {
            UserSession Session = null;
            try
            {
                logger.Info(String.Format("RegisterProtocol: '{0}' на '{1}'", Request.Session.UserAccount, Request.Session.UserComputer));
                Session = OpenUserSession();
                return Core.RegisterProtocol(Session, cardId, tempName, Request.Session.EmployeeId).ToString();
            }
            catch (Exception Ex)
            {
                logger.Warn("RegisterProtocol: " + Ex.Message);
                logger.ErrorException("RegisterProtocol", Ex);
                logger.Error("cardId: {0}; tempName: {1}", cardId, tempName);
                return Boolean.FalseString;
            }
            finally { CloseUserSession(Session); }
        }
        /// <summary>
        /// Загружает из указанной папки документы калибровки.
        /// </summary>
        /// <param name="cardId"></param>
        /// <param name="tempName"></param>
        /// <returns></returns>
        [ExtensionMethod]
        public string LoadCalibrationDocuments(String LoadFolderPath, String ArchiveFolderPath, String PassportFolderID, String TemplateCardID)
        {
            UserSession Session = null;
            try
            {
                logger.Info(String.Format("LoadCalibrationDocuments: '{0}' на '{1}'", Request.Session.UserAccount, Request.Session.UserComputer));
                Session = OpenUserSession();
                return DocumentsRegistrar.LoadDocuments(LoadFolderPath, ArchiveFolderPath, PassportFolderID, TemplateCardID, Session);
            }
            catch (Exception Ex)
            {
                logger.Warn("LoadCalibrationDocuments: " + Ex.Message);
                logger.ErrorException("LoadCalibrationDocuments", Ex);
                return Ex.Message;
            }
            finally { CloseUserSession(Session); }
        }
        /// <summary>
        /// Отправляет уведомление Германюк о завершении запроса.
        /// </summary>
        /// <param name="subject"></param>
        /// <param name="body"></param>
        /// <param name="cardId"></param>
        /// <param name="folderId"></param>
        /// <returns></returns>
        [ExtensionMethod]
        public String SendNoticeOfExecution(String subject, String body, String cardId, String folderId)
        {
            try
            {
                return Core.SendNoticeOfExecution(Request.Session, CreateSession(true), subject, body, cardId, folderId).ToString();
            }
            catch (Exception Ex)
            {
                logger.Warn("RegisterProtocol: " + Ex.Message);
                logger.ErrorException("SendNoticeOfExecution", Ex);
                logger.Error("subject: {0}; body: {1}; cardId: {2}; folderId: {3}", subject, body, cardId, folderId);
                return Boolean.FalseString;
            }
        }
        /// <summary>
        /// Возвращает информацию о среде выполнения.
        /// </summary>
        [ExtensionMethod]
        public String GetExecutionInfo()
        {
            try { return Core.GetExecutionInfo(); }
            catch (Exception ex)
            {
                logger.Warn("GetExecutionInfo: " + ex.Message);
                logger.ErrorException("GetExecutionInfo", ex);
                return Boolean.FalseString;
            }
        }
        /// <summary>
        /// Возвращает путь к папке архива.
        /// </summary>
        [ExtensionMethod]
        public String GetArchivePath ()
        {
            try
            {
                logger.Info(String.Format("GetArchivePath: '{0}' on '{1}'", Request.Session.UserAccount, Request.Session.UserComputer));
                return Core.ArchivePath;
            }
            catch (Exception Ex)
            {
                logger.Warn("GetArchivePath: " + Ex.Message);
                logger.ErrorException("GetArchivePath", Ex);
                return String.Empty;
            }
        }
        /// <summary>
        /// Возвращает путь к папке архива для временных файлов.
        /// </summary>
        [ExtensionMethod]
        public String GetArchiveTempPath ()
        {
            try
            {
                logger.Info(String.Format("GetArchiveTempPath: '{0}' on '{1}'", Request.Session.UserAccount, Request.Session.UserComputer));
                return Core.ArchiveTempPath;
            }
            catch (Exception Ex)
            {
                logger.Warn("GetArchiveTempPath: " + Ex.Message);
                logger.ErrorException("GetArchiveTempPath", Ex);
                return String.Empty;
            }
        }
        /// <summary>
        /// Возвращает путь к папке архива для подготовки к удалению.
        /// </summary>
        [ExtensionMethod]
        public String GetArchiveDeletePath ()
        {
            try
            {
                logger.Info(String.Format("GetArchiveDeletePath: '{0}' on '{1}'", Request.Session.UserAccount, Request.Session.UserComputer));
                return Core.ArchiveDeletePath;
            }
            catch (Exception Ex)
            {
                logger.Warn("GetArchiveDeletePath: " + Ex.Message);
                logger.ErrorException("GetArchiveDeletePath", Ex);
                return String.Empty;
            }
        }
        /// <summary>
        /// Возвращает путь к матрице прав доступа.
        /// </summary>
        [ExtensionMethod]
        public String GetMatrixPath ()
        {
            try
            {
                logger.Info(String.Format("GetMatrixPath: '{0}' on '{1}'", Request.Session.UserAccount, Request.Session.UserComputer));
                return Settings.MatrixPath;
            }
            catch (Exception Ex)
            {
                logger.Warn("GetMatrixPath: " + Ex.Message);
                logger.ErrorException("GetMatrixPath", Ex);
                return String.Empty;
            }
        }
        /// <summary>
        /// Запускает задание
        /// </summary>
        [ExtensionMethod]
        public void StartTask(String taskId)
        {
            UserSession Session = null;
            
            try
            {
                logger.Info(String.Format("StartTask: '{0}' on '{1}'", Request.Session.UserAccount, Request.Session.UserComputer));
                Guid taskID = new Guid(taskId);
                Session = OpenUserSession();
                ObjectContext context = Session.CreateContext();
                ITaskService taskService = context.GetService<ITaskService>();
                Task task = context.GetObject<Task>(taskID);

                if (task != null)
                {
                    taskService.StartTask(task);
                    MyHelper.ChangeState(context, Session.CardManager.GetCardData(taskID), TaskStateNames.Started.GetDescription());
                }
            }
            catch (Exception Ex)
            {
                logger.Warn("StartTask: " + Ex.Message);
                logger.ErrorException("StartTask", Ex);
            }
            finally { CloseUserSession(Session); };
        }

        private UserSession OpenUserSession ()
        {
            SessionManager Manager = SessionManager.CreateInstance();
            Manager.Connect(Server, String.Empty, Domain + "\\" + User, Password);
            return Manager.CreateSession();
        }

        private void CloseUserSession (UserSession Session)
        {
            if (Session != null)
                Session.Close();
        }

        /// <summary>
        /// Выдает конкретный номер Внутреннему документу
        /// </summary>
        [ExtensionMethod]
        public Guid GetNumberID(string NumeratorName, string Number = "0")
        {
            logger.Info("GetNumberID start...");

            UserSession Session = null;
            try
            {
                logger.Info(String.Format("GetNumberID: '{0}' on '{1}'", Request.Session.UserAccount, Request.Session.UserComputer));
                Session = OpenUserSession();
                ObjectContext context = Session.CreateContext();
                logger.Info(String.Format("NumeratorName: '{0}'; Number: '{1}'.", NumeratorName, Number));
                Guid NewNumberID;
                if (Number == "0")
                    NewNumberID = MyHelper.GetNumberID(context, NumeratorName, MyHelper.GetNumber(context, NumeratorName));
                else
                    NewNumberID = MyHelper.GetNumberID(context, NumeratorName, Convert.ToInt32(Number));
                logger.Info(String.Format("NewNumberID: '{0}'.", NewNumberID));
                return NewNumberID;
            }
            catch (Exception Ex)
            {
                logger.Warn("GetNumberID: " + Ex.Message);
                logger.ErrorException("GetNumberID", Ex);
                return Guid.Empty;
            }
            finally { CloseUserSession(Session); };
        }
        /// <summary>
        /// Выдает номер Внутреннему документу по указанному нумератору
        /// </summary>
        [ExtensionMethod]
        public Int32 GetNumber(string NumeratorName)
        {
            logger.Info("GetNumber start...");

            UserSession Session = null;
            try
            {
                logger.Info(String.Format("GetNumber: '{0}' on '{1}'", Request.Session.UserAccount, Request.Session.UserComputer));
                Session = OpenUserSession();
                ObjectContext context = Session.CreateContext();
                logger.Info(String.Format("NumeratorName: '{0}'", NumeratorName));
                int NewNumberID = MyHelper.GetNumber(context, NumeratorName);
                logger.Info(String.Format("NewNumberID: '{0}'.", NewNumberID));
                return NewNumberID;
            }
            catch (Exception Ex)
            {
                logger.Warn("GetNumberID: " + Ex.Message);
                logger.ErrorException("GetNumberID", Ex);
                return 0;
            }
            finally { CloseUserSession(Session); };
        }
        /// <summary>
        /// Освобождает номер Внутреннего документа
        /// </summary>
        [ExtensionMethod]
        public bool ReleaseNumber(string NumeratorName, string NumberID)
        {
            logger.Info("ReleaseNumber start...");

            UserSession Session = null;
            try
            {
                Guid Number = new Guid(NumberID);
                logger.Info(String.Format("ReleaseNumber: '{0}' on '{1}'", Request.Session.UserAccount, Request.Session.UserComputer));
                Session = OpenUserSession();
                ObjectContext context = Session.CreateContext();
                logger.Info(String.Format("NumeratorName: '{0}'; Number: '{1}'.", NumeratorName, NumberID));

                CardData NumeratorCardData = null;
                NumeratorCardData = Session.CardManager.GetCardData(Number, DocsVision.Platform.Cards.Constants.NumeratorCard.BusyNumbers.ID);

                RowData Row = NumeratorCardData.Sections[DocsVision.Platform.Cards.Constants.NumeratorCard.BusyNumbers.ID].GetRow(Number);
                NumeratorCard NumeratorCard = (NumeratorCard)Session.CardManager.GetCard(NumeratorCardData.Id);
                NumeratorZone numeratorZone = NumeratorCard.Zones[Row.SubSection.ParentRow.Id];

                numeratorZone.ReleaseNumber(Number);
                return true;
            }
            catch (Exception Ex)
            {
                logger.Warn("ReleaseNumber: " + Ex.Message);
                logger.ErrorException("ReleaseNumber", Ex);
                return false;
            }
            finally { CloseUserSession(Session); };
        }
        /// <summary>
        /// Архивирует файл в файловом архиве
        /// </summary>
        [ExtensionMethod]
        public bool ArchivingFile(String FileCardId, String ArchivePath)
        {
            logger.Info("ArchivingFile start...");
            try
            {
                UserSession Session = OpenUserSession();
                VersionedFileCard FileCard = (VersionedFileCard)Session.CardManager.GetCard(new Guid(FileCardId));
                String FilePath = Path.Combine(ArchivePath, FileCard.CurrentVersion.Name);
                FileCard.CurrentVersion.Download(FilePath);
                logger.Info("Archived successfully! " + FilePath);
                return true;
            }
            catch (Exception Ex)
            {
                logger.Warn("Archived error! "  + Ex.Message);
                logger.ErrorException("Archived error! ", Ex);
                return false;
            }
        }
    }
}
using SKB.Base;
using System;
using System.IO;
using System.Globalization;
using System.Text.RegularExpressions;
using DocsVision.Platform.ObjectManager;

namespace SKB.UploadExtension
{
    /// <summary>
    /// Представляет документ протокола калибровки.
    /// </summary>
    class Protocol
    {
        #region Fields

        /// <summary>
        /// Шаблон для разбора имени документа протокола (ПК-01.05.11-18-11-ПКВМ7).
        /// </summary>
        private static readonly Regex RgxPattern =
            new Regex(@"(?<Type>\w{2})-(?<Date>\d{2}.\d{2}.\d{2})-(?<Number>.*)-(?<Year>\d{2})-(?<Unit>.*)");

        /// <summary>
        /// Строка соответствия литеры в заводском номере и года выпуска прибора.
        /// </summary>
        private static string MatchingString = "B = 2013; C = 2014; D = 2015; H = 2016; I = 2017; J = 2018; K = 2019; L = 2020; M = 2021; N = 2022; O = 2023; P = 2024; Q = 2025; R = 2026; S = 2027; T = 2028; U = 2029; V = 2030; W = 2031; X = 2032; Y = 2033; Z = 2034";

        /// <summary>
        /// Объект файла на ФС.
        /// </summary>
        private FileInfo physicalFile;

        /// <summary>
        /// Дата начала испытаний.
        /// </summary>
        private DateTime unitDate;

        /// <summary>
        /// Название прибора.
        /// </summary>
        private string dvUnitName;

        /// <summary>
        /// Шаблон для проверки прибора на существование.
        /// </summary>
        private string regexPatternUnit;

        /// <summary>
        /// Шаблон для проверки прибора на существование.
        /// </summary>
        private string regexPatternParty;

        #endregion

        #region Properties

        /// <summary>
        /// Дата начала испытаний (строка).
        /// </summary>
        public string StringDate { get; private set; }

        /// <summary>
        /// Дата начала испытаний.
        /// </summary>
        public DateTime Date
        {
            get
            {
                IFormatProvider culture = CultureInfo.CreateSpecificCulture("ru-RU");
                DateTimeStyles styles = DateTimeStyles.None;
                DateTime.TryParse(StringDate, culture, styles, out unitDate);
                return unitDate;
            }
        }

        /// <summary>
        /// Год выпуска прибора (11).
        /// </summary>
        public string ShortYear { get; private set; }

        /// <summary>
        /// Год выпуска прибора (2011).
        /// </summary>
        public string Year
        {
            get 
            {
                if (Convert.ToInt32(ShortYear) < 90)
                    return "20" + ShortYear;
                else
                    return "19" + ShortYear;
            }
        }

        /// <summary>
        /// Заводской номер прибора.
        /// </summary>
        public string Number { get; private set; }
        /// <summary>
        /// Тип документа (краткое обозначение).
        /// </summary>
        public string DocumentType { get; private set; }
        /// <summary>
        /// Тип документа.
        /// </summary>
        public string DocumentTypeID
        {
            get
            {
                string documentTypeID = "";
                switch (DocumentType)
                {
                    case "ДИ":
                        documentTypeID = "{7CD55E06-7BA9-467D-8A3A-89AEC914B5BF}";  // Категория "ДИ – Данные измерений"
                        break;
                    case "ПК":
                        documentTypeID = "{937151F3-A501-4DE0-991E-59594D73CBE2}";  // Категория "ПК - Протоколы приемосдаточных испытаний"
                        break;
                    case "ПР":
                        documentTypeID = "{3F290A81-39F4-4DEA-83B4-B26F1B569B73}";  // Категория "ПР – Протокол калибровки"
                        break;
                    case "СП":
                        documentTypeID = "{991867CF-8D3E-4A3F-B319-8AB8CDF63739}";  // Категория "СП – Свидетельство о поверке"
                        break;
                    case "ПВ":
                        documentTypeID = "{CD7A90AB-4BD5-4F15-B5F7-A695911E594F}";  // Категория "СП – Свидетельство о поверке"
                        break;
                }
                return documentTypeID;
            }
        }

        public int NumberDigit { get { return int.Parse(Number); } }

        /// <summary>
        /// Название прибора.
        /// </summary>
        public string UnitName { get; private set; }

        /// <summary>
        /// Название прибора в справочнике DV.
        /// </summary>
        public string DVUnitName
        {
            get
            {
                if (dvUnitName == null)
                {
                    switch (UnitName)
                    {
                        case "МИКО1":
                            dvUnitName = "МИКО-1";
                            break;
                        case "МИКО21":
                            dvUnitName = "МИКО-21";
                            break;
                        case "МИКО2.2":
                            dvUnitName = "МИКО-2.2";
                            break;
                        case "МИКО2.3":
                            dvUnitName = "МИКО-2.3";
                            break;
                        case "МИКО7":
                            dvUnitName = "МИКО-7";
                            break;
                        case "МИКО7М":
                            dvUnitName = "МИКО-7М";
                            break;
                        case "МИКО7МА":
                            dvUnitName = "МИКО-7МА";
                            break;
                        case "МИКО8":
                            dvUnitName = "МИКО-8";
                            break;
                        case "МИКО8М":
                            dvUnitName = "МИКО-8М";
                            break;
                        case "МИКО8МА":
                            dvUnitName = "МИКО-8МА";
                            break;
                        case "МИКО9":
                            dvUnitName = "МИКО-9";
                            break;
                        case "МИКО9А":
                            dvUnitName = "МИКО-9А";
                            break;
                        case "МИКО10":
                            dvUnitName = "МИКО-10";
                            break;
                        case "ПКВМ5":
                            dvUnitName = "ПКВ/М5";
                            break;
                        case "ПКВМ5А":
                            dvUnitName = "ПКВ/М5А";
                            break;
                        case "ПКВМ5Н":
                            dvUnitName = "ПКВ/М5Н";
                            break;
                        case "ПКВМ6":
                            dvUnitName = "ПКВ/М6";
                            break;
                        case "ПКВМ6Н":
                            dvUnitName = "ПКВ/М6Н";
                            break;
                        case "ПКВМ7":
                            dvUnitName = "ПКВ/М7";
                            break;
                        case "ПКВУ3.0":
                            dvUnitName = "ПКВ/У3.0";
                            break;
                        case "ПКВУ3.1":
                            dvUnitName = "ПКВ/У3.1";
                            break;
                        case "ПКВУ3.0-01":
                            dvUnitName = "ПКВ/У3.0-01";
                            break;
                        case "ПКВ35":
                            dvUnitName = "ПКВ-35";
                            break;
                        case "ПУВ35":
                            dvUnitName = "ПУВ-35";
                            break;
                        case "ПКР1":
                            dvUnitName = "ПКР-1";
                            break;
                        case "ПКР2":
                            dvUnitName = "ПКР-2";
                            break;
                        case "ПКР2М":
                            dvUnitName = "ПКР-2М";
                            break;
                        case "ПУВ10":
                            dvUnitName = "ПУВ-10";
                            break;
                        case "ПУВ50":
                            dvUnitName = "ПУВ-50";
                            break;
                        case "ПУВрегулятор":
                            dvUnitName = "ПУВ-регулятор";
                            break;
                        case "ТК021":
                            dvUnitName = "ТК-021";
                            break;
                        case "ТК026":
                            dvUnitName = "ТК-026";
                            break;
                        case "ПКВУ2":
                            dvUnitName = "ПКВ/У2";
                            break;
                        case "ПКВУ1":
                            dvUnitName = "ПКВ/У1";
                            break;
                        case "ПКВВ3":
                            dvUnitName = "ПКВ/В3";
                            break;
                        case "ПКВВ2":
                            dvUnitName = "ПКВ/В2";
                            break;
                        case "ПКВВ1":
                            dvUnitName = "ПКВ/В1";
                            break;
                        case "ПКВВ3А":
                            dvUnitName = "ПКВ/В3А";
                            break;
                        case "ПКВМ1":
                            dvUnitName = "ПКВ/М1";
                            break;
                        case "ПКВМ2":
                            dvUnitName = "ПКВ/М2";
                            break;
                        case "ПКВМ3":
                            dvUnitName = "ПКВ/М3";
                            break;
                        case "ПКВМ4":
                            dvUnitName = "ПКВ/М4";
                            break;
                        case "ПКВМ16":
                            dvUnitName = "ПКВ/М16";
                            break;
                        default:
                            dvUnitName = UnitName;
                            break;
                    }
                }
                return dvUnitName;
            }
        }

        /// <summary>
        /// Шаблон для проверки существования карточки прибора по дайджесту.
        /// </summary>
        public string RegexPatternUnit
        {
            get
            {
                // ПКВ/М7 № 19/2010 из партии ПКВ/М7 - 50 - 2/2011
                return string.IsNullOrEmpty(regexPatternUnit)
                           ? string.Format("{0} № {1}/{2}", DVUnitName, Number, Year)
                           : regexPatternUnit;
            }
        }

        /// <summary>
        /// Шаблон для нахождения записи по названию в справочнике партий.
        /// </summary>
        public string RegexPatternParty
        {
            get
            {
                // Мико-2.3 - 6 - 12/2010
                return string.IsNullOrEmpty(regexPatternParty)
                           ? string.Format("{0}.*{1}", DVUnitName, Year)
                           : regexPatternParty;
            }
        }

        /// <summary>
        /// Указывает успешно ли прошел разбор имени файла.
        /// </summary>
        public bool IsParsed { get; private set; }

        public FileInfo PhysicalFile
        {
            get { return physicalFile; }
        }

        #endregion

        #region Constructor

        /// <summary>
        /// Создает новый экземпляр протокола калибровки.
        /// </summary>
        /// <param name="fi"></param>
        public Protocol(FileInfo fi)
        {
            physicalFile = fi;
            regexPatternUnit = string.Empty;
            regexPatternParty = string.Empty;
            Parse();

            if (Number != null)
            {
                if (Regex.IsMatch(Number, @"\D"))
                {
                    if (Regex.IsMatch(Number, @"\d{3}\D\Z"))
                    {
                        if (MatchingString.IndexOf(Number.Substring(3) + " = " + Year) < 0) IsParsed = false;
                    }
                    else
                    { if ((DVUnitName != "ТК-021") && (DVUnitName != "ТК-026")) IsParsed = false; }
                }
                else
                { if (Number.Length > 3) IsParsed = false; }
            }
        }

        #endregion

        #region Methods

        private void Parse()
        {
            // ReSharper disable AssignNullToNotNullAttribute
            var mh = RgxPattern.Match(Path.GetFileNameWithoutExtension(physicalFile.Name));
            // ReSharper restore AssignNullToNotNullAttribute
            if (mh.Groups.Count == 6)
            {
                StringDate = mh.Groups["Date"].Value;
                ShortYear = mh.Groups["Year"].Value;
                Number = mh.Groups["Number"].Value;
                UnitName = mh.Groups["Unit"].Value;
                DocumentType = mh.Groups["Type"].Value;
                IsParsed = true;
            }
        }

        #endregion
    }
    public class RawView
    {
        public string Description { get; set; }
        public Guid? InstanceId { get; set; }
    }

    public static class RegexEngine
    {
        public static bool IsMatch(string input, string pattern)
        {
            return Regex.IsMatch(input, pattern, RegexOptions.IgnoreCase);
        }
    }

    public class UnitParty
    {
        private readonly RowData row;
        private const string partyName = @"(.*) - (.*) - (.*)";

        public string UnitName { get; set; }
        public int UnitCount { get; set; }
        public string Date { get; set; }
        public ushort Month { get; set; }

        public UnitParty(RowData row)
        {
            this.row = row;
            string name = row.GetString("Name");
            var mh = Regex.Match(name, partyName);

            if (mh.Groups.Count == 4)
            {
                ushort month;
                UnitName = mh.Groups[1].Value;
                UnitCount = int.Parse(mh.Groups[2].Value);
                Date = mh.Groups[3].Value;

                if (ushort.TryParse(this.Date.Split('/')[0], out month))
                    Month = month;
            }
        }

        public RowData GetRowData() { return this.row; }
    }
}
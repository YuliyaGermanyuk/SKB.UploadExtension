using SKB.Base;
using System;
using System.IO;
using System.Globalization;
using System.Text.RegularExpressions;
using DocsVision.Platform.ObjectManager;

namespace SKB.UploadExtension
{
    /// <summary>
    /// ������������ �������� ��������� ����������.
    /// </summary>
    class Protocol
    {
        #region Fields

        /// <summary>
        /// ������ ��� ������� ����� ��������� ��������� (��-01.05.11-18-11-����7).
        /// </summary>
        private static readonly Regex RgxPattern =
            new Regex(@"(?<Type>\w{2})-(?<Date>\d{2}.\d{2}.\d{2})-(?<Number>.*)-(?<Year>\d{2})-(?<Unit>.*)");

        /// <summary>
        /// ������ ������������ ������ � ��������� ������ � ���� ������� �������.
        /// </summary>
        private static string MatchingString = "B = 2013; C = 2014; D = 2015; H = 2016; I = 2017; J = 2018; K = 2019; L = 2020; M = 2021; N = 2022; O = 2023; P = 2024; Q = 2025; R = 2026; S = 2027; T = 2028; U = 2029; V = 2030; W = 2031; X = 2032; Y = 2033; Z = 2034";

        /// <summary>
        /// ������ ����� �� ��.
        /// </summary>
        private FileInfo physicalFile;

        /// <summary>
        /// ���� ������ ���������.
        /// </summary>
        private DateTime unitDate;

        /// <summary>
        /// �������� �������.
        /// </summary>
        private string dvUnitName;

        /// <summary>
        /// ������ ��� �������� ������� �� �������������.
        /// </summary>
        private string regexPatternUnit;

        /// <summary>
        /// ������ ��� �������� ������� �� �������������.
        /// </summary>
        private string regexPatternParty;

        #endregion

        #region Properties

        /// <summary>
        /// ���� ������ ��������� (������).
        /// </summary>
        public string StringDate { get; private set; }

        /// <summary>
        /// ���� ������ ���������.
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
        /// ��� ������� ������� (11).
        /// </summary>
        public string ShortYear { get; private set; }

        /// <summary>
        /// ��� ������� ������� (2011).
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
        /// ��������� ����� �������.
        /// </summary>
        public string Number { get; private set; }
        /// <summary>
        /// ��� ��������� (������� �����������).
        /// </summary>
        public string DocumentType { get; private set; }
        /// <summary>
        /// ��� ���������.
        /// </summary>
        public string DocumentTypeID
        {
            get
            {
                string documentTypeID = "";
                switch (DocumentType)
                {
                    case "��":
                        documentTypeID = "{7CD55E06-7BA9-467D-8A3A-89AEC914B5BF}";  // ��������� "�� � ������ ���������"
                        break;
                    case "��":
                        documentTypeID = "{937151F3-A501-4DE0-991E-59594D73CBE2}";  // ��������� "�� - ��������� ��������������� ���������"
                        break;
                    case "��":
                        documentTypeID = "{3F290A81-39F4-4DEA-83B4-B26F1B569B73}";  // ��������� "�� � �������� ����������"
                        break;
                    case "��":
                        documentTypeID = "{991867CF-8D3E-4A3F-B319-8AB8CDF63739}";  // ��������� "�� � ������������� � �������"
                        break;
                    case "��":
                        documentTypeID = "{CD7A90AB-4BD5-4F15-B5F7-A695911E594F}";  // ��������� "�� � ������������� � �������"
                        break;
                }
                return documentTypeID;
            }
        }

        public int NumberDigit { get { return int.Parse(Number); } }

        /// <summary>
        /// �������� �������.
        /// </summary>
        public string UnitName { get; private set; }

        /// <summary>
        /// �������� ������� � ����������� DV.
        /// </summary>
        public string DVUnitName
        {
            get
            {
                if (dvUnitName == null)
                {
                    switch (UnitName)
                    {
                        case "����1":
                            dvUnitName = "����-1";
                            break;
                        case "����21":
                            dvUnitName = "����-21";
                            break;
                        case "����2.2":
                            dvUnitName = "����-2.2";
                            break;
                        case "����2.3":
                            dvUnitName = "����-2.3";
                            break;
                        case "����7":
                            dvUnitName = "����-7";
                            break;
                        case "����7�":
                            dvUnitName = "����-7�";
                            break;
                        case "����7��":
                            dvUnitName = "����-7��";
                            break;
                        case "����8":
                            dvUnitName = "����-8";
                            break;
                        case "����8�":
                            dvUnitName = "����-8�";
                            break;
                        case "����8��":
                            dvUnitName = "����-8��";
                            break;
                        case "����9":
                            dvUnitName = "����-9";
                            break;
                        case "����9�":
                            dvUnitName = "����-9�";
                            break;
                        case "����10":
                            dvUnitName = "����-10";
                            break;
                        case "����5":
                            dvUnitName = "���/�5";
                            break;
                        case "����5�":
                            dvUnitName = "���/�5�";
                            break;
                        case "����5�":
                            dvUnitName = "���/�5�";
                            break;
                        case "����6":
                            dvUnitName = "���/�6";
                            break;
                        case "����6�":
                            dvUnitName = "���/�6�";
                            break;
                        case "����7":
                            dvUnitName = "���/�7";
                            break;
                        case "����3.0":
                            dvUnitName = "���/�3.0";
                            break;
                        case "����3.1":
                            dvUnitName = "���/�3.1";
                            break;
                        case "����3.0-01":
                            dvUnitName = "���/�3.0-01";
                            break;
                        case "���35":
                            dvUnitName = "���-35";
                            break;
                        case "���35":
                            dvUnitName = "���-35";
                            break;
                        case "���1":
                            dvUnitName = "���-1";
                            break;
                        case "���2":
                            dvUnitName = "���-2";
                            break;
                        case "���2�":
                            dvUnitName = "���-2�";
                            break;
                        case "���10":
                            dvUnitName = "���-10";
                            break;
                        case "���50":
                            dvUnitName = "���-50";
                            break;
                        case "������������":
                            dvUnitName = "���-���������";
                            break;
                        case "��021":
                            dvUnitName = "��-021";
                            break;
                        case "��026":
                            dvUnitName = "��-026";
                            break;
                        case "����2":
                            dvUnitName = "���/�2";
                            break;
                        case "����1":
                            dvUnitName = "���/�1";
                            break;
                        case "����3":
                            dvUnitName = "���/�3";
                            break;
                        case "����2":
                            dvUnitName = "���/�2";
                            break;
                        case "����1":
                            dvUnitName = "���/�1";
                            break;
                        case "����3�":
                            dvUnitName = "���/�3�";
                            break;
                        case "����1":
                            dvUnitName = "���/�1";
                            break;
                        case "����2":
                            dvUnitName = "���/�2";
                            break;
                        case "����3":
                            dvUnitName = "���/�3";
                            break;
                        case "����4":
                            dvUnitName = "���/�4";
                            break;
                        case "����16":
                            dvUnitName = "���/�16";
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
        /// ������ ��� �������� ������������� �������� ������� �� ���������.
        /// </summary>
        public string RegexPatternUnit
        {
            get
            {
                // ���/�7 � 19/2010 �� ������ ���/�7 - 50 - 2/2011
                return string.IsNullOrEmpty(regexPatternUnit)
                           ? string.Format("{0} � {1}/{2}", DVUnitName, Number, Year)
                           : regexPatternUnit;
            }
        }

        /// <summary>
        /// ������ ��� ���������� ������ �� �������� � ����������� ������.
        /// </summary>
        public string RegexPatternParty
        {
            get
            {
                // ����-2.3 - 6 - 12/2010
                return string.IsNullOrEmpty(regexPatternParty)
                           ? string.Format("{0}.*{1}", DVUnitName, Year)
                           : regexPatternParty;
            }
        }

        /// <summary>
        /// ��������� ������� �� ������ ������ ����� �����.
        /// </summary>
        public bool IsParsed { get; private set; }

        public FileInfo PhysicalFile
        {
            get { return physicalFile; }
        }

        #endregion

        #region Constructor

        /// <summary>
        /// ������� ����� ��������� ��������� ����������.
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
                    { if ((DVUnitName != "��-021") && (DVUnitName != "��-026")) IsParsed = false; }
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
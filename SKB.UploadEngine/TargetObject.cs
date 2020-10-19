using System.Collections.Generic;
using System.Security.AccessControl;
using DocsVision.Platform.Security.AccessControl;

namespace RightsAssigner
{
    /// <summary>
    /// ������������ ������, �� ������� ����� ��������� �����.
    /// </summary>
    public class TargetObject
    {
        /// <summary>
        /// ��� �������.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// ��������� ����� �� ����������� ����� �� �������� ��������.
        /// </summary>
        public bool NeedInherit { get; set; }

        /// <summary>
        /// ��������� ����� �� ��������� ����� �� ������.
        /// </summary>
        public bool NeedAssign { get; set; }

        public InheritanceFlags Inheritance
        {
            get { return NeedInherit ? InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit : InheritanceFlags.None; }
        }

        public PropagationFlags Propagation
        {
            get { return PropagationFlags.None; }
        }

        /// <summary>
        /// ������ ����.
        /// </summary>
        public List<Group> Groups { get; private set; }

        /// <summary>
        /// ������� ��������� TargetObject.
        /// </summary>
        /// <param name="name">
        /// ��� �������.
        /// </param>
        /// <param name="needInherit">
        /// ��������� ����� �� ����������� ����� �� �������� ��������.
        /// </param>
        /// <param name="needAssign">
        /// ��������� ����� �� ��������� ����� �� ������.
        /// </param>
        /// <param name="groups">
        /// ������ ����.
        /// </param>
        public TargetObject(string name, bool needInherit, bool needAssign, List<Group> groups)
        {
            this.Name        = name;
            this.NeedInherit = needInherit;
            this.NeedAssign  = needAssign;
            this.Groups      = groups;
        }
    }


    /// <summary>
    /// ������������ ������ ������������.
    /// </summary>
    public class Group
    {
        private string name;
        private string rights;

        /// <summary>
        /// ��� ������.
        /// </summary>
        public string Name
        {
            get { return @"SKB\" + name.Trim(); }
        }

        /// <summary>
        /// ��������� ����� �� ��������� ����� �� ���������� � ������.
        /// </summary>
        public bool CanReadDirectory
        {
            get { return this.rights.Contains("+"); }
        }

        /// <summary>
        /// ����� ����.
        /// </summary>
        public CardDataRights Rights
        {
            get
            {
                if (rights == "�" || rights == "�+")
                    return CardDataRights.ReadData;
                if (this.rights == "�+; ��")
                    return CardDataRights.ReadData | CardDataRights.WriteData | CardDataRights.CreateChildObjects;
                if (this.rights == "�+; ��; �")
                    return CardDataRights.ReadData | CardDataRights.WriteData | CardDataRights.CreateChildObjects | CardDataRights.Copy;
                if (this.rights == "�+; ��; �")
                    return CardDataRights.ReadData | CardDataRights.WriteData | CardDataRights.CreateChildObjects | CardDataRights.Delete;
                if (this.rights == "��")
                    return CardDataRights.FullControl;

                return 0;
            }
        }

        /// <summary>
        /// ������� ��������� Group.
        /// </summary>
        /// <param name="name">
        /// ��� ������.
        /// </param>
        /// <param name="rights">
        /// ��������� ������������� ������ ����.
        /// </param>
        public Group(string name, string rights)
        {
            this.name = name;
            this.rights = rights;
        }
    }
}
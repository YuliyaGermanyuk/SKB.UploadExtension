using System.Collections.Generic;
using System.Security.AccessControl;
using DocsVision.Platform.Security.AccessControl;

namespace RightsAssigner
{
    /// <summary>
    /// Представляет объект, на который нужно назначить права.
    /// </summary>
    public class TargetObject
    {
        /// <summary>
        /// Имя объекта.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Указывает нужно ли наслдеовать права на дочерние элементы.
        /// </summary>
        public bool NeedInherit { get; set; }

        /// <summary>
        /// Указывает нужно ли назначить права на объект.
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
        /// Группы прав.
        /// </summary>
        public List<Group> Groups { get; private set; }

        /// <summary>
        /// Создать экземпляр TargetObject.
        /// </summary>
        /// <param name="name">
        /// Имя объекта.
        /// </param>
        /// <param name="needInherit">
        /// Указывает нужно ли наслдеовать права на дочерние элементы.
        /// </param>
        /// <param name="needAssign">
        /// Указывает нужно ли назначить права на объект.
        /// </param>
        /// <param name="groups">
        /// Группы прав.
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
    /// Представляет группу безопасности.
    /// </summary>
    public class Group
    {
        private string name;
        private string rights;

        /// <summary>
        /// Имя группы.
        /// </summary>
        public string Name
        {
            get { return @"SKB\" + name.Trim(); }
        }

        /// <summary>
        /// Указывает нужно ли назначить права на директорию в архиве.
        /// </summary>
        public bool CanReadDirectory
        {
            get { return this.rights.Contains("+"); }
        }

        /// <summary>
        /// Набор прав.
        /// </summary>
        public CardDataRights Rights
        {
            get
            {
                if (rights == "Ч" || rights == "Ч+")
                    return CardDataRights.ReadData;
                if (this.rights == "Ч+; ИС")
                    return CardDataRights.ReadData | CardDataRights.WriteData | CardDataRights.CreateChildObjects;
                if (this.rights == "Ч+; ИС; К")
                    return CardDataRights.ReadData | CardDataRights.WriteData | CardDataRights.CreateChildObjects | CardDataRights.Copy;
                if (this.rights == "Ч+; ИС; У")
                    return CardDataRights.ReadData | CardDataRights.WriteData | CardDataRights.CreateChildObjects | CardDataRights.Delete;
                if (this.rights == "ПП")
                    return CardDataRights.FullControl;

                return 0;
            }
        }

        /// <summary>
        /// Создать экземпляр Group.
        /// </summary>
        /// <param name="name">
        /// Имя группы.
        /// </param>
        /// <param name="rights">
        /// Строковое представление набора прав.
        /// </param>
        public Group(string name, string rights)
        {
            this.name = name;
            this.rights = rights;
        }
    }
}
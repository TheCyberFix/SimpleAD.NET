using System.Collections.Generic;
using System.DirectoryServices;

namespace NCS.ActiveDirectory
{
    public class ADGroup : ADObject
    {
        public ADGroup(string _searchName)
        {
            this.ActiveDirectorySchemaClassSearchType = ActiveDirectorySchemaClassEnum.group;
            this.Name = _searchName;
        }

        public override string ToString()
        {
            return this.Name;
        }

        //public List<ADGroupMember> getMembersOfGroupDirect()
        //{
        //    List<ADGroupMember> _list = new List<ADGroupMember>();

        //    if (this.Name != null && this.Name != "")
        //    {
        //        foreach (object dn in this.ActiveDirectoryEntry.Properties["member"])
        //        {
        //            _list.Add(getGroupMember(dn));
        //        }
        //    }

        //    return _list;
        //}

        public List<ADGroupMember> getMembersOfGroupDirect()
        {
            List<ADGroupMember> _list = new List<ADGroupMember>();

            if (this.Name != null && this.Name != "")
            {
                getMembers(_list, this.ActiveDirectoryEntry, false);
            }

            _list.Sort((x, y) => string.Compare(x.ADObject.Name, y.ADObject.Name));

            return _list;
        }

        public List<ADGroupMember> getMembersOfGroupRecursive()
        {
            List<ADGroupMember> _list = new List<ADGroupMember>();

            if (this.Name != null && this.Name != "")
            {
                getMembers(_list, this.ActiveDirectoryEntry, true);
            }

            //_list.Sort((x, y) => string.Compare(x.ADObject.Name, y.ADObject.Name));

            return _list;
        }

        private void getMembers(List<ADGroupMember> _list, DirectoryEntry entry, bool recursive)
        {
            var tempList = new List<ADGroupMember>();

            foreach (object dn in entry.Properties["member"])
            {
                var member = getGroupMember(dn);
                tempList.Add(member);
            }

            tempList.Sort((x, y) => string.Compare(x.ADObject.Name, y.ADObject.Name));

            if (recursive)
            {
                tempList.ForEach(item =>
                {
                    if (item.Type == ActiveDirectorySchemaClassEnum.group)
                    {
                        _list.Add(item);
                        getMembers(_list, item.ADObject.ActiveDirectoryEntry, true);
                    }
                    else
                    {
                        _list.Add(item);
                    }
                });
            }
            else
                tempList.ForEach(item => _list.Add(item));
        }

        //private void getMembers(List<ADGroupMember> _list, DirectoryEntry entry, bool recursive)
        //{
        //    foreach (object dn in entry.Properties["member"])
        //    {
        //        var member = getGroupMember(dn);
        //        _list.Add(member);

        //        if (recursive)
        //        {
        //            if (member.Type == ActiveDirectorySchemaClassEnum.group)
        //                getMembers(_list, member.ADObject.ActiveDirectoryEntry, true);
        //        }

        //    }
        //}

        private ADGroupMember getGroupMember(object dn)
        {
            return new ADGroupMember(dn.ToString());// { distinguishedName = dn.ToString() };
        }
    }
}
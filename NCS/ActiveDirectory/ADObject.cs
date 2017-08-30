using System;
using System.Collections.Generic;
using System.Configuration;
using System.DirectoryServices;
using System.Reflection;

namespace NCS.ActiveDirectory
{
    ////Good Reads :-)
    ////http://msdn.microsoft.com/en-us/library/windows/desktop/aa746407(v=vs.85).aspx
    ////http://msdn.microsoft.com/en-us/library/aa746475(v=vs.85).aspx
    ////http://stackoverflow.com/questions/6252819/find-recursive-group-membership-active-directory-using-c-sharp
    ////http://support.microsoft.com/kb/914828   <-- see the examples section

    public class ADObject : IDisposable
    {
        private static string _activeDirectoryBasePathWithServer;
        private static string _activeDirectoryBaseDN;
        private static string _aDAdminUserPassword;
        private static string _aDAdminUsername;
        private static bool _isConfigLoaded = false;

        private string _DistinguishedName;
        private DirectoryEntry _DirectoryEntry;
        private bool _disposed;
        private object _lockObjfor_DirectoryEntry = new object();

        ///types of searches
        //The Glue for the ADActions class
        public enum ActiveDirectorySchemaClassEnum
        {
            user, contact, group, computer
        }

        public ADObject()
        {
            if (!_isConfigLoaded)
                _LoadConfigInternal();
        }

        public string Name
        {
            get;
            set;
        }

        public string SchemaClassName
        {
            get
            {
                return (string)this.ActiveDirectoryEntry.SchemaClassName;
            }
            private set
            {
                notImplementedException();
            }
        }

        public string Mail
        {
            get
            {
                return (string)this.ActiveDirectoryEntry.Properties["mail"].Value ?? string.Empty;
            }

            private set
            {
                notImplementedException();
            }
        }

        public List<string> ProxyAddresses
        {
            get
            {
                List<string> returnList = new List<string>();

                /// if it has at least one entry
                if (this.ActiveDirectoryEntry.Properties["proxyAddresses"].Count > 0)
                {
                    // if only one entry return just the first.
                    if (this.ActiveDirectoryEntry.Properties["proxyAddresses"].Count == 1)
                    {
                        returnList.Add(this.ActiveDirectoryEntry.Properties["proxyAddresses"][0].ToString());

                        return returnList;
                    }
                    // else return all of them
                    else
                    {
                        foreach (var item in this.ActiveDirectoryEntry.Properties["proxyAddresses"].Value as object[])
                        {
                            returnList.Add(item.ToString());
                        }

                        return returnList;
                    }
                }
                // else there is nothing to return
                else
                    return returnList;
            }

            private set
            {
                notImplementedException();
            }
        }

        public ActiveDirectorySchemaClassEnum ActiveDirectorySchemaClassSearchType { get; set; }

        /// <summary>
        /// returns true if the object can be found in the Directory
        /// </summary>
        public bool isFoundInDirectory
        {
            get
            {
                if (!string.IsNullOrWhiteSpace(this.DistinguishedName))
                {
                    return true;
                }

                return false;
            }

            private set { }
        }

        public string Description
        {
            get
            {
                return (string)this.ActiveDirectoryEntry.Properties["description"].Value ?? string.Empty;
            }

            set
            {
    
                if (!string.Equals(this.ActiveDirectoryEntry.Properties["description"].Value, value))
                {
                    if (string.IsNullOrWhiteSpace(value))
                    {
                        this.ActiveDirectoryEntry.Properties["description"].Clear();
                    }
                    else
                    {
                        this.ActiveDirectoryEntry.Properties["description"].Value = value;
                    }

                    this.ActiveDirectoryEntry.CommitChanges();
                }
            }
        }

        public string DistinguishedName
        {
            get
            {
                if (_DistinguishedName != null)
                    return _DistinguishedName;

                else
                {
                    _DistinguishedName = GetObjectDistinguishedName(this.ActiveDirectorySchemaClassSearchType, this.Name);
                    return _DistinguishedName;
                }
            }
            set { _DistinguishedName = value; }
        }

        public string ObjectGuid
        {
            get
            {
                return this.ActiveDirectoryEntry.Guid.ToString();
            }
            private set
            {
                notImplementedException();
            }
        }

        public string ObjectSID
        {
            get
            {
                byte[] x = (byte[])this.ActiveDirectoryEntry.Properties["objectSid"].Value;

                System.Security.Principal.SecurityIdentifier sid = new System.Security.Principal.SecurityIdentifier(x, 0);

                return sid.ToString();
            }
            private set
            {
                notImplementedException();
            }
        }

        public DirectoryEntry ActiveDirectoryEntry
        {
            get
            {
                lock (_lockObjfor_DirectoryEntry)
                {
                    if (_DirectoryEntry != null)
                        return _DirectoryEntry;

                    else
                    {
                        _DirectoryEntry = _getDirectoryEntryWithAuthInternal(string.Format("LDAP://{0}", this.DistinguishedName));
                        return _DirectoryEntry;
                    }
                }
            }

            private set
            {
            }
        }

        public DateTime WhenCreated
        {
            get
            {
                DateTime _temp = (DateTime)this.ActiveDirectoryEntry.Properties["whenCreated"].Value;
                return _temp.ToLocalTime();
            }

            private set
            {
                notImplementedException();
            }
        }

        public List<ADGroup> getGroupMembershipsDirect()
        {
            return _getGroupMembersDirectInternal();
        }

        private List<ADGroup> _getGroupMembersDirectInternal()
        {
            List<ADGroup> _memberOfGroups = new List<ADGroup>();

            using (DirectoryEntry entry = this.ActiveDirectoryEntry)
            {
                if (this.Name != null && this.Name != "")
                {
                    foreach (object dn in this.ActiveDirectoryEntry.Properties["memberOf"])
                    {
                        using (DirectoryEntry tempEntry = _getDirectoryEntryWithAuthInternal(string.Format("LDAP://{0}", dn)))
                        {
                            //tempEntry.Name = 'CN=Group Name Here'
                            //Quick Remove of string 'cn=' is faster than calling AD
                            _memberOfGroups.Add(new ADGroup(tempEntry.Name.Remove(0, 3)));
                        }
                    }
                }
            }

            //Sort the list by name.
            _memberOfGroups.Sort(delegate(ADGroup group1, ADGroup group2)
            {
                return group1.Name.CompareTo(group2.Name);
            });

            return _memberOfGroups;
        }

        public List<ADGroup> getGroupMembershipsRecursive()
        {
            return _getGroupMembersRecursiveInternal();
        }

        private List<ADGroup> _getGroupMembersRecursiveInternal()
        {
            List<ADGroup> _memberOfGroups = new List<ADGroup>();

            using (DirectoryEntry entry = _getDirectoryEntryWithAuthInternal(string.Format("{0}/{1}", _activeDirectoryBasePathWithServer, _activeDirectoryBaseDN)))
            {
                using (DirectorySearcher mySearcher = new DirectorySearcher(entry))
                {
                    string _thisObjCN = this.DistinguishedName;

                    mySearcher.Filter = String.Format("member:1.2.840.113556.1.4.1941:={0}", _thisObjCN);
                    mySearcher.SearchScope = SearchScope.Subtree;
                    mySearcher.PropertiesToLoad.Add("Name");

                    using (SearchResultCollection results = mySearcher.FindAll())
                    {
                        for (int i = 0; i < results.Count; i++)
                        {
                            _memberOfGroups.Add(new ADGroup(results[i].Properties["Name"][0] as string));
                        }
                    }
                }
            }

            //Sort the list by name.
            _memberOfGroups.Sort(delegate(ADGroup group1, ADGroup group2)
            {
                return group1.Name.CompareTo(group2.Name);
            });

            return _memberOfGroups;
        }

        public bool isMemberOfGroup(string activeDirectoryGroupName)
        {
            using (var theGroup = new ActiveDirectory.ADGroup(activeDirectoryGroupName))
            {
                return _isMemberOfGroupInternal(theGroup);
            }
        }

        public bool isMemberOfGroup(ADGroup ActiveDirectoryGroup)
        {
            return _isMemberOfGroupInternal(ActiveDirectoryGroup);
        }

        private bool _isMemberOfGroupInternal(ADGroup _thegroup)
        {
            if (!_thegroup.isFoundInDirectory) // protects against null DistinguishedName below
                return false;

            using (DirectoryEntry entry = _getDirectoryEntryWithAuthInternal(string.Format("{0}/{1}", _activeDirectoryBasePathWithServer, this.DistinguishedName)))
            {
                using (DirectorySearcher mySearcher = new DirectorySearcher(entry))
                {
                    string _groupCN = _thegroup.DistinguishedName;

                    mySearcher.Filter = String.Format("(memberof:1.2.840.113556.1.4.1941:={0})", _groupCN);
                    mySearcher.SearchScope = SearchScope.Base;
                    mySearcher.PropertiesToLoad.Add("CN");

                    using (SearchResultCollection results = mySearcher.FindAll())
                    {
                        if (results.Count > 0)
                            return true;
                        else
                            return false;
                    }
                }
            }
        }

        public void addToGroup(ADGroup _GroupObject)
        {
            try
            {
                //The userDN must not have the 'LDAP://' in front ... see the attributes of a group in AD.
                _GroupObject.ActiveDirectoryEntry.Properties["member"].Add(this.DistinguishedName);
                _GroupObject.ActiveDirectoryEntry.CommitChanges();
            }
            catch (System.DirectoryServices.DirectoryServicesCOMException E)
            {
                E.Message.ToString();
            }
        }

        public void removeFromGroup(ADGroup _GroupObject)
        {
            try
            {
                //The userDN must not have the 'LDAP://' in front ... see the attributes of a group in AD.
                _GroupObject.ActiveDirectoryEntry.Properties["member"].Remove(this.DistinguishedName);
                _GroupObject.ActiveDirectoryEntry.CommitChanges();
            }
            catch (System.DirectoryServices.DirectoryServicesCOMException E)
            {
                E.Message.ToString();
            }
        }

        /// <summary>
        /// Creates an AD User
        /// </summary>
        /// <param name="_newUsersAMAccountName"></param>
        /// <param name="_newUserUserPrincipalName"></param>
        /// <param name="_newUserDirectoryLocationDN"></param>
        /// <param name="_newUserFullName"></param>
        /// <param name="_newUserFirstName"></param>
        /// <param name="_newUserLastName"></param>
        /// <param name="_newUserInitials"></param>
        /// <param name="_newUserDescription"></param>
        protected static void CreateADUser(string _newUsersAMAccountName, string _newUserUserPrincipalName, string _newUserDirectoryLocationDN,
            string _newUserFullName, string _newUserFirstName, string _newUserLastName, string _newUserInitials,
            string _newUserDescription)
        {
            using (DirectoryEntry dirEntry = _getDirectoryEntryWithAuthInternal(string.Format("LDAP://{0}", _newUserDirectoryLocationDN)))//new DirectoryEntry(connectionPrefix))
            using (DirectoryEntry newUser = dirEntry.Children.Add("CN=" + _newUserFullName, "user"))
            {
                //Pre-Windows 2000 Account Name
                newUser.Properties["samAccountName"].Value = _newUsersAMAccountName; //this.AD_accountName.ToLower();
                //Windows 2000+ Domain Name
                newUser.Properties["userPrincipalName"].Value = _newUserUserPrincipalName;//this.AD_accountName.ToLower() + "@example.com";
                //AD Object Name
                newUser.Properties["name"].Value = _newUserFullName;

                //Display name in Exchange and Lync, etc
                newUser.Properties["displayName"].Value = _newUserFullName;
                //given Name : first name
                newUser.Properties["givenName"].Value = _newUserFirstName;
                //sur name : last name
                newUser.Properties["sn"].Value = _newUserLastName;

                if (!string.IsNullOrWhiteSpace(_newUserDescription))
                {
                    newUser.Properties["description"].Value = _newUserDescription;
                }

                //Middle Initial
                if (!string.IsNullOrWhiteSpace(_newUserInitials))
                {
                    newUser.Properties["initials"].Value = _newUserInitials;
                }

                ///Create the User.
                newUser.CommitChanges();
            }
        }

        private static string GetObjectDistinguishedName(ActiveDirectorySchemaClassEnum objectCls, string _objectName)
        {
            using (DirectoryEntry entry = _getDirectoryEntryWithAuthInternal(_activeDirectoryBasePathWithServer))
            {
                using (DirectorySearcher mySearcher = new DirectorySearcher(entry))
                {
                    switch (objectCls)
                    {
                        case ActiveDirectorySchemaClassEnum.user:
                            mySearcher.Filter = string.Format("(&(objectCategory=person)(objectClass=user)(sAMAccountName={0}))", _objectName);
                            break;

                        case ActiveDirectorySchemaClassEnum.contact:
                            mySearcher.Filter = string.Format("(&(objectCategory=person)(objectClass=contact)(Name={0}))", _objectName);
                            break;

                        case ActiveDirectorySchemaClassEnum.group:
                            mySearcher.Filter = string.Format("(&(objectCategory=group)(Name={0}))", _objectName);
                            break;

                        case ActiveDirectorySchemaClassEnum.computer:
                            mySearcher.Filter = string.Format("(&(objectCategory=computer)(Name={0}))", _objectName);
                            break;
                    }

                    ///specifying/limiting the requested information via PropsToLoad is MUCH FASTER
                    mySearcher.PropertiesToLoad.Add("distinguishedName");

                    SearchResult result = mySearcher.FindOne();

                    if (result == null) return null;

                    ///Using the index of the SearchResult is faster than 'for-eaching' through each new DiretoryEntry.
                    return result.Properties["distinguishedName"][0] as string;
                }
            }
        }

        protected SearchResultCollection searchForADObjects(ADObject.ActiveDirectorySchemaClassEnum objectCls, string _searchTerm)
        {
            using (DirectoryEntry entry = _getDirectoryEntryWithAuthInternal(_activeDirectoryBasePathWithServer))
            {
                using (DirectorySearcher mySearcher = new DirectorySearcher(entry))
                {
                    switch (objectCls)
                    {
                        case ActiveDirectorySchemaClassEnum.user:
                            if (string.IsNullOrWhiteSpace(_searchTerm))
                                mySearcher.Filter = "(&(objectCategory=person)(objectClass=user))";
                            else
                                mySearcher.Filter = string.Format("(&(objectCategory=person)(objectClass=user)(sAMAccountName=*{0}*))", _searchTerm);

                            mySearcher.PropertiesToLoad.Add("sAMAccountName");
                            break;

                        case ActiveDirectorySchemaClassEnum.contact:
                            if (string.IsNullOrWhiteSpace(_searchTerm))
                                mySearcher.Filter = "(&(objectCategory=person)(objectClass=contact))";
                            else
                                mySearcher.Filter = string.Format("(&(objectCategory=person)(objectClass=contact)(name=*{0}*))", _searchTerm);

                            mySearcher.PropertiesToLoad.Add("name");
                            break;

                        case ActiveDirectorySchemaClassEnum.group:
                            if (string.IsNullOrWhiteSpace(_searchTerm))
                                mySearcher.Filter = "((objectCategory=group))";

                            else
                                mySearcher.Filter = string.Format("(&(objectCategory=group)(Name=*{0}*))", _searchTerm);

                            mySearcher.PropertiesToLoad.Add("Name");
                            break;

                        case ActiveDirectorySchemaClassEnum.computer:
                            if (string.IsNullOrWhiteSpace(_searchTerm))
                                mySearcher.Filter = "((objectCategory=computer))";
                            else
                                mySearcher.Filter = string.Format("(&(objectCategory=computer)(Name=*{0}*))", _searchTerm);

                            mySearcher.PropertiesToLoad.Add("Name");
                            break;
                    }
                    SearchResultCollection results = mySearcher.FindAll();
                    return results;
                }
            }
        }

        ///Delete a object from AD.
        public void deleteActiveDirectoryObject(bool reallyDeleteForSure)
        {
            if (reallyDeleteForSure)
            {
                try
                {
                    switch (this.ActiveDirectorySchemaClassSearchType)
                    {
                        case ActiveDirectorySchemaClassEnum.user:
                            using (DirectoryEntry parent = this.ActiveDirectoryEntry.Parent)
                            {
                                parent.Children.Remove(this.ActiveDirectoryEntry);
                                parent.CommitChanges();
                            }
                            break;

                        case ActiveDirectorySchemaClassEnum.group:
                            using (DirectoryEntry parent = this.ActiveDirectoryEntry.Parent)
                            {
                                parent.Children.Remove(this.ActiveDirectoryEntry);
                                parent.CommitChanges();
                            }
                            break;

                        case ActiveDirectorySchemaClassEnum.computer:
                            //Since a computer object is actually a container the entire TREE must be removed.
                            this.ActiveDirectoryEntry.DeleteTree();
                            this.ActiveDirectoryEntry.CommitChanges();
                            break;

                        default:
                            break;
                    }
                }
                catch (Exception)
                {
                    throw new Exception("Unable To Delete Object!");
                }
            }
        }

        public void moveActiveDirectoryObject(string newLocationDistingishedName)
        {
            if (this.isFoundInDirectory && DirectoryEntry.Exists(string.Format("LDAP://{0}", newLocationDistingishedName)))
            {
                using (DirectoryEntry nLocation = _getDirectoryEntryWithAuthInternal(string.Format("LDAP://{0}", newLocationDistingishedName)))
                {
                    string newName = this.ActiveDirectoryEntry.Name;
                    this.ActiveDirectoryEntry.MoveTo(nLocation, newName);

                    //Reset the internal backing property to re-get from AD
                    _DistinguishedName = null;
                }
            }
            else
            {
                throw new NullReferenceException
                ("unable to LOCATE the the object in the domain OR unable to LOCATE the new Location: " + newLocationDistingishedName);
            }
        }

        private static DirectoryEntry _getDirectoryEntryWithAuthInternal(string _path)
        {
            if (_aDAdminUserPassword == string.Empty)
                return new DirectoryEntry(_path);
            else
                return new DirectoryEntry(_path, _aDAdminUsername, _aDAdminUserPassword);
        }

        private static void _LoadConfigInternal()
        {

            Configuration x = GetDllConfiguration();

            _activeDirectoryBasePathWithServer = x.AppSettings.Settings["DirectoryBasePathWithServer"].Value;
            _activeDirectoryBaseDN = x.AppSettings.Settings["DirectoryBaseDN"].Value;
            _aDAdminUsername = x.AppSettings.Settings["DirectoryConnectUsername"].Value ?? string.Empty;
            _aDAdminUserPassword = x.AppSettings.Settings["DirectoryConnectPassword"].Value ?? string.Empty;

            _isConfigLoaded = true;
        }

        private static Configuration GetDllConfiguration()
        {
            var configFile = Assembly.GetCallingAssembly().Location + ".config";

            var map = new ExeConfigurationFileMap
            {
                ExeConfigFilename = configFile
            };
            return ConfigurationManager.OpenMappedExeConfiguration(map, ConfigurationUserLevel.None);
        }

        private static void notImplementedException()
        {
            throw new NotImplementedException();
        }

        // The finalizer will call the internal dispose method, telling it to free managed resources.
        //a finalizer is a special method that is executed when an object is garbage collected
        ~ADObject()
        {
            this.Dispose(false);
        }

        public void Dispose()
        {
            this.Dispose(true);

            // Use SupressFinalize in case a subclass of this type implements a finalizer.
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            // If you need thread safety, use a lock around these
            // operations, as well as in your methods that use the resource.
            lock (_lockObjfor_DirectoryEntry)
            {
                if (!_disposed)
                {
                    if (disposing)
                    {
                        if (_DirectoryEntry != null)
                        {
                            _DirectoryEntry.Close();
                            _DirectoryEntry.Dispose();
                        }

                        //Debug.WriteLine("Disposing:" + this.Name);
                    }

                    // Indicate that the instance has been disposed.
                    ActiveDirectoryEntry = null;
                    _DirectoryEntry = null;
                    _disposed = true;
                }
            }
        }
    }
}
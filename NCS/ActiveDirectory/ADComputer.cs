using System;
using ActiveDs;

namespace NCS.ActiveDirectory
{
    public class ADComputer : ADObject
    {
        //Account Control Const in AD
        private readonly long ADS_UF_ACCOUNTDISABLE = 0x0002;

        public ADComputer(string _searchName)
        {
            this.ActiveDirectorySchemaClassSearchType = ActiveDirectorySchemaClassEnum.computer;
            this.Name = _searchName;
        }

        public override string ToString()
        {
            return this.Name;
        }

        /// <summary>
        /// Returns the 'operatingSystem' attribute from the object in the directory (or an empty string if null)
        /// </summary>
        public string operatingSystemName
        {
            get
            {
                return (string)this.ActiveDirectoryEntry.Properties["operatingSystem"].Value ?? string.Empty;
            }
            private set { }
        }

        /// <summary>
        /// Returns the 'operatingSystemVersion' attribute from the object in the directory (or an empty string if null)
        /// </summary>
        public string operatingSystemVersion
        {
            get
            {
                return (string)this.ActiveDirectoryEntry.Properties["operatingSystemVersion"].Value ?? string.Empty;
            }
            private set { }
        }

        /// <summary>
        /// Returns the 'operatingSystemServicePack' attribute from the object in the directory (or an empty string if null)
        /// </summary>
        public string operatingSystemServicePack
        {
            get
            {
                return (string)this.ActiveDirectoryEntry.Properties["operatingSystemServicePack"].Value ?? string.Empty;
            }
            private set { }
        }

        /// <summary>
        /// The lastLogontimeStamp attribute is not updated every time a user or computer logs on to the domain. The decision to update the value is based on the current date minus the value of the (ms-DS-Logon-Time-Sync-Interval attribute minus a random percentage of 5).
        ///  see http://blogs.technet.com/b/askds/archive/2009/04/15/the-lastlogontimestamp-attribute-what-it-was-designed-for-and-how-it-works.aspx
        /// </summary>
        public DateTime lastLogonTimestamp
        {
            //http://blogs.technet.com/b/askds/archive/2009/04/15/the-lastlogontimestamp-attribute-what-it-was-designed-for-and-how-it-works.aspx
            get
            {
                if (this.ActiveDirectoryEntry.Properties["lastLogonTimestamp"].Value != null)
                {
                    IADsLargeInteger plsVal;
                    long winFileTime;

                    plsVal = (IADsLargeInteger)this.ActiveDirectoryEntry.Properties["lastLogonTimestamp"].Value;
                    winFileTime = plsVal.HighPart * 4294967296 + plsVal.LowPart;
                    return DateTime.FromFileTime(winFileTime);
                }
                else
                    return new DateTime();
            }

            private set { notImplementedException(); }
        }

        /// <summary>
        /// DISABLES the Computer Account in Active Directory
        /// </summary>
        public void disableComputerAccount()
        {
            int val = (int)this.ActiveDirectoryEntry.Properties["userAccountControl"].Value;
            this.ActiveDirectoryEntry.Properties["userAccountControl"].Value = val | 0x2;
            //ADS_UF_ACCOUNTDISABLE;

            this.ActiveDirectoryEntry.CommitChanges();
        }

        /// <summary>
        /// ENABLES the Computer Account in Active Directory
        /// </summary>
        public void enableComputerAccount()
        {
            try
            {
                int val = (int)this.ActiveDirectoryEntry.Properties["userAccountControl"].Value;
                this.ActiveDirectoryEntry.Properties["userAccountControl"].Value = val & ~0x2;
                this.ActiveDirectoryEntry.CommitChanges();
            }
            catch (System.DirectoryServices.DirectoryServicesCOMException E)
            {
                E.Message.ToString();
            }
        }

        /// <summary>
        /// Checks to see if the account is DISABLED in the directory
        /// </summary>
        /// <returns>True if the account is disabled</returns>
        public bool isDisabledAccount()
        {
            if (this.ActiveDirectoryEntry.Properties["userAccountControl"].Value != null)
            {
                int currentValue = (int)this.ActiveDirectoryEntry.Properties["userAccountControl"].Value;

                //Is Account disabled?
                if (Convert.ToBoolean(currentValue & ADS_UF_ACCOUNTDISABLE))
                {
                    return true;
                }
            }
            return false;
        }

        private static void notImplementedException()
        {
            throw new NotImplementedException();
        }
    }
}
using System;

namespace NCS.ActiveDirectory
{
    public class ADContact : ADObject
    {
        public ADContact(string _searchName)
        {
            this.ActiveDirectorySchemaClassSearchType = ActiveDirectorySchemaClassEnum.contact;
            this.Name = _searchName;
        }

        public override string ToString()
        {
            return this.Name;
        }

        /// <summary>
        /// Returns the 'displayName' attribute from the object in the directory (or an empty string if null)
        /// Value usually looks like 'Sam W. Vista'
        /// </summary>
        public string displayName
        {
            get
            {
                return (string)this.ActiveDirectoryEntry.Properties["displayName"].Value ?? string.Empty;
            }

            private set { notImplementedException(); }
        }

        /// <summary>
        /// Returns the 'givenName' attribute from the object in the directory (or an empty string if null)
        /// Value usually looks like 'Sam'  (object's first name)
        /// </summary>
        public string firstNameGivenName
        {
            get
            {
                return (string)this.ActiveDirectoryEntry.Properties["givenName"].Value ?? string.Empty;
            }

            private set { notImplementedException(); }
        }

        /// <summary>
        /// Returns the 'sn' attribute from the object in the directory (or an empty string if null)
        /// Value usually looks like 'Vista'  (object's last name)
        /// </summary>
        public string lastNameSurName
        {
            get
            {
                return (string)this.ActiveDirectoryEntry.Properties["sn"].Value ?? string.Empty;
            }

            private set { notImplementedException(); }
        }

        /// <summary>
        /// Returns the 'initials' attribute from the object in the directory (or an empty string if null)
        /// Value usually looks like 'W'  (Just an inital)
        /// </summary>
        public string initials
        {
            get
            {
                return (string)this.ActiveDirectoryEntry.Properties["initials"].Value ?? string.Empty;
            }

            private set { notImplementedException(); }
        }

        /// <summary>
        /// Checks to see if the account is enabled for Lync in the directory
        /// </summary>
        /// <returns>True if the account is enabled for Lync</returns>
        public bool msRTCSIPUserEnabled
        {
            get
            {
                if (this.ActiveDirectoryEntry.Properties["msRTCSIP-UserEnabled"].Value != null)
                    return (bool)this.ActiveDirectoryEntry.Properties["msRTCSIP-UserEnabled"].Value;
                else
                    return false;
            }

            private set { notImplementedException(); }
        }

        /// <summary>
        /// Returns the 'msRTCSIP-Line' attribute from the object in the directory (or an empty string if null)
        /// The Lync Line number ex: 'tel:+17777777777;ext=5424'
        /// </summary>
        public string msRTCSIPLine
        {
            get
            {
                return (string)this.ActiveDirectoryEntry.Properties["msRTCSIP-Line"].Value ?? string.Empty;
            }

            private set { notImplementedException(); }
        }

        /// <summary>
        /// Returns the 'company' attribute from the object in the directory (or an empty string if null)
        /// </summary>
        public string company
        {
            get
            {
                
                return (string)this.ActiveDirectoryEntry.Properties["company"].Value ?? string.Empty;
            }
            set
            {
                
                if (!string.Equals(this.ActiveDirectoryEntry.Properties["company"].Value, value))
                {
                    this.ActiveDirectoryEntry.Properties["company"].Value = value;
                    this.ActiveDirectoryEntry.CommitChanges();
                }
            }
        }

        /// <summary>
        /// Returns the 'employeeNumber' attribute from the object in the directory (or an empty string if null)
        /// </summary>
        public string employeeNumber
        {
            get
            {
                return (string)this.ActiveDirectoryEntry.Properties["employeeNumber"].Value ?? string.Empty;
            }
            set
            {
                if (!string.Equals(this.ActiveDirectoryEntry.Properties["employeeNumber"].Value, value))
                {
                    this.ActiveDirectoryEntry.Properties["employeeNumber"].Value = value;
                    this.ActiveDirectoryEntry.CommitChanges();
                }
            }
        }

        /// <summary>
        /// Returns the 'department' attribute from the object in the directory (or an empty string if null)
        /// </summary>
        public string department
        {
            get
            {
                if (this.ActiveDirectoryEntry.Properties["department"].Value != null)
                    return this.ActiveDirectoryEntry.Properties["department"].Value.ToString();
                else
                    return string.Empty;
                //obscure issue
                //return (string)this.ActiveDirectoryEntry.Properties["department"].Value ?? string.Empty;
            }
            set
            {
                if (!string.Equals(this.ActiveDirectoryEntry.Properties["department"].Value, value))
                {
                    this.ActiveDirectoryEntry.Properties["department"].Value = value;
                    this.ActiveDirectoryEntry.CommitChanges();
                }
            }
        }

        /// <summary>
        /// Returns the 'title' attribute from the object in the directory (or an empty string if null)
        /// </summary>
        public string title
        {
            get
            {
                return (string)this.ActiveDirectoryEntry.Properties["title"].Value ?? string.Empty;
            }
            set
            {
                if (!string.Equals(this.ActiveDirectoryEntry.Properties["title"].Value, value.Trim()))
                {
                    this.ActiveDirectoryEntry.Properties["title"].Value = value.Trim();
                    this.ActiveDirectoryEntry.CommitChanges();
                }
            }
        }

        /// <summary>
        /// Returns the 'telephoneNumber' attribute from the object in the directory (or an empty string if null)
        /// ex: '+17777777777'
        /// </summary>
        public string telephoneNumber
        {
            get
            {
                //I can't just return value ?? string.empty because you need the null check on the
                //if (this.ActiveDirectoryEntry.Properties["telephoneNumber"].Value != null)
                return (string)this.ActiveDirectoryEntry.Properties["telephoneNumber"].Value ?? string.Empty;
                //else
                //  return string.Empty;
            }
            set
            {
                this.ActiveDirectoryEntry.Properties["telephoneNumber"].Value = value.Trim();
                this.ActiveDirectoryEntry.CommitChanges();
            }
        }

        /// <summary>
        /// Returns the 'physicalDeliveryOfficeName' attribute from the object in the directory (or an empty string if null)
        /// </summary>
        public string physicalDeliveryOfficeName
        {
            get
            {
                //I can't just return value ?? string.empty because you need the null check on the
                //if (this.ActiveDirectoryEntry.Properties["telephoneNumber"].Value != null)
                return (string)this.ActiveDirectoryEntry.Properties["physicalDeliveryOfficeName"].Value ?? string.Empty;
                //else
                //  return string.Empty;
            }
            set
            {
                this.ActiveDirectoryEntry.Properties["physicalDeliveryOfficeName"].Value = value.Trim();
                this.ActiveDirectoryEntry.CommitChanges();
            }
        }

        /// <summary>
        /// Returns an ADUser object created from the 'manager' attribute in the directory (or NULL if no value)
        /// </summary>
        public ADUser manager
        {
            get
            {
                //I can't just return value ?? string.empty because you need the null check on the
                if (this.ActiveDirectoryEntry.Properties["manager"].Value != null)
                {
                    using (ADObject obj = new ADObject())
                    {
                        obj.DistinguishedName = (string)this.ActiveDirectoryEntry.Properties["manager"].Value;

                        string _tempName = obj.ActiveDirectoryEntry.Properties["sAMAccountName"].Value as string;

                        return new ADUser(_tempName);
                    }
                }
                else
                    return null;
            }
            set
            {
                this.ActiveDirectoryEntry.Properties["manager"].Value = value.DistinguishedName; //makse sure it has no e 'LDAP://' in front.
                this.ActiveDirectoryEntry.CommitChanges();
            }
        }

        private static void notImplementedException()
        {
            throw new NotImplementedException();
        }

        ////Account Control Const in AD
        //private readonly long ADS_UF_ACCOUNTDISABLE = 0x0002;

        ///// <summary>
        ///// Internal pratical AD conversion
        ///// </summary>
        ///// <param name="Li"></param>
        ///// <returns></returns>
        //private static long GetLongFromLargeInteger(IADsLargeInteger Li)
        //{
        //    long retval = Li.HighPart;
        //    retval <<= 32;
        //    retval |= (uint)Li.LowPart;
        //    return retval;
        //}
    }
}
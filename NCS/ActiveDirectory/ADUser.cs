using System;
using System.Linq;
using ActiveDs;

namespace NCS.ActiveDirectory
{
    public class ADUser : ADObject
    {
        /// <summary>
        /// Finds an existing user in Active Directory
        /// </summary>
        /// <param name="_searchName">the account name to look for with no spaces.  Ex. 'firstnamelastname'</param>
        public ADUser(string _searchName)
        {
            this.ActiveDirectorySchemaClassSearchType = ActiveDirectorySchemaClassEnum.user;
            this.Name = _searchName;
        }

        public override string ToString()
        {
            return this.sAMAccountName;
        }

        /// <summary>
        /// Creates a new User Account in Active Directory. Creates the full name for you.
        /// </summary>
        /// <param name="_newUsersAMAccountName">REQUIRED: Pre-Windows 2000 Account Name; ex. 'firstnamelastname'</param>
        /// <param name="_newUserUserPrincipalName">REQUIRED: Windows 2000+ Domain Name; ex. 'firstnamelastname@domain.com</param>
        /// <param name="_newUserDirectoryLocationDN">REQUIRED: the DN of the OU that you want to create this user in ex. '</param>
        /// <param name="_newUserFirstName">REQUIRED: The user's first name</param>
        /// <param name="_newUserLastName">REQUIRED: The user's last name</param>
        /// <param name="_newUserMiddleInitial">The user's initials</param>
        /// <param name="_newUserDescription">The user's description</param>
        public ADUser(string _newUsersAMAccountName, string _newUserUserPrincipalName, string _newUserDirectoryLocationDN,
             string _newUserFirstName, string _newUserLastName, string _newUserMiddleInitial,
            string _newUserDescription)
        {
            this.ActiveDirectorySchemaClassSearchType = ActiveDirectorySchemaClassEnum.user;
            this.Name = _newUsersAMAccountName;

            ///generate the 'full name'
            System.Text.StringBuilder fullNameBuilder = new System.Text.StringBuilder();

            //add the first name
            fullNameBuilder.Append(_newUserFirstName);

            if (!string.IsNullOrWhiteSpace(_newUserMiddleInitial))
            {
                fullNameBuilder.Append(" " + _newUserMiddleInitial.Trim() + ".");
            }

            if (!string.IsNullOrWhiteSpace(_newUserLastName))
            {
                fullNameBuilder.Append(" " + _newUserLastName);
            }

            CreateADUser(_newUsersAMAccountName.ToLower(), ///to lower is not required but often a standard
                         _newUserUserPrincipalName,
                         _newUserDirectoryLocationDN,
                         fullNameBuilder.ToString(),
                         _newUserFirstName,
                         _newUserLastName,
                         _newUserMiddleInitial,
                         _newUserDescription
            );
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
        /// Returns the 'sAMAccountName' attribute from the object in the directory (or an empty string if null)
        /// Value usually looks like 'samvista'
        /// </summary>
        public string sAMAccountName
        {
            get
            {
                return (string)this.ActiveDirectoryEntry.Properties["sAMAccountName"].Value ?? string.Empty;
            }

            private set { notImplementedException(); }
        }

        /// <summary>
        /// Returns the 'userPrincipalName' attribute from the object in the directory (or an empty string if null)
        /// Value usually looks like 'samvista@domain.com'
        /// </summary>
        public string userPrincipalName
        {
            get
            {
                return (string)this.ActiveDirectoryEntry.Properties["userPrincipalName"].Value ?? string.Empty;
            }

            private set { notImplementedException(); }
        }

        /// <summary>
        /// Returns the date when the current password was set.
        /// </summary>
        public DateTime passwordLastSetTime
        {
            get
            {
                IADsLargeInteger plsVal;
                long winFileTime;

                plsVal = (IADsLargeInteger)this.ActiveDirectoryEntry.Properties["pwdLastSet"].Value;
                winFileTime = plsVal.HighPart * 4294967296 + plsVal.LowPart;
                return DateTime.FromFileTime(winFileTime);
            }

            private set { notImplementedException(); }
        }

        /// <summary>
        /// Returns the time of the last attempt to log in with an incorrect password.
        /// </summary>
        public DateTime badPasswordTime
        {
            get
            {
                if (this.ActiveDirectoryEntry.Properties["BadPasswordTime"].Value != null)
                {
                    IADsLargeInteger plsVal;
                    long filetime;

                    plsVal = (IADsLargeInteger)this.ActiveDirectoryEntry.Properties["BadPasswordTime"].Value;
                    filetime = plsVal.HighPart * 4294967296 + plsVal.LowPart;
                    return DateTime.FromFileTime(filetime);
                }
                else
                    return new DateTime();
            }
            private set { notImplementedException(); }
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



        /// <summary>
        /// Checks to see if the account password is set to never expire in the directory
        /// </summary>
        /// <returns>True if the account password is set to never expire</returns>
        public bool isPasswordNeverExpireSet()
        {
            if (this.ActiveDirectoryEntry.Properties["userAccountControl"].Value != null)
            {
                int currentValue = (int)this.ActiveDirectoryEntry.Properties["userAccountControl"].Value;

                //Is Account password set to never expire?
                if (Convert.ToBoolean(currentValue & ADS_UF_DONT_EXPIRE_PASSWD))
                {
                    return true;
                }
            }
            return false;
        }




        /// <summary>
        /// Checks to see if the account is locked out in the directory
        /// </summary>
        /// <returns>True if the account is locked out</returns>
        public bool isLockedOutAccount()
        {
            if (this.ActiveDirectoryEntry.Properties["lockoutTime"].Value != null)
            {
                IADsLargeInteger currentValue = (IADsLargeInteger)this.ActiveDirectoryEntry.Properties["lockoutTime"].Value;
                if (GetLongFromLargeInteger(currentValue) != 0 )
                {
                    return true;
                }
            }

            return false;
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
                    if (string.IsNullOrWhiteSpace(value))
                    {
                        this.ActiveDirectoryEntry.Properties["company"].Clear();
                    }
                    else
                    {
                        this.ActiveDirectoryEntry.Properties["company"].Value = value;
                    }

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
                commitPropValueInternal("employeeNumber", value);

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
                commitPropValueInternal("department", value);
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
                commitPropValueInternal("title", value);
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
                commitPropValueInternal("telephoneNumber", value);
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
                commitPropValueInternal("physicalDeliveryOfficeName", value);
            }
        }


        /// <summary>
        /// checks for same value, null, empty, and whitespace before either clearing the value or setting value - then commits to AD
        /// </summary>
        /// <param name="_property"></param>
        /// <param name="_value"></param>
        private void commitPropValueInternal(string _property, string _value)
        {
            if (!string.Equals(this.ActiveDirectoryEntry.Properties[_property].Value, _value))
            {
                if (string.IsNullOrWhiteSpace(_value))
                {
                    this.ActiveDirectoryEntry.Properties[_property].Clear();
                }
                else
                {
                    this.ActiveDirectoryEntry.Properties[_property].Value = _value.Trim();
                }

                this.ActiveDirectoryEntry.CommitChanges();
            }
        }

        /// <summary>
        /// Returns the 'msExchUMRecipientDialPlanLink' attribute from the object in the directory (or an empty string if null)
        /// </summary>
        public string msExchUMRecipientDialPlanLink
        {
            get
            {
                //I can't just return value ?? string.empty because you need the null check on the
                //if (this.ActiveDirectoryEntry.Properties["telephoneNumber"].Value != null)
                return (string)this.ActiveDirectoryEntry.Properties["msExchUMRecipientDialPlanLink"].Value ?? string.Empty;
                //else
                //  return string.Empty;
            }
            private set { notImplementedException(); }
        }

        /// <summary>
        /// Returns the 'msExchUMEnabledFlags' attribute from the object in the directory (or an empty string if null)
        /// </summary>
        public int? msExchUMEnabledFlags
        {
            get
            {
                //I can't just return value ?? string.empty because you need the null check on the
                //if (this.ActiveDirectoryEntry.Properties["telephoneNumber"].Value != null)
                return (int?)this.ActiveDirectoryEntry.Properties["msExchUMEnabledFlags"].Value;
                //else
                //  return string.Empty;
            }
            private set { notImplementedException(); }
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
                //this.ActiveDirectoryEntry.Properties["manager"].Value = value.DistinguishedName; //makse sure it has no e 'LDAP://' in front.
                //this.ActiveDirectoryEntry.CommitChanges();

                commitPropValueInternal("manager", value.DistinguishedName);

            }
        }

        public void clearManager()
        {
            commitPropValueInternal("manager", string.Empty);
        }

        /// <summary>
        /// The thumbnailPhoto property of the AD User in a byte array.  Returns null if value is empty.
        /// </summary>
        public byte[] thumbnailPhoto
        {
            get
            {
                if ((byte[])this.ActiveDirectoryEntry.Properties["thumbnailPhoto"].Value != null)
                    return (byte[])this.ActiveDirectoryEntry.Properties["thumbnailPhoto"].Value;
                else
                    return null;
            }

            set
            {
                bool _isSameThumbnailPhoto = false;

                if ((byte[])this.ActiveDirectoryEntry.Properties["thumbnailPhoto"].Value == null)
                    _isSameThumbnailPhoto = false;
                else
                    _isSameThumbnailPhoto = value.SequenceEqual((byte[])this.ActiveDirectoryEntry.Properties["thumbnailPhoto"].Value);///returns true if same

                // if value is not null and is not the same photo, insert the new photo
                if (value != null && _isSameThumbnailPhoto)
                {
                    this.ActiveDirectoryEntry.Properties["thumbnailPhoto"].Clear();
                    this.ActiveDirectoryEntry.Properties["thumbnailPhoto"].Insert(0, value);
                    this.ActiveDirectoryEntry.CommitChanges();
                }
            }
        }

        /// <summary>
        /// Disables the User Account in Active Directory; the user cannot logon to the domain when the account is disabled.
        /// </summary>
        public void disableUserAccount()
        {
            int val = (int)this.ActiveDirectoryEntry.Properties["userAccountControl"].Value;
            this.ActiveDirectoryEntry.Properties["userAccountControl"].Value = val | 0x2;
            //ADS_UF_ACCOUNTDISABLE;

            this.ActiveDirectoryEntry.CommitChanges();
        }

        /// <summary>
        /// Resets the user's password in the directory
        /// </summary>
        /// <param name="_NewPassword">The new password.  It must meet the password complexity requirements for the domain</param>
        public void resetUserPassword(string _NewPassword)
        {
            try
            {
                this.ActiveDirectoryEntry.Invoke("SetPassword", new object[] { "" + _NewPassword + "" }); //set password
                this.ActiveDirectoryEntry.CommitChanges();
            }
            catch (Exception e)
            {
                throw new NullReferenceException
              ("unable to RESET PASSWORD for Object " + this.DistinguishedName + " in the domain. " + e);
            }
        }

        /// <summary>
        /// Sets the PwdLastSet attribute in the directory to force the user to change their password on next interactive logon.
        /// </summary>
        public void setChangePasswordAtNextLogon()
        {
            try
            {
                this.ActiveDirectoryEntry.Properties["PwdLastSet"].Value = 0; //Change Password at next logon
                this.ActiveDirectoryEntry.CommitChanges();
            }
            catch (Exception e)
            {
                throw new NullReferenceException
              ("unable to set change pw at next logon for Object " + this.DistinguishedName + " in the domain. " + e);
            }
        }

        /// <summary>
        /// unlocks the user's account in the directory
        /// </summary>
        public void unlockUserAccount()
        {
            try
            {
                this.ActiveDirectoryEntry.Properties["LockOutTime"].Value = 0; //unlock account
                this.ActiveDirectoryEntry.CommitChanges();
            }
            catch (Exception e)
            {
                throw new NullReferenceException
              ("unable to UNLOCK ACCOUNT for Object " + this.DistinguishedName + " in the domain. " + e);
            }
        }

        /// <summary>
        /// ENABLES the users's account in the directory
        /// </summary>
        public void enableUserAccount()
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

        private static void notImplementedException()
        {
            throw new NotImplementedException();
        }

        //Account Control Const in AD
        private readonly long ADS_UF_ACCOUNTDISABLE = 0x0002;
        private readonly long ADS_UF_DONT_EXPIRE_PASSWD = 0x10000;

        /// <summary>
        /// Internal pratical AD conversion
        /// </summary>
        /// <param name="Li"></param>
        /// <returns></returns>
        private static long GetLongFromLargeInteger(IADsLargeInteger Li)
        {
            long retval = Li.HighPart;
            retval <<= 32;
            retval |= (uint)Li.LowPart;
            return retval;
        }
    }
}
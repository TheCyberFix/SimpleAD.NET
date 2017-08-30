using System.Collections.Generic;
using System.DirectoryServices;

namespace NCS.ActiveDirectory
{
    public class ADSearches : ADObject
    {
        public List<ADUser> findUsers(string searchByName)
        {
            using (SearchResultCollection results = searchForADObjects(ADObject.ActiveDirectorySchemaClassEnum.user, searchByName))
            {
                List<ADUser> objList = new List<ADUser>();

                foreach (SearchResult result in results)
                {
                    objList.Add(new ADUser(result.Properties["sAMAccountName"][0] as string));
                }

                return objList;
            }
        }

        public List<ADGroup> findGroups(string searchByName)
        {
            using (SearchResultCollection results = searchForADObjects(ADObject.ActiveDirectorySchemaClassEnum.group, searchByName))
            {
                List<ADGroup> objList = new List<ADGroup>();

                foreach (SearchResult result in results)
                {
                    objList.Add(new ADGroup(result.Properties["Name"][0] as string));
                }

                return objList;
            }
        }

        public List<ADComputer> findComputers(string searchByName)
        {
            using (SearchResultCollection results = searchForADObjects(ADObject.ActiveDirectorySchemaClassEnum.computer, searchByName))
            {
                List<ADComputer> objList = new List<ADComputer>();

                foreach (SearchResult result in results)
                {
                    objList.Add(new ADComputer(result.Properties["Name"][0] as string));
                }

                return objList;
            }
        }

        public List<ADContact> findContacts(string searchByName)
        {
            using (SearchResultCollection results = searchForADObjects(ADObject.ActiveDirectorySchemaClassEnum.contact, searchByName))
            {
                List<ADContact> objList = new List<ADContact>();

                foreach (SearchResult result in results)
                {
                    objList.Add(new ADContact(result.Properties["Name"][0] as string));
                }

                return objList;
            }
        }
    }
}
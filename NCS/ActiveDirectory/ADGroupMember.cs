namespace NCS.ActiveDirectory
{
    public class ADGroupMember
    {
        public ADObject ADObject { get; set; }

        public ADObject.ActiveDirectorySchemaClassEnum Type { get; private set; }

        public ADGroupMember(string _distinguishedName)
        {
            ADObject = new ADObject();
            ADObject.DistinguishedName = _distinguishedName;

            string _temp = ADObject.ActiveDirectoryEntry.SchemaClassName;

            switch (_temp)
            {
                case ("user"):

                    this.Type = ADObject.ActiveDirectorySchemaClassEnum.user;
                    ADObject.Name = ADObject.ActiveDirectoryEntry.Properties["sAMAccountName"].Value.ToString();

                    break;

                case ("group"):

                    this.Type = ADObject.ActiveDirectorySchemaClassEnum.group;
                    ADObject.Name = ADObject.ActiveDirectoryEntry.Properties["name"].Value.ToString();
                    break;

                case ("computer"):

                    this.Type = ADObject.ActiveDirectorySchemaClassEnum.computer;
                    ADObject.Name = ADObject.ActiveDirectoryEntry.Properties["name"].Value.ToString();
                    break;

                case ("contact"):

                    this.Type = ADObject.ActiveDirectorySchemaClassEnum.contact;
                    ADObject.Name = ADObject.ActiveDirectoryEntry.Properties["name"].Value.ToString();

                    break;
            }
        }

        //public DirectoryEntry DirectoryEntry
        //{
        //    get
        //    {
        //        ADObject x = new ADObject();
        //        x.DistinguishedName = this.distinguishedName;
        //        return x.ActiveDirectoryEntry;
        //    }

        //}
    }
}
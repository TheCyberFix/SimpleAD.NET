# SimpleAD.NET

The goal of this .NET library is not performance, but simplicity.  We attempt to abstract away some of the necessary operations required to make changes (via system.directoryservices.directoryentry) to user, group, computer and contact objects in Active Directory while still allowing for direct directoryentry modification, when necessary.


## Requirements to Build
- .NET  4.5+  (VS Project currently targets 4.7)
- assemblyref://System.DirectoryServices
- assemblyref://System.Configuration (for the .config file)
- comref://Interop.ActiveDs (for certain AD Types)


## Works on

Windows 7+
Server 2008 R2+


## Use

1. Download .dll & .dll.config file (see releases)
2. Edit .dll.config file
3. Add .dll as reference to your VS project (or LINQpad)


## Examples:

```csharp

using NCS.ActiveDirectory

////Examples

//Find Computer Objects
using(var x = new ADSearches())
{
    var list = x.findComputers("SERVER");
}

//Modify title/department/manager attributes
using (var user = new ADUser("johnsmith"))
{
    user.title = "SVP of Data");
    user.department = "IT";
    user.manager = new ADUser("ExistingUser")
}

//Find all Domain Admins with recursive group membership search
using (ADGroup group = new ADGroup("Domain Admins"))
{
    var mems = group.getMembersOfGroupRecursive();
    mems.ForEach(x => Console.WriteLine(x.ADObject.Name));
}



//Find locked out user accounts (and unlock them)
using (var x = new ADSearches())
{
    var list = x.findUsers("");

    list.ForEach(user =>
    {
        if (user.isLockedOutAccount())
        {
            user.unlockUserAccount(); 
        }
    }
}
 


//Remove user from group
using (var user = new ADUser("johnsmith"))
using (var group = new ADGroup("Domain Admins"))
{
    if (user.isFoundInDirectory && group.isFoundInDirectory)
    {
        user.removeFromGroup(group);
    }
}





///direct directoryentry modification

using System.DirectoryServices

//modify user title attribute
using (var user = new ADUser("johnsmith"))
{

    string newTitle = "VP of Helpdesk";

    if (string.IsNullOrWhiteSpace(newTitle))
    {
        user.ActiveDirectoryEntry.Properties["title"].Clear();
    }
    else
    {
        user.ActiveDirectoryEntry.Properties["title"].Value = newTitle.Trim();
    }

    user.ActiveDirectoryEntry.CommitChanges();
}

```

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;

namespace ActiveDirectorySearchDemo
{
    public static class AccountManagementExtensions
    {

        public static String GetProperty(this Principal principal, string property)
        {
            DirectoryEntry directoryEntry = principal.GetUnderlyingObject() as DirectoryEntry;
            if (directoryEntry.Properties.Contains(property))
                return directoryEntry.Properties[property].Value.ToString();
            else
                return String.Empty;
            
        }

        public static string GetCompany(this Principal principal)
        {
            return principal.GetProperty("company").Substring(principal.GetProperty("company").Length - 4);
        }

        public static string GetDepartment(this Principal principal)
        {
            return principal.GetProperty("department");
        }

        public static string GetCurrentFullUserName(this Principal principal)
        {
            return principal.Name;
        }

        public static string GetCurrentUserName(this Principal principal)
        {
            return principal.SamAccountName;
        }

    }
}

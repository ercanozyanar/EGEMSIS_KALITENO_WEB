using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Security;
using Uretim_Kalite.Models.Entity;

namespace Uretim_Kalite.Roller
{
    public class KullaniciRolleri : RoleProvider
    {
        public override string ApplicationName { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public override void AddUsersToRoles(string[] usernames, string[] roleNames)
        {
            throw new NotImplementedException();
        }

        public override void CreateRole(string roleName)
        {
            throw new NotImplementedException();
        }

        public override bool DeleteRole(string roleName, bool throwOnPopulatedRole)
        {
            throw new NotImplementedException();
        }

        public override string[] FindUsersInRole(string roleName, string usernameToMatch)
        {
            throw new NotImplementedException();
        }

        public override string[] GetAllRoles()
        {
            throw new NotImplementedException();
        }


        EGEM2021Entities1 db = new EGEM2021Entities1();
        public override string[] GetRolesForUser(string username)
        {
            var kullanici= db.EGEM_KALITE_USER.FirstOrDefault(x => x.AD_SOYAD == username);
            return new string[] { kullanici.ROL };
        
        }




        public override string[] GetUsersInRole(string roleName)
        {
            throw new NotImplementedException();
        }

        public override bool IsUserInRole(string username, string roleName)
        {
            throw new NotImplementedException();
        }

        public override void RemoveUsersFromRoles(string[] usernames, string[] roleNames)
        {
            throw new NotImplementedException();
        }

        public override bool RoleExists(string roleName)
        {
            throw new NotImplementedException();
        }
    }
}
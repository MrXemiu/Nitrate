using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SqlServer.Management.Smo;

namespace Nitrate.Plugins.SqlServer
{
    public static class ExtensionMethods
    {
        /// <summary>
        /// Gets the database user for a given log in.
        /// </summary>
        /// <param name="db">The parent of a UserCollection</param>
        /// <param name="loginName">A Login.Name</param>
        /// <returns>The database user for a login name if one exists, null otherwise.</returns>
        public static User UserForLogin(this Database db, string loginName)
        {
            for (var i = 0; i < db.Users.Count; i++)
            {
                if (db.Users[i].Login.Equals(loginName, StringComparison.OrdinalIgnoreCase))
                    return db.Users[i];
            }

            return null;
        }
    }
}

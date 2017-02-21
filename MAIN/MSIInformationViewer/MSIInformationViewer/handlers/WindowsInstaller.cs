using System;
using System.Linq;
using System.Windows;
using Microsoft.Deployment.WindowsInstaller;

namespace MSIInformationViewer.handlers
{
    internal class MsiInformation
    {
        public string ProductName { get; set; }

        public string ProductVersion { get; set; }

        public string ProductCode { get; set; }

        public string AllUsers { get; set; }

        public string InstallLevel { get; set; }

        public string Manufacturer { get; set; }

        public string ReinstallMode { get; set; }

        public string UpgradeCode { get; set; }
    }


    internal static class WindowsInstaller
    {
        internal static bool IsMsiReadable(string pathToMsi, string pathToMst)
        {
            try
            {
                using (var database = new Database(pathToMsi, DatabaseOpenMode.ReadOnly))
                {
                    if (pathToMst != string.Empty)
                    {
                        database.ApplyTransform(pathToMst);    
                    } 

                    return true;
                }
            }
            catch (Exception)
            {
                return false;
            }
        }

        internal static bool HasTable(string pathToMsi, string pathToMst, string tableName)
        {
            using (var database = new Database(pathToMsi, DatabaseOpenMode.ReadOnly))
            {
                if (pathToMst != string.Empty)
                {
                    database.ApplyTransform(pathToMst);
                }

                return database.Tables.Contains(tableName);
            }
        }


        internal static MsiInformation GetMsiInformation(string pathToMsi, string pathToMst)
        {
            var returnDatabase = new MsiInformation();

            try
            {
                using (var database = new Database(pathToMsi, DatabaseOpenMode.ReadOnly))
                {
                    if (pathToMst != string.Empty)
                    {
                        database.ApplyTransform(pathToMst);
                    }

                    returnDatabase.ProductName = database.ExecutePropertyQuery("ProductName");
                    returnDatabase.ProductVersion = database.ExecutePropertyQuery("ProductVersion");
                    returnDatabase.ProductCode = database.ExecutePropertyQuery("ProductCode");
                    returnDatabase.AllUsers = database.ExecutePropertyQuery("ALLUSERS");
                    returnDatabase.InstallLevel = database.ExecutePropertyQuery("INSTALLLEVEL");
                    returnDatabase.Manufacturer = database.ExecutePropertyQuery("Manufacturer");
                    returnDatabase.ReinstallMode = database.ExecutePropertyQuery("REINSTALLMODE");
                    returnDatabase.UpgradeCode = database.ExecutePropertyQuery("UpgradeCode");
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Error while reading MSI Informations! Application will close.");
                Environment.Exit(13);
            }

            return returnDatabase;
        }

        internal static int GetShortcutsCount(string pathToMsi, string pathToMst)
        {
            var count = new int();

            try
            {
                using (var database = new Database(pathToMsi, DatabaseOpenMode.ReadOnly))
                {
                    if (pathToMst != string.Empty)
                    {
                        database.ApplyTransform(pathToMst);
                    }

                    if (database.Tables.Contains("Shortcut"))
                    {
                        using (var view = database.OpenView(database.Tables["Shortcut"].SqlSelectString))
                        {
                            view.Execute();
                            count += view.Count();
                        }
                    }
                    else
                    {
                        return 0;
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show( e + " Error while counting shortcuts!  Application will close.");
                Environment.Exit(14);
            }

            return count;
        }

        internal static int GetFilesCount(string pathToMsi, string pathToMst)
        {
            var count = new int();

            try
            {
                using (var database = new Database(pathToMsi, DatabaseOpenMode.ReadOnly))
                {
                    if (pathToMst != string.Empty)
                    {
                        database.ApplyTransform(pathToMst);
                    }

                    if (database.Tables.Contains("File"))
                    {
                        using (var view = database.OpenView(database.Tables["File"].SqlSelectString))
                        {
                            view.Execute();
                            count += view.Count();
                        }
                    }
                    else
                    {
                        return 0;
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Error while counting files! Application will close.");
                Environment.Exit(15);
            }

            return count;
        }

        internal static int GetRegistrysCount(string pathToMsi, string pathToMst)
        {
            var count = new int();

            try
            {
                using (var database = new Database(pathToMsi, DatabaseOpenMode.ReadOnly))
                {
                    if (pathToMst != string.Empty)
                    {
                        database.ApplyTransform(pathToMst);
                    }

                    if (database.Tables.Contains("Registry"))
                    {
                        using (var view = database.OpenView(database.Tables["Registry"].SqlSelectString))
                        {
                            view.Execute();
                            count += view.Count();
                        }
                    }
                    else
                    {
                        return 0;
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Error while counting registrys! Application will close.");
                Environment.Exit(16);
            }

            return count;
        }
    }
}

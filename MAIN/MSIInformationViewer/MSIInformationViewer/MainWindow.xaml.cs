using System;
using System.IO;
using System.Windows;
using Microsoft.Win32;
using System.Diagnostics;

namespace MsiInformationViewer
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private static string _pathToMsi = string.Empty;
        private static string _pathToMst = string.Empty;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void SetMsiInfo()
        {
            var databaseInfromations = handlers.WindowsInstaller.GetMsiInformation(_pathToMsi, _pathToMst);

            ProductNameTextBox.Text = databaseInfromations.ProductName;
            ProductVersionTextBox.Text = databaseInfromations.ProductVersion;
            ProductCodeTextBox.Text = databaseInfromations.ProductCode;
            AllUsersTextBox.Text = databaseInfromations.AllUsers;
            InstalllevelTextBox.Text = databaseInfromations.InstallLevel;
            ManufacturerTextBox.Text = databaseInfromations.Manufacturer;
            ReinstallmodeTextBox.Text = databaseInfromations.ReinstallMode;
            UpgradeCodeTextBox.Text = databaseInfromations.UpgradeCode;
        }

        private void MainWindow_OnLoaded(object sender, RoutedEventArgs e)
        {
            if (!Debugger.IsAttached)
            {
                CheckExtensions();

                if (_pathToMsi == string.Empty)
                {
                    Environment.Exit(10);
                }

                if (_pathToMst == string.Empty)
                {
                    MstStackpanel.Visibility = Visibility.Collapsed;
                }

                if (handlers.WindowsInstaller.IsMsiReadable(_pathToMsi, _pathToMst))
                {
                    FilePathMsiTextBox.Text = _pathToMsi;
                    FilePathMstTextBox.Text = _pathToMst;

                    SetMsiInfo();
                    ShortcutsTextBox.Text = handlers.WindowsInstaller.GetShortcutsCount(_pathToMsi, _pathToMst).ToString();
                    RegistryTextBox.Text = handlers.WindowsInstaller.GetRegistrysCount(_pathToMsi, _pathToMst).ToString();
                    FilesTextBox.Text = handlers.WindowsInstaller.GetFilesCount(_pathToMsi, _pathToMst).ToString();
                }
                else
                {
                    MessageBox.Show("Cannot read MSI. Is MSI database open in another application?");
                    Environment.Exit(11);
                }
            }
        }

        private void CheckExtensions()
        {
            foreach (var argument in Environment.GetCommandLineArgs())
            {
                if (File.Exists(argument))
                {
                    var extension = Path.GetExtension(argument);
                    if (extension == ".msi")
                    {
                        _pathToMsi = argument;
                    }

                    if (extension == ".mst")
                    {
                        _pathToMst = argument;
                        GetMsi(Path.GetDirectoryName(argument));
                    }
                }
            }
        }

        private void GetMsi(string initialDirectory)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Windows Installer File (.msi)|*.msi",
                FilterIndex = 1,
                InitialDirectory = initialDirectory,
                Multiselect = true
            };

            if (openFileDialog.ShowDialog() == true)
            {
                _pathToMsi = openFileDialog.FileName;
            }
            else
            {
                Environment.Exit(12);
            }
        }
    }
}

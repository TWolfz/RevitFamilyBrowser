using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Autodesk.Revit.UI;
using zRevitFamilyBrowser.Revit_Classes;
using System.IO;
using System.Windows.Data;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Linq;
using System.Collections.Specialized;
using TWolfz.Revit;
using System.ComponentModel;

namespace zRevitFamilyBrowser.WPF_Classes
{
    public partial class DockPanel : UserControl, IDockablePaneProvider
    {
        private ExternalEvent m_ExEvent;
        private SingleInstallEvent m_Handler;

        private string temp = string.Empty;
        public string Temp
        {
            get => temp;
            set => temp = value;
        }
        private string collectedData = string.Empty;
        private int ImageListLength = 0;

        private string tempFamilyPath = string.Empty;
        private string tempFamilySymbol = string.Empty;
        private string tempFamilyName = string.Empty;
        private FolderSelect folderSelect = new FolderSelect();
        private void RemoveItem(object sender, RoutedEventArgs e)
        {
            var foldersToRemove = Properties.Settings.Default.FamilyFolderPath.Cast<string>()
                                   .Where(folderPath => folderPath.Contains(label_dfsdsds.SelectedItem.ToString())).First();
            Properties.Settings.Default.FamilyFolderPath.Remove(foldersToRemove);
            Properties.Settings.Default.Save();
            Properties.Settings.Default.Reload();
            List<string> familyNameFolder = new List<string>();
            foreach (string familyPath in Properties.Settings.Default.FamilyFolderPath)
            {
                string folderName = Path.GetFileName(familyPath);
                familyNameFolder.Add(folderName);
            }
            label_dfsdsds.ItemsSource = familyNameFolder;
        }

        public DockPanel(ExternalEvent exEvent, SingleInstallEvent handler)
        {
            InitializeComponent();
            if (Properties.Settings.Default.FamilyFolderPath == null)
            {
                Properties.Settings.Default.FamilyFolderPath = new StringCollection();
                Properties.Settings.Default.Save();
            }

            List<string> familyNameFolder = new List<string>();
            foreach (string familyPath in Properties.Settings.Default.FamilyFolderPath)
            {
                string folderName = Path.GetFileName(familyPath);
                familyNameFolder.Add(folderName);
            }

            //label_dfsdsds.ItemsSource = Properties.Settings.Default.FamilyFolderPath;
            label_dfsdsds.ItemsSource = familyNameFolder;
            m_ExEvent = exEvent;
            m_Handler = handler;

            System.Windows.Threading.DispatcherTimer dispatcherTimer = new System.Windows.Threading.DispatcherTimer();
            dispatcherTimer.Tick += new EventHandler(dispatcherTimer_Tick);
            dispatcherTimer.Interval = new TimeSpan(0, 0, 0, 1);
            dispatcherTimer.Start();

            CreateEmptyFamilyImage();
        }

        public DockPanel()
        {
            InitializeComponent();
        }

        public void SetupDockablePane(DockablePaneProviderData data)
        {
            data.FrameworkElement = this;
            data.InitialState.DockPosition = DockPosition.Left;
        }

        private void dispatcherTimer_Tick(object sender, EventArgs e)
        {
            GenerateGrid();
            GenerateHistoryGrid();
        }

        public void GenerateHistoryGrid()
        {
            string[] ImageList = Directory.GetFiles(System.IO.Path.GetTempPath() + "FamilyBrowser\\");

            if (collectedData != Properties.Settings.Default.CollectedData || ImageList.Length != ImageListLength)
            {
                ImageListLength = ImageList.Length;
                collectedData = Properties.Settings.Default.CollectedData;
                ObservableCollection<FamilyData> collectionData = new ObservableCollection<FamilyData>();
                List<string> listData = new List<string>(collectedData.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries));
                DirectoryInfo di = new DirectoryInfo(System.IO.Path.GetTempPath() + "FamilyBrowser\\RevitLogo.bmp");
                foreach (var item in listData)
                {
                    int index = item.IndexOf('#');
                    string[] symbols = item.Substring(index + 1).Split('#');
                    foreach (var symbol in symbols)
                    {
                        FamilyData projectInstance = new FamilyData();
                        projectInstance.Name = symbol;
                        projectInstance.img = new Uri(di.ToString());

                        try
                        {
                            projectInstance.FamilyName = item.Substring(0, index);
                        }
                        catch (Exception)
                        {
                            projectInstance.FamilyName = "NO FAMILY NAME";
                        }

                        foreach (var imageName in ImageList)
                        {
                            if (imageName.Contains(projectInstance.Name))
                            {
                                projectInstance.img = new Uri(imageName);
                            }
                        }
                        collectionData.Add(projectInstance);
                    }
                }

                collectionData = new ObservableCollection<FamilyData>(collectionData.Reverse());

                foreach (var symbol in collectionData)
                {
                    if (symbol.img == new Uri(di.ToString()))
                        foreach (var item in collectionData)
                        {
                            if (item.FamilyName == symbol.FamilyName && item.img != new Uri(di.ToString()))
                                symbol.img = item.img;
                        }
                }
                ListCollectionView collectionProject = new ListCollectionView(collectionData);
                collectionProject.GroupDescriptions.Add(new PropertyGroupDescription("FamilyName"));
                dataGridHistory.ItemsSource = collectionProject;
            }
        }

        public void GenerateGrid()
        {
            string[] ImageList = Directory.GetFiles(System.IO.Path.GetTempPath() + "FamilyBrowser\\");

            if (temp != Properties.Settings.Default.SymbolList)
            {
                temp = Properties.Settings.Default.SymbolList;
                string category = Properties.Settings.Default.RootFolder;
                label_CategoryName.Content = " " + category.Substring(category.LastIndexOf("\\") + 1);

                List<string> list = new List<string>(temp.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries));
                ObservableCollection<FamilyData> fi = new ObservableCollection<FamilyData>();

                foreach (var item in list)
                {
                    FamilyData instance = new FamilyData();
                    int index = item.IndexOf(' ');
                    instance.Name = item.Substring(0, index);
                    instance.FullName = item.Substring(index + 1);

                    string Name = item.Substring(index + 1);
                    Name = Name.Substring(Name.LastIndexOf("\\") + 1);
                    Name = Name.Substring(0, Name.IndexOf('.'));
                    instance.FamilyName = Name;

                    foreach (var imageName in ImageList)
                    {
                        if (imageName.Contains(instance.Name.TrimEnd()))
                        {
                            instance.img = new Uri(imageName);
                        }
                    }
                    fi.Add(instance);
                }

                //------Collection to sort data in XAML------
                ListCollectionView collection = new ListCollectionView(fi);
                collection.GroupDescriptions?.Add(new PropertyGroupDescription("FamilyName"));
                dataGrid.ItemsSource = collection;
            }
        }

        private void dataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (dataGrid.Items.Count <= 0) return;
            var instance = dataGrid.SelectedItem as FamilyData;
            SetProperty(instance);
            m_ExEvent.Raise();
        }

        private void dataGridHistory_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (dataGridHistory.Items.Count <= 0) return;
            var instance = dataGridHistory.SelectedItem as FamilyData;
            SetHistoryProperty(instance);
            //Properties.Settings.Default.FamilyPath = string.Empty;
            //Properties.Settings.Default.FamilySymbol = instance.Name;
            //Properties.Settings.Default.FamilyName = instance.FamilyName;
            m_ExEvent.Raise();
        }

        private void dataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (dataGrid.Items.Count <= 0) return;
            var instance = dataGrid.SelectedItem as FamilyData;
            SetProperty(instance);
        }

        private void dataGridHistory_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (dataGridHistory.Items.Count <= 0) return;
            var instance = dataGridHistory.SelectedItem as FamilyData;
            SetHistoryProperty(instance);
        }

        private void dataGrid_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (dataGrid.Items.Count <= 0) return;
            var instance = dataGrid.SelectedItem as FamilyData;
            SetProperty(instance);

            tempFamilyPath = instance.FullName;
            tempFamilySymbol = instance.Name;
            tempFamilyName = instance.FamilyName;

            DragDrop.DoDragDrop(dataGrid, instance, DragDropEffects.Copy);
        }

        private void dataGridHistory_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (dataGridHistory.Items.Count <= 0) return;
            var instance = dataGridHistory.SelectedItem as FamilyData;
            SetHistoryProperty(instance);
            SetHistoryTemp(instance);
            DragDrop.DoDragDrop(dataGridHistory, instance, DragDropEffects.Copy);
        }

        private void dataGridHistory_MouseLeave(object sender, MouseEventArgs e)
        {
            if (!(string.IsNullOrEmpty(tempFamilyPath) &&
                  string.IsNullOrEmpty(tempFamilySymbol) &&
                  string.IsNullOrEmpty(tempFamilyName)))
                m_ExEvent.Raise();
        }

        private void dataGrid_MouseLeave(object sender, MouseEventArgs e)
        {
            if (!(string.IsNullOrEmpty(tempFamilyPath) &&
                string.IsNullOrEmpty(tempFamilySymbol) &&
                string.IsNullOrEmpty(tempFamilyName)))
                m_ExEvent.Raise();
        }

        private void dataGrid_MouseEnter(object sender, MouseEventArgs e)
        {
            ClearTemp();
        }

        private void dataGridHistory_MouseEnter(object sender, MouseEventArgs e)
        {
            ClearTemp();
        }

        private void dataGrid_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            ClearTemp();
        }

        private void dataGridHistory_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            ClearTemp();
        }

        private void ClearTemp()
        {
            tempFamilyPath = string.Empty;
            tempFamilySymbol = string.Empty;
            tempFamilyName = string.Empty;
        }

        private void SetHistoryTemp(FamilyData instance)
        {
            tempFamilyPath = string.Empty;
            tempFamilySymbol = instance.Name;
            tempFamilyName = instance.FamilyName;
        }

        private void SetProperty(FamilyData instance)
        {
            Properties.Settings.Default.FamilyPath = instance.FullName;
            Properties.Settings.Default.FamilySymbol = instance.Name;
            Properties.Settings.Default.FamilyName = instance.FamilyName;
        }

        private void SetHistoryProperty(FamilyData instance)
        {
            Properties.Settings.Default.FamilyPath = string.Empty;
            Properties.Settings.Default.FamilySymbol = instance.Name;
            Properties.Settings.Default.FamilyName = instance.FamilyName;
        }


        private void CreateEmptyFamilyImage()
        {
            //TODO optimise creating
            string TempImgFolder = System.IO.Path.GetTempPath() + "FamilyBrowser\\";
            if (!System.IO.Directory.Exists(TempImgFolder))
            {
                System.IO.Directory.CreateDirectory(TempImgFolder);
            }
            ImageConverter converter = new ImageConverter();
            DirectoryInfo di = new DirectoryInfo(System.IO.Path.GetTempPath() + "FamilyBrowser\\RevitLogo.bmp");
            if (!System.IO.File.Exists(di.ToString()))
            {
                File.WriteAllBytes(di.ToString(), (byte[])converter.ConvertTo(Properties.Resources.RevitLogo, typeof(byte[])));
            }
        }

        private void SelectItem(object sender, SelectionChangedEventArgs e)
        {
            // Check if an item is selected
            if (label_dfsdsds.SelectedItem != null)
            {
                Properties.Settings.Default.SymbolList = string.Empty;
                string folderPath = Properties.Settings.Default.FamilyFolderPath.Cast<string>()
                                   .Where(folderName => folderName.Contains(label_dfsdsds.SelectedItem.ToString())).First();
                string[] rfaFiles = Directory.GetFiles(folderPath, "*.rfa");

                foreach (string filePath in rfaFiles)
                {
                    string fileName = Path.GetFileName(filePath);
                    string fileNameWithoutEx = Path.GetFileNameWithoutExtension(filePath);
                    string fullPath = Path.Combine(folderPath, fileName);
                    Properties.Settings.Default.SymbolList += $"{fileNameWithoutEx} {fullPath}" + "\n";
                    Properties.Settings.Default.Save();
                    Properties.Settings.Default.Reload();
                }
            }
        }
        }
}

using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using Tekla.Structures;
using Tekla.Structures.Model;
using TSM = Tekla.Structures.Model;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Xml.Linq;

namespace Jungle_WPF_MultiReport
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        
        TSM.Model model = new TSM.Model();

        

        public ObservableCollection<ReportClass> ListReport { get; set; }

        public List<AttributeClass> ListAttribute { get; set; }

        List<string> AttributesDirectory { get; set; }

        public MainWindow()
        {
            InitializeComponent();
           

            if (!model.GetConnectionStatus())
                MessageBox.Show("Подключиться не удалось");
            else
            {
                ModelInfo modelInfo = model.GetInfo();
                List<string> TemplateDirectory;
                ArrayList ArrayListOfTemplteXLS;

                TemplateDirectory = TemplateDirectoryMethod();
                ListReport = GetCollectionOfTemplteXLS(TemplateDirectory);


                listBoxReport.ItemsSource = ListReport;


                AttributesDirectory = AttributesDirectoryMethod();
                List<AttributeClass> ListAttributes = GetCollectionOfAttributes(AttributesDirectory);

                
                if (ListAttributes.Contains(new AttributeClass("standard_Jungle_MultiReport.xml")))
                {
                    var standardAtributePath = ListAttributes.First(r => r.NameTag == "standard").FullName;
                    load_xml_file(standardAtributePath);
                }
                    

                cm_attributes.ItemsSource = ListAttributes;
            }
            
        }

        private void StackPanel_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            this.DragMove();
        }

        /// <summary>
        /// Возвращает массив путей директорий, где находятся атрибуты
        /// </summary>
        /// <returns></returns>
        public List<string> AttributesDirectoryMethod()
        {
            ModelInfo modelInfo = model.GetInfo();
            var pathAttributesDirectory = modelInfo.ModelPath + "\\attributes\\";
            string tempSystemDirectory = "";
            char[] Delim = { ';' };
            TeklaStructuresSettings.GetAdvancedOption("XS_SYSTEM", ref tempSystemDirectory);
            string SystemDirectory = pathAttributesDirectory + ";" + tempSystemDirectory;
            string[] AllAttributesDirectory = SystemDirectory.Split(Delim, System.StringSplitOptions.RemoveEmptyEntries);
            List<string> AttributesDirectoryPathList = AllAttributesDirectory.ToList();
            return AttributesDirectoryPathList;
        }

        /// <summary>
        /// Возвращает список атрибутов
        /// </summary>
        /// <param name="AttributesDirectory"></param>
        /// <returns></returns>
        public List<AttributeClass> GetCollectionOfAttributes(List<string> AttributesDirectory)
        {
            List<string> collectionAttributesFiles = new List<string>();

            foreach(string AttributeDirectory in AttributesDirectory)
            {
                if(Directory.Exists(AttributeDirectory))
                {
                    string[] files = Directory.GetFiles(AttributeDirectory);
                    
                    List<string> d = files.Where(f => f.Contains("_Jungle_MultiReport.xml")).ToList();
                    collectionAttributesFiles.AddRange(d);
                }
            }

            List<AttributeClass> attributeClasses = new List<AttributeClass>();
            foreach(string AttributeFile in collectionAttributesFiles)
            {                
                AttributeClass fileAttribute = new AttributeClass(AttributeFile);
                if(!attributeClasses.Contains(fileAttribute))
                    attributeClasses.Add(fileAttribute);
            }

            return attributeClasses;
        }

        /// <summary>
        /// Возвращает массив путей директорий, где находятся шаблоны
        /// </summary>
        /// <returns></returns>
        public List<string> TemplateDirectoryMethod()
        {
            
            char[] Delim = { ';' };

            ///Получаем директории из Папки шаблонов (XS_TEMPLATE_DIRECTORY)
            string TemplateDirectory = "";
            TeklaStructuresSettings.GetAdvancedOption("XS_TEMPLATE_DIRECTORY", ref TemplateDirectory);
            string[] TemplateDirectoryPath = TemplateDirectory.Split(Delim, System.StringSplitOptions.RemoveEmptyEntries);
            List<string> TemplateDirectoryPathList = TemplateDirectoryPath.ToList();

            ///Получаем директорию модели
            string ModelPath = model.GetInfo().ModelPath + "\\";


            ///Получаем директории из Папки проекта (XS_PROJECT)
            string ProjectDirectory = "";
            TeklaStructuresSettings.GetAdvancedOption("XS_PROJECT", ref ProjectDirectory);
            string[] ProjectDirectoryPathTemp = ProjectDirectory.Split(Delim, System.StringSplitOptions.RemoveEmptyEntries);
            List<string> ProjectDirectoryPathList = new List<string>();
            foreach(string path in ProjectDirectoryPathTemp)
            {
                ProjectDirectoryPathList.Add(path);
                ProjectDirectoryPathList.Add(path + "templates\\");
                
            }

            ///Получаем директории из Папки компании (XS_FIRM)
            string FirmDirectory = "";
            TeklaStructuresSettings.GetAdvancedOption("XS_FIRM", ref FirmDirectory);
            string[] FirmDirectoryPathTemp = FirmDirectory.Split(Delim, System.StringSplitOptions.RemoveEmptyEntries);
            List<string> FirmDirectoryPathList = new List<string>();
            foreach (string path in FirmDirectoryPathTemp)
            {
                FirmDirectoryPathList.Add(path);
                FirmDirectoryPathList.Add(path + "templates\\");
                
            }

            ///Получаем директории из папки системных шаблонов для данной среды (XS_TEMPLATE_DIRECTORY_SYSTEM)
            string TemplateDirectorySystem = "";
            TeklaStructuresSettings.GetAdvancedOption("XS_TEMPLATE_DIRECTORY_SYSTEM", ref TemplateDirectorySystem);
            string[] TemplateDirectorySystemPathTemp = TemplateDirectorySystem.Split(Delim, System.StringSplitOptions.RemoveEmptyEntries);
            List<string> TemplateDirectorySystemPathList = TemplateDirectorySystemPathTemp.ToList();



            ///Получаем директории из папки системной папки (XS_SYSTEM)
            string SystemDirectory = "";
            TeklaStructuresSettings.GetAdvancedOption("XS_SYSTEM", ref SystemDirectory);
            string[] SystemDirectoryPathTemp = SystemDirectory.Split(Delim, System.StringSplitOptions.RemoveEmptyEntries);
            List<string> SystemDirectoryPathList = SystemDirectoryPathTemp.ToList();
            
            ///Получаем общий список директорий
            List<string> DirAll = new List<string>();
            DirAll.AddRange(TemplateDirectoryPathList);
            DirAll.Add(ModelPath);
            DirAll.AddRange(ProjectDirectoryPathList);
            DirAll.AddRange(FirmDirectoryPathList);
            DirAll.AddRange(TemplateDirectorySystemPathList);
            DirAll.AddRange(SystemDirectoryPathList);

            List<string> TemplateDirectoryAll = DirAll.Distinct().ToList(); 
            return TemplateDirectoryAll;
        }

        /// <summary>
        /// Возвращает коллекцию шаблонов формата xls.rpt
        /// </summary>
        /// <param name="TemplateDirectory"></param>
        /// <returns></returns>
        public ObservableCollection<ReportClass> GetCollectionOfTemplteXLS(List<string> TemplateDirectory)
        {
            ObservableCollection<ReportClass> ArrayListOfTemplteXLS = new ObservableCollection<ReportClass>();
            foreach (string pathCurrentDirectory in TemplateDirectory)
            {

                if (Directory.Exists(pathCurrentDirectory))
                {
                    string[] files = Directory.GetFiles(pathCurrentDirectory);
                    foreach (string s in files)
                    {
                        if (s.Contains("xls.rpt"))
                        {
                            FileInfo file = new FileInfo(s);
                            string name = file.Name;
                            name = name.Replace(".rpt", String.Empty);
                            ArrayListOfTemplteXLS.Add(new ReportClass { Flag=false, Name=name });
                        }
                    }
                }
            }

            ArrayListOfTemplteXLS.Sort();
            return ArrayListOfTemplteXLS;
        }


        /// <summary>
        /// Создает отчет Excel
        /// </summary>
        /// <param name="pathReports"></param>
        /// <param name="pathReportsExcel"></param>
        /// <param name="nameReport"></param>
        public void CreateReportsExcel(string pathReports, string pathReportsExcel, string nameReport)
        {
            Excel.Application ex = new Microsoft.Office.Interop.Excel.Application();
            ex.Visible = false;
            string fileName = pathReports + "\\" + nameReport;
            string fileNameCheck = pathReportsExcel + "\\" + nameReport.Replace(".xls", ".xlsx");
            FileInfo fileCheck = new FileInfo(fileNameCheck);
            if (fileCheck.Exists)
            {
                fileCheck.Delete();
            }
            Excel.Workbook wrkb = ex.Workbooks.Open(fileName);
            //string newPath = pathReportsExcel + "\\" + wrkb.Name.Replace(".xls", string.Empty);
            string newPath = pathReportsExcel + "\\" + wrkb.Name.Replace(".xls", ".xlsx");
            wrkb.SaveAs(Filename: newPath, FileFormat: Excel.XlFileFormat.xlWorkbookDefault, ReadOnlyRecommended: false);
            wrkb.Close();
            wrkb = null;
            ex.Quit();
            ex = null;
            GC.Collect();
        }

        private void btn_select_all_Click(object sender, RoutedEventArgs e)
        {
            foreach(ReportClass report in ListReport)
                report.Flag = true;

            listBoxReport.Items.Refresh();
        }

        private void btn_clear_all_Click(object sender, RoutedEventArgs e)
        {
            foreach (ReportClass report in ListReport)
                report.Flag = false;
            listBoxReport.Items.Refresh();
        }

        

        private void btn_minimize_Click(object sender, RoutedEventArgs e)
        {
            this.WindowState = WindowState.Minimized;
        }

        private void btn_maximize_Click(object sender, RoutedEventArgs e)
        {
            
            if (this.WindowState == WindowState.Maximized)
            {
                this.WindowState = WindowState.Normal;
                maximizeImage.Source = new BitmapImage(new Uri("pack://application:,,,/Image/maximize.png"));
            }                
            else
            {
                this.WindowState = WindowState.Maximized;
                maximizeImage.Source = new BitmapImage(new Uri("pack://application:,,,/Image/normalize.png"));
            }
                
        }

        

        private void btn_close_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void btn_create_report_Click(object sender, RoutedEventArgs e)
        {

            ObservableCollection<ReportClass> reportAll = listBoxReport.ItemsSource as ObservableCollection<ReportClass>;
            
            List<string> reportTrueName = reportAll.Where(report=>report.Flag == true).Select(report=>report.Name).ToList();



            TSM.ModelInfo modelInfo = model.GetInfo();
            var pathReports = modelInfo.ModelPath + "\\Reports";
            var pathReportsExcel = pathReports + "\\Excel";


            try
            {
                DirectoryInfo dirInfo = new DirectoryInfo(pathReports);
                DirectoryInfo dirInfoReportExcel = new DirectoryInfo(pathReportsExcel);

                if (!dirInfo.Exists)
                {
                    dirInfo.Create();

                }
                if (!dirInfoReportExcel.Exists)
                {
                    dirInfoReportExcel.Create();
                }


                ///Определение префикса и постфикса
                string prefix;
                if (tb_prefix.Text == String.Empty)
                    prefix = string.Empty;
                else prefix = tb_prefix.Text + "_";

                string postfix;
                if (tb_postfix.Text == String.Empty)
                    postfix = string.Empty;
                else postfix = "_" + tb_postfix.Text;

                ///Создание отчетов
                if((bool)radioAll.IsChecked)
                {
                    foreach(string name in reportTrueName)
                    {
                        string nameReport = prefix + name.Remove(name.Length-4) + postfix + ".xls";
                        TSM.Operations.Operation.CreateReportFromAll(name, nameReport, "", "", "");

                        if((bool)radioExcel.IsChecked)
                        {
                            CreateReportsExcel(pathReports: pathReports, pathReportsExcel: pathReportsExcel, nameReport: nameReport);
                        }
                    }
                }
                else
                {
                    foreach (string name in reportTrueName)
                    {
                        string nameReport = prefix + name.Remove(name.Length - 4) + postfix + ".xls";
                        TSM.Operations.Operation.CreateReportFromSelected(name, nameReport, "", "", "");
                        if ((bool)radioExcel.IsChecked)
                        {
                            CreateReportsExcel(pathReports: pathReports, pathReportsExcel: pathReportsExcel, nameReport: nameReport);
                        }
                    }
                }



            }
            catch (System.Runtime.InteropServices.COMException)
            {
                MessageBox.Show("Excel не установлен");
            }
            catch
            {
                MessageBox.Show("Исключение!");
            }

        }
        


        private void btn_Support_Click(object sender, RoutedEventArgs e)
        {
            Process.Start("https://support.tekla.com/ru/tekla-structures");
        }

        private void btn_save_as_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(tb_save_as.Text))
            {
                string nameAttributeReport = tb_save_as.Text + "_Jungle_MultiReport.xml";
                ModelInfo modelInfo = model.GetInfo();
                var pathReports = modelInfo.ModelPath + "\\attributes";
                var pathAttributeReport = pathReports + "\\" + nameAttributeReport;                               
                            
                save_xml_file(pathAttributeReport);
                cm_attributes.ItemsSource = GetCollectionOfAttributes(AttributesDirectory);
            }
        }

        

        private void btn_save_Click(object sender, RoutedEventArgs e)
        {
            if (!(cm_attributes.SelectedItem is null))
            {
                var nameSelected = ((AttributeClass)cm_attributes.SelectedItem).NameTag;
                string nameAttributeReport = nameSelected + "_Jungle_MultiReport.xml";
                var pathReports = model.GetInfo().ModelPath + "\\attributes";
                var pathAttributeReport = pathReports + "\\" + nameAttributeReport;              
                
                save_xml_file(pathAttributeReport);
                cm_attributes.ItemsSource = GetCollectionOfAttributes(AttributesDirectory);
            }    
           
        }

        private void btn_load_Click(object sender, RoutedEventArgs e)
        {
            if (!(cm_attributes.SelectedItem is null))
            {
                foreach (ReportClass report in ListReport)
                    report.Flag = false;
                

                var nameSelected = ((AttributeClass)cm_attributes.SelectedItem).FullName;

                load_xml_file(nameSelected);
            }
        }


        /// <summary>
        /// Сохраняет настройки приложения в файл xml
        /// </summary>
        /// <param name="path"></param>
        public void save_xml_file(string path)
        {
            XDocument xDoc = new XDocument();

            XElement radioSelectNode;
            XElement radioWebNode;
            XElement prefixNode;
            XElement postfixNode;
            XElement reportsNode = new XElement("reports");

            if ((bool)radioSelect.IsChecked)
                radioSelectNode = new XElement("radioSelect", "1");
            else
                radioSelectNode = new XElement("radioSelect", "0");

            if ((bool)radioWeb.IsChecked)
                radioWebNode = new XElement("radioWeb", "1");
            else
                radioWebNode = new XElement("radioWeb", "0");

            if (tb_prefix.Text == String.Empty)
                prefixNode = new XElement("prefix", string.Empty);
            else
                prefixNode = new XElement("prefix", tb_prefix.Text);


            if (tb_postfix.Text == String.Empty)
                postfixNode = new XElement("postfix", string.Empty);
            else
                postfixNode = new XElement("postfix", tb_postfix.Text);


            XElement settingNode = new XElement("setting");
            settingNode.Add(radioSelectNode);
            settingNode.Add(radioWebNode);
            settingNode.Add(prefixNode);
            settingNode.Add(postfixNode);


            ObservableCollection<ReportClass> reportAll = listBoxReport.ItemsSource as ObservableCollection<ReportClass>;
            List<string> reportTrueNameList = reportAll.Where(report => report.Flag == true).Select(report => report.Name).ToList();
            foreach (string report in reportTrueNameList)
            {
                XElement reportNode = new XElement("report", report);
                reportsNode.Add(reportNode);
            }
            settingNode.Add(reportsNode);
            xDoc.Add(settingNode);
            xDoc.Save(path);


        }



        public void load_xml_file(string path)
        {
            XDocument xDoc = XDocument.Load(path);

            var radioSelect_query = xDoc.Element("setting").Element("radioSelect").Value;
            if (radioSelect_query == "1")
            {
                radioAll.IsChecked = false;
                radioSelect.IsChecked = true;
            }
            else
            {
                radioAll.IsChecked = true;
                radioSelect.IsChecked = false;
            }

            var radioWeb_query = xDoc.Element("setting").Element("radioWeb").Value;
            if (radioWeb_query == "1")
            {
                radioWeb.IsChecked = true;
                radioExcel.IsChecked = false;
            }
            else
            {
                radioWeb.IsChecked = false;
                radioExcel.IsChecked = true;
            }

            var prefix_query = xDoc.Element("setting").Element("prefix").Value;
            tb_prefix.Text = prefix_query;

            var postfix_query = xDoc.Element("setting").Element("postfix").Value;
            tb_postfix.Text = postfix_query;


            var reports_query = xDoc.Element("setting").Element("reports").Elements("report");
            
            List<string> reportsNameList = new List<string>();
            foreach (XElement reportNode in reports_query)
            {
                string file = reportNode.Value;
                ObservableCollection<ReportClass> reportAll = listBoxReport.ItemsSource as ObservableCollection<ReportClass>;
                ReportClass reportTrueName = reportAll.Where(report => report.Name == file).First();
                reportTrueName.Flag = true;
            }

            listBoxReport.Items.Refresh();
            cm_attributes.Items.Refresh();
        }

        
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using System.Windows.Input;
using System.Xml.Linq;
using ThirdPartyLibSearcher.Annotations;

namespace ThirdPartyLibSearcher
{

    internal class ViewModel : INotifyPropertyChanged
    {
        private string _installDir;

        public string InstallDirText
        {
            get { return _installDir; }
            set
            {
                _installDir = value;
                OnPropertyChanged();
            }
        }

        public ViewModel()
        {
            Generate = new SimpleDelegateCommand(DoGenerate);
            SelectDir = new SimpleDelegateCommand(DoSelectDir);
        }

        public ICommand Generate { get; private set; }

        public ICommand SelectDir { get; private set; }

        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }

        private void DoGenerate()
        {
            if (!Directory.Exists(InstallDirText))
                MessageBox.Show(String.Format("No such directory: {0}", InstallDirText));
            else
            {
                var assemblyList = Directory.GetFiles(InstallDirText, "*.*", SearchOption.AllDirectories).
                    Where(s => s.EndsWith(".dll") || s.EndsWith(".exe")).
                    Where(s => !Path.GetFileName(s).Contains("JetBrains"));
                var assemblyDict = new SortedDictionary<string, string>();
                foreach (var assPath in assemblyList)
                {
                    if (!assemblyDict.ContainsKey(Path.GetFileName(assPath)))
                        assemblyDict.Add(Path.GetFileName(assPath), assPath);
                }

                using (var brwsr = new FolderBrowserDialog() {Description = "Choose where to save the output"})
                {
                    const string thirdPartyFileName = "Third_Party_Libs";
                    if (brwsr.ShowDialog() == DialogResult.Cancel) return;


                    string saveDirectoryPath = brwsr.SelectedPath;
                    string fileName = Path.Combine(saveDirectoryPath, thirdPartyFileName + ".xml");

                    var thirdPartyTopic = new XDocument();
                    var thirdPartyRootElement = new XElement("ThirdPartyLibs");

                    foreach (var assPath in assemblyDict)
                    {
                        var assemblyPath = assPath.Value;
                        var assemblyFileName = Path.GetFileName(assemblyPath);
                        
                        var info = FileVersionInfo.GetVersionInfo(assemblyPath);

                        var assemblyTitle = Path.GetFileNameWithoutExtension(assemblyPath);
                        var version = info.FileVersion ?? "Unknown";
                        var companyName = info.CompanyName ?? "Unknown";
                        var copyright = info.LegalCopyright ?? "Unknown";
                        var description = info.FileDescription ?? "Unknown";
                        var product = info.ProductName ?? "Unknown";

                        try
                        {
                            Assembly assembly = Assembly.LoadFrom(assemblyPath);
                            object[] titleAttribs = assembly.GetCustomAttributes(typeof(AssemblyTitleAttribute), true);
                            if (titleAttribs.Length > 0)
                            {
                                assemblyTitle = ((AssemblyTitleAttribute)titleAttribs[0]).Title;
                            } 
                            object[] companyAttribs = assembly.GetCustomAttributes(typeof (AssemblyCompanyAttribute), true);
                            if (companyAttribs.Length > 0)
                            {
                                companyName = ((AssemblyCompanyAttribute)companyAttribs[0]).Company;
                            }
                            object[] versionAttribs = assembly.GetCustomAttributes(typeof(AssemblyVersionAttribute), true);
                            if (companyAttribs.Length > 0)
                            {
                                version = ((AssemblyVersionAttribute) versionAttribs[0]).Version;
                            }
                            object[] copyrightAttribs = assembly.GetCustomAttributes(typeof(AssemblyCopyrightAttribute), true);
                            if (companyAttribs.Length > 0)
                            {
                                copyright = ((AssemblyCopyrightAttribute)copyrightAttribs[0]).Copyright;
                            }
                            object[] descriptionAttribs = assembly.GetCustomAttributes(typeof(AssemblyDescriptionAttribute), true);
                            if (companyAttribs.Length > 0)
                            {
                                description = ((AssemblyDescriptionAttribute)descriptionAttribs[0]).Description;
                            }
                            object[] productAttribs = assembly.GetCustomAttributes(typeof(AssemblyProductAttribute), true);
                            if (companyAttribs.Length > 0)
                            {
                                product = ((AssemblyProductAttribute)productAttribs[0]).Product;
                            }

                        }
                        catch (Exception ex)
                        {
                           
                        }
                        if (assemblyFileName.Contains("JetBrains") || assemblyTitle.Contains("JetBrains") || copyright.Contains("JetBrains")) continue;

                        thirdPartyRootElement.Add(new XElement("Assembly",
                            new XAttribute("Title", assemblyTitle),
                            new XAttribute("Version", version),
                            new XElement("Description", description),
                            new XElement("AssemblyFile", assemblyFileName),
                            new XElement("Company", companyName),
                            new XElement("Copyright", copyright),
                            new XElement("Product", product)
                            ));
                    }

                    thirdPartyTopic.Add(thirdPartyRootElement);
                    thirdPartyTopic.Save(fileName);
                    MessageBox.Show("Third-party assemblies are successfully exported to " + thirdPartyFileName + ".xml");
                }
            }
        }

        private void DoSelectDir()
        {
            using (var brwsr = new FolderBrowserDialog() {Description = "Select installation directory of the product"})
            {
                if (brwsr.ShowDialog() == DialogResult.Cancel) return;
                _installDir = brwsr.SelectedPath;
            }
        }
    }


    internal class SimpleDelegateCommand : ICommand
    {
        private readonly Action _action;

        public SimpleDelegateCommand(Action action)
        {
            _action = action;
        }

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            _action();
        }

        public event EventHandler CanExecuteChanged;
    }
}

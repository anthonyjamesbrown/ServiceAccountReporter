using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Threading;
using System.Windows.Controls;
using System.Windows.Data;
using System.DirectoryServices;
using System.ComponentModel;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System.Globalization;

namespace ServiceAccountReporter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    /// Add comment
    public partial class MainWindow : System.Windows.Window
    {
        private List<ServiceAccount> accounts = new List<ServiceAccount>();

        public MainWindow()
        {
            InitializeComponent();
            TextBox1.IsReadOnly = true;
        }

        public void DoEvents()
        {
            var frame = new DispatcherFrame();
            Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Background, new DispatcherOperationCallback(ExitFrame), frame);
            Dispatcher.PushFrame(frame);
        }

        public object ExitFrame(Object f)
        {
            ((DispatcherFrame)f).Continue = false;
            return null;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            bttnQueryAD.IsEnabled = false;
            lblStatus.Content = $"Quering AD...";

            dataGridAccounts.Items.Clear();

            DoEvents();

            GetServiceAccounts();
            
            TextBox1.IsReadOnly = false;
        }

        public void GetServiceAccounts()
        {
            string[] domainList = new string[] 
            {
                "LDAP://DC=int,DC=asurion,DC=com",
                "LDAP://DC=newcorp,DC=com",
                "LDAP://DC=icmasu,DC=icm",
                "LDAP://DC=asurion,DC=org"
            };

            string[] propList = new string[]
            {
                "cn",
                "sAMAccountName",
                "employeeType",
                "msExchRecipientTypeDetails",
                "description",
                "adminDescription",
                "memberOf",
                "DistinguishedName",
                "whenCreated",
                "userAccountControl",
                "pwdLastSet",
                "lastLogonTimeStamp"
            };

            foreach (string domainPath in domainList)
            {
                SearchResultCollection searchResultCollection;

                var domain = GetDomain(domainPath);
                var directoryEntry = new DirectoryEntry(domainPath);
                var ldapFilter = "(&(objectClass=user)(employeeType=service))";

                var search = new DirectorySearcher(directoryEntry)
                {
                    Filter = ldapFilter,
                    PageSize = 5000
                };

                search.PropertiesToLoad.AddRange(propList);

                using (searchResultCollection = search.FindAll())
                {
                    foreach (SearchResult item in searchResultCollection)
                    {
                        ResultPropertyCollection resultPropColl = item.Properties;

                        ResultPropertyValueCollection accountName = resultPropColl["sAMAccountName"];
                        ResultPropertyValueCollection employeeType = resultPropColl["employeeType"];
                        ResultPropertyValueCollection recipientTypeDetails = resultPropColl["msExchRecipientTypeDetails"];
                        ResultPropertyValueCollection description = resultPropColl["description"];
                        ResultPropertyValueCollection adminDescription = resultPropColl["adminDescription"];
                        ResultPropertyValueCollection memberOf = resultPropColl["memberOf"];
                        ResultPropertyValueCollection dn = resultPropColl["DistinguishedName"];
                        ResultPropertyValueCollection created = resultPropColl["whenCreated"];
                        ResultPropertyValueCollection pwLastSet = resultPropColl["pwdLastSet"];
                        ResultPropertyValueCollection uAC = resultPropColl["userAccountControl"];
                        ResultPropertyValueCollection lastLogonTimeStamp = resultPropColl["lastLogonTimeStamp"];

                        var allowLogon = IsInGroup(memberOf, "Allow Log On");
                        var denyLogon = IsInGroup(memberOf, "Deny Log On");

                        int uACvalue = Convert.ToInt32(uAC[0].ToString());

                        var createString = created[0].ToString();                       
                        var lastLogonTimeStampString = lastLogonTimeStamp.Count > 0 ? lastLogonTimeStamp[0].ToString() : "";

                        long pwDateString = Convert.ToInt64(pwLastSet[0].ToString());

                        DateTime createDate = DateTime.ParseExact(createString, "G", CultureInfo.CurrentCulture);
                        DateTime pwDate = DateTime.FromFileTimeUtc(pwDateString);


                        if (!String.IsNullOrWhiteSpace(lastLogonTimeStampString))
                        {
                            lastLogonTimeStampString = DateTime.FromFileTimeUtc(Convert.ToInt64(lastLogonTimeStampString)).ToString();
                        }

                        var interactiveStatus = "";
                        if(allowLogon && denyLogon)
                        {
                            interactiveStatus = "both";
                        }
                        else if (allowLogon && !denyLogon)
                        {
                            interactiveStatus = "allow";
                        }
                        else if (!allowLogon && denyLogon)
                        {
                            interactiveStatus = "deny";
                        }
                        else if (!allowLogon && !denyLogon)
                        {
                            interactiveStatus = "neither";
                        }

                        string recipType = (recipientTypeDetails.Count > 0) ? recipientTypeDetails[0].ToString() : "";
                        string desc = (description.Count > 0) ? description[0].ToString() : "";
                        string securityApproval = (adminDescription.Count > 0) ? adminDescription[0].ToString() : "false";

                        if (!String.IsNullOrEmpty(accountName[0].ToString()))
                        {
                            accounts.Add(new ServiceAccount()
                            {
                                Name = accountName[0].ToString(),
                                EmployeeType = employeeType[0].ToString(),
                                RecipientTypeDetails = recipType,
                                InteractiveStatus = interactiveStatus,
                                SecurityApproval = securityApproval,
                                Description = desc,                                
                                LockedOut = (uACvalue & 16) == 16 ? true : false,
                                PasswordExpired = (uACvalue & 8388608) == 8388608 ? true : false,
                                PasswordNeverExpires = (uACvalue & 65536) == 65536 ? true : false,
                                PasswordNotRequired = (uACvalue & 32) == 32 ? true : false,
                                Created = createDate,
                                PasswordLastSet = pwDate,
                                LastLogonTimeStamp = lastLogonTimeStampString,
                                Domain = domain,
                                DN = dn[0].ToString()
                            });
                        }                       
                    }
                }
                ICollectionView cvAccounts = CollectionViewSource.GetDefaultView(accounts);

                if (cvAccounts != null)
                {
                    dataGridAccounts.AutoGenerateColumns = true;
                    dataGridAccounts.ItemsSource = cvAccounts;
                    cvAccounts.Filter = TextFilter;
                    lblStatus.Content = $"Accounts returned: {dataGridAccounts.Items.Count}";
                    BttnExport.IsEnabled = true;
                }
            }
        }

        public bool TextFilter(object o)
        {
            ServiceAccount p = (o as ServiceAccount);
            if (p == null)
            {
                return false;
            }

            if (p.Name.ToLower().Contains(TextBox1.Text.ToLower()) || 
                p.InteractiveStatus.ToLower().Contains(TextBox1.Text.ToLower()) ||
                p.Description.ToLower().Contains(TextBox1.Text.ToLower()) ||
                p.SecurityApproval.ToLower().Contains(TextBox1.Text.ToLower()) ||
                p.Domain.ToLower().Contains(TextBox1.Text.ToLower())
                )
            {
                return true;
            }               
            else
            {
                return false;
            }                
        }

        private void TextBox1_TextChanged(object sender, TextChangedEventArgs e)
        {
            ICollectionView cvAccounts = CollectionViewSource.GetDefaultView(accounts);

            if (cvAccounts != null)
            {
                dataGridAccounts.AutoGenerateColumns = true;
                dataGridAccounts.ItemsSource = cvAccounts;
                cvAccounts.Filter = TextFilter;
                lblStatus.Content = $"Accounts returned: {dataGridAccounts.Items.Count}";
                BttnExport.IsEnabled = true;
            }
        }

        public static bool IsInGroup(ResultPropertyValueCollection memberOf, string groupName)
        {
            if (memberOf.Count > 0 & !String.IsNullOrWhiteSpace(groupName))
            {
                foreach(String group in memberOf)
                {
                    if(group.ToString().Contains(groupName))
                    {
                        return true;
                    }
                }
                return false;
            }
            else
            {
                return false;
            }
        }

        public string GetDomain(string ldapPath)
        {
            if (ldapPath == "LDAP://DC=int,DC=asurion,DC=com")
            {
                return "HQDomain";
            }
            else if (ldapPath == "LDAP://DC=newcorp,DC=com")
            {
                return "NEW_STERLING";
            }
            else if (ldapPath == "LDAP://DC=icmasu,DC=icm")
            {
                return "ICMASU";
            }
            else if (ldapPath == "LDAP://DC=asurion,DC=org")
            {
                return "ASURION";
            }
            return "unknown";
        }

        private void BttnExport_Click(object sender, RoutedEventArgs e)
        {
            if(dataGridAccounts.Items.Count > 0)
            {
                var misValue = System.Reflection.Missing.Value;
                var sPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                var xlApp = new Microsoft.Office.Interop.Excel.Application();
                Workbook xlWorkBook;
                Worksheet xlWorkSheet;
                var currentDate = DateTime.Now;
                var fileName = $"Server Accounts {currentDate.ToShortDateString().Replace('/', '-')}";

                xlWorkBook = xlApp.Workbooks.Add(misValue);
                //xlWorkSheet = xlWorkBook.Sheets.Add("sheet1");
                xlWorkSheet = (Worksheet)xlWorkBook.Sheets[1];
                xlWorkSheet.Name = "Service Accounts";

                xlWorkSheet.Range["A1:O1"].Font.Bold = true;
                xlWorkSheet.Range["A1:O1"].Font.Size = 10;
                xlWorkSheet.Range["A1:O1"].RowHeight = 15;

                xlWorkSheet.Cells[1, 1] = "Account Name";
                xlWorkSheet.Cells[1, 2] = "Employee Type";
                xlWorkSheet.Cells[1, 3] = "Recipient Type";
                xlWorkSheet.Cells[1, 4] = "Description";
                xlWorkSheet.Cells[1, 5] = "Interactive Status";
                xlWorkSheet.Cells[1, 6] = "Security Approval";
                xlWorkSheet.Cells[1, 7] = "Locked Out";
                xlWorkSheet.Cells[1, 8] = "Password Expired";
                xlWorkSheet.Cells[1, 9] = "Password Never Expires";
                xlWorkSheet.Cells[1, 10] = "Password Not Required";
                xlWorkSheet.Cells[1, 11] = "Created On";
                xlWorkSheet.Cells[1, 12] = "Password Last Set";
                xlWorkSheet.Cells[1, 13] = "LastLogonTimeStamp";
                xlWorkSheet.Cells[1, 14] = "Domain";
                xlWorkSheet.Cells[1, 15] = "DN";

                xlWorkSheet.Columns["A"].ColumnWidth = 35;
                xlWorkSheet.Columns["B:C"].ColumnWidth = 15;
                xlWorkSheet.Columns["D:N"].ColumnWidth = 25;
                xlWorkSheet.Columns["D"].ColumnWidth = 50;
                xlWorkSheet.Columns["O"].ColumnWidth = 120;

                int row = -1;
                Type type = null;

                int rows = dataGridAccounts.Items.Count;
                int cols = dataGridAccounts.Columns.Count;

                var arr = new object[rows, cols];

                foreach(ServiceAccount sa in dataGridAccounts.ItemsSource)
                {
                    ++row;
                    type = sa.GetType();
                    for(int col = 0; col < cols; ++col)
                    {
                        var column_name = (string)dataGridAccounts.Columns[col].Header;
                        var value = type.GetProperty(column_name).GetValue(sa).ToString();
                        arr.SetValue(value, row, col);
                    }
                }
                xlWorkSheet.Range["A2"].Resize[rows, cols].Value = arr;

                xlWorkSheet.EnableAutoFilter = true;
                xlWorkSheet.Range[$"A1:O{dataGridAccounts.Items.Count}"].AutoFilter(Field:1, Operator:XlAutoFilterOperator.xlFilterValues);

                var dlg = new SaveFileDialog() {
                    Filter = "Excel Files (*.xlsx)|*.xlsx",
                    FilterIndex = 1,
                    InitialDirectory = sPath,
                    FileName = fileName
                };
                var excelFile = "";
                Nullable<bool> result = dlg.ShowDialog();

                if (result == true)
                {
                    excelFile = dlg.FileName;
                    xlWorkSheet.SaveAs(excelFile);
                }
                xlWorkBook.Close();
                xlApp.Quit();
            }
        }
    }

    public class ServiceAccount
    {
        public string Name { get; set; }
        public string EmployeeType { get; set; }
        public string RecipientTypeDetails { get; set; }
        public string Description { get; set; }
        public string InteractiveStatus { get; set; }
        public string SecurityApproval { get; set; }
        public bool LockedOut { get; set; }
        public bool PasswordExpired { get; set; }
        public bool PasswordNeverExpires { get; set; }
        public bool PasswordNotRequired { get; set; }
        public DateTime Created { get; set; }
        public DateTime PasswordLastSet { get; set; }
        public string LastLogonTimeStamp { get; set; }
        public string Domain { get; set; }
        public string DN { get; set; }
    }
}

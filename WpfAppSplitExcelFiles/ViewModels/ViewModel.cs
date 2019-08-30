using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

// -- https://www.add-in-express.com/creating-addins-blog/2013/11/05/release-excel-com-objects/ --

namespace SplitExcelFiles
{
    public class ViewModel : ViewModelBase
    {
        #region --   -- IsFileNameEnabled
        private Workbook sourceWorkbook;
        private Worksheet sourceWorksheet;
        Workbook tempWorkbook;
        Worksheet tempWorksheet;
        private object[,] rowsArray;

        private int rowCount = 0;
        private int counter = 1;
        string[] rows;
        private object[,] valueArray;
        private int dataRowIndex = 0;
        private int ColKeyIndex = 0;
        private int EmailColIndex = 0;
        Dictionary<String, RowRange> allRowsId = new Dictionary<string, RowRange>();
        #endregion

        #region - EXCEL APPLICATION -
        // -- Demarage de l'application Excel --
        Application _excelApp;
        #endregion

        #region --  Attributs / Properties  -- 
        #region --  Double Attributs / Properties  --
        private double _progressBarValue;
        public double ProgressBarValue
        {
            get { return _progressBarValue; }
            set { _progressBarValue = value; NotifyPropertyChanged("ProgressBarValue"); }
        }
        #endregion

        #region --  Int Attributs / Properties  --
        private string _mailBody;
        public string MailBody
        {
            get { return _mailBody; }
            set { _mailBody = value; NotifyPropertyChanged("MailBody"); }
        }

        private int _totalFiles;
        public int TotalFiles
        {
            get { return _totalFiles; }
            set { _totalFiles = value; NotifyPropertyChanged("TotalFiles"); }
        }
        #endregion

        #region --  String Attributs / Properties  --
        private string _fileName;
        public string FileName
        {
            get { return _fileName; }
            set
            {
                _fileName = value;
                if (!string.IsNullOrWhiteSpace(FirstCell) && !string.IsNullOrWhiteSpace(OutputFolder) && !string.IsNullOrWhiteSpace(FileName))
                    IsSplitCmdVisible = true;
                else
                    IsSplitCmdVisible = false;
                NotifyPropertyChanged("FileName");
            }
        }

        private string _keyColName;
        public string KeyColName
        {
            get { return _keyColName; }
            set { _keyColName = value; NotifyPropertyChanged("KeyColName"); }
        }

        private string _sheetName;
        public string SheetName
        {
            get { return _sheetName; }
            set
            {
                _sheetName = value; NotifyPropertyChanged("SheetName");

                // -- Appel de la fonction qui récupère le nom des colomnnes --
                GetSheetColumnName();
            }
        }

        private string _otherFileName;
        public string OtherFileName
        {
            get { return _otherFileName; }
            set { _otherFileName = value; NotifyPropertyChanged("OtherFileName"); }
        }

        private string _filePath;
        public string FilePath
        {
            get { return _filePath; }
            set { _filePath = value; NotifyPropertyChanged("FilePath"); }
        }

        private string _prefixe;
        public string Prefixe
        {
            get { return _prefixe; }
            set { _prefixe = value; NotifyPropertyChanged("Prefixe"); }
        }

        private string _suffixe;
        public string Suffixe
        {
            get { return _suffixe; }
            set { _suffixe = value; NotifyPropertyChanged("Suffixe"); }
        }

        private string _selExtension;
        public string SelExtension
        {
            get { return _selExtension; }
            set { _selExtension = value; NotifyPropertyChanged("SelExtension"); }
        }

        private string _outputFolder;
        public string OutputFolder
        {
            get { return _outputFolder; }
            set
            {
                _outputFolder = value;
                if (!string.IsNullOrWhiteSpace(FirstCell) && !string.IsNullOrWhiteSpace(OutputFolder) && !string.IsNullOrWhiteSpace(FileName))
                    IsSplitCmdVisible = true;
                else
                    IsSplitCmdVisible = false;
                NotifyPropertyChanged("OutputFolder");
            }
        }

        private string _senderMail;
        public string SenderMail
        {
            get { return _senderMail; }
            set
            {
                _senderMail = value;
                // --    --
                CheckNotEmptyEmailFeel();
                NotifyPropertyChanged("SenderMail");
            }
        }

        private string _senderName;
        public string SenderName
        {
            get { return _senderName; }
            set
            {
                _senderName = value;
                // --    --
                CheckNotEmptyEmailFeel();
                NotifyPropertyChanged("SenderName");
            }
        }

        private string _object;
        public string Object
        {
            get { return _object; }
            set
            {
                _object = value;
                // --    --
                CheckNotEmptyEmailFeel();
                NotifyPropertyChanged("Object");
            }
        }

        private string _ccMail;
        public string CcMail
        {
            get { return _ccMail; }
            set { _ccMail = value; NotifyPropertyChanged("CcMail"); }
        }

        private string _bccMail;
        public string BccMail
        {
            get { return _bccMail; }
            set { _bccMail = value; NotifyPropertyChanged("BccMail"); }
        }

        private string _recipientEmail;
        public string RecipientEmail
        {
            get { return _recipientEmail; }
            set { _recipientEmail = value; NotifyPropertyChanged("RecipientEmail"); }
        }

        private string _firstCell;
        public string FirstCell
        {
            get { return _firstCell; }
            set
            {
                _firstCell = value;
                if (string.IsNullOrWhiteSpace(value))
                    DisplayErrorMessage("Vous devez indiquer la celulle de debut des données !");
                if (!string.IsNullOrWhiteSpace(FirstCell) && !string.IsNullOrWhiteSpace(OutputFolder) && !string.IsNullOrWhiteSpace(FileName))
                    IsSplitCmdVisible = true;
                else
                    IsSplitCmdVisible = false;
                NotifyPropertyChanged("FirstCell");
            }
        }

        private string _mailBodyPath;
        public string MailBodyPath
        {
            get { return _mailBodyPath; }
            set
            {
                _mailBodyPath = value;
                // --    --
                CheckNotEmptyEmailFeel();
                NotifyPropertyChanged("mailBodyPath");
            }
        }

        #endregion

        #region --  IEnnumerable Attributs / Properties  --
        private IEnumerable<string> _sheetNames;
        public IEnumerable<string> SheetNames
        {
            get { return _sheetNames; }
            set { _sheetNames = value; NotifyPropertyChanged("SheetNames"); }
        }

        private IEnumerable<string> _keyColNames;
        public IEnumerable<string> KeyColNames
        {
            get { return _keyColNames; }
            set { _keyColNames = value; NotifyPropertyChanged("KeyColNames"); }
        }

        private IEnumerable<string> _allRecipientsEmails;
        public IEnumerable<string> AllRecipientsEmails
        {
            get { return _allRecipientsEmails; }
            set { _allRecipientsEmails = value; NotifyPropertyChanged("AllRecipientsEmails"); }
        }
        #endregion

        private List<String> _extensions = new List<string>() { ".xlsm", ".xlsx", ".xls", ".csv", ".xlsb" };
        public List<String> Extensions
        {
            get { return _extensions; }
            set { _extensions = value; NotifyPropertyChanged("Extensions"); }
        }
        #endregion

        #region --  Visibility attributes / properties  --                     
        private bool _isSendByMailCmdVisible;
        public bool IsSendByMailCmdVisible
        {
            get { return _isSendByMailCmdVisible; }
            set { _isSendByMailCmdVisible = value; NotifyPropertyChanged("IsSendByMailCmdVisible"); }
        }

        private bool _isSendPreviewCmdVisible;
        public bool IsSendPreviewCmdVisible
        {
            get { return _isSendPreviewCmdVisible; }
            set { _isSendPreviewCmdVisible = value; NotifyPropertyChanged("IsSendPreviewCmdVisible"); }
        }

        private bool _isBackToFirstFormCmdVisible;
        public bool IsBackToFirstFormCmdVisible
        {
            get { return _isBackToFirstFormCmdVisible; }
            set { _isBackToFirstFormCmdVisible = value; NotifyPropertyChanged("IsBackToFirstFormCmdVisible"); }
        }

        private bool _isBackToSplitFormCmdVisible;
        public bool IsBackToSplitFormCmdVisible
        {
            get { return _isBackToSplitFormCmdVisible; }
            set { _isBackToSplitFormCmdVisible = value; NotifyPropertyChanged("IsBackToSplitFormCmdVisible"); }
        }

        private bool _isLoadingFormVisible;
        public bool IsLoadingFormVisible
        {
            get { return _isLoadingFormVisible; }
            set { _isLoadingFormVisible = value; NotifyPropertyChanged("IsLoadingFormVisible"); }
        }

        private bool _isSplitFormVisible;
        public bool IsSplitFormVisible
        {
            get { return _isSplitFormVisible; }
            set { _isSplitFormVisible = value; NotifyPropertyChanged("IsSplitFormVisible"); }
        }

        private bool _isSendMailFormVisible;
        public bool IsSendMailFormVisible
        {
            get { return _isSendMailFormVisible; }
            set { _isSendMailFormVisible = value; NotifyPropertyChanged("IsSendMailFormVisible"); }
        }

        private bool _isSplitCmdVisible;
        public bool IsSplitCmdVisible
        {
            get { return _isSplitCmdVisible; }
            set { _isSplitCmdVisible = value; NotifyPropertyChanged("IsSplitCmdVisible"); }
        }

        private bool _isFirstFormVisible;
        public bool IsFirstFormVisible
        {
            get { return _isFirstFormVisible; }
            set { _isFirstFormVisible = value; NotifyPropertyChanged("IsFirstFormVisible"); }
        }

        private bool _isShowSendMailFormCmdVisible;
        public bool IsShowSendMailFormCmdVisible
        {
            get { return _isShowSendMailFormCmdVisible; }
            set { _isShowSendMailFormCmdVisible = value; NotifyPropertyChanged("IsShowSendMailFormCmdVisible"); }
        }

        private bool _isProgressBarVisibility;
        public bool IsProgressBarVisibility
        {
            get { return _isProgressBarVisibility; }
            set { _isProgressBarVisibility = value; NotifyPropertyChanged("IsProgressBarVisibility"); }
        }
        #endregion  

        #region --  Constructor  --
        #region --      *******************  Progress bar  ********************      --
        private BackgroundWorker worker;
        private readonly System.Windows.Input.ICommand instigateWorkCommand;

        // your UI binds to this command in order to kick off the work
        public System.Windows.Input.ICommand InstigateWorkCommand
        {
            get { return this.instigateWorkCommand; }
        }

        private double _currentProgress;
        public double CurrentProgress
        {
            get { return _currentProgress; }
            set
            {
                if (_currentProgress != value)
                {
                    _currentProgress = value;
                    NotifyPropertyChanged("CurrentProgress");
                }
            }
        }

        #endregion

        public ObservableCollection<string> ColumnNamesList { get; set; }
        public ViewModel()
        {
            #region --  TEST  --
            //string s = Regex.Replace("gh g hghj+-**/*'1155bjb", "[^0-9A-Za-z]+", "");

            //var client = new SmtpClient("smtp.gmail.com", 587)
            //{
            //    Credentials = new NetworkCredential("mamessageri@gmail.com", "0677300811"),
            //    EnableSsl = true
            //};
            //client.Send("mamessageri@gmail.com", "citoyenlamda@gmail.com", "test", "testbody");
            ////Console.WriteLine("Sent");

            #endregion
            ;

            // --  Manage background worker ProgressBar  --
            this.instigateWorkCommand = new DelegateCommand(o => this.worker.RunWorkerAsync(), o => !this.worker.IsBusy);

            #region --  Manage visibility  --
            //IsFirstFormVisible = true;
            //IsSplitFormVisible = false; ;

            IsSplitCmdVisible = false;
            IsSplitFormVisible = true;
            IsLoadingFormVisible = false;
            IsSendMailFormVisible = false;
            IsSendByMailCmdVisible = false;
            IsSendPreviewCmdVisible = false;
            IsProgressBarVisibility = false;
            IsShowSendMailFormCmdVisible = false;
            IsBackToSplitFormCmdVisible = false;
            IsBackToFirstFormCmdVisible = false;
            #endregion

            // --  Set default extension  --
            SelExtension = Extensions.First();

            // --    --
            this.ColumnNamesList = new ObservableCollection<string>();

            #region --  All commands  --
            //// --  Back to first form command  --
            //BackToFirstFormCmd = new DelegateCommand(x => BackToFirstForm());

            //// --  Upload document  --
            //UploadDocument = new DelegateCommand((x) => System.Diagnostics.Process.Start((string)x));     

            // --  Show send mail form command  --            
            ShowSendMailFormCmd = new DelegateCommand(x => ShowSendMailForm());

            // --  Back to split form command  --
            BackToSplitFormCmd = new DelegateCommand(x => BackToSplitForm());

            // --  Get Excel file command  --
            BrowerFileCmd = new DelegateCommand(x => GetExcelFile());

            // --  Choose mailbody command  --
            ChooseMailBodyCmd = new DelegateCommand(x => ChooseMailBody());

            // --  Choose optional files command  --
            AddOptionalFilesCde = new DelegateCommand(x => ChooseAddFile());

            // --  Send all Excel files by mail to users command  --
            //SendByMailCmd = new DelegateCommand(x => SendMailProperties());
            SendByMailCmd = new DelegateCommand(x => Mail());

            // --  Send mail preview command  --
            //SendPreviewCmd = new DelegateCommand(x => SendMailPreview());
            SendPreviewCmd = new DelegateCommand(x => SendMailPreview());

            // --  Build Copie Split Excel files command  --
            SplitCmd = new DelegateCommand(x => Split());

            // --  Choose folder directory command  --
            ChooseFolderDirectoryCmd = new DelegateCommand(x => GetOutPutFolder());
            #endregion
        }
        #endregion

        #region -- Button command PROPERTIES --                 
        public DelegateCommand SplitCmd { get; set; }
        public DelegateCommand BrowerFileCmd { get; set; }
        public DelegateCommand SendByMailCmd { get; set; }
        public DelegateCommand SendPreviewCmd { get; set; }
        public DelegateCommand UploadDocument { get; set; }
        public DelegateCommand ChooseMailBodyCmd { get; set; }
        //public DelegateCommand BackToFirstFormCmd { get; set; }
        public DelegateCommand BackToSplitFormCmd { get; set; }
        public DelegateCommand ShowSendMailFormCmd { get; set; }
        public DelegateCommand AddOptionalFilesCde { get; set; }
        public DelegateCommand ChooseFolderDirectoryCmd { get; set; }
        #endregion

        /// <summary>
        /// -- Open send mail form --
        /// </summary>
        private void ShowSendMailForm()
        {
            IsSplitFormVisible = false;
            IsSendMailFormVisible = true;
            IsBackToSplitFormCmdVisible = true;
        }

        /// <summary>
        /// --  Back to split form  --
        /// </summary>
        private void BackToSplitForm()
        {
            IsSplitCmdVisible = true;
            IsSplitFormVisible = true;
            IsSendMailFormVisible = false;
            IsBackToSplitFormCmdVisible = false;
        }
        
        #region --     --
        /// <summary>
        /// -- Récupère les feuilles du fichier choisie par user et selectionne la prémière  --
        /// </summary>
        private void ChooseExcelFile()
        {
            // -- Le contenu du fichier choisi est mi dans l'objet thisBook --
            sourceWorkbook = _excelApp.Workbooks.Open(FileName);

            // -- Selection du nom des feuilles et stokage ds SheetNames --
            SheetNames = from sheet in sourceWorkbook.Sheets.Cast<Worksheet>()
                         select sheet.Name;

            // -- Extraction du nom d'une feuille  et affichage dans le combobox --
            SheetName = SheetNames.First();
        }

        /// <summary>
        /// -- Fonction qui convertie les index de feuille en nom de colonne --
        /// </summary>
        /// <returns></returns>
        private string GetSheetColumnName()
        {
            if (sourceWorkbook != null)
            {
                // - Création de la range de la feuille choisie par user a partir de l'index de celle-ci -
                Range thisRange = sourceWorkbook.Sheets[SheetName].UsedRange.SpecialCells(XlCellType.xlCellTypeLastCell);

                // - Get the total of column of the range -
                int lastColumn = thisRange.Column;

                // --      --
                int lastRow = thisRange.Row;
                rows = new string[lastRow];
                for (int i = 0; i < lastRow; i++)
                {
                    rows[i] = i.ToString();
                }

                string[] columns = new string[lastColumn];
                for (int i = 0; i < lastColumn; i++)
                    // - Appel de la fonction qui convertie un index en nom de colonne -
                    columns[i] = ColumnIndexToColumnLetter(i + 1);

                #region - Récupère toutes les colonnes du fichier choisi pour la page de split et charge le combobox pour la colonne de découpe -        
                //// --   --       
                //KeyColNames = columns;

                ColumnNamesList = GetColumnNames(columns);

                // -- Récupère et affiche le 1er nom de la colonne  --
                KeyColName = ColumnNamesList.First();
                #endregion

                #region - Récupère toutes les colonnes du fichier choisi pour la page de split et charge le combobox TO pour la colonne d'email  -
                // --   --
                AllRecipientsEmails = ColumnNamesList;

                // -- Récupère et affiche le 1er nom de la colonne  --
                RecipientEmail = AllRecipientsEmails.First();
                #endregion
            }
            return KeyColName;
        }

        #region --  Get column name cellName based on column name   --
        /// <summary>
        /// --  Get column name cellName based on column name   --
        /// </summary>
        /// <param name="columns"></param>
        /// <returns></returns>
        private ObservableCollection<string> GetColumnNames(string[] columns)
        {
            // -- Check the state of SourceWorkbook close? open: null --
            if (sourceWorksheet == null)
                sourceWorkbook = _excelApp.Workbooks.Open(FileName);

            string cellValue = null;
            sourceWorksheet = sourceWorkbook.Sheets[SheetName];

            // --  Clear Column name list  --  
            ColumnNamesList.Clear();
            #region MyRegion
            for (int i = 1; i < rows.Length; i++)
            {
                string col = null;
                foreach (var item in columns)
                {
                    object rangeObject = sourceWorksheet.Cells[i, GetColumnIndexByName(item)];
                    Range range = (Range)rangeObject;
                    object rangeValue = range.Value2;
                    if (rangeValue != null)
                    {
                        cellValue = rangeValue.ToString();
                        if (!ColumnNamesList.Contains(item + "  -  " + cellValue))
                        {
                            if (ColumnNamesList.Count() == columns.Length)
                            {
                                FirstCell = item + "" + i;
                                string de = i + "" + item;
                                break;
                            }
                            else
                            {
                                col = item;
                                ColumnNamesList.Add(item + "  -  " + cellValue);
                            }
                        }
                    }
                    else
                    {
                        ColumnNamesList.Clear();
                        break;
                    }
                }
                if (ColumnNamesList.Count() == columns.Length)
                {
                    FirstCell = "A" + "" + (i + 1);
                    break;
                }

                #region MyRegion
                //// -- Check the state of SourceWorkbook close? open: null --
                //if (sourceWorksheet == null)
                //    sourceWorkbook = _excelApp.Workbooks.Open(FileName);

                //sourceWorksheet = sourceWorkbook.Sheets[SheetName];
                //foreach (var item in columns)
                //{
                //    int colIndex = GetColumnIndexByName(item);

                //    //object rangeObject = sourceWorksheet.Cells[1, colIndex];
                //    //Range range = (Range)rangeObject;
                //    //object rangeValue = range.Value2;
                //    //string cellValue = rangeValue.ToString();
                //    //// --   --                      
                //    //ColumnNamesList.Add(item + "  -  " + cellValue);
                //    // --   --                      
                //    ColumnNamesList.Add(item);
                //}
                //return ColumnNamesList; 
                #endregion
            }
            #endregion
            return ColumnNamesList;
        }

        /// <summary>
        /// --  Get column index based on colum name  --
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public static int GetColumnIndexByName(string name)
        {
            int number = 0;
            int pow = 1;
            for (int i = name.Length - 1; i >= 0; i--)
            {
                number += (name[i] - 'A' + 1) * pow;
                pow *= 26;
            }

            return number;
        }
        #endregion

        /// <summary>
        /// --     --
        /// </summary>
        /// <returns></returns>
        private bool CheckNotEmptyEmailFeel()
        {
            if (!string.IsNullOrWhiteSpace(SenderName) && !string.IsNullOrWhiteSpace(Object) && !string.IsNullOrWhiteSpace(SenderMail) && !string.IsNullOrWhiteSpace(MailBody))
            {
                IsSendByMailCmdVisible = true;
                IsSendPreviewCmdVisible = true;
                return true;
            }
            else
            {
                IsSendByMailCmdVisible = false;
                IsSendPreviewCmdVisible = false;
                return false;
            }
        }

        /// <summary>
        /// --  Cooose mail body function  --
        /// </summary>
        private void ChooseMailBody()
        {
            OpenFileDialog opfd = new OpenFileDialog();
            #region - MyRegion 1 -
            // "Excel Files |*.xlsx;*.xls;*.xlsm;*.csv;*,*.xlsb";
            opfd.Filter = "html files |*.html;*.htm";
            /*indique l'index à laquelle ton filtre du dessus se place (Ici All files)*/
            opfd.Title = "Select a File";
            //openFileDialog.FilterIndex = 1;
            opfd.RestoreDirectory = true;

            if (opfd.ShowDialog() == true)
            {
                MailBody = opfd.FileName;
                MailBodyPath = Path.GetDirectoryName(MailBody);
            }
            #endregion
        }

        /// <summary>
        /// --  Choose optional files  --
        /// </summary>
        private void ChooseAddFile()
        {
            OpenFileDialog ofd = new OpenFileDialog();

            ofd.Filter = "All files (*.*)|*.*";
            ofd.Title = "Please Select Source File(s)";

            List<string> myAddFilesList = new List<string>();

            // - Add many files in the textBox control -
            if (ofd.ShowDialog() == true)
                OtherFileName += ofd.FileName.ToString() + "\n";
        }
        #endregion

        #region --  Methodes  --
        /// <summary>
        /// -- Fonction qui convertie les index de feuille en nom de colonne --
        /// </summary>
        /// <param name="colIndex"></param>
        private string ColumnIndexToColumnLetter(int colIndex)
        {
            if (colIndex == 0) throw new ArgumentNullException("Entrez l'index de la colonne");

            int div = colIndex;
            string colLetter = String.Empty;
            int mod = 0;

            while (div > 0)
            {
                mod = (div - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                div = (int)((div - mod) / 26);
            }
            return colLetter;
        }

        /// <summary>
        /// -- Fonction qui convertie le nom de colonne de feuille en index  --
        /// </summary>
        /// <param name="columnName"></param>
        private int ExcelColNameToNumber(string columnName)
        {
            if (string.IsNullOrEmpty(columnName))
                throw new ArgumentNullException("Entrez le nom de colonne");
            columnName = columnName.ToUpperInvariant();

            int sum = 0;
            for (int i = 0; i < columnName.Length; i++)
            {
                sum *= 26;
                sum += (columnName[i] - 'A' + 1);
            }
            return sum;
        }

        bool IsDigitsOnly(string str)
        {
            foreach (char c in str)
            {
                if (c < '0' || c > '9')
                    return false;
            }
            return true;
        }
        #endregion

        #region --  Debut processus  --
        /// <summary>
        /// -- Selectionne le fichier excele choisi par use --
        /// </summary>
        private void GetExcelFile()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files |*.xlsx;*.xls;*.xlsm;*.csv;*,*.xlsb";
            openFileDialog.FilterIndex = 1;     // --> indique l'index à laquelle ton filtre du dessus se place (Ici All files) --
            openFileDialog.RestoreDirectory = true;

            if (openFileDialog.ShowDialog() == true)
            {
                FileName = openFileDialog.FileName;
                string filePath = Path.GetDirectoryName(FileName);

                #region  - Récupère le nom de toutes les feuilles du fichier choisi -
                if (FileName != null)
                {
                    // -- Get file extension --
                    SelExtension = Path.GetExtension(FileName);

                    #region  - Check file extension -
                    if (Extensions.Contains(SelExtension) == true)
                    {
                        // -- Ouverture de l'application Excel --
                        _excelApp = new Application();
                        // -- Choix de la feuille --
                        ChooseExcelFile();
                    }
                    else
                        DisplayErrorMessage("Le fichier choisi ne correspond pas aux extensions attendues !");
                    #endregion
                }
                #endregion
            }
        }

        /// <summary>
        /// -- Choose folder --
        /// </summary>
        private void GetOutPutFolder()
        {
            System.Windows.Forms.FolderBrowserDialog FBD = new System.Windows.Forms.FolderBrowserDialog();
            FBD.Description = "Choississez le dossier de destination ! ";
            FBD.ShowNewFolderButton = false;
            if (FBD.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                OutputFolder = FBD.SelectedPath;
        }

        /// <summary>
        /// -- Ouvre le workBook recupère le contenu de la feuille ds un tableau multidimensionnel --
        /// </summary>
        private void MyArray()
        {
            // --  Call restart Excel com object  --
            RestartExcelComObject();

            // -- Check the state of SourceWorkbook close? open: null --
            if (sourceWorksheet == null)
                sourceWorkbook = _excelApp.Workbooks.Open(FileName);

            sourceWorksheet = sourceWorkbook.Sheets[SheetName];

            // -- Take the used range of the sheet. Finally, get an object array of all of the cells in the sheet (their values). --
            Range excelRange = sourceWorksheet.UsedRange;

            object[,] valArray = (object[,])excelRange.get_Value(XlRangeValueDataType.xlRangeValueDefault);

            // -- Do something with the data in the array with a custom method --
            rowsArray = valArray;

            // -- Compte le nbre de ligne ds la tableau multidimentionnel rowArray --
            int rowCount = rowsArray.GetLength(0);
        }

        /// <summary>
        /// --  Get start data index function  --
        /// </summary>
        private int GetStartDataRowIndex()
        {
            dataRowIndex = Convert.ToInt32(FirstCell.Substring(1, FirstCell.Length - 1));

            return dataRowIndex;
        }

        /// <summary>
        /// -- Fonction qui récupère l'index de la colonne de debut des données --
        /// </summary>
        private int GetStartDataColIndex()
        {
            int startDataColIndex = (int)FirstCell[0] - 64;

            return startDataColIndex;
        }

        /// <summary>
        /// --  Fonction qui récupère l'index de la colonne de découpage des données --
        /// </summary>
        private int GetStartKeyColIndex()
        {
            string[] keyCNames = KeyColName.Split('-');

            int KeyColIndex = ExcelColNameToNumber(keyCNames[0].Trim());

            return KeyColIndex;
        }

        /// <summary>
        /// --  Fonction qui récupère l'index de la colonne d'email --
        /// </summary>
        private int GetEmailColIndex()
        {
            string[] emails = RecipientEmail.Split('-');

            EmailColIndex = ExcelColNameToNumber(emails[0].Trim());

            return EmailColIndex;
        }

        /// <summary>
        /// --  Check first cell function  --
        /// </summary>
        /// <returns></returns>
        private bool CheckFirstCell()
        {
            FirstCell = RemoveSpecialCharacters(FirstCell);

            string fCell = FirstCell.Substring(1, FirstCell.Count() - 1);
            char Cel = FirstCell[0];

            if (!Char.IsLetter(Cel))
            {
                //string Cell = Regex.Replace(Convert.ToString(Cel), "[^a-zA-Z_]+", "");
                DisplayErrorMessage("Dans la colonne de debut des données,\n\r Le prémier caractère doit être une lettre.\n\r  Veuillez modifier puis recommencez !");
                return false;
            }
            else
            {
                fCell = Regex.Replace(fCell, "[^0-9_]+", "", RegexOptions.Compiled);

                #region - Check if cell IsDigit or not -
                if (IsDigitsOnly(fCell))
                {
                    FirstCell = RemoveSpecialCharacters(FirstCell);
                    string fCel = Regex.Replace(FirstCell.Substring(1, FirstCell.Count() - 1), "[^0-9_]+", "", RegexOptions.Compiled);

                    int FCell = Convert.ToInt32(fCel);

                    // -- Appel de lan fonction qui Ouvre le workBook recupère le contenu de la feuille ds un tableau multidimensionnel --
                    MyArray();

                    // --  Obtient le total des lignes du tableaux multidimensionnel -- 
                    rowCount = rowsArray.GetLength(0);
                    // -  - 
                    if (FCell < rowCount)
                    {
                        FirstCell = FirstCell[0] + "" + FCell;
                        //Console.WriteLine(FCell);
                        return true;
                    }
                    else
                        DisplayErrorMessage("Dans la colonne de debut des données,\n\r Vous avez saisi plus de ligne que le fichier ne contient !");

                    // --   --
                    ReleaseExcelComObject();
                    CloseEXCELAPP();
                }
                #endregion
            }
            return false;
        }
        #endregion

        #region --  Split back ground  --
        /// <summary>
        /// -- Function qui gere le background worker pendant la découpe  -- 
        /// </summary>
        public void Split()
        {
            worker = new BackgroundWorker();
            worker.DoWork += worker_DoWork;
            worker.ProgressChanged += worker_ProgressChanged;
            worker.ProgressChanged += Worker_ProgressChanged;
            worker.RunWorkerAsync();
            worker.RunWorkerCompleted += worker_RunWorkerCompleted;
        }

        // --    --
        void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            // -- Appel de la fonction de découpage suivant la clef saisie par user -- 
            SplitExcelFile(GetStartKeyColIndex());
        }

        // --     --
        void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            this.CurrentProgress = e.ProgressPercentage;
            if (e.UserState != null)
                Console.WriteLine(e.UserState);
        }

        // --    --
        void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            System.Windows.MessageBox.Show(System.Windows.Application.Current.MainWindow, "La découpe s'est bien déroulée.",
                "XLSXDécoupeFiles", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);

            // --  Show SendMailForm  --
            IsShowSendMailFormCmdVisible = true;
        }

        private void Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            CurrentProgress = e.ProgressPercentage;
        }
        #endregion

        #region --  Send mail Background worker  --
        /// <summary>
        /// -- Function qui gere le background worker pendant l'envoie des mails  -- 
        /// </summary>
        public void Mail()
        {
            worker = new BackgroundWorker();
            worker.DoWork += DoWork;
            worker.ProgressChanged += ProgressChanged;
            worker.ProgressChanged += ProgressChanged;
            worker.RunWorkerAsync();
            worker.RunWorkerCompleted += RunWorkerCompleted;
        }

        void DoWork(object sender, DoWorkEventArgs e)
        {
            // -- Appel de la fonction de découpage suivant la clef saisie par user -- 
            SendMail();
        }

        void ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            this.CurrentProgress = e.ProgressPercentage;
            if (e.UserState != null)
                Console.WriteLine(e.UserState);
        }

        void RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //System.Windows.MessageBox.Show(System.Windows.Application.Current.MainWindow, "Les fichiers ont bien étés envoyés par mail aux destinataires.",
            //    "XLSXDécoupeFiles", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
        }
        #endregion

        #region - Debut Trie et découpage -
        /// <summary>
        /// -- Fonction qui trie et découpe les données suivant la clef choisie par user et enregistre  --
        /// </summary>
        /// <param name="columnID"></param>
        private void SplitExcelFile(int columnID)
        {
            #region - Declare... -
            string srcFilePathSortSheet = null, destFilePathSortSheet = null;
            dataRowIndex = GetStartDataRowIndex();
            ColKeyIndex = GetStartKeyColIndex();
            #endregion

            try
            {
                #region --  Restart Excel Com Objject  --
                // --  Call restart Excel com object  --
                RestartExcelComObject();

                if (sourceWorkbook == null)
                    sourceWorkbook = _excelApp.Workbooks.Open(FileName);
                #endregion

                if (CheckFirstCell())
                {
                    TotalFiles = 0; counter = 0;

                    // --     --
                    IsProgressBarVisibility = true;

                    // --     --
                    IsSplitCmdVisible = false;
                    //IsLoadingFormVisible = true;
                    IsShowSendMailFormCmdVisible = false;
                    // --     --
                    allRowsId.Clear();

                    Console.WriteLine("Lecture des différentes valeurs ...");

                    #region - getsion du trie du fichier choisi par user -            
                    // -- Récupération de la plage de donnée à trier -- 
                    Range sortRange = sourceWorksheet.Range[FirstCell, sourceWorksheet.UsedRange.SpecialCells(XlCellType.xlCellTypeLastCell)];

                    // -- Appel de la fonction de tri d'excel par colonne --
                    sortRange.Sort(sortRange.Columns[ColKeyIndex], XlSortOrder.xlAscending);

                    // -- Récupération de la range triée ds un tableau --
                    valueArray = (object[,])sortRange.get_Value(XlRangeValueDataType.xlRangeValueDefault);
                    rowsArray = valueArray;
                    int dataRowCount = rowsArray.GetLength(0);
                    #endregion

                    #region - Début alimentation du dictionaire -
                    for (int i = 1; i <= dataRowCount; i++)
                    {
                        object keyObject = rowsArray[i, columnID];
                        string keyString = keyObject == null ? null : keyObject.ToString();

                        if (keyString != null)
                        {
                            keyString = keyString.Trim().ToUpper();
                            if (!allRowsId.ContainsKey(keyString))
                                allRowsId.Add(keyString, new RowRange(i + dataRowIndex - 1));
                            allRowsId[keyString].End++;
                        }
                    }

                    //// - Retrive the first key of allRowsId -
                    #endregion - Fin alimentation du dictionaire -

                    #region - Création du chemin des fichiers triés en fonction de l'extension de l'extension du fichier - 
                    // - Gere l'affichage de l'extension du fichier de sortie en fonction fichier du fichier d'entré (Par defaut) -
                    string fileExtension = Path.GetExtension(FileName);

                    srcFilePathSortSheet = Path.Combine(OutputFolder, "OrigSortFile" + fileExtension);
                    destFilePathSortSheet = Path.Combine(OutputFolder, "CopieSortFile" + fileExtension);
                    #endregion

                    // -- Sauvegarde d'une copie du fichier trié -- 
                    _excelApp.DisplayAlerts = false;
                    sourceWorksheet.SaveAs(srcFilePathSortSheet);
                    _excelApp.DisplayAlerts = true;

                    // -- Restauration de la copie du fichier trié -- 
                    File.Copy(srcFilePathSortSheet, destFilePathSortSheet, true);

                    // -- Open thisWorkbook --
                    if (tempWorkbook == null)
                        tempWorkbook = _excelApp.Workbooks.Open(destFilePathSortSheet);

                    // -- Get sortWorkbook --
                    tempWorksheet = tempWorkbook.Sheets[SheetName];

                    string outputFileName = null;

                    #region - Begining cutting -
                    foreach (string key in allRowsId.Keys)
                    {
                        Range sourceRange = sourceWorksheet.UsedRange;

                        String KEY;
                        Console.WriteLine("Key : " + counter++ + " - " + key);

                        #region - Cutting -  
                        // -- Suppression des lignes après la sélection --
                        if (allRowsId[key].End < rowCount)
                        {
                            sourceRange = sourceWorksheet.get_Range("A" + (allRowsId[key].End), "A" + rowCount + dataRowIndex);
                            Range entireRowTop = sourceRange.EntireRow;
                            entireRowTop.Delete(Type.Missing);
                        }

                        // -- Suppression des lignes avant la sélection --
                        if (allRowsId[key].Begin > dataRowIndex)
                        {
                            sourceRange = sourceWorksheet.get_Range("A" + dataRowIndex, "A" + (allRowsId[key].Begin - 1));
                            Range entireRowBottom = sourceRange.EntireRow;
                            entireRowBottom.Delete(Type.Missing);
                        }
                        #endregion - Cutting -

                        // -- Gestion de l'enregistrement -- 
                        KEY = Regex.Replace(key, "[^a-zA-Z0-9_.]+", "_");
                        Prefixe = Prefixe == null ? null : RemoveSpecialCharacters(Prefixe);
                        Suffixe = Suffixe == null ? null : RemoveSpecialCharacters(Suffixe);

                        string OutputfolderName = Prefixe + KEY + Suffixe + SelExtension;
                        outputFileName = Path.Combine(OutputFolder, OutputfolderName);

                        if (File.Exists(outputFileName))
                            File.Delete(outputFileName);

                        // -- Set the focus to the first cell --
                        sourceWorksheet.Range["A1"].Copy();
                        sourceWorksheet.Range["A1"].PasteSpecial();

                        #region - Save -  
                        if (SelExtension == ".xls")
                            sourceWorkbook.SaveAs(outputFileName, XlFileFormat.xlExcel8);
                        if (SelExtension == ".xlsx")
                            sourceWorkbook.SaveAs(outputFileName, XlFileFormat.xlOpenXMLWorkbook);
                        if (SelExtension == ".xlsm")
                            sourceWorkbook.SaveAs(outputFileName, XlFileFormat.xlOpenXMLWorkbookMacroEnabled);
                        if (SelExtension == ".xlsb")
                            sourceWorkbook.SaveAs(outputFileName, XlFileFormat.xlExcel12);
                        if (SelExtension == ".csv")
                            sourceWorkbook.SaveAs(outputFileName, XlFileFormat.xlCSVWindows);
                        #endregion - Save -

                        // -- Ouvrir un worksheet a patir de son index --
                        tempWorksheet.UsedRange.Copy();

                        Range tempRange = tempWorksheet.UsedRange;
                        tempRange = tempWorksheet.get_Range(FirstCell, tempWorksheet.UsedRange.SpecialCells(XlCellType.xlCellTypeLastCell));
                        sourceRange = sourceWorksheet.get_Range(FirstCell, sourceWorksheet.UsedRange.SpecialCells(XlCellType.xlCellTypeLastCell));
                        tempRange.Copy();
                        sourceRange.PasteSpecial();
                    }
                    // --    --
                    TotalFiles = counter;

                    Console.WriteLine("Fin de la lecture des valeurs.");
                    #endregion - End cut -                                      

                    // -- Manage visibility --
                    IsSplitCmdVisible = true;
                    IsLoadingFormVisible = false;
                    IsProgressBarVisibility = false;
                    IsShowSendMailFormCmdVisible = true;
                }
            }
            #region - Catch exception and Close enstenses -
            catch (Exception ex)
            {
                Console.WriteLine("Erreur :" + ex.Message);
                string erreur = ex.ToString();
            }
            finally
            {
                if (TotalFiles > 0)
                {
                    //IsShowSendMailFormCmdVisible = true;
                    IsSplitCmdVisible = false;
                }
                else
                {
                    IsSplitCmdVisible = true;
                    IsShowSendMailFormCmdVisible = false;
                }

                // --  Call release excel object  --
                ReleaseExcelComObject();

                // --     --
                CloseEXCELAPP();

                // -- Pour suprimer les fichiers temporaires crées --
                if (File.Exists(srcFilePathSortSheet) || File.Exists(destFilePathSortSheet))
                {
                    File.Delete(srcFilePathSortSheet);
                    File.Delete(destFilePathSortSheet);
                }
            }
            #endregion
        }
        #endregion

        #region - Send mail -
        /// <summary>
        /// --  Check email format function  --
        /// </summary>
        /// <param name="mailAd"></param>
        private void CheckEmailFormat(string mailAd)
        {
            string pattern = null;
            pattern = "^([0-9a-zA-Z]([-\\.\\w]*[0-9a-zA-Z])*@([0-9a-zA-Z][-\\w]*[0-9a-zA-Z]\\.)+[a-zA-Z]{2,9})$";

            if (!Regex.IsMatch(mailAd, pattern))
                DisplayErrorMessage("L'adresse email : " + mailAd + " ne respecte pas le format requis ");
        }

        /// <summary>
        /// --  Send mail  --
        /// </summary>
        private void SendMail()
        {
            int nbr = 0;

            EmailColIndex = GetEmailColIndex();

            try
            {
                Task.Run(() =>
                {
                    // --     --
                    IsSplitFormVisible = false;
                    IsLoadingFormVisible = true;
                    IsSendMailFormVisible = false;
                    IsBackToSplitFormCmdVisible = false;

                    #region - Début alimentation du dictionaire et envoie des mails -
                    // -- Création d'un dictionnaire pour y stocker l'email et la cléf de découpage --
                    Dictionary<string, string> keyEmailDictionary = new Dictionary<string, string>();

                    for (int i = dataRowIndex; i < valueArray.GetLength(0); i++)
                    {
                        object keyCell = valueArray[i, ColKeyIndex];

                        if (keyCell != null)
                        {
                            string keyValue = keyCell == null ? null : keyCell.ToString().ToUpper();
                            object emailCell = valueArray[i, EmailColIndex];

                            if (emailCell != null)
                            {
                                #region - Check email format before store in my dictionary -
                                string email = emailCell == null ? null : emailCell.ToString();

                                if (email != null)
                                {
                                    email = email.Trim();
                                    if (!keyEmailDictionary.ContainsKey(keyValue))
                                        keyEmailDictionary.Add(keyValue, email);
                                }
                                #endregion
                            }
                        }
                    }
                    #endregion

                    #region - Gestion de l'envoie des mails avec récupération du nom associé à la key et à l'email -
                    // -- Création du template de l'email avec encodage par défaut --
                    foreach (KeyValuePair<string, RowRange> keyVP in allRowsId)
                    {
                        string email = keyEmailDictionary[keyVP.Key];
                        // -- Gestion du format de l'amail -- 
                        email = email.Contains(",") == true ? email.Replace(",", ".") : email;

                        string fileName = Prefixe + keyVP.Key + Suffixe + SelExtension;
                        string filePath = Path.Combine(OutputFolder, fileName);

                        // -- Send mail --
                        if (!SendMail(email, filePath, MailBody, valueArray, keyVP.Key))
                            DisplayErrorMessage("Erreur durant l'envoie des fichiers!");
                        else
                            nbr++;
                    }

                    // --  Set message after sending emails  --
                    if (nbr == allRowsId.Count)
                        DisplayMessage("Tous les fichiers ont bien été envoyés aux destinataires.");
                    #endregion
                    // --  Manage visibility  --
                    IsSplitFormVisible = false;
                    IsLoadingFormVisible = false;
                    IsSendMailFormVisible = true;
                    IsBackToSplitFormCmdVisible = true;
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine("Erreur :" + ex.ToString());
            }
        }

        /// <summary>
        /// -- Fontion d'envoie de mail contenant les différents fichiers --
        /// </summary>
        private bool SendMail(string toEmail, string file, string htmlTemplate, object[,] valueArray, string key)
        {
            bool result = false;

            #region - CHECK EMAIL -
            CheckEmailFormat(SenderMail);
            CheckEmailFormat(toEmail);
            #endregion

            #region --  New with images  --
            //using (MailClient mailClient = new MailClient())
            //using (MailTemplate mailTemplate = new MailTemplate(htmlTemplate, Encoding.Default))
            //{
            //    // - Set SMTP client credentials -
            //    mailClient.Credentials = CredentialCache.DefaultNetworkCredentials;

            //    // - Load template for the first time -
            //    mailTemplate.LoadTemplate();

            //    #region - Creation du corps du mail -
            //    // - Parcours des lignes du tableau multidimensionnel pour récupérer celle correspondant à chaque key -
            //    for (int row = dataRowIndex; row < valueArray.GetLength(0); row++)
            //    {
            //        // - Définission de la cellule -
            //        string keyCell = valueArray[row, ColKeyIndex].ToString().ToUpper();
            //        if (keyCell == key)
            //        {
            //            // - Parcours des colonnes du tableau afin d'effectuer les remplacements -
            //            for (int cIndex = 1; cIndex < KeyColNames.Count() + 1; cIndex++)
            //            {
            //                // - Conversion de l'index de la colonne en nom nom colonne -
            //                string colName = ColumnIndexToColumnLetter(cIndex);
            //                string cName = "{" + colName + "}";

            //                object cellValue = valueArray[row, cIndex];
            //                string cellText = cellValue == null ? null : cellValue.ToString();

            //                // - Remplacement des champs de fusions -
            //                mailTemplate.ReplaceText(cName, cellText);
            //            }
            //            break;
            //        }
            //    }
            //    #endregion

            //    #region - Check CC and Bcc -
            //    if (!string.IsNullOrWhiteSpace(CcMail))
            //    {
            //        mailTemplate.CC.Add(new MailAddress(CcMail));
            //    }
            //    if (!string.IsNullOrWhiteSpace(BccMail))
            //    {
            //        mailTemplate.Bcc.Add(new MailAddress(BccMail));
            //    }
            //    #endregion

            //    // - Add main excel file in attachments -
            //    mailTemplate.AddAttachment(file, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

            //    #region - Check optional file -
            //    if (!string.IsNullOrWhiteSpace(OtherFileName))
            //    {
            //        // -Tableau des diférents fichiers selectionnés et découpé par la fonction split -
            //        string[] files = OtherFileName.Split(new char[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);

            //        for (int i = 0; i < files.Length; i++)
            //        {
            //            string filePath = files[i];

            //            if (File.Exists(filePath) == true)
            //            {
            //                mailTemplate.AddAttachment(filePath, "application/octet-stream");
            //            }
            //        }
            //    }
            //    #endregion

            //    mailTemplate.LoadData();

            //    Console.WriteLine("Sending an e-mail message to {0} and {1}.", SenderMail, mailTemplate.CC.ToString());

            //    mailClient.SendMailProperties(mailTemplate, SenderMail, SenderName, toEmail, Object);
            //} 
            #endregion

            #region --  OLD  --
            MailAddress from = new MailAddress(SenderMail, SenderName);
            MailAddress to = new MailAddress(toEmail);
            MailMessage message = new MailMessage(from, to);
            message.Subject = Object;

            try
            {
                #region - Creation du corps du mail -
                htmlTemplate = File.ReadAllText(MailBody, System.Text.Encoding.Default);
                string body = htmlTemplate;
                int row = 0;

                string cName = null;

                // -- Parcours des lignes du tableau multidimensionnel pour récupérer celle correspondant à chaque key --
                for (row = dataRowIndex; row < valueArray.GetLength(0); row++)
                {
                    // -- Définission de la cellule --
                    string keyCell = valueArray[row, ColKeyIndex].ToString().ToUpper();
                    if (keyCell == key)
                    {
                        // -- Parcours des colonnes du tableau afin d'effectuer les remplacements --
                        //for (int cIndex = 1; cIndex < KeyColNames.Count() + 1; cIndex++)
                        for (int cIndex = 1; cIndex < ColumnNamesList.Count() + 1; cIndex++)
                        {
                            // -- Conversion de l'index de la colonne en nom nom colonne --
                            string colName = ColumnIndexToColumnLetter(cIndex);
                            cName = "{" + colName + "}";

                            object cellValue = valueArray[row, cIndex];
                            string cellText = cellValue == null ? null : cellValue.ToString();

                            // -- Remplacement des champs de fusions --
                            body = body.Replace(cName, cellText);
                        }
                        break;
                    }
                }
                #endregion                

                message.BodyEncoding = System.Text.Encoding.Default;
                message.IsBodyHtml = true;
                message.Body = body;

                #region - Check CC and Bcc -
                if (!string.IsNullOrWhiteSpace(CcMail))
                {
                    MailAddress cc = new MailAddress(CcMail);
                    message.CC.Add(cc);
                }
                if (!string.IsNullOrWhiteSpace(BccMail))
                {
                    MailAddress bcc = new MailAddress(BccMail);
                    message.Bcc.Add(bcc);
                }
                #endregion

                #region - Your log file path -
                Attachment attachment1 = new Attachment(file);
                message.Attachments.Add(attachment1);
                #endregion

                #region - Check optional file -
                if (!string.IsNullOrWhiteSpace(OtherFileName))
                {
                    // -- Tableau des diférents fichiers selectionnés et découpé par la fonction split --
                    string[] files = OtherFileName.Split(new char[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);

                    for (int i = 0; i < files.Length; i++)
                    {
                        string filePath = files[i];

                        if (File.Exists(filePath) == true)
                        {
                            Attachment attachment2 = new Attachment(filePath);
                            message.Attachments.Add(attachment2);
                        }
                    }
                }
                #endregion

                #region - Client send -
                //// -- Client send Message --
                //SmtpClient client = new SmtpClient("fr-appsp04.hilti.com");
                //client.Credentials = System.Net.CredentialCache.DefaultNetworkCredentials;
                //Console.WriteLine("Sending an e-mail message to {0} and {1}.", to, message.CC.ToString());
                Console.WriteLine("Sending an e-mail message to {0} and {1}.", to, message.CC.ToString());
                try
                {
                    var client = new SmtpClient("smtp.gmail.com", 587)
                    {
                        Credentials = new NetworkCredential("mamessageri@gmail.com", "0677300811"),
                        EnableSsl = true
                    };

                    #region -- Test !!!!!!!!!!!!!!!!!!!!!!!!!!!!  --
                    //Console.WriteLine("Destiné à : " + message.To.FirstOrDefault());
                    //message.Subject += " - To : " + message.To.FirstOrDefault();
                    //message.To.Clear();
                    //message.To.Add("MABOMIC@hilti.com");
                    //message.CC.Clear();
                    //message.Bcc.Clear();
                    //message.Bcc.Add("MABOMIC@hilti.com");
                    //client.Send(message);
                    #endregion
                    client.Send(message);
                    result = true;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Exception caught in CreateBccTestMessage(): {0}", ex.ToString());
                    result = false;
                }
                #endregion
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception caught. ", ex.ToString());
                result = false;
            }
            #endregion
            return result;
        }

        /// <summary>
        /// --  Function qui permet d'envoyer les fichiers par mail  --
        /// </summary>
        private void SendMailPreview()
        {
            try
            {
                Task.Run(() =>
                {
                    // --     --
                    IsSplitFormVisible = false;
                    IsLoadingFormVisible = true;
                    IsSendMailFormVisible = false;
                    IsBackToSplitFormCmdVisible = false;

                    // -- Définition du fichier correspond à la prémiere key -- 
                    string attachement = Path.Combine(OutputFolder, Prefixe + allRowsId.Keys.First() + Suffixe + SelExtension);

                    // -- Envoie du mail test avec le fichier correspond à la prémiere key -- 
                    if (SendMail("citoyenlamda@gmail.com", attachement, MailBody, valueArray, allRowsId.Keys.First()))
                        DisplayMessage("Les fichiers ont bien étés envoyés par mail aux destinataires.");
                    else
                        DisplayErrorMessage("Erreur durant l'envoie des fichiers par mail !");

                    // --  Manage visibility  --
                    IsSplitFormVisible = false;
                    IsLoadingFormVisible = false;
                    IsSendMailFormVisible = true;
                    IsBackToSplitFormCmdVisible = true;
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine("Erreur :" + ex.ToString());
            }
        }
        #endregion

        #region --   Manage Excel Com Object application   --
        /// <summary>
        /// -- Fonction qui ferme l'application Excel --
        /// </summary>
        private void CloseEXCELAPP()
        {
            if (_excelApp != null)
            {
                // --  Allow excel to display message   --
                _excelApp.DisplayAlerts = true;

                // --  Release Excel instance  --
                _excelApp.Quit();
                Marshal.ReleaseComObject(_excelApp);
                _excelApp = null;
            }
        }

        /// <summary>
        /// --  Restart Excel Com object  --
        /// </summary>
        private void RestartExcelComObject()
        {
            if (_excelApp == null)
                _excelApp = new Application();
        }

        /// <summary>
        /// -- Fonction qui ferme l'application Excel --
        /// </summary>
        private void ReleaseExcelComObject()
        {
            _excelApp.DisplayAlerts = false;

            if (sourceWorksheet != null)
            {
                Marshal.ReleaseComObject(sourceWorksheet);
                sourceWorksheet = null;
            }

            if (tempWorksheet != null)
            {
                Marshal.ReleaseComObject(tempWorksheet);
                tempWorksheet = null;
            }

            if (sourceWorkbook != null)
            {
                sourceWorkbook.Close(false);
                Marshal.ReleaseComObject(sourceWorkbook);
                sourceWorkbook = null;
            }

            if (tempWorkbook != null)
            {
                tempWorkbook.Close(false);
                Marshal.ReleaseComObject(tempWorkbook);
                tempWorkbook = null;
            }
        }
        #endregion
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace WasteDepartments
{
    enum Product {Mixtures,Fabric,Steel,Total}//קוד זיהוי עבור איזה טבלה דרוש המידע
    public partial class Form1 : Form
    {
        DataTable WasteTableMixures = new DataTable();//טבלת מיקסרים
        DataTable WasteTableFabric = new DataTable();//טבלת בדים
        DataTable WasteTableSteel = new DataTable();//טבלת פלדה
        DataTable WasteTableTotal = new DataTable();//טבלת ריכוזי פסולות
        DataTable TableForUpdate = new DataTable();//טבלה עבור עדכונים של הכמויות
        DataTable TableFor1DetailedCell;//טבלה שנועדה עבור פירוט פסולות של תא אחד -נלקח מתוך DataGrid_Desc
        DataTable DescriptionWasteTable;// תיאור מפורט של פסולת בטב עדכון
        DataTable DataOfCellTable = new DataTable();//טבלת פירוט על כל תא של קודי פסולת
        DBService DBS = new DBService();
        Waste WasteObject = new Waste();
        int RowsPosition = -1;//מיקום השורה בא ארצה להכניס את כל נתוני עדכוני כמויות
        int RowsPositionForDetailedCell = -1;//מיקום השורה עבור עדכוני כמויות אבל של תא ספציפי
        int PreviousValue;//ערך קודם של תא לפני ששונה
        int eRowIndex, eColumnIndex;//לדעת באיזה תא מדובר בשביל אחרי העדכון לצבוע אותו בטורקיז
        int LeftOver = 0;//משתנה פירוט פסולת,אומר כמה נשאר לפרט
        int TotalAmount = 0;//כמות כוללת שיש מאותו יום
        int PreviousNumOfcell;//מספר קודם אם מעדכן מתוך טבלת DataGrid_Desc
        int IndexOfDatagrid;
        bool Start = true;//האם מדובר בהצגת נתונים של חודש נוכחי או לא
        bool ValueChanged = false;//חיווי אם היה שינוי כלשהו במידה ויסגור תוכנית בלי לשמור נתונים
        bool ValueNew = true;//האם מדובר בתא שמעדכנים כמות
        bool CellUpdated = false;//בדיקה אם מדובר ברשומה שעודכנה או שלא
        bool CboCodeWaste = false;//האם קומבובוקס פירוט קוד פסולת השתנה,עבור שחרור כפתור אשר
        bool SaveData = false;//יחסום צביעת תא אם לא שמרנו נתונים
        bool btnUpdatePress = false;//לדעת אם לשחרר את tab page waste update
        bool CellEmpty = true;//אם מדובר בתא ריק (עבור עדכון פסולות של תא )נועד לדעת אם למחוק את הקודם
        bool AskToSave = true;//האם המשתמש סימן שלא רוצה לשמור פירוט עבור תא
        bool Tourqise = false;//אם מדובר בערך טורקיז או לא
        TabPage TabPageToReturn;//לאיזה טב לחזור לאחר שעדכנתי בפירוט פסולות
        List<int> IndexHeaders = new List<int>();//מיקומים עבור כותרות -משמש לתת כותרות של משמרות בוקר ערב ולילה
        List<DataGridView> dataGridViewsList = new List<DataGridView>();
        List<CellPaint> cellPaintsListMixures = new List<CellPaint>();
        List<CellPaint> cellPaintsListFabric = new List<CellPaint>();
        List<CellPaint> cellPaintsListSteel = new List<CellPaint>();
        Dictionary<int, bool> IsPainted = new Dictionary<int, bool>();//האם נצבע בעבר או לא?
        bool NewPaint = true;//מונע איטיות,צובע רק פעם אחת את התאים


        public Form1()
        {   
            InitializeComponent();
            WasteObject.CreateMember();
            CheckS400FieldsExist();//בדיקה קיום שדות באס400 לפני שמריצים את התוכנית לפי הנחייתו של אלי
            ShowWaste();
            Start = false;//אחרי שסיימנו להציג, ההצגה הבאה תהיה רק לפי בחירת חודשים לפי בחירת משתמש
            StartPositionFunc();
            CreateTableForUpdate();//יוצר טבלת שתיועד לעדכוני כמויות        
        }



        private async void Form1_Load(object sender, EventArgs e)
        {
            this.Location = new Point(0, 0);
            this.Size = Screen.PrimaryScreen.WorkingArea.Size;
            this.FormBorderStyle = FormBorderStyle.Fixed3D;
            this.WindowState = FormWindowState.Maximized;
            try
            {
                 await InitSelectTabs();
            }
            catch (Exception ex)
            {
                //Handle Exception
            }
            

            NewPaint = false;
            //עוזר לי לתת כותרות
            this.DGV_Fabric.AllowUserToAddRows = false;
            this.DGV_Fabric.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
            this.DGV_Fabric.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
            this.DGV_Fabric.ColumnHeadersHeight = this.DGV_Fabric.ColumnHeadersHeight * 2;
            this.DGV_Fabric.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
            this.DGV_Steel.AllowUserToAddRows = false;
            this.DGV_Steel.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
            this.DGV_Steel.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing;
            this.DGV_Steel.ColumnHeadersHeight = (int)(this.DGV_Fabric.ColumnHeadersHeight * 1.3);
            this.DGV_Steel.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
        }

        /// <summary>
        /// ישר בעליית התוכנית מכין את כל ההדטה גריד
        /// </summary>
        public async Task<bool> InitSelectTabs()
        {
            await Task.Run(() => 
            {
                    
                if (tabControl1.InvokeRequired)
                {
                    tabControl1.Invoke(new MethodInvoker(delegate
                    {
                        DgvCellPainting(dataGridViewsList.IndexOf(DGVֹ_Mixtures),"תערובות");
                        DgvCellPainting(dataGridViewsList.IndexOf(DGV_Fabric),"בדים");
                        DgvCellPainting(dataGridViewsList.IndexOf(DGV_Steel),"פלדה");
                        int i = 0;
                        foreach (TabPage p in tabControl1.TabPages)
                        {
                            tabControl1.SelectedTab = p;
                            CellPaintRound2(i);
                            i++;
                        }
                        tabControl1.SelectedTab = TabPageMixures;
                        ;
                        
                    }));
                }
            });

            NewPaint = false;
            return true;
        }

        /// <summary>
        ///מילוי קומבובוקס חודשים- מסך התחלתי
        /// </summary>
        private void StartPositionFunc()
        {
            //רשימת datagrid
            dataGridViewsList.Add(DGVֹ_Mixtures);
            dataGridViewsList.Add(DGV_Fabric);
            dataGridViewsList.Add(DGV_Steel);
            //המשך איפוס מסך התחלתי
            for (int i = DateTime.Now.Year; i >= 2010; i--)
            {
                Cbo_YearReport.Items.Add(i);
            }
            Cbo_YearReport.SelectedIndex = 0;
            Cbo_MonthReport.Items.AddRange(CultureInfo.CurrentCulture.DateTimeFormat.MonthNames);//הוספת שמות חודשים
            Cbo_MonthReport.SelectedIndex = DateTime.Now.Month - 1;//חודש נוכחי בקומבובוקס
            btn_save.Enabled = false;//לא יוכל לשמור נתונים כל עוד לא ערך תא
            ReadOnlyRowsDataGrid();
            DataTable GetWasteCodes = WasteObject.GetWasteCodes();
            cbo_CodeWaste.DataSource = GetWasteCodes;
            cbo_CodeWaste.DisplayMember = "desc";
            cbo_CodeWaste.SelectedIndex = -1;
            for (int i = 0; i < dataGridViewsList.Count; i++)
            {
                dataGridViewsList[i].Columns[0].Width = 60;
                dataGridViewsList[i].Columns[1].Width = 90;
            }
            DGV_Steel.EnableHeadersVisualStyles = false;
            DGV_Steel.Columns["חוטי פלדה גולמיים"].HeaderCell.Style.BackColor = Color.PaleGreen;
            DGV_Steel.Columns["חישוקים"].HeaderCell.Style.BackColor = Color.LimeGreen;
            lbl_LastUpdate.Text = DateTime.Now.ToString("dd/MM/yyyy HH:mm");
        }

        /// <summary>
        /// הצגת טבלת פסולת לפי חודשים
        /// </summary>
        private void ShowWaste()
        {
            string MonthNumber;
            int Days, Year;

            //בתחילת התוכנית הצגה של החודש הנוכחי
            if (Start)
            {
                MonthNumber = DateTime.Now.Month.ToString();
                if (int.Parse(MonthNumber) < 10)
                    MonthNumber = "0" + MonthNumber;//הוספת 0 לחודש חד ספרתי
                Days = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month);//כמה ימים בחודש העכשווי
                WasteObject.InsertDatesValues(MonthNumber, DateTime.Now.Year, Days);//הכנסת תאריכים לשדות
                WasteObject.FillWasteTable();//שליפת נתוני פסולות מs400
                //שיבוץ בטבלאות תערובות בדים ופלדה
                WasteTableMixures = WasteObject.CreateTableMixures();//דיפולט-חודש ושנה של היום
                WasteTableFabric = WasteObject.CreateTableFabric();
                WasteTableSteel = WasteObject.CreateTableSteel();
                WasteTableTotal = WasteObject.CreateTableTotal();
            }

            else
            {
                string Month = Cbo_MonthReport.SelectedItem.ToString();
                //בניית תאריך רצוי עבור השאילתה
                Year = int.Parse(Cbo_YearReport.SelectedItem.ToString());//שנה וחודש לפי בחירת המשתמש
                MonthNumber = DateTime.ParseExact(Month, "MMMM", CultureInfo.CurrentCulture).Month.ToString();
                if (int.Parse(MonthNumber) < 10)
                    MonthNumber = "0" + MonthNumber;//הוספת 0 לחודש חד ספרתי
                Days = DateTime.DaysInMonth(Year, int.Parse(MonthNumber));//כמה ימים בחודש הנבחר
                WasteObject.InsertDatesValues(MonthNumber, Year, Days);//הכנסת תאריכים לשדות
                WasteObject.FillWasteTable();//שליפת נתוני פסולות מs400
                //מילוי כל הטבלאות
                WasteTableMixures = WasteObject.CreateTableMixures();
                WasteTableFabric = WasteObject.CreateTableFabric();
                WasteTableSteel = WasteObject.CreateTableSteel();
                WasteTableTotal = WasteObject.CreateTableTotal();

                //לבקשת בומה אם מדובר בחודשים שעברו אי אפשר לעדכן
                string MonthNumber2 = DateTime.ParseExact(Month, "MMMM", CultureInfo.CurrentCulture).Month.ToString();
                if (int.Parse(DateTime.Today.Month.ToString()) > int.Parse(MonthNumber2) || Year < DateTime.Now.Year)
                {
                    //אם אנחנו לפני ה5 בחודש מותר לעדכן על חודש לפני
                    if (int.Parse(DateTime.Today.Month.ToString()) - int.Parse(MonthNumber2) == 1 && Year == DateTime.Now.Year && DateTime.Today.Day <= 5)
                    {
                        DGVֹ_Mixtures.ReadOnly = false;
                        DGV_Fabric.ReadOnly = false;
                        DGV_Steel.ReadOnly = false;
                    }
                    else
                    {
                        DGVֹ_Mixtures.ReadOnly = true;
                        DGV_Fabric.ReadOnly = true;
                        DGV_Steel.ReadOnly = true;
                        DGVֹ_Mixtures.Columns["יום"].ReadOnly = true;
                        DGVֹ_Mixtures.Columns["תאריך"].ReadOnly = true;
                        DGV_Fabric.Columns["יום"].ReadOnly = true;
                        DGV_Fabric.Columns["תאריך"].ReadOnly = true;
                        DGV_Steel.Columns["יום"].ReadOnly = true;
                        DGV_Steel.Columns["תאריך"].ReadOnly = true;
                    }
                }
                else
                {
                    DGVֹ_Mixtures.ReadOnly = false;
                    DGV_Fabric.ReadOnly = false;
                    DGV_Steel.ReadOnly = false;
                    DGVֹ_Mixtures.Columns["יום"].ReadOnly = true;
                    DGVֹ_Mixtures.Columns["תאריך"].ReadOnly = true;
                    DGV_Fabric.Columns["יום"].ReadOnly = true;
                    DGV_Fabric.Columns["תאריך"].ReadOnly = true;
                    DGV_Steel.Columns["יום"].ReadOnly = true;
                    DGV_Steel.Columns["תאריך"].ReadOnly = true;
                }

            }
            //צביעת תאים שעודכנו לפי בדיקה מול טבלת MSVQTP בs400
            DataOfCellTable = WasteObject.GetCellsData();
            DGVֹ_Mixtures.DataSource = WasteTableMixures;
            DGV_Fabric.DataSource = WasteTableFabric;
            DGV_Steel.DataSource = WasteTableSteel;
            DGV_Total.DataSource = WasteTableTotal;
            ChangeColumnHeadersFabricTableName();
        }



        //פונקציות מעטפת

        /// <summary>
        /// שינוי שמות לכותרות של הטבלת בדים
        /// </summary>
        private void ChangeColumnHeadersFabricTableName()
        {
            DGV_Fabric.Columns["בוקר חתכן 1"].HeaderText = "ב";
            DGV_Fabric.Columns["ערב חתכן 1"].HeaderText = "ע";
            DGV_Fabric.Columns["לילה חתכן 1"].HeaderText = "ל";
            DGV_Fabric.Columns["ערב חתכן 3"].HeaderText = "ע";
            DGV_Fabric.Columns["בוקר חתכן 3"].HeaderText = "ב";
            DGV_Fabric.Columns["ערב חתכן 5"].HeaderText = "ע";
            DGV_Fabric.Columns["בוקר חתכן 5"].HeaderText = "ב";
            DGV_Fabric.Columns["לילה חתכן 5"].HeaderText = "ל";
            DGV_Total.Columns["FB-40002"].HeaderText = "תערובות";
            DGV_Total.Columns["SC-FC010"].HeaderText = "בדים";
            DGV_Total.Columns["SC-FC050"].HeaderText = "פלדה";
            DGV_Total.Columns["SC-BE200"].HeaderText = "חישוקים";
            DGV_Total.Columns["SC-SC100"].HeaderText = "חוטי פלדה גולמיים";
        }


        /// <summary>
        /// חסימת שורות של תאריכים עתידיים
        /// </summary>
        private void ReadOnlyRowsDataGrid()
        {
            DGVֹ_Mixtures.Columns["יום"].ReadOnly = true;
            DGVֹ_Mixtures.Columns["תאריך"].ReadOnly = true;
            DGV_Fabric.Columns["יום"].ReadOnly = true;
            DGV_Fabric.Columns["תאריך"].ReadOnly = true;
            DGV_Steel.Columns["יום"].ReadOnly = true;
            DGV_Steel.Columns["תאריך"].ReadOnly = true;
            for (int i = 0; i < dataGridViewsList.Count; i++)
            {
                dataGridViewsList[i].Rows[dataGridViewsList[i].Rows.Count - 2].ReadOnly = true;
                dataGridViewsList[i].Rows[dataGridViewsList[i].Rows.Count - 1].ReadOnly = true;
            }
        }

        /// <summary>
        /// יצירת טבלה עבור עדכוני כמויות 
        /// </summary>
        private void CreateTableForUpdate()
        {
            TableForUpdate.Columns.Add("תאריך");
            TableForUpdate.Columns.Add("מספר רץ");
            TableForUpdate.Columns.Add("קוד פריט");
            TableForUpdate.Columns.Add("מחלקה");
            TableForUpdate.Columns.Add("מכונה");
            TableForUpdate.Columns.Add("משמרת");
            TableForUpdate.Columns.Add("מרכז עבודה");
            TableForUpdate.Columns.Add("כמות");
            TableForUpdate.Columns.Add("ערך חדש");
            TableForUpdate.Columns.Add("סוג תנועה");//בעיקרון הכל px אלא אם יש עדכון ואז צריך לבטל ערך קודם עם p5
            TableForUpdate.Columns.Add("יחידת משקל");
            TableForUpdate.Columns.Add("קוד תקלה");
            TableForUpdate.Columns.Add("הערות");
        }


        /// <summary>
        /// כפתור הצג לפי חודשים
        /// </summary>
        private void btn_show_Click(object sender, EventArgs e)
        {
            WasteObject = new Waste();
            DataTable GetWasteCodes = WasteObject.GetWasteCodes();
            ShowWaste();            
        }



       /// <summary>
       /// איוונטים בדטה גרידים
       /// </summary>
 
            //תחילת עריכת תא


        private void DGVֹ_Mixtures_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            DgvCellBeginEdit(dataGridViewsList.IndexOf(DGVֹ_Mixtures), e.RowIndex, e.ColumnIndex);        
        }


        private void DGV_Fabric_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            DgvCellBeginEdit(dataGridViewsList.IndexOf(DGV_Fabric), e.RowIndex, e.ColumnIndex);
        }

        private void DGV_Steel_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            DgvCellBeginEdit(dataGridViewsList.IndexOf(DGV_Steel), e.RowIndex, e.ColumnIndex);
        }

        /// <summary>
        /// תחילת עריכת תא,נועד בשביל לבדוק שערך ריק או לא ואם לא ריק שומר ערך קודם
        /// </summary>
        private void DgvCellBeginEdit(int IndexOfDataGrid, int RowIndex, int ColumnIndex)
        {
            if (!string.IsNullOrEmpty(dataGridViewsList[IndexOfDataGrid].Rows[RowIndex].Cells[ColumnIndex].Value.ToString() as string))
            {
                ValueNew = false;//מדובר בערך ישן              
                PreviousValue = int.Parse(dataGridViewsList[IndexOfDataGrid].Rows[RowIndex].Cells[ColumnIndex].Value.ToString());
            }
            lbl_tap.Visible = true;
            lbl_tap.Text ="הינך מקליד עבור "+ dataGridViewsList[IndexOfDataGrid].Rows[RowIndex].Cells["תאריך"].Value.ToString();
        }

               ///אחרי עריכת תא 
       
            
        /// <summary>
        ///  mixtures עריכת כמות תערובות בתוך תא מסוים
        /// </summary>
        private void DGVֹ_Mixtures_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            DgvCellEndEdit(dataGridViewsList.IndexOf(DGVֹ_Mixtures),e.RowIndex,e.ColumnIndex);         
        }

        /// <summary>
        /// עריכת כמות תערובות בתוך תא מסוים fabric
        /// </summary>
        private void DGV_Fabric_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            DgvCellEndEdit(dataGridViewsList.IndexOf(DGV_Fabric), e.RowIndex, e.ColumnIndex);
            string str = DGV_Fabric.Columns[e.ColumnIndex].DataPropertyName;//השם המקורי של הטור
        }


        /// <summary>
        /// עריכת כמות תערובות בתוך תא מסוים steel
        /// </summary>
        private void DGV_Steel_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            DgvCellEndEdit(dataGridViewsList.IndexOf(DGV_Steel), e.RowIndex, e.ColumnIndex);
        }

        /// <summary>
        /// אחרי כתיבת כמות מסוימת בתא-מוסיף לטבלת העדכונים 
        /// </summary>
        private void DgvCellEndEdit(int IndexOfDataGrid, int RowIndex, int ColumnIndex)
        {
            int Quantity;
            string UnitValue;//יחידות משקל
            bool successfullyParsed = int.TryParse(dataGridViewsList[IndexOfDataGrid].Rows[RowIndex].Cells[ColumnIndex].Value.ToString(), out Quantity);
            WasteObject.CatalogNumber = dataGridViewsList[IndexOfDataGrid].AccessibleDescription;//קוד פריט של דטה גריד
            WasteObject.catalogNumToFind = WasteObject.ListRowCatalogNum.Find(x => x.CatalogNumber == WasteObject.CatalogNumber && x.Description == dataGridViewsList[IndexOfDataGrid].Columns[ColumnIndex].DataPropertyName);//  מחפש ברשימה של מספרים קטלוגיים את כל הפרטים של אותו טור שרשמתי בו לפי שם טור וטבלה ספציפית
            if ((!successfullyParsed) && (!string.IsNullOrEmpty(dataGridViewsList[IndexOfDataGrid].Rows[RowIndex].Cells[ColumnIndex].Value.ToString())))//לא הקליד מספר
            {
                MessageBox.Show("נא הקלד מספר", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (PreviousValue == 0)
                {
                    dataGridViewsList[IndexOfDataGrid].Rows[RowIndex].Cells[ColumnIndex].Value = "";
                }
                else
                {
                    dataGridViewsList[IndexOfDataGrid].Rows[RowIndex].Cells[ColumnIndex].Value = PreviousValue;
                }
            }

            else if (string.IsNullOrEmpty(dataGridViewsList[IndexOfDataGrid].Rows[RowIndex].Cells[ColumnIndex].Value.ToString()))//נכנס לתא ויצא ממנו בלי לרשום כלום
            {
                if (PreviousValue == 0)
                {
                    dataGridViewsList[IndexOfDataGrid].Rows[RowIndex].Cells[ColumnIndex].Value = "";
                }
                else
                {
                    dataGridViewsList[IndexOfDataGrid].Rows[RowIndex].Cells[ColumnIndex].Value = PreviousValue;
                }
            }


            else if (dataGridViewsList[IndexOfDataGrid].Rows[RowIndex].Cells[ColumnIndex].Value.ToString() != PreviousValue.ToString()) //הקליד ערך תקין 
            {   

                WasteObject.UnitMessaure.TryGetValue(WasteObject.CatalogNumber, out UnitValue);//מוסיף שזה יחידות מידה של קילו
                if (UnitValue == null) UnitValue = "KG";

                for (int i = 0; i < TableForUpdate.Rows.Count; i++)//אם משתמש התחרט ומכניס כמות אחרת,דורס את הקודם
                {
                    if (dataGridViewsList[IndexOfDataGrid].Rows[RowIndex].Cells["תאריך"].Value.ToString() == TableForUpdate.Rows[i]["תאריך"].ToString()
                       && WasteObject.catalogNumToFind.Department.ToString() == TableForUpdate.Rows[i]["מחלקה"].ToString()
                       && WasteObject.catalogNumToFind.Machine.ToString() == TableForUpdate.Rows[i]["מכונה"].ToString() && WasteObject.catalogNumToFind.Shift.ToString() == TableForUpdate.Rows[i]["משמרת"].ToString()
                       && TableForUpdate.Rows[i]["סוג תנועה"].ToString() == "PX" && WasteObject.catalogNumToFind.CatalogNumSon == TableForUpdate.Rows[i]["קוד פריט"].ToString())//שדה ייחודי של תאריך ומחלקה ככה נדע שהתבצע תיקון
                    {
                        if (Quantity == 0)//אם ביטל את הפעולה שרשם ממש עכשיו
                        {
                            TableForUpdate.Rows[i].Delete();
                            RowsPosition--;
                        }
                        else
                        {
                            TableForUpdate.Rows[i]["כמות"] = Quantity;
                            ValueNew = true;//אם מדובר בערך חדש שעדכנו פעמיים מאפס את הערך של התא למצב דיפולט
                            btn_save.Enabled = true;//עכשיו יוכל לשמור נתונים
                        }
                        PreviousValue = 0;
                        return;
                    }
                }
                RowsPosition++;
                //ערך חדש
                if (ValueNew == true)//אם מדובר בערך חדש לגמרי
                {
                    InsertTableForUpdate(true);
                    btn_save.Enabled = true;//עכשיו יוכל לשמור נתונים

                }
                else //ערך ישן שעדכנו
                {
                    if (PreviousValue != 0)
                    {
                        //שינוי 28.4 -במקום P5 לעשות PX עם מינוס
                        //ביטול הערך הקודם עם p5 בסימון מינוס הסכון
                        InsertTableForUpdate(false);
                    }


                    if (dataGridViewsList[IndexOfDataGrid].Rows[RowIndex].Cells[ColumnIndex].Value.ToString() != "0")//אם הקליד את המספר 0-אבל מעדכן רשומה שהייתה קיימת,לא יוסיף רשומה חדשה אלא רק יבטל את הקודמת                        
                    {

                        //תנועה חדשה
                        if (PreviousValue != 0) RowsPosition++;
                        InsertTableForUpdate(true);              
                    }

                    ValueNew = true;//איפוס ערך בוליאני של תא השתנה
                    btn_save.Enabled = true;//עכשיו יוכל לשמור נתונים
                }


            }
            PreviousValue = 0;

           void InsertTableForUpdate(bool New)
            {
                TableForUpdate.Rows.Add(dataGridViewsList[IndexOfDataGrid].Rows[RowIndex].Cells["תאריך"].Value.ToString());
                TableForUpdate.Rows[RowsPosition]["מספר רץ"] = RowsPosition + 1;
                TableForUpdate.Rows[RowsPosition]["קוד פריט"] = WasteObject.catalogNumToFind.CatalogNumSon;// "FB-40002";
                TableForUpdate.Rows[RowsPosition]["מחלקה"] = WasteObject.catalogNumToFind.Department;
                TableForUpdate.Rows[RowsPosition]["מכונה"] = WasteObject.catalogNumToFind.Machine;
                TableForUpdate.Rows[RowsPosition]["משמרת"] = WasteObject.catalogNumToFind.Shift;
                TableForUpdate.Rows[RowsPosition]["מרכז עבודה"] = WasteObject.catalogNumToFind.WorkCenter;
                if(New)//ערך חדש
                    TableForUpdate.Rows[RowsPosition]["כמות"] = Quantity;
                else
                    TableForUpdate.Rows[RowsPosition]["כמות"] = -PreviousValue;
                TableForUpdate.Rows[RowsPosition]["ערך חדש"] = false;
                //TableForUpdate.Rows[RowsPosition]["סוג תנועה"] = "P5";//
                TableForUpdate.Rows[RowsPosition]["סוג תנועה"] = "PX";
                TableForUpdate.Rows[RowsPosition]["יחידת משקל"] = UnitValue;
                TableForUpdate.Rows[RowsPosition]["קוד תקלה"] = "01";
            }
        }


                    ///אחרי לחיצה כפולה על תא        
 
        private void DGVֹ_Mixtures_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            TabPageToReturn = TabPageMixures;
            DgvCellDoubleClick(dataGridViewsList.IndexOf(DGVֹ_Mixtures), e.RowIndex, e.ColumnIndex,"");
        }

        private void DGV_Fabric_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            TabPageToReturn = TabPageFabric;
            DgvCellDoubleClick(dataGridViewsList.IndexOf(DGV_Fabric), e.RowIndex, e.ColumnIndex,"");
        }

        private void DGV_Steel_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            TabPageToReturn = TabPageSteel;
            string s = WasteObject.GetCatalogNumForsteel(DGV_Steel.Columns[e.ColumnIndex].HeaderText);
            DgvCellDoubleClick(dataGridViewsList.IndexOf(DGV_Steel), e.RowIndex, e.ColumnIndex,s);
          
        }
        /// <summary>
        /// חדש-עדכון כמות מתוך מסך עדכון פסולות
        /// </summary>
        private void DgvCellDoubleClick(int IndexOfDataGrid, int RowIndex, int ColumnIndex,string CatalogNumForsteel)
        {
            if (ColumnIndex == 0 || ColumnIndex == 1 || RowIndex==-1) return;
            if (RowIndex >= dataGridViewsList[IndexOfDataGrid].RowCount - 2) return;
            if (TableForUpdate.Rows.Count>0)//לפני עדכון תא שמירת נתונים עד עכשיו
            {
                SaveTableForUpdate();
            }
            WasteObject.CatalogNumber = dataGridViewsList[IndexOfDataGrid].AccessibleDescription;//קוד פריט של דטה גריד
            WasteObject.catalogNumToFind = WasteObject.ListRowCatalogNum.Find(x => x.CatalogNumber == WasteObject.CatalogNumber && x.Description == dataGridViewsList[IndexOfDataGrid].Columns[ColumnIndex].DataPropertyName);//  מחפש ברשימה של מספרים קטלוגיים את כל הפרטים של אותו טור שרשמתי בו לפי שם טור וטבלה ספציפית
            eRowIndex = RowIndex;//אחרי שמירת הנתונים תופיע הכמות בטבלה
            eColumnIndex = ColumnIndex;
            IndexOfDatagrid = IndexOfDataGrid;
            lblUpdate.Visible = false;
            txt_blue.Visible = false;
            AskToSave = true;
            btnUpdatePress = true;//בשביל שחרור טב פייג' של עדכון פסולת
            tabControl1.SelectedTab = TabPageWasteUpdate;//מעבר לעדכון 
            ClearScreen();
  
            lbl_WasteUpdate.Text = WasteObject.catalogNumToFind.TableType;//מספר קטלוגי ותיאור
            lbl_department.Text = dataGridViewsList[IndexOfDataGrid].Columns[ColumnIndex].DataPropertyName;
            lbl_dateWantedUpdate.Text = dataGridViewsList[IndexOfDataGrid].Rows[RowIndex].Cells["תאריך"].Value.ToString();//תאריך נבחר לעדכון
            WasteObject.DateOfCellUpdate = dataGridViewsList[IndexOfDataGrid].Rows[RowIndex].Cells["תאריך"].Value.ToString();//תאריך נבחר לעדכון
            if (!string.IsNullOrEmpty(dataGridViewsList[IndexOfDataGrid].Rows[RowIndex].Cells[ColumnIndex].Value.ToString()))//אם הערך עודכן בעבר
            {
                WasteObject.AmountPerDay = int.Parse(dataGridViewsList[IndexOfDataGrid].Rows[RowIndex].Cells[ColumnIndex].Value.ToString());//כמה יש מאותו יום ש
                WasteObject.PreviousAmount = WasteObject.AmountPerDay;
                LeftOver = WasteObject.AmountPerDay;//חזרה מאפס כמה יש מאותו יום שאנחנו רוצים לעדכן בו כמות פסולות
                CellEmpty = false;
                if (WasteObject.AmountPerDay == 0) CellEmpty = true;
            }
            else//תא ריק
            {
                CellEmpty = true;
            }
            TotalAmount = WasteObject.AmountPerDay;//כמה נשאר לפרט,בהתחלה נשאר הכל
            lbl_total.Text = Convert.ToDecimal(TotalAmount).ToString("#,#");
            if(TotalAmount==0)
            lbl_total.Text = TotalAmount.ToString();

            //טבלה חדשה עבור פירוט כמויות פסולת
            DescriptionWasteTable = new DataTable();
            DescriptionWasteTable.Columns.Add("סידורי");
            DescriptionWasteTable.Columns.Add("כמות");
            DescriptionWasteTable.Columns.Add("קוד");
            DescriptionWasteTable.Columns.Add("תיאור תקלה");
            DescriptionWasteTable.Columns.Add("הערות");
            DataGrid_Desc.DataSource = DescriptionWasteTable;
            this.DataGrid_Desc.Columns["סידורי"].ReadOnly = true;
            this.DataGrid_Desc.Columns["קוד"].ReadOnly = true;
            this.DataGrid_Desc.Columns["תיאור תקלה"].ReadOnly = true;
            this.DataGrid_Desc.Columns["הערות"].ReadOnly = true;

            if (dataGridViewsList[IndexOfDataGrid].Rows[RowIndex].Cells[ColumnIndex].Style.BackColor == Color.Turquoise)//אם מדובר בתא שמעודכן כבר-מעודכן צבוע בטורקיז
            {
                lbl_AmountWaste.Text = Convert.ToDecimal(WasteObject.AmountPerDay).ToString("#,#");//כמות פסולת כוללת מתוך אותו מק"ט באותו יום
                Tourqise = true;
                string WhichDay = "";
                int RowsUpdate = 0;//כמה שורות בטבלה
                for (int i = 0; i < DataOfCellTable.Rows.Count; i++)
                {
                    //DateTime d = DateTime.ParseExact(DataOfCellTable.Rows[i]["Date"].ToString(), "yyMMdd", CultureInfo.InvariantCulture);
                    //WhichDay = d.ToString("dd/MM/yy");
                    WhichDay = DataOfCellTable.Rows[i]["Date"].ToString().Substring(4, 2);
                    if (WasteObject.catalogNumToFind.CatalogNumSon == DataOfCellTable.Rows[i]["CatalogNumber"].ToString().Trim() && dataGridViewsList[IndexOfDataGrid].Rows[RowIndex].Cells["תאריך"].Value.ToString().Split('/').First() == WhichDay
                        && WasteObject.catalogNumToFind.Department.ToString() == DataOfCellTable.Rows[i]["Department"].ToString() && DataOfCellTable.Rows[i]["Quantity"].ToString()!="0"
                         && WasteObject.catalogNumToFind.Machine.Trim() == DataOfCellTable.Rows[i]["Machine"].ToString().Trim() && WasteObject.catalogNumToFind.Shift.ToString()== DataOfCellTable.Rows[i]["Shift"].ToString().Trim())//מק"ט רלוונטי שתואם תערובות ותאריך שתואם את תאריך שבדטה בייס
                    {
                        DataRow row = DescriptionWasteTable.NewRow();//הכנסה לטבלה את כל העדכונים שנעשו
                        DescriptionWasteTable.Rows.Add(row);
                        DescriptionWasteTable.Rows[RowsUpdate]["כמות"] = int.Parse(double.Parse(DataOfCellTable.Rows[i]["Quantity"].ToString()).ToString());
                        DescriptionWasteTable.Rows[RowsUpdate]["קוד"] = DataOfCellTable.Rows[i]["ReasonCode"].ToString();
                        DescriptionWasteTable.Rows[RowsUpdate]["סידורי"] = RowsUpdate + 1;
                        DescriptionWasteTable.Rows[RowsUpdate]["תיאור תקלה"] = DataOfCellTable.Rows[i]["Description"].ToString();
                        DescriptionWasteTable.Rows[RowsUpdate]["הערות"] = DataOfCellTable.Rows[i]["Comment"].ToString();
                        LeftOver -= int.Parse(double.Parse(DataOfCellTable.Rows[i]["quantity"].ToString()).ToString());//מעדכן כמה נשאר מהטוטל
                        RowsUpdate++;
                    }
                }
                DataGrid_Desc.DataSource = DescriptionWasteTable;
                btnSaveDescForCell.Enabled = true;
                RowsUpdate = 0;
                CellUpdated = true;//מעדכן לרשומה שעודכנה במידה וירצה לשנות אותה באירוע cellEditEnd
            }
            else//אם לא טורקיז-עדכון חדש
            {
                CellUpdated = false;
                Tourqise = false;
            }
            btnUpdatePress = false;
            btnSaveDescForCell.Visible = true;//כפתור שמירת נתוני פסולות מפורט יופיע פתאום
            TableFor1DetailedCell = new DataTable();//נוצרת טבלה עדכון פסולות חדשה עבור התא
            TableFor1DetailedCell = TableForUpdate.Clone();//אותו מבנה
        }

        public bool check(int i)
        {
            if (DataOfCellTable.Rows[i]["Shift"].ToString() == "32")
                return true;
            else return false;
        }


        //צביעת תאים


        private void DGVֹ_Mixtures_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex >= DGVֹ_Mixtures.Rows.Count - 2 )
            {
                e.CellStyle.Font = new Font("Tahoma", 12, FontStyle.Bold);
                e.CellStyle.BackColor = Color.FromArgb(255, 128, 0);
                if(e.ColumnIndex == 0)
                e.AdvancedBorderStyle.Right = DataGridViewAdvancedCellBorderStyle.None;
                DGVֹ_Mixtures.Rows[DGVֹ_Mixtures.Rows.Count - 2].Height = 35;
                DGVֹ_Mixtures.Rows[DGVֹ_Mixtures.Rows.Count - 1].Height = 35;
            }
        }


        private void DGV_Fabric_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex >= DGV_Fabric.Rows.Count - 2)
            {
                e.CellStyle.Font = new Font("Tahoma", 12, FontStyle.Bold);
                e.CellStyle.BackColor = Color.Yellow;
                if (e.ColumnIndex == 0)
                    e.AdvancedBorderStyle.Right = DataGridViewAdvancedCellBorderStyle.None;
                DGV_Fabric.Rows[DGV_Fabric.Rows.Count - 2].Height = 35;
                DGV_Fabric.Rows[DGV_Fabric.Rows.Count - 1].Height = 35;
            }
        }

        private void DGV_Steel_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex >= DGV_Steel.Rows.Count - 2)
            {
                e.CellStyle.Font = new Font("Tahoma", 12, FontStyle.Bold);
                e.CellStyle.BackColor = Color.YellowGreen;
                if (e.ColumnIndex == 0)
                    e.AdvancedBorderStyle.Right = DataGridViewAdvancedCellBorderStyle.None;
                DGV_Steel.Rows[DGV_Fabric.Rows.Count - 2].Height = 35;
                DGV_Steel.Rows[DGV_Fabric.Rows.Count - 1].Height = 35;
            }
        }

        /// <summary>
        /// מכניס את התאיםלרשימה של אובייקטים מסוג CELLPAINT שצריכים להיצבע בתכלת--תאים מפורטים
        /// </summary>
        /// <param name="IndexOfDataGrid"></param>
        private void DgvCellPainting(int IndexOfDataGrid,string TableType)
        {
            CellPaint cellPaint;
            string ForRowFilter= "CatalogNumber IN(";
            if (NewPaint)
            {
                //צביעת תאים שעודכנו כבר
                //הם יכנסו לתוך רשימה של אובייקטים מסוג cellpaint
                int WhichDay;//יום להכניס-אינדקס שורה
                string ColumnName;//שם טור-אינדקס עמודה

                WasteObject.CatalogFilterTable.RowFilter = $@"TableType ='{TableType}'";//מסנן רשומות של תערובות/בדים/פלדה
                //לוקח את כל המקטים הבנים של הסוג הרלוונטי תערובות/בדים/פלדה 
                WasteObject.CatalogSons =WasteObject.CatalogFilterTable.ToTable(true, "CatalogNumberSon");//עובר על כל המקטים הבנים של המקט תערובות ומכניס אותם לטבלת התערובות
                foreach (DataRow rowView in WasteObject.CatalogSons.Rows)
                {
                    ForRowFilter+= $@"'{rowView["CatalogNumberSon"]}',";//משרשר את כל המקטים הבנים שישמשו לROW FILTER
                }
                ForRowFilter= ForRowFilter.Remove(ForRowFilter.Length - 1);
                ForRowFilter += ")";

                DataView filterTable = new DataView(DataOfCellTable);
                filterTable.RowFilter = ForRowFilter;

                foreach (DataRowView row in filterTable)
                {                      
                    DateTime d = DateTime.ParseExact(row["date"].ToString(), "yyMMdd", CultureInfo.InvariantCulture);//יום
                    WhichDay = int.Parse(d.ToString("yyMMdd").Substring(4).TrimStart(new Char[] { '0' }));//מוריד 0 מוביל
                    WasteObject.catalogNumToFind = WasteObject.ListRowCatalogNum.Find(x => x.CatalogNumSon == row["CatalogNumber"].ToString().Trim() && x.Department == int.Parse(row["Department"].ToString()) && x.Machine.Trim() == row["Machine"].ToString().Trim() && x.Shift == int.Parse(row["Shift"].ToString()));
                    if (WasteObject.catalogNumToFind != null)
                    {
                        ColumnName = WasteObject.catalogNumToFind.Description;//מחפש את שם הטור שאנחנו צריכים עבור האינדקס

                        if (row["ReasonCode"].ToString() != "01" && row["Quantity"].ToString() != "0")//אם הכמות 0 לא צריך להיצבע. אם הסיבת תקלה היא רגילה 01 אז היא לא תיצבע
                        {
                            switch (IndexOfDataGrid)
                            {
                                case 0:
                                    cellPaint = new CellPaint(WhichDay - 1, dataGridViewsList[IndexOfDataGrid].Columns[ColumnName].Index);
                                    cellPaintsListMixures.Add(cellPaint);
                                    break;

                                case 1:
                                    cellPaint = new CellPaint(WhichDay - 1, dataGridViewsList[IndexOfDataGrid].Columns[ColumnName].Index);
                                    cellPaintsListFabric.Add(cellPaint);
                                    break;

                                case 2:
                                    cellPaint = new CellPaint(WhichDay - 1, dataGridViewsList[IndexOfDataGrid].Columns[ColumnName].Index);
                                    cellPaintsListSteel.Add(cellPaint);
                                    break;
                            }
                        }
                    }

                }
             
            }

            //צביעת תא שעודכן עכשיו
            if (CellUpdated && SaveData)
                {
                    if (string.IsNullOrEmpty(DataGrid_Desc.Rows[0].Cells[0].EditedFormattedValue as string))//אם טבלה ריקה זה אומר שלא היה עדכון
                        dataGridViewsList[IndexOfDataGrid].Rows[eRowIndex].Cells[eColumnIndex].Style.BackColor = Color.White;
                    else
                        dataGridViewsList[IndexOfDataGrid].Rows[eRowIndex].Cells[eColumnIndex].Style.BackColor = Color.Turquoise;
                    SaveData = false;//יחסום צביעת תא אם לא שמרנו נתונים
                }

            
                      
        }

        ///// <summary>
        ///// סבב שני של צביעה אחרי ששמר את התאים הצבועים
        ///// </summary>
        ///// <param name="IndexOfDataGrid"></param>
        public void CellPaintRound2(int IndexOfDataGrid)
        {
            switch (IndexOfDataGrid)
            {
                case 0:
                    for (int i = 0; i < cellPaintsListMixures.Count; i++)
                    {
                        DGVֹ_Mixtures.Rows[cellPaintsListMixures[i].Row].Cells[cellPaintsListMixures[i].Column].Style.BackColor = Color.Turquoise;
                    }
                    break;

                case 1:
                    for (int i = 0; i < cellPaintsListFabric.Count; i++)
                    {
                        DGV_Fabric.Rows[cellPaintsListFabric[i].Row].Cells[cellPaintsListFabric[i].Column].Style.BackColor = Color.Turquoise;
                    }
                    break;

                case 2:
                    for (int i = 0; i < cellPaintsListSteel.Count; i++)
                    {
                        DGV_Steel.Rows[cellPaintsListSteel[i].Row].Cells[cellPaintsListSteel[i].Column].Style.BackColor = Color.Turquoise;
                    }
                    break;
            }
        }






        //חסימת תאים

        /// <summary>
        /// חסימת עריכת תאים במידה ומדובר בתאריכים עתידיים
        /// </summary>
        private void DGVֹ_Mixtures_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridViewsList.Count>0)
                DgvCellEnter(dataGridViewsList.IndexOf(DGVֹ_Mixtures), e.RowIndex, e.ColumnIndex);           
        }


        private void DGV_Fabric_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridViewsList.Count > 0)
                DgvCellEnter(dataGridViewsList.IndexOf(DGV_Fabric), e.RowIndex, e.ColumnIndex);
        }

        private void DGV_Steel_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridViewsList.Count > 0)
                DgvCellEnter(dataGridViewsList.IndexOf(DGV_Steel), e.RowIndex, e.ColumnIndex);
        }


        private void DgvCellEnter(int IndexOfDataGrid, int RowIndex, int ColumnIndex)
        {
            if (dataGridViewsList[IndexOfDataGrid].Rows[RowIndex].Cells[ColumnIndex].Style.BackColor == Color.Turquoise)//מתחיל לערוך כמות של תא מפורט כבר
            {
                dataGridViewsList[IndexOfDataGrid].Rows[RowIndex].Cells[ColumnIndex].ReadOnly = true;
                return;
            }
            if (string.IsNullOrEmpty(dataGridViewsList[IndexOfDataGrid].Rows[RowIndex].Cells["תאריך"].Value.ToString())) return;
            if (DateTime.Now < DateTime.Parse(dataGridViewsList[IndexOfDataGrid].Rows[RowIndex].Cells["תאריך"].Value.ToString()))
                dataGridViewsList[IndexOfDataGrid].Rows[RowIndex].ReadOnly = true;
        }


        //שינוי ערך תא קיים

        private void DGVֹ_Mixtures_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            ValueChanged = true;
        }


        private void DGV_Fabric_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            ValueChanged = true;
        }

        private void DGV_Steel_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            ValueChanged = true;
        }

                //יצירת כותרות שתופסות על כמה טורים

        /// <summary>
        /// עיצוב של תתי כותרות עם rectangle
        /// </summary>
        private void DGV_Fabric_Paint(object sender, PaintEventArgs e)
        {
            Rectangle r1, r2, r3;
            int CountRectangle = 0;//מיועד לשמות המלבנים שיצרתי,כל מלבן שם אחר
            string str = "";//שמות מלבנים
            for (int i = 0; i < DGV_Fabric.ColumnCount; i++)
            {
                string headerCaptionText = DGV_Fabric.Columns[i].HeaderText;
                if (headerCaptionText.Length == 1)//קיצור של ב, ע, ל
                {
                    r1 = this.DGV_Fabric.GetCellDisplayRectangle(i, -1, true);
                    if (DGV_Fabric.Columns[i + 1].HeaderText.Length == 1)//גם הבא מכיל רק אות אחת
                    {
                        r2 = this.DGV_Fabric.GetCellDisplayRectangle(i + 1, -1, true);
                        r1.Width += r2.Width;
                        i++;
                        CountRectangle++;//מלבן ראשון שם ראשון וכך הלאה
                    }
                    if (DGV_Fabric.Columns[i + 1].HeaderText.Length == 1)
                    {
                        r3 = this.DGV_Fabric.GetCellDisplayRectangle(i + 2, -1, true);
                        r1.Width += r3.Width;
                        i++;
                    }
                    //r1.Width += r2.Width + r3.Width;

                    r1.X += 1;
                    r1.Y += 1;
                    r1.Height = r1.Height / 2 - 2;
                    r1.Width -= 2;

                    //using (Brush BackColor = new SolidBrush(this.DGV_Fabric.ColumnHeadersDefaultCellStyle.BackColor))
                    SolidBrush BackColor = new SolidBrush(Color.Yellow);
                    using (Brush ForeColor = new SolidBrush(this.DGV_Fabric.ColumnHeadersDefaultCellStyle.ForeColor))
                    using (Pen p = new Pen(this.DGV_Fabric.GridColor))

                    using (StringFormat format = new StringFormat())
                    {
                        switch (CountRectangle)
                        {
                            case 1:
                                str = "חתכן 1 בנדים";
                                break;

                            case 2:
                                str = "חתכן 3";
                                break;

                            case 3:
                                str = "חתכן 5 אופקי";
                                break;
                        }

                        format.Alignment = StringAlignment.Center;
                        format.LineAlignment = StringAlignment.Center;
                        e.Graphics.FillRectangle(BackColor, r1);
                        e.Graphics.DrawRectangle(p, r1);
                        e.Graphics.DrawString(str, this.DGV_Fabric.ColumnHeadersDefaultCellStyle.Font, ForeColor, r1, format);
                    }
                }
            }       
        }

        /// עיצוב של תתי כותרות עם rectangle
        private void DGV_Steel_Paint(object sender, PaintEventArgs e)
        {
            Rectangle r1, r2, r3, r4;
            for (int i = 0; i < DGV_Steel.ColumnCount; i++)
            {
                string headerCaptionText = DGV_Steel.Columns[i].HeaderText;
                if (headerCaptionText == "תקלות")
                {
                    r1 = this.DGV_Steel.GetCellDisplayRectangle(i, -1, true);
                    r2 = this.DGV_Steel.GetCellDisplayRectangle(i + 1, -1, true);
                    r3 = this.DGV_Steel.GetCellDisplayRectangle(i + 2, -1, true);
                    r4 = this.DGV_Steel.GetCellDisplayRectangle(i + 3, -1, true);
                    r1.Width += r2.Width + r3.Width + r4.Width;

                    r1.X += 1;
                    r1.Y += 1;
                    r1.Height = r1.Height / 2 - 2;
                    r1.Width -= 2;

                    //using (Brush BackColor = new SolidBrush(this.DGV_Fabric.ColumnHeadersDefaultCellStyle.BackColor))
                    SolidBrush BackColor = new SolidBrush(Color.GreenYellow);
                    using (Brush ForeColor = new SolidBrush(this.DGV_Steel.ColumnHeadersDefaultCellStyle.ForeColor))
                    using (Pen p = new Pen(this.DGV_Steel.GridColor))

                    using (StringFormat format = new StringFormat())
                    {
                        string str = "ספדון+VMI";

                        format.Alignment = StringAlignment.Center;
                        format.LineAlignment = StringAlignment.Center;
                        e.Graphics.FillRectangle(BackColor, r1);
                        e.Graphics.DrawRectangle(p, r1);
                        e.Graphics.DrawString(str, this.DGV_Steel.ColumnHeadersDefaultCellStyle.Font, ForeColor, r1, format);
                    }
                    break;
                }
            }
        }

        /// <summary>
        /// בשביל עיצוב הכותרות בדטה גריד
        /// </summary>
        private void InvalidateHeader()
        {
            Rectangle rtHeader = this.DGV_Fabric.DisplayRectangle;
            rtHeader.Height = this.DGV_Fabric.ColumnHeadersHeight / 2;
            this.DGV_Fabric.Invalidate(rtHeader);
        }

        private void DGV_Fabric_Resize(object sender, EventArgs e)
        {
            this.InvalidateHeader();
        }

        private void DGV_Fabric_ColumnWidthChanged(object sender, DataGridViewColumnEventArgs e)
        {
            this.InvalidateHeader();
        }

        private void DGV_Fabric_Scroll(object sender, ScrollEventArgs e)
        {
            this.InvalidateHeader();
        }


        ///כפתור שמירת נתונים


        private void btn_save_Click(object sender, EventArgs e)
        {
                //פירוט פסולות מפורט עבור תא
                if (tabControl1.SelectedTab == TabPageWasteUpdate)
                {
                    SaveCellData();                     
                }

                //שמירת נתונים כלליים 
                else
                {
                   SaveTableForUpdate();
                }

            timer1.Start();//מחכה 10 שניות ושולף טבלת ריכוזי פסולות חדשה
            lbl_LastUpdate.Text = DateTime.Now.ToString("dd/MM/yyyy HH:mm");
            lbl_LoadData.Visible = true;
        }

        /// <summary>
        /// שמירת טבלת עדכונים
        /// </summary>
        public void SaveTableForUpdate()
        {
            //DialogResult dialogResult = MessageBox.Show("? האם ברצונך לשמור נתונים", "שמירת נתונים", MessageBoxButtons.YesNo);
            //if (dialogResult == DialogResult.Yes)
            //{
                bool Saved = WasteObject.InsertDataToNin(TableForUpdate);//הכנסה לקובץ ITH ומשתנה האם הצליחה השמירה
                btn_save.Enabled = false;
                if (Saved)//איפוס טבלת עדכונים
                {
                    TableForUpdate.Rows.Clear();
                    RowsPosition = -1;
                }

            //}
        }

        //מסך עדכון מפורט של פסולות

        private void cbo_CodeWaste_SelectedIndexChanged(object sender, EventArgs e)
        {
            CboCodeWaste = true;
            if (cbo_CodeWaste.SelectedIndex == 0) //הערה
            {
                txt_comment.Visible = true;
                lbl_comment.Visible = true;
            }
            else
            {
                txt_comment.Visible = false;
                lbl_comment.Visible = false;
            }
            //if (RealeseComboWasteCode)
            //{
            //    if (!string.IsNullOrEmpty(cbo_CodeWaste.Text))
            //    {
            //        WasteObject.CodeAndDescWaste.TryGetValue(cbo_CodeWaste.Text, out value);
            //        //txt_CodeWaste.Text = value.Trim();
            //    }
            //}
        }

        private void txt_AmountSpesific_TextChanged(object sender, EventArgs e)
        {
            btn_confirm.Enabled = true;
        }


        /// <summary>
        /// הוסף רשומת פירוט פסולת לטבלה
        /// </summary>
        private void btn_confirm_Click(object sender, EventArgs e)
        {
            try
            {
            if (CboCodeWaste == false || txt_AmountSpesific.Text == "") return;
            bool isNumeric = int.TryParse(txt_AmountSpesific.Text, out int n);                
            if (!isNumeric) { MessageBox.Show("נא הקלד מספר");return; } 
            if(cbo_CodeWaste.SelectedIndex==-1) { MessageBox.Show("נא בחר קוד תקלה "); return; }
            DataRow row = DescriptionWasteTable.NewRow();
            DescriptionWasteTable.Rows.Add(row);
            DescriptionWasteTable.Rows[DescriptionWasteTable.Rows.Count - 1]["סידורי"] = DescriptionWasteTable.Rows.Count.ToString();
            DescriptionWasteTable.Rows[DescriptionWasteTable.Rows.Count - 1]["כמות"] = txt_AmountSpesific.Text;
            DescriptionWasteTable.Rows[DescriptionWasteTable.Rows.Count - 1]["קוד"] = cbo_CodeWaste.Text.Substring(0,2);
            DescriptionWasteTable.Rows[DescriptionWasteTable.Rows.Count - 1]["תיאור תקלה"] = cbo_CodeWaste.Text.Substring(3);
            DescriptionWasteTable.Rows[DescriptionWasteTable.Rows.Count - 1]["הערות"] = txt_comment.Text;
            DataGrid_Desc.DataSource = DescriptionWasteTable;


                //חישוב כמה נשאר
                LeftOver -= int.Parse(DescriptionWasteTable.Rows[DescriptionWasteTable.Rows.Count - 1]["כמות"].ToString());
                if (LeftOver == 0)
                    lbl_total.Text = "0";
                else
                    lbl_total.Text = Convert.ToDecimal(LeftOver).ToString("#,#");

                if (!CellEmpty)//אם היה מדובר בתא לא ריק
                {
                    WasteObject.AmountPerDay -= int.Parse(DescriptionWasteTable.Rows[DescriptionWasteTable.Rows.Count - 1]["כמות"].ToString());
                }

                //חישוב כמה יש
                TotalAmount += int.Parse(DescriptionWasteTable.Rows[DescriptionWasteTable.Rows.Count - 1]["כמות"].ToString());
                if (TotalAmount == 0)
                    lbl_total.Text = "0";
                else
                    lbl_total.Text = Convert.ToDecimal(LeftOver).ToString("#,#");

                //הוספת לטבלת השמירה
                AddRowToTableForUpdate("PX",false, DescriptionWasteTable.Rows.Count - 1);
            }
            catch (Exception ex)
            {
                MessageBox.Show("שגיאה");
            }
            
            btnSaveDescForCell.Enabled = true;//כעת ניתן לשמור נתונים
            ClearScreen();
        }

        /// <summary>
        /// 11.04 מוסיף שורה לטבלת העדכונים
        /// </summary>
        private void AddRowToTableForUpdate(string TransactionType,bool PreviousAmount,int RowIndex)
        {
            int Quantity=0;
            bool successfullyParsed;
            string ReasonCode= DescriptionWasteTable.Rows[RowIndex]["קוד"].ToString();
            string Comment= DescriptionWasteTable.Rows[RowIndex]["הערות"].ToString();
            WasteObject.UnitMessaure.TryGetValue(WasteObject.CatalogNumber, out string value);
            if (TransactionType == "CANCEL" && PreviousAmount == false)//ערך ישן
                Quantity = -PreviousNumOfcell;
            else if (TransactionType == "PX")//רגיל
                successfullyParsed = int.TryParse(DescriptionWasteTable.Rows[RowIndex]["כמות"].ToString(), out Quantity);
            else if (TransactionType == "CANCEL" && PreviousAmount == true)//ביטול תא שלא פורט ועכשיו כן
            {
                Quantity = -WasteObject.PreviousAmount;
                ReasonCode = "01";
                Comment = "";
            }
            
            RowsPositionForDetailedCell++;
            //ערך חדש
            //if (ValueNew == true)//אם מדובר בערך חדש לגמרי
            //{
                TableFor1DetailedCell.Rows.Add(WasteObject.DateOfCellUpdate);
                TableFor1DetailedCell.Rows[RowsPositionForDetailedCell]["מספר רץ"] = RowsPositionForDetailedCell + 1;
                TableFor1DetailedCell.Rows[RowsPositionForDetailedCell]["קוד פריט"] = WasteObject.catalogNumToFind.CatalogNumSon;//"FB-40002";
                TableFor1DetailedCell.Rows[RowsPositionForDetailedCell]["מחלקה"] = WasteObject.catalogNumToFind.Department;//אם מדובר במילה ממיר למספר,לדוגמא ברומו שרוף
                TableFor1DetailedCell.Rows[RowsPositionForDetailedCell]["מכונה"] = WasteObject.catalogNumToFind.Machine;
                TableFor1DetailedCell.Rows[RowsPositionForDetailedCell]["משמרת"] = WasteObject.catalogNumToFind.Shift;
                TableFor1DetailedCell.Rows[RowsPositionForDetailedCell]["מרכז עבודה"] = WasteObject.catalogNumToFind.WorkCenter;
                TableFor1DetailedCell.Rows[RowsPositionForDetailedCell]["כמות"] = Quantity;
                TableFor1DetailedCell.Rows[RowsPositionForDetailedCell]["ערך חדש"] = true;
                TableFor1DetailedCell.Rows[RowsPositionForDetailedCell]["סוג תנועה"] = TransactionType;
                TableFor1DetailedCell.Rows[RowsPositionForDetailedCell]["יחידת משקל"] = value;
                TableFor1DetailedCell.Rows[RowsPositionForDetailedCell]["קוד תקלה"] = ReasonCode;
                TableFor1DetailedCell.Rows[RowsPositionForDetailedCell]["הערות"] = Comment;
            //RowsPositionForDetailedCell++; 
            //}
            Comment = "";
        }


        /// <summary>
        /// מנקה מסך קודם אחרי דאבל קליק של תא
        /// </summary>
        private void ClearScreen()
        {
            txt_AmountSpesific.Text = "";
            cbo_CodeWaste.SelectedIndex = -1;
            txt_comment.Text = "";
            lbl_AmountWaste.Text = "";
            lbl_total.Text = "";
            LeftOver = 0;
            WasteObject.AmountPerDay=0;
        }



        private void DataGrid_Desc_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            PreviousNumOfcell = int.Parse(DataGrid_Desc.Rows[e.RowIndex].Cells["כמות"].Value.ToString());
        }

        /// <summary>
        /// עדכון של תא בפירוטי פסולות
        /// </summary>
        private void DataGrid_Desc_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                int num;//אם שינה כמות בטבלה
                if (e.ColumnIndex == 0 || e.ColumnIndex == 2) return;
                for (int i = 0; i < DataGrid_Desc.RowCount - 1; i++)
                {
                    bool isNumeric = int.TryParse(DataGrid_Desc.Rows[i].Cells["כמות"].Value.ToString(), out num);
                    if (isNumeric)
                    {
                        LeftOver -= num;
                    }
                    else
                    {
                        MessageBox.Show("אנא כתוב מספר לבד בעמודת כמות");
                        DataGrid_Desc.Rows[i].Cells["כמות"].Value = PreviousNumOfcell;//מספר שהיה מקודם אם לא הקליד טוב
                        return;
                    }
                }
                AddRowToTableForUpdate("CANCEL",false,e.RowIndex);//הוספת רשומה של ביטול
                AddRowToTableForUpdate("PX",false,e.RowIndex);//הוספת רשומה חדשה
                lbl_total.Text = Convert.ToDecimal(LeftOver).ToString("#,#");
                if (LeftOver < 0) lbl_total.ForeColor = System.Drawing.Color.Red;
            }
            catch(Exception ex)
            {
                MessageBox.Show("שגיאה בהזנת נתונים");
                    
            }
        }


        private void DataGrid_Desc_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(DataGrid_Desc.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString())) DataGrid_Desc.Rows[e.RowIndex].Cells[e.ColumnIndex].ReadOnly = true;
            }
            catch (Exception ex)
            {

            }
        }


        private void DataGrid_Desc_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
            LeftOver += int.Parse(DataGrid_Desc.Rows[e.Row.Index].Cells["כמות"].Value.ToString());
            lbl_total.Text = LeftOver.ToString();
            if (LeftOver > 0) lbl_total.ForeColor = System.Drawing.Color.Black;
            else if (LeftOver < 0) lbl_total.ForeColor = System.Drawing.Color.Red;
            if(CellUpdated)
            {
                RowsPositionForDetailedCell++;
                TableFor1DetailedCell.Rows.Add(WasteObject.DateOfCellUpdate);
                TableFor1DetailedCell.Rows[RowsPositionForDetailedCell]["מספר רץ"] = RowsPositionForDetailedCell + 1;
                TableFor1DetailedCell.Rows[RowsPositionForDetailedCell]["קוד פריט"] = WasteObject.catalogNumToFind.CatalogNumSon;//"FB-40002";
                TableFor1DetailedCell.Rows[RowsPositionForDetailedCell]["מחלקה"] = WasteObject.catalogNumToFind.Department;
                TableFor1DetailedCell.Rows[RowsPositionForDetailedCell]["מכונה"] = WasteObject.catalogNumToFind.Machine;
                TableFor1DetailedCell.Rows[RowsPositionForDetailedCell]["משמרת"] = WasteObject.catalogNumToFind.Shift;
                TableFor1DetailedCell.Rows[RowsPositionForDetailedCell]["מרכז עבודה"] = WasteObject.catalogNumToFind.WorkCenter;
                if(double.Parse(DataGrid_Desc.Rows[e.Row.Index].Cells["כמות"].Value.ToString()) >0)
                    TableFor1DetailedCell.Rows[RowsPositionForDetailedCell]["כמות"] = "-"+ DataGrid_Desc.Rows[e.Row.Index].Cells["כמות"].Value.ToString();
                else
                    TableFor1DetailedCell.Rows[RowsPositionForDetailedCell]["כמות"] = double.Parse( DataGrid_Desc.Rows[e.Row.Index].Cells["כמות"].Value.ToString())*-1;
                TableFor1DetailedCell.Rows[RowsPositionForDetailedCell]["ערך חדש"] = false;
                TableFor1DetailedCell.Rows[RowsPositionForDetailedCell]["סוג תנועה"] = "PX";
                WasteObject.UnitMessaure.TryGetValue(WasteObject.CatalogNumber, out string value);
                TableFor1DetailedCell.Rows[RowsPositionForDetailedCell]["יחידת משקל"] = value;
                TableFor1DetailedCell.Rows[RowsPositionForDetailedCell]["קוד תקלה"] = DataGrid_Desc.Rows[e.Row.Index].Cells["קוד"].Value;

            }
        }

        /// <summary>
        /// כפתור שמירת נתונים שמיועד רק עבור פירוט של תא מסוים(כפתור שמירת נתונים רגיל נמצא מתחתיו)
        /// </summary>
        private void btnSaveDescForCell_Click(object sender, EventArgs e)
        {
            SaveCellData();
            lbl_LoadData.Visible = true;
        }

        /// <summary>
        /// שמירת נתונים עבור תא מסוים
        /// </summary>
        private void SaveCellData()
        {
            try
            {
                if (!CellEmpty && WasteObject.AmountPerDay > 0 && Tourqise)//לא עדכנו כמות תואמת למה שהיה לפני הפירוט פסולות
                {
                    DialogResult d = MessageBox.Show($@"לשמור שינויים?", "אישור נתונים", MessageBoxButtons.YesNo);// {WasteObject.AmountPerDay} קג האם ברצונך לשמור בכל זאת נתונים
                    if (d == DialogResult.No)
                    {
                        AskToSave = false;
                        return;
                    }
                }

                SaveData = true;//לצביעת תא לטורקיז
                lblUpdate.Visible = true;
                txt_blue.Visible = true;
                btnSaveDescForCell.Visible = false;
                dataGridViewsList[IndexOfDatagrid].Rows[eRowIndex].Cells[eColumnIndex].Value = "מעדכן נתונים";
                dataGridViewsList[IndexOfDatagrid].Rows[eRowIndex].Cells[eColumnIndex].Style.BackColor = Color.Turquoise;
                //אם פירטנו תא שלא היה מפורט לפני נמחק אותו
                if (!CellEmpty)
                {
                    bool exist = false;
                    for (int i = 0; i < TableForUpdate.Rows.Count; i++)//אם מדובר בתא חדש ולא שמור לא צריך לעשות p5
                    {
                        if (TableForUpdate.Rows[i]["קוד פריט"].ToString() == WasteObject.catalogNumToFind.CatalogNumSon && TableForUpdate.Rows[i]["מחלקה"].ToString() == WasteObject.catalogNumToFind.Department.ToString()
                            && TableForUpdate.Rows[i]["תאריך"].ToString()==WasteObject.DateOfCellUpdate && TableForUpdate.Rows[i]["משמרת"].ToString()== WasteObject.catalogNumToFind.Shift.ToString()
                            && TableForUpdate.Rows[i]["מכונה"].ToString()== WasteObject.catalogNumToFind.ToString())
                        {
                            exist = true;//קיים תא שנרשם בו אך לא נשמר עדיין
                        }
          
                    }
                    if (!exist && !Tourqise)//אם היה קיים צריך לבטל&&אם הוא טורקיז אז הביטול לא מתבצע בטבלה הגדולה אלא במפורטת
                    AddRowToTableForUpdate("CANCEL", true,0);//0 כי זה לא משנה לא מפורט תא בכלל
                }


                bool Saved = WasteObject.InsertDataToNin(TableFor1DetailedCell);//הכנסה לקובץ ITH ומשתנה האם הצליחה השמירה
                //btn_save.Enabled = false;
                if (Saved)//איפוס טבלת עדכונים
                {
                    TableFor1DetailedCell.Rows.Clear();
                    RowsPositionForDetailedCell = -1;
                }
                tabControl1.SelectedTab = TabPageToReturn;
                timer1.Start();

                //הוספת תא בצבע כחול

                CellPaint cellPaint = new CellPaint(eRowIndex, eColumnIndex);
                switch (IndexOfDatagrid)
                {
                    case 0:
                        cellPaintsListMixures.Add(cellPaint);
                        break;

                    case 1:
                        cellPaintsListFabric.Add(cellPaint);
                        break;

                    case 2:
                        cellPaintsListSteel.Add(cellPaint);
                        break;
                }



            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            //שמירת נתונים
        }

        /// <summary>
        /// עדכון של תא
        /// </summary>
        public void UpdateCell()
        {
            DBService dbs = new DBService();
            string InsertValues = "";
            //שמירה לs400
            for (int i = 0; i < DataGrid_Desc.RowCount - 1; i++)
            {
                //לפי הסדר: מספר קטלוגי,תאריך דיווח,קוד תקלה,כמות,שם מדווח(כרגע ריק) ,תאריך עדכון היום
                InsertValues = $@"insert into  MSK.MSVQTP values ('{WasteObject.CatalogNumber.Trim()}','{WasteObject.Machine}',{WasteObject.Department},{DateTime.Parse(lbl_dateWantedUpdate.Text).ToString("yyMMdd")}
                                         ,{DataGrid_Desc.Rows[i].Cells["קוד"].Value.ToString()},{DataGrid_Desc.Rows[i].Cells["כמות"].Value.ToString()},'{WasteObject.Shift}','{WasteObject.NameUser}',{DateTime.Now.ToString("yyMMdd")})";// ('test11',123456,123,100,'test1',87654321)"
                dbs.executeInsertQuery(InsertValues);
            }
        }




        //פונקציות עקיפות-פחות חשובות

        /// <summary>
        /// בדיקת קיום שדות לפי דרישתו של אלי
        /// </summary>
        private void CheckS400FieldsExist()
        {
            DataTable CheckTable = new DataTable();
            //קיום פריט
            string qry = $@"SELECT *
                          FROM  BPCSFV30.IIML01 
                          WHERE   IPROD='FB-40002'";
            CheckTable= DBS.executeSelectQueryNoParam(qry);
            if(CheckTable.Rows.Count==0)
                MessageBox.Show("בדוק עם מנהל המערכת", "שגיאת מערכת 1",MessageBoxButtons.OK, MessageBoxIcon.Error);
            //בדיקת מחסן
            qry = $@"SELECT *
                   FROM  BPCSfV30.ILML01  
                   WHERE WWHS='SC' ";
            CheckTable = DBS.executeSelectQueryNoParam(qry);
            if (CheckTable.Rows.Count == 0)
                MessageBox.Show("בדוק עם מנהל המערכת", "שגיאת מערכת 2", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //בדיקת מחלקות
            qry = $@"SELECT    CDEPT
                   FROM  BPCSFV30. CDPL01";
            CheckTable = DBS.executeSelectQueryNoParam(qry);
            WasteObject.DepartmentsList = CheckTable.Rows.OfType<DataRow>()
                                            .Select(dr => dr.Field<decimal>("CDEPT")).ToList();
            //בדיקת קיום סוגי תנועות
            qry = $@"SELECT  TTYPE
                   FROM BPCSFV30.ITEL01 
                   WHERE  TTYPE IN('PX','P5')";
            CheckTable = DBS.executeSelectQueryNoParam(qry);
            if(CheckTable.Rows.Count!=2)
                MessageBox.Show(" אחד מסוגי ההתנועות התחלפו, בדוק עם מנהל המערכת", "שגיאת מערכת 3", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //בדיקת קיום סוגי תקלה
            qry = $@"SELECT DISTINCT  substring(PKEY,5,2)  
                   FROM BPCSFV30.ZPAL01   
                   WHERE substring(PKEY,5,2) in ('PX','P5')";
            CheckTable = DBS.executeSelectQueryNoParam(qry);
            if(CheckTable.Rows.Count!=2)
                MessageBox.Show(" אחד מסוגי התקלה התחלפו, בדוק עם מנהל המערכת", "שגיאת מערכת 4", MessageBoxButtons.OK, MessageBoxIcon.Error);

        }




   
        private void panel1_Paint(object sender, PaintEventArgs e)
        {
            if (panel1.BorderStyle == BorderStyle.FixedSingle)
            {
                int thickness = 5;//it's up to you
                int halfThickness = thickness / 2;
                using (Pen p = new Pen(Color.Black, thickness))
                {
                    e.Graphics.DrawRectangle(p, new Rectangle(halfThickness,
                                                              halfThickness,
                                                              panel1.ClientSize.Width - thickness,
                                                              panel1.ClientSize.Height - thickness));
                }
            }
        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {
            if (panel2.BorderStyle == BorderStyle.FixedSingle)
            {
                int thickness = 5;//it's up to you
                int halfThickness = thickness / 2;
                using (Pen p = new Pen(Color.Black, thickness))
                {
                    e.Graphics.DrawRectangle(p, new Rectangle(halfThickness,
                                                              halfThickness,
                                                              panel2.ClientSize.Width - thickness,
                                                              panel2.ClientSize.Height - thickness));
                }
            }
        }


        private void panel3_Paint(object sender, PaintEventArgs e)
        {
            if (panel3.BorderStyle == BorderStyle.FixedSingle)
            {
                int thickness = 5;//it's up to you
                int halfThickness = thickness / 2;
                using (Pen p = new Pen(Color.Black, thickness))
                {
                    e.Graphics.DrawRectangle(p, new Rectangle(halfThickness,
                                                              halfThickness,
                                                              panel3.ClientSize.Width - thickness,
                                                              panel3.ClientSize.Height - thickness));
                }
            }
        }

 

        private void TabPageWasteUpdate_Paint(object sender, PaintEventArgs e)
        {
            base.OnPaint(e);
            Pen arrow = new Pen(Brushes.Black, 4);
            arrow.EndCap = System.Drawing.Drawing2D.LineCap.ArrowAnchor;
            e.Graphics.DrawLine(arrow, tabControl1.Location.X-100,140, tabControl1.Location.X+tabControl1.Size.Width+300, 140);
            
            arrow.Dispose();
        }





        private void tabControl1_Selecting(object sender, TabControlCancelEventArgs e)
        {

            if (e.TabPage == TabPageWasteUpdate && !btnUpdatePress)
            {
                e.Cancel = true;
            }
            if (e.TabPage != TabPageWasteUpdate)
            {
                btnSaveDescForCell.Visible = false;
                btnSaveDescForCell.Enabled = false;
            }
            if (e.TabPage == TabPageMixures)
            {
                btn_save.BackColor = panel1.BackColor;
            }
            else if (e.TabPage == TabPageFabric)
            {
                btn_save.BackColor = panel2.BackColor;
            }
            else if (e.TabPage == TabPageSteel)
            {
                btn_save.BackColor = panel3.BackColor;
            }
            if (e.TabPage!=TabPageWasteUpdate && TableFor1DetailedCell!=null && AskToSave)//עבר טב ופירט פסולות
            {
                if (TableFor1DetailedCell.Rows.Count > 0)
                {
                    DialogResult dialogResult = MessageBox.Show("?האם ברצונך לשמור פירוט פסולות עבור התא הנוכחי", "שמירת נתונים", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        SaveCellData();
                    }
                    else
                    {
                        TableFor1DetailedCell = new DataTable();
                        RowsPosition = -1;
                    }
                }
            }
             if (e.TabPage != TabPageWasteUpdate && TableFor1DetailedCell != null && !AskToSave)
            {
                if (TableFor1DetailedCell.Rows.Count > 0)
                {
                    TableFor1DetailedCell = new DataTable();
                    RowsPosition = -1;
                }
            }

            // if(e.TabPage== TabPageTotal)
            //{
            //    btn_refresh.Visible = true;
            //}
            //else
            //{
            //    btn_refresh.Visible = false;
            //}
        }

        private void pictureBox1_Paint(object sender, PaintEventArgs e)
        {
            Rectangle r = new Rectangle(0, 0, btn_save.Width, btn_save.Height);
            System.Drawing.Drawing2D.GraphicsPath gp = new System.Drawing.Drawing2D.GraphicsPath();
            int d = 50;
            gp.AddArc(r.X, r.Y, d, d, 180, 90);
            gp.AddArc(r.X + r.Width - d, r.Y, d, d, 270, 90);
            gp.AddArc(r.X + r.Width - d, r.Y + r.Height - d, d, d, 0, 90);
            gp.AddArc(r.X, r.Y + r.Height - d, d, d, 90, 90);
            btn_save.Region = new Region(gp);
        }


        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (btn_save.Enabled)
            {
                if (ValueChanged)
                {
                    DialogResult dialogResult = MessageBox.Show("? האם ברצונך לשמור נתונים לפני יציאה", "סגירת תוכנית", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        WasteObject.InsertDataToNin(TableForUpdate);
                    }
                }
            }
            Application.Exit();
        }


        /// <summary>
        /// היה עדכון תא וצריך להתעדכן עד שפי 5 יכנס לith
        /// </summary>
        private void timer1_Tick(object sender, EventArgs e)
        {
            //הבעיה: בומה שומר נתונים ממשיך להקליד ואז זה נשמר בטבלה. אחר כך מתרענן קורא מאס400 ומעלים לו את הנתונים ואז נרשם לו פעמיים
            //לנסות לא להכניס את הנתונים עד שמסיים לגמרי להקליד
            bool nin=WasteObject.CheckNinData();
            if (!nin)//אם הנתונים מנין עברו לITH
            {
                lbl_LoadData.Text = "נתונים הוזנו במערכת";
                timer1.Stop();
                tabControl1.SelectedTab = tabControl1.SelectedTab;
            }         
        }

        //אחרי שמירת נתונים יחכה 10 שניות ויעדכן טבלת ריכוזי פסולות
        private void timer2_Tick(object sender, EventArgs e)
        {
            
            ////טבלת ריכוזי פסולות חדשה
            //WasteTableTotal = new DataTable();
            //WasteObject.NewTotal();
            //WasteTableTotal = WasteObject.CreateTableTotal();
            //DGV_Total.DataSource = WasteTableTotal;
            //timer2.Stop();
        }


        private void DGV_Fabric_CellClick(object sender, DataGridViewCellEventArgs e)
        {
                foreach (DataGridViewColumn column in DGV_Fabric.Columns)
                {
                    column.SortMode = DataGridViewColumnSortMode.NotSortable;
                }          
        }

        private void DGV_Steel_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            foreach (DataGridViewColumn column in DGV_Steel.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
        }

        private void DGV_Total_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            //עיצוב שורה אחרונה של טוטל כללי
            if (e.RowIndex == DGV_Total.Rows.Count - 3 ||e.RowIndex== DGV_Total.Rows.Count -2)
            {
                e.AdvancedBorderStyle.Top = DGV_Total.AdvancedCellBorderStyle.Top;
                e.AdvancedBorderStyle.Bottom = DGV_Total.AdvancedCellBorderStyle.Bottom;
                e.CellStyle.Font = new Font("Tahoma", 20, FontStyle.Bold);
                e.CellStyle.BackColor = Color.LightBlue;
                DGV_Total.Rows[e.RowIndex].Height = 42;
            }
        }

        private void panel4_Paint(object sender, PaintEventArgs e)
        {
            if (panel4.BorderStyle == BorderStyle.FixedSingle)
            {
                int thickness = 5;//it's up to you
                int halfThickness = thickness / 2;
                using (Pen p = new Pen(Color.Black, thickness))
                {
                    e.Graphics.DrawRectangle(p, new Rectangle(halfThickness,
                                                              halfThickness,
                                                              panel4.ClientSize.Width - thickness,
                                                              panel4.ClientSize.Height - thickness));
                }
            }
        }

  
        /// <summary>
        /// ריענון טבלאות פסולות
        /// </summary>
        private void btn_refresh_Click(object sender, EventArgs e)
        {
            WasteTableMixures = new DataTable();//טבלת מיקסרים
            WasteTableFabric = new DataTable();//טבלת בדים
            WasteTableSteel = new DataTable();//טבלת פלדה
            WasteTableTotal = new DataTable();//טבלה ריכוזי פסולות
            WasteObject.NewDataTable();
            ShowWaste();
            DataOfCellTable = WasteObject.GetCellsData();
            int i = 0;
            foreach (TabPage p in tabControl1.TabPages)
            {
                tabControl1.SelectedTab = p;
                CellPaintRound2(i);
                i++;
            }

            lbl_LoadData.Visible = false;
            lbl_LoadData.Text = "...מכניס נתונים";
            ////טבלת ריכוזי פסולות חדשה
            //WasteTableTotal = new DataTable();
            //WasteObject.NewTotal();
            //WasteTableTotal=WasteObject.CreateTableTotal();
            //DGV_Total.DataSource = WasteTableTotal;
        }

        /// <summary>
        /// פלט לאקסל
        /// </summary>
        private void MakeExcellRepBTN_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            Excel.Application xlexcel;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;

            object misValue = System.Reflection.Missing.Value;
            xlexcel = new Excel.Application();
            xlWorkBook = xlexcel.Workbooks.Add(misValue);


            //xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            Excel.Worksheet active = (Excel.Worksheet)xlexcel.ActiveSheet;
            xlWorkSheet = xlWorkBook.ActiveSheet as Excel.Worksheet;
            xlWorkSheet.Name = "תערובות";        
            active.DisplayRightToLeft = false;
    


            //טוטל נכון לחודש
            string Month = CultureInfo.CurrentCulture.DateTimeFormat.GetAbbreviatedMonthName(int.Parse(WasteObject.MonthNumber)).ToString(CultureInfo.InvariantCulture);
            Excel.Range chartTotalRange1 = xlWorkSheet.get_Range("AG4:AK5");
            xlWorkSheet.Cells[4, "AG"] = $@"Month {Month} - {WasteObject.Year.ToString()} ";
            chartTotalRange1.Merge();//מיזוג תאים
            chartTotalRange1.Font.Bold = true;
            chartTotalRange1.Font.Size = 14;
            chartTotalRange1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
            chartTotalRange1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט

 

            //כותרות דטה גריד
            int x = 1;
            for (x = 1; x < DGVֹ_Mixtures.Columns.Count + 1; x++)
            {
                switch(DGVֹ_Mixtures.Columns[x - 1].HeaderText)
                {
              
                    case "מיקסר":
                        xlWorkSheet.Cells[6, x + 1] = "Compound";
                        break;

                    case "קלנדר":
                        xlWorkSheet.Cells[6, x + 1] = "Calender";
                        break;

                    case "ברומו שרוף":
                        xlWorkSheet.Cells[6, x + 1] = "scorched Bromo";
                        break;

                    default:
                        xlWorkSheet.Cells[6, x + 1] = DGVֹ_Mixtures.Columns[x - 1].HeaderText;
                        break;
                }
                xlWorkSheet.Cells[7, x + 1] = DGVֹ_Mixtures.Columns[x - 1].HeaderText;
            }
            xlWorkSheet.get_Range("A7:A7").EntireRow.Hidden = true;//בלי עברית כרגע-מדובר בכותרות
            xlWorkSheet.get_Range("A8:A8").EntireRow.Hidden = true;
            xlWorkSheet.get_Range("A1:A2").EntireRow.Hidden = true;
            string ColumnLetter = ColumnIndexToColumnLetter(x);//מחליף את המספר לאות
            xlWorkSheet.Rows[6].WrapText = true;

            //מעל דטה גריד
            xlWorkSheet.Cells[7, 1] = "Topic subject";
            xlWorkSheet.Cells[7, 1].Font.Size = 14;
            xlWorkSheet.Cells[7, 1].Font.Bold = true;

            Excel.Range chartTotalRange2 = xlWorkSheet.get_Range("A6:" + ColumnLetter + "7");
            chartTotalRange2.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
            chartTotalRange2.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble;
            chartTotalRange2.Font.Bold = true;
            chartTotalRange2.Font.Size = 14;
            chartTotalRange2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט


            //כותרת ראשית
            xlWorkSheet.Cells[3, 2] = "פסולות-תערובות";
            Excel.Range chartTotalRange3 = xlWorkSheet.get_Range("B3:" + ColumnLetter + "3");
            chartTotalRange3.Merge();//מיזוג תאים
            chartTotalRange3.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
            chartTotalRange3.Font.Bold = true;
            chartTotalRange3.Font.Size = 24;
            chartTotalRange3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט
            chartTotalRange3.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble;//גבולות

            //תאריך עדכון אחרון
            xlWorkSheet.Cells[4, 2] = "Month Report";
            xlWorkSheet.Cells[4, 2].Font.Size = 14;
            //xlWorkSheet.get_Range("A4").RowHeight = 50;
            xlWorkSheet.Cells[4, 4] = $@"Month {Month} - {WasteObject.Year.ToString()} ";
            xlWorkSheet.Cells[4, 4].Font.Size = 14;

            xlWorkSheet.get_Range("C4:" + ColumnLetter+"4").Merge();//מיזוג תאים
            xlWorkSheet.get_Range("B4:" + ColumnLetter+"4").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
            xlWorkSheet.get_Range("C4:"+ ColumnLetter+"4").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט
            xlWorkSheet.get_Range("B4:"+ ColumnLetter+"3").Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble;//גבולות
            xlWorkSheet.get_Range("B4:"+ ColumnLetter+"3").Font.Bold = true;

            DataObject dataObj = null;
            DGVֹ_Mixtures.SelectAll();
            dataObj = DGVֹ_Mixtures.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
            Excel.Range CR = (Excel.Range)xlWorkSheet.Cells[9, 1];//טווח מילוי הטבלה שורה 9 טור 1
            CR.Select();
            xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);

            //עיצוב דטה גריד
            int RowsBorder = DGVֹ_Mixtures.RowCount + 8;//-יתחיל בשורה 9 גבולות תא של דטה גריד באקסל
            Excel.Range chartTotalRange4 = xlWorkSheet.get_Range("B8:" + ColumnLetter + RowsBorder);
            chartTotalRange4.Borders.Color = System.Drawing.Color.Black.ToArgb();
            chartTotalRange4.Font.Bold = true;
            chartTotalRange4.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט
            chartTotalRange4.Font.Size = 12;
            chartTotalRange4.RowHeight = 15;
            xlWorkSheet.get_Range("A5:A5").RowHeight = 6;
            chartTotalRange4.ColumnWidth = 13.25;
            xlWorkSheet.get_Range("D:F").ColumnWidth = 7.5;
            chartTotalRange4.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט

            //טוטל
            xlWorkSheet.get_Range("B"+(RowsBorder-1)+":" + ColumnLetter + RowsBorder).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);

            xlWorkSheet.get_Range("A:A").EntireColumn.Hidden = true;


            //בדיםםםםםםםםםםםםםםםםםםםםםםם
            var xlSheets = xlWorkBook.Sheets as Excel.Sheets;
            var xlNewSheet2 = (Excel.Worksheet)xlSheets.Add(xlSheets[1], Type.Missing, Type.Missing, Type.Missing);
            xlNewSheet2.Name = "בדים";
            Excel.Worksheet active2 = (Excel.Worksheet)xlexcel.ActiveSheet;
            xlNewSheet2 = xlWorkBook.ActiveSheet as Excel.Worksheet;
            active2.DisplayRightToLeft = false;


            //טוטל נכון לחודש
            string Month2 = CultureInfo.CurrentCulture.DateTimeFormat.GetAbbreviatedMonthName(int.Parse(WasteObject.MonthNumber)).ToString(CultureInfo.InvariantCulture);
            xlNewSheet2.Cells[4, "AG"] = $@"Month {Month} - {WasteObject.Year.ToString()} ";
            Excel.Range chartTotalRange5 = xlNewSheet2.get_Range("AG4:AK5");
            chartTotalRange5.Merge();//מיזוג תאים
            chartTotalRange5.Font.Bold = true;
            chartTotalRange5.Font.Size = 14;
            chartTotalRange5.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
            chartTotalRange5.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט



            //כותרות דטה גריד

            ////הורדת עברית בגלל זה בהערות
            x = 1;
            for (x = 1; x < DGV_Fabric.Columns.Count + 1; x++)
            {                
                xlNewSheet2.Cells[7, x + 1] = DGV_Fabric.Columns[x - 1].HeaderText;
            }
            ColumnLetter = ColumnIndexToColumnLetter(x);//מחליף את המספר לאות

            //כותרות איפה שיש משמרות
            xlNewSheet2.Cells[7, "D"] = "71";
            xlNewSheet2.Cells[7, "E"] = "160";
            xlNewSheet2.Cells[7, "F"] = "Calender";
            xlNewSheet2.Cells[7, "L"] = "Cutter Breaker";
            xlNewSheet2.Cells[7, "P"] = "RJS cutter 2";
            xlNewSheet2.Cells[7, "Q"] = "Textile Salvage";
            xlNewSheet2.Cells[7, "R"] = "PET Salvage";
            xlNewSheet2.Cells[7, "S"] = "Fabric Calender";
            xlNewSheet2.Cells[7, "T"] = "Surplus Production & building";
            xlNewSheet2.Cells[7, "U"] = "Trials";

            xlNewSheet2.Cells[6, "G"] = "Bands cutter 1";
            xlNewSheet2.get_Range("G6:I6").Merge();//מיזוג תאים
            xlNewSheet2.get_Range("G6:I6").Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble;
            xlNewSheet2.Cells[6, "J"] = "Cutter 3";
            xlNewSheet2.get_Range("J6:K6").Merge();
            xlNewSheet2.get_Range("J6:K6").Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble;
            xlNewSheet2.Cells[6, "M"] = "Cutter 5 horizontal";
            xlNewSheet2.get_Range("M6:O6").Merge();
            xlNewSheet2.get_Range("D6:U6").Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble;          
            xlNewSheet2.Rows[6].WrapText = true;
            xlNewSheet2.Rows[7].WrapText = true;
            xlNewSheet2.get_Range("A8:A8").EntireRow.Hidden = true;
            xlNewSheet2.get_Range("A1:A2").EntireRow.Hidden = true;

            //מעל דטה גריד
            xlNewSheet2.Cells[7, 1] = "Topic subject";
            xlNewSheet2.Cells[7, 1].Font.Size = 14;
            xlNewSheet2.Cells[7, 1].Font.Bold = true;
            xlNewSheet2.get_Range("A6:" + ColumnLetter + "7").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
            xlNewSheet2.get_Range("A7:" + ColumnLetter + "7").Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble;
            xlNewSheet2.get_Range("A5:" + ColumnLetter + "7").Font.Bold = true;
            xlNewSheet2.get_Range("A6:" + ColumnLetter + "7").Font.Size = 12;
            xlNewSheet2.get_Range("A5:" + ColumnLetter + "7").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט


            //כותרת ראשית
            xlNewSheet2.Cells[3, 2] = "פסולות-בדים";
            Excel.Range chartTotalRange6 = xlNewSheet2.get_Range("B3:" + ColumnLetter + "3");
            chartTotalRange6.Merge();//מיזוג תאים
            chartTotalRange6.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
            chartTotalRange6.Font.Bold = true;
            chartTotalRange6.Font.Size = 24;
            chartTotalRange6.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט
            chartTotalRange6.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble;//גבולות

            //תאריך עדכון אחרון
            xlNewSheet2.Cells[4, 2] = "Month Report";
            xlNewSheet2.Cells[4, 2].Font.Size = 14;
            //xlNewSheet2.get_Range("A4").RowHeight = 30;
            xlNewSheet2.Cells[4, 4] = $@"Month {Month} - {WasteObject.Year.ToString()} ";
            xlNewSheet2.Cells[4, 4].Font.Size = 14;
            xlNewSheet2.get_Range("C4:" + ColumnLetter + "4").Merge();//מיזוג תאים
            xlNewSheet2.get_Range("B4:" + ColumnLetter + "4").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
            xlNewSheet2.get_Range("C4:" + ColumnLetter + "4").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט
            xlNewSheet2.get_Range("B4:" + ColumnLetter + "3").Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble;//גבולות
            xlNewSheet2.get_Range("B4:" + ColumnLetter + "3").Font.Bold = true;

            dataObj = null;
            DGV_Fabric.SelectAll();
            dataObj = DGV_Fabric.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
             CR = (Excel.Range)xlNewSheet2.Cells[9, 1];//טווח מילוי הטבלה שורה 9 טור 1
            CR.Select();
            xlNewSheet2.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);

            //עיצוב דטה גריד
             RowsBorder = DGVֹ_Mixtures.RowCount + 8;//-יתחיל בשורה 9 גבולות תא של דטה גריד באקסל
            Excel.Range chartTotalRange7 = xlNewSheet2.get_Range("B8:" + ColumnLetter + RowsBorder);

            chartTotalRange7.Borders.Color = System.Drawing.Color.Black.ToArgb();
            chartTotalRange7.Font.Bold = true;
            chartTotalRange7.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט
            chartTotalRange7.Font.Size = 12;
            chartTotalRange7.RowHeight = 15;
            xlNewSheet2.get_Range("A5:A5").RowHeight = 6;
            chartTotalRange7.ColumnWidth = 8;
            xlNewSheet2.get_Range("B:C").ColumnWidth = 13;
            xlNewSheet2.get_Range("P:T").ColumnWidth = 12.5;
            chartTotalRange7.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט

            //טוטל
            xlNewSheet2.get_Range("B" + (RowsBorder - 1) + ":" + ColumnLetter + RowsBorder).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);

            xlNewSheet2.get_Range("A:A").EntireColumn.Hidden = true;

           


            //פלדהההההההההה
            var xlNewSheet3 = (Excel.Worksheet)xlSheets.Add(xlSheets[1], Type.Missing, Type.Missing, Type.Missing);
            xlNewSheet3.Name = "פלדה";
            Excel.Worksheet active3 = (Excel.Worksheet)xlexcel.ActiveSheet;
            xlNewSheet3 = xlWorkBook.ActiveSheet as Excel.Worksheet;
            active3.DisplayRightToLeft = false;


            //טוטל נכון לחודש
            string Month3 = CultureInfo.CurrentCulture.DateTimeFormat.GetAbbreviatedMonthName(int.Parse(WasteObject.MonthNumber)).ToString(CultureInfo.InvariantCulture);
            xlNewSheet2.Cells[4, "AG"] = $@"Month {Month} - {WasteObject.Year.ToString()} ";
            Excel.Range chartTotalRange8 = xlNewSheet2.get_Range("AG4:AK5");
            chartTotalRange8.Merge();//מיזוג תאים
            chartTotalRange8.Font.Bold = true;
            chartTotalRange8.Font.Size = 14;
            chartTotalRange8.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.YellowGreen);
            chartTotalRange8.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט



            //כותרות דטה גריד
            x = 1;
            for (x = 1; x < DGV_Steel.Columns.Count + 1; x++)
            {
                if (DGV_Fabric.Columns[x - 1].HeaderText == "ספדון")
                    xlNewSheet3.Cells[7, x + 1] = "ספדון+Vmi";
                else
                    xlNewSheet3.Cells[7, x + 1] = DGV_Steel.Columns[x - 1].HeaderText;
            }

            xlNewSheet3.Cells[6, "D"] = "Berstof";
            xlNewSheet3.Cells[6, "E"] = "Spedon+Vmi";
            xlNewSheet3.Cells[6, "F"] = "160";
            xlNewSheet3.Cells[6, "G"] = "71";
            xlNewSheet3.Cells[6, "H"] = "Steel Leftover";
            xlNewSheet3.Cells[6, "I"] = "Bead Scrap";
            xlNewSheet3.Cells[6, "J"] = "Steel Cord";
   
            ColumnLetter = ColumnIndexToColumnLetter(x);//מחליף את המספר לאות

            //מעל דטה גריד
            xlNewSheet3.Cells[7, 1] = "Topic subject";
            xlNewSheet3.Cells[7, 1].Font.Size = 14;
            xlNewSheet3.Cells[7, 1].Font.Bold = true;
            Excel.Range chartTotalRange9 = xlNewSheet3.get_Range("A6:" + ColumnLetter + "7");
            chartTotalRange9.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.YellowGreen);
            xlNewSheet3.get_Range("C6:" + ColumnLetter + "7").Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble;
            chartTotalRange9.Font.Bold = true;
            chartTotalRange9.Font.Size = 14;
            chartTotalRange9.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט


            //כותרת ראשית
            xlNewSheet3.Cells[3, 2] = "פסולות-פלדה";
            Excel.Range chartTotalRange10 = xlNewSheet3.get_Range("B3:" + ColumnLetter + "3");
            chartTotalRange10.Merge();//מיזוג תאים
            chartTotalRange10.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.YellowGreen);
            chartTotalRange10.Font.Bold = true;
            chartTotalRange10.Font.Size = 24;
            chartTotalRange10.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט
            chartTotalRange10.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble;//גבולות

            //תאריך עדכון אחרון
            xlNewSheet3.Cells[4, 2] = "Month Report";
            xlNewSheet3.Cells[4, 2].Font.Size = 14;
            //xlNewSheet3.get_Range("A4").RowHeight = 50;
            xlNewSheet3.Cells[4, 4] = $@"Month {Month} - {WasteObject.Year.ToString()} ";
            xlNewSheet3.Cells[4, 4].Font.Size = 14;
            Excel.Range chartTotalRange11 = xlNewSheet3.get_Range("B4:" + ColumnLetter + "4");

            xlNewSheet3.get_Range("C4:" + ColumnLetter + "4").Merge();//מיזוג תאים
            chartTotalRange11.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.YellowGreen);
            xlNewSheet3.get_Range("C4:" + ColumnLetter + "4").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט
            chartTotalRange11.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble;//גבולות
            chartTotalRange11.Font.Bold = true;

            dataObj = null;
            DGV_Steel.SelectAll();
            dataObj = DGV_Steel.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
            CR = (Excel.Range)xlNewSheet3.Cells[9, 1];//טווח מילוי הטבלה שורה 9 טור 1
            CR.Select();
            xlNewSheet3.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);

            //עיצוב דטה גריד
            RowsBorder = DGV_Fabric.RowCount + 8;//-יתחיל בשורה 9 גבולות תא של דטה גריד באקסל
            Excel.Range chartTotalRange12 = xlNewSheet3.get_Range("B8:" + ColumnLetter + RowsBorder);
            chartTotalRange12.Borders.Color = System.Drawing.Color.Black.ToArgb();
            chartTotalRange12.Font.Bold = true;
            chartTotalRange12.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט
            chartTotalRange12.Font.Size = 12;
            chartTotalRange12.RowHeight = 15;
            xlNewSheet3.get_Range("A5:A5").RowHeight = 6;
            chartTotalRange12.ColumnWidth = 13;
            chartTotalRange12.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט
            xlNewSheet3.Rows[6].WrapText = true;

            //טוטל
            xlNewSheet3.get_Range("B" + (RowsBorder - 1) + ":" + ColumnLetter + RowsBorder).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.YellowGreen);

            xlNewSheet3.get_Range("A:A").EntireColumn.Hidden = true;
            xlNewSheet3.get_Range("A7:A7").EntireRow.Hidden = true;
            xlNewSheet3.get_Range("A1:A2").EntireRow.Hidden = true;


            //חישובי אחוזים עבור פלדה לאלחנן
            DataTable SpedonTable = new DataTable();//סכומים של תקלות ספדון
            SpedonTable = WasteObject.GetSumSteelForExcel();
            xlNewSheet3.Cells[RowsBorder + 2, "B"] = "ברסטוף";
            xlNewSheet3.Cells[RowsBorder + 2, "C"] = "ספדון+vmi";
            xlNewSheet3.Cells[RowsBorder + 3, "C"] = "גימום לא תקין";
            xlNewSheet3.Cells[RowsBorder + 3, "D"] = "החלפות";
            xlNewSheet3.Cells[RowsBorder + 3, "E"] = "החלפת זוויות";
            xlNewSheet3.Cells[RowsBorder + 3, "F"] = "תקלות";
            xlNewSheet3.Cells[RowsBorder + 2, "G"] = "160";
            xlNewSheet3.Cells[RowsBorder + 2, "H"] = "71";
            xlNewSheet3.get_Range("C" + (RowsBorder + 2) + ":F" + (RowsBorder + 2)).Merge();//  מיזוג ספדון
            xlNewSheet3.get_Range("B" + (RowsBorder + 2) + ":B" + (RowsBorder + 3)).Merge();//מיזוג ברסטוף
            xlNewSheet3.get_Range("G" + (RowsBorder + 2) + ":G" + (RowsBorder + 3)).Merge();//מיזוג 160
            xlNewSheet3.get_Range("H" + (RowsBorder + 2) + ":H" + (RowsBorder + 3)).Merge();//מיזוג 71
            xlNewSheet3.get_Range("B" + (RowsBorder + 2) + ":H" + (RowsBorder + 3)).Borders.Color = System.Drawing.Color.Black.ToArgb();
            xlNewSheet3.get_Range("B" + (RowsBorder + 2) + ":H" + (RowsBorder + 3)).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            xlNewSheet3.get_Range("B" + (RowsBorder + 2) + ":H" + (RowsBorder + 3)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט




            int SumReason = Summ(0, SpedonTable.Rows.Count - 1);//סכום רקורסיה-הסכום של פירוטי ספדון
            int Summ(int sum, int Rows)
            {
                if (Rows == -1)
                {
                    return sum;
                }
                else
                {
                    return sum += Summ(int.Parse(SpedonTable.Rows[Rows]["Quantity"].ToString()), Rows - 1);                   
                }
            }
            if (SumReason != 0)
            {
                int SumReasonAndBerstof = SumReason +int.Parse(WasteTableSteel.Rows[WasteTableSteel.Rows.Count - 2]["ברסטוף"].ToString());
                int SumReasonBerstofAnd16071 = SumReasonAndBerstof + int.Parse(WasteTableSteel.Rows[WasteTableSteel.Rows.Count - 2]["160"].ToString()) + int.Parse(WasteTableSteel.Rows[WasteTableSteel.Rows.Count - 2]["71"].ToString());
                //רק התקלות
                for (int i = 0; i < SpedonTable.Rows.Count; i++)
                {
                    //שורה ראשונה רק חישובי אחוזים של פירוטי ספדון
                    xlNewSheet3.Cells[RowsBorder + 4, 3 + i] = (double.Parse(SpedonTable.Rows[i]["Quantity"].ToString()) / SumReason).ToString("p2");//מתחיל מטור C
                    //שורה שניה פירוטי ספדון+ברסטוף
                    xlNewSheet3.Cells[RowsBorder + 5, 3 + i]= (double.Parse(SpedonTable.Rows[i]["Quantity"].ToString()) / SumReasonAndBerstof).ToString("p2");
                    //שורה שלישית פירוטי ספדון+ברסטוף+160+71
                    xlNewSheet3.Cells[RowsBorder + 6, 3 + i] = (double.Parse(SpedonTable.Rows[i]["Quantity"].ToString()) / SumReasonBerstofAnd16071).ToString("p2");
                }
                xlNewSheet3.Cells[RowsBorder + 5, 2] = (double.Parse(WasteTableSteel.Rows[WasteTableSteel.Rows.Count - 2]["ברסטוף"].ToString()) / SumReasonAndBerstof).ToString("p2");//חלק ברסטוף משורה שניה
                xlNewSheet3.Cells[RowsBorder + 6, 2] = (double.Parse(WasteTableSteel.Rows[WasteTableSteel.Rows.Count - 2]["ברסטוף"].ToString()) / SumReasonBerstofAnd16071).ToString("p2");//חלק ברסטוף משורה שניה
                xlNewSheet3.Cells[RowsBorder + 6, "G"] = (double.Parse(WasteTableSteel.Rows[WasteTableSteel.Rows.Count - 2]["160"].ToString()) / SumReasonBerstofAnd16071).ToString("p2");//חלק ברסטוף משורה שניה
                xlNewSheet3.Cells[RowsBorder + 6, "H"] = (double.Parse(WasteTableSteel.Rows[WasteTableSteel.Rows.Count - 2]["71"].ToString()) / SumReasonBerstofAnd16071).ToString("p2");//חלק ברסטוף משורה שניה

            }

            //ריכוז פסולותתת
            var xlNewSheet4 = (Excel.Worksheet)xlSheets.Add(xlSheets[1], Type.Missing, Type.Missing, Type.Missing);
            xlNewSheet4.Name = "ריכוז פסולות";
            Excel.Worksheet active4 = (Excel.Worksheet)xlexcel.ActiveSheet;
            xlNewSheet4 = xlWorkBook.ActiveSheet as Excel.Worksheet;
            active4.DisplayRightToLeft = false;


            //טוטל נכון לחודש
            string Month4 = CultureInfo.CurrentCulture.DateTimeFormat.GetAbbreviatedMonthName(int.Parse(WasteObject.MonthNumber)).ToString(CultureInfo.InvariantCulture);
            xlNewSheet4.Cells[4, "AG"] = $@"Month {Month} - {WasteObject.Year.ToString()} ";
            Excel.Range chartTotalRange13 = xlNewSheet4.get_Range("AG4:AK5");

            chartTotalRange13.Merge();//מיזוג תאים
            chartTotalRange13.Font.Bold = true;
            chartTotalRange13.Font.Size = 14;
            chartTotalRange13.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
            chartTotalRange13.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט



            //כותרות דטה גריד
            x = 1;
            for (x = 1; x < DGV_Total.Columns.Count + 1; x++)
            {
                xlNewSheet4.Cells[7, x + 1] = DGV_Total.Columns[x - 1].HeaderText;
            }

            xlNewSheet4.Cells[6, "D"] = "Compounds";
            xlNewSheet4.Cells[6, "E"] = "Fabric";
            xlNewSheet4.Cells[6, "F"] = "Steel";
            xlNewSheet4.Cells[6, "G"] = "Bead Scrap";
            xlNewSheet4.Cells[6, "H"] = "Steel Cord";

            ColumnLetter = ColumnIndexToColumnLetter(x);//מחליף את המספר לאות
            xlNewSheet4.Rows[6].WrapText = true;

            //מעל דטה גריד
            xlNewSheet4.Cells[7, 1] = "Topic subject";
            xlNewSheet4.Cells[7, 1].Font.Size = 14;
            xlNewSheet4.Cells[7, 1].Font.Bold = true;
            Excel.Range chartTotalRange14 = xlNewSheet4.get_Range("A6:" + ColumnLetter + "7");

            chartTotalRange14.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
            chartTotalRange14.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble;
            chartTotalRange14.Font.Bold = true;
            chartTotalRange14.Font.Size = 14;
            chartTotalRange14.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט


            //כותרת ראשית
            xlNewSheet4.Cells[3, 2] = $@"ריכוז פסולות -נכון לתאריך {lbl_LastUpdate.Text}";
            Excel.Range chartTotalRange15 = xlNewSheet4.get_Range("B3:" + ColumnLetter + "3");

            chartTotalRange15.Merge();//מיזוג תאים
            chartTotalRange15.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
            chartTotalRange15.Font.Bold = true;
            chartTotalRange15.Font.Size = 24;
            chartTotalRange15.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט
            chartTotalRange15.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble;//גבולות

            //תאריך עדכון אחרון
            xlNewSheet4.Cells[4, 2] = "Month Report";
            xlNewSheet4.Cells[4, 2].Font.Size = 14;
            //xlNewSheet4.get_Range("A4").RowHeight = 25;
            xlNewSheet4.Cells[4, 4] = $@"Month {Month} - {WasteObject.Year.ToString()} ";
            xlNewSheet4.Cells[4, 4].Font.Size = 14;
            Excel.Range chartTotalRange16 = xlNewSheet4.get_Range("B4:" + ColumnLetter + "4");
            xlNewSheet4.get_Range("C4:" + ColumnLetter + "4").Merge();//מיזוג תאים
            chartTotalRange16.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
            xlNewSheet4.get_Range("C4:" + ColumnLetter + "4").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט
            chartTotalRange16.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble;//גבולות
            chartTotalRange16.Font.Bold = true;

            dataObj = null;
            DGV_Total.SelectAll();
            dataObj = DGV_Total.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
            CR = (Excel.Range)xlNewSheet4.Cells[9, 1];//טווח מילוי הטבלה שורה 9 טור 1
            CR.Select();
            xlNewSheet4.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);

            //עיצוב דטה גריד
            RowsBorder = DGV_Total.RowCount + 8;//-יתחיל בשורה 9 גבולות תא של דטה גריד באקסל
            Excel.Range chartTotalRange17 = xlNewSheet4.get_Range("B8:" + ColumnLetter + RowsBorder);

            chartTotalRange17.Borders.Color = System.Drawing.Color.Black.ToArgb();
            chartTotalRange17.Font.Bold = true;
            chartTotalRange17.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט
            chartTotalRange17.Font.Size = 12;
            chartTotalRange17.RowHeight = 15;
            xlNewSheet4.get_Range("A5:A5").RowHeight = 6;
            chartTotalRange17.ColumnWidth = 13;
            xlNewSheet4.get_Range("D:D").ColumnWidth = 15.38;
            xlNewSheet4.get_Range("H:H").ColumnWidth = 18;


            xlNewSheet4.get_Range("B8:" + ColumnLetter + RowsBorder).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט

            //טוטל
            xlNewSheet4.get_Range("B" + (RowsBorder - 2) + ":" + ColumnLetter + RowsBorder).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);

            xlNewSheet4.get_Range("A:A").EntireColumn.Hidden = true;

            //xlNewSheet4.get_Range("A2:A2").Application.ActiveWindow.FreezePanes = true;
            Excel.Window xlWnd3 = xlexcel.ActiveWindow;
            xlNewSheet4.get_Range("A1", "A1").get_Offset(1, 0).EntireRow.Select();
            xlWnd3.FreezePanes = true;



            //טבלת הערות מתחת לריכוז פסולות
            //כותרת הערות
            xlNewSheet4.Cells[6, "L"] = "הערות/תקלות";
            Excel.Range chartTotalRange18 = xlNewSheet4.get_Range("L6:P6");
            chartTotalRange18.Merge();//מיזוג תאים
            chartTotalRange18.Font.Bold = true;
            chartTotalRange18.Font.Size = 14;
            chartTotalRange18.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            chartTotalRange18.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט
            xlNewSheet4.Cells[7, "P"] = "תאריך";
            xlNewSheet4.Cells[7, "O"] = "פריט";
            xlNewSheet4.Cells[7, "N"] = "מחלקה";
            xlNewSheet4.Cells[7, "M"] = "כמות";
            xlNewSheet4.Cells[7, "L"] = "הערה";
            //טבלת הערות
            int RowInsert=8;
            string Comment;
            for (int i = 0; i < DataOfCellTable.Rows.Count; i++)
            {
                if (DataOfCellTable.Rows[i]["ReasonCode"].ToString() != "01" && DataOfCellTable.Rows[i]["Quantity"].ToString() != "0")//שונה מכניסת פסולת רגילה
                {
                    RowInsert ++;
                    xlNewSheet4.Cells[RowInsert, "P"] = DataOfCellTable.Rows[i]["Date"].ToString().Substring(2, 2) + "/" + DataOfCellTable.Rows[i]["Date"].ToString().Substring(4, 2);
                    switch (DataOfCellTable.Rows[i]["CATALOGNUMBER"].ToString().Trim())
                    {
                        case "FB-40002":
                            xlNewSheet4.Cells[RowInsert, "O"] = "תערובות";
                            break;

                        case "SC-FC010":
                            xlNewSheet4.Cells[RowInsert, "O"] = "בדים";
                            break;

                        case "SC-FC050":
                            xlNewSheet4.Cells[RowInsert, "O"] = "פלדה";
                            break;

                        case "SC-BE200":
                            xlNewSheet4.Cells[RowInsert, "O"] = "חישוקים";
                            break;

                        case "SC-FC055":
                            xlNewSheet4.Cells[RowInsert, "O"] = "עודפי פלדה";
                            break;

                        case "SC-SC100":
                            xlNewSheet4.Cells[RowInsert, "O"] = "חוטי פלדה גולמיים";
                            break;
                    }
                    if (DataOfCellTable.Rows[i]["Department"].ToString() == "57")
                        xlNewSheet4.Cells[RowInsert, "N"]=  WasteObject.GetColumnFabricName57(DataOfCellTable.Rows[i]["Machine"].ToString().Trim(), DataOfCellTable.Rows[i]["shift"].ToString().Trim());
                    else
                        xlNewSheet4.Cells[RowInsert, "N"] = WasteObject.GetStringDepartmentOrMachine(DataOfCellTable.Rows[i]["Department"].ToString().Trim(), DataOfCellTable.Rows[i]["CatalogNumber"].ToString().Trim(), DataOfCellTable.Rows[i]["Machine"].ToString().Trim());

                    xlNewSheet4.Cells[RowInsert, "M"] = DataOfCellTable.Rows[i]["Quantity"].ToString();
                    if (DataOfCellTable.Rows[i]["Description"].ToString() == "הערות")
                    {
                        Comment = DataOfCellTable.Rows[i]["Comment"].ToString();
                    }
                    else
                    {
                        Comment = DataOfCellTable.Rows[i]["Description"].ToString();
                    }
                    xlNewSheet4.Cells[RowInsert, "L"] = Comment;
                }
            }
            Excel.Range chartTotalRange19 = xlNewSheet4.get_Range("L7:P" + RowInsert);
            chartTotalRange19.Borders.Color = System.Drawing.Color.Black.ToArgb();
            chartTotalRange19.Font.Bold = true;
            chartTotalRange19.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט
            chartTotalRange19.Font.Size = 12;
            chartTotalRange19.RowHeight = 15;
            xlNewSheet4.get_Range("M7:P" + RowInsert).ColumnWidth = 13;
            chartTotalRange19.ColumnWidth = 25;//הערות
            chartTotalRange19.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט
            xlNewSheet4.get_Range("A8:A8").EntireRow.Hidden = true;
            xlNewSheet4.get_Range("A1:A2").EntireRow.Hidden = true;


            xlexcel.Visible = true;
            Cursor.Current = Cursors.Default;
            //xlWorkBook.Close();
            releaseObject(xlexcel);
            releaseObject(xlWorkBook);
            releaseObject(xlWorkSheet);

        }


        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        static string ColumnIndexToColumnLetter(int colIndex)
        {
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



        private void button1_Click(object sender, EventArgs e)
        {
            DGV_Fabric.Rows[5].Cells[5].Value = "TEST";
            //DGV_Fabric.Rows[5].DefaultCellStyle.BackColor = Color.Red;
            DGV_Fabric.Rows[5].Cells[5].Style.BackColor = Color.Turquoise;
            DGV_Fabric.Invalidate();
        }

        private void txt_comment_TextChanged(object sender, EventArgs e)
        {
            if(txt_comment.Text.Length>=15)
            {
                MessageBox.Show("הערה יכולה להכיל עד 14 תווים");
                txt_comment.Text = txt_comment.Text.Substring(0, 14);
                txt_comment.Focus();
                txt_comment.SelectionStart = txt_comment.Text.Length;
            }

        }


        private void DGVֹ_Mixtures_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            foreach (DataGridViewColumn column in DGVֹ_Mixtures.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
        }
    }
}

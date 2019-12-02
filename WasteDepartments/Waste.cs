using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WasteDepartments
{
    class Waste
    {
        public string CatalogNumber { get; set; }// זיהוי קטלוג-בתוכנית הזאת לפי הלשוניות של טב קונטרול
        public int AmountPerDay { get; set; }//כמות של יום ספציפי
        public int PreviousAmount { get; set; }//כמות קודמת
        public string Department { get; set; }//מחלקה ששייך אליה הפסולת
        public string Machine { get; set; }//מכונה שמשוייכת לפסולת
        public string Shift { get; set; }//משמרת
        public string WorkCenter { get; set; }//מרכז עבודה
        public int DaysInMonth { get; set; }//כמה ימים בחודש העכשווי
        public string DayNumber { get; set; }
        public string DayLetter { get; set; }
        public int Year { get; set; }//שנה נבחרת לדוח
        public string MonthNumber { get; set; }// חודש נבחר לדוח
        public string DateByS400Start { get; set; }//תאריך התחלה בs400 לדוגמא 901101\
        public string DateByS400End { get; set; }//תאריך סיום לפי s400
        public int YearS400 { get; set; }//שנה לפי s400
        public string DateOfCellUpdate { get; set; }//תאריך על עדכון פסולות מפורט עבור תא מסוים
        public string ForQueryCatalogNum { get; set; }//שרשור גדול של כל המספרים הקטלוגיים עבור שאילתת השליפה
        public string ForQueryDepartment{ get; set; }//שרשור גדול של כל המחלקות עבור שאילתת השליפה
        public List<CatalogNum> ListRowCatalogNum { get; set; }//רשימה של המספרים הקטלוגיים
        CatalogNum CatalogNumObj;//אובייקט מספר קטלוגי לקוח מטבלת SQL
        public CatalogNum catalogNumToFind{ get; set; }
        public DataTable CatalogSons { get; set; }//טבלה שלוקחת את המקטים בנים של המקט העיקרי

        DataTable DtMixuresGeneral = new DataTable();//טבלת תערובות
        DataTable DtFabricGeneral = new DataTable();//טבלת בדים
        DataTable DtSteelGeneral = new DataTable();//טבלת פלדה
        DataTable DtTotal = new DataTable();//טבלת סיכומים
        DataTable DtTotalSql = new DataTable();//איזה פריטים צריך לעשות סיכום
        DataTable dataTableWaste;//נתוני כל הפסולות מs400
        DataTable dataTableTotal;//סיכומים של כל הפסולות עבור טבלת dtTotal
        public DataView CatalogFilterTable{ get; set; }//טבלה סינון מספרי קטלוגים מטבלת SQL -נועד לcreate
       


        public string NameUser { get; set; }
        public string MemberNameUser { get; set; }//תווים 10 עם ריפוד של

        DBService DBS = new DBService();
        public Dictionary<string, string> CodeAndDescWaste { get; set; }//קוד פסולת ותיאורו
        public Dictionary<string, string> UnitMessaure { get; set; }//מידת יחידה עבור פריט-תערובות בקילו לדוגמא
        public List<decimal> DepartmentsList { get; set; }

        bool ExistMember = false;//אם הייתה בעיה בשליחת נתונים שלא יקרא שוב לממבר כי זה יזרוק אותי מהתוכנית
        public Waste()
        {
            DayLetter = "";
            DepartmentsList = new List<decimal>();//בדיקת קיום מחלקות
            BringUnitMessaure();////-נגיד תערובות זה קילו מביא יחידות מידה עבור הפריטים
            NameUser = Environment.UserName.ToUpper();//מקבל שם יוזר עבור ממבר
            //הממבר חייב להיות בגודל 10
            if (NameUser.Length<10)
                 MemberNameUser = NameUser.PadRight(10, ' ');
            else if(NameUser.Length>10)
            {
                NameUser = NameUser.Substring(0, 10);
                MemberNameUser = NameUser;
            }
            ListRowCatalogNum = new List<CatalogNum>();
            CatalogNum catalogNumToFind = new CatalogNum();
            BringAllCatalogNumber();//שולף מטבלת SQL את כל המק"טים והבנים שלהם
        }



        /// <summary>
        ///מקים חדש בדיקה אם קיים הממבר בקובץ נין

        /// </summary>
        public void CreateMember()
        {
            string qry = $@"call BPCSOALI.SFC002('NIN       ','BPCSDV30  ','{MemberNameUser}')";//יבדוק אם יש ממבר אם לא יצור חדש, לבדוק אם יש רווחים/תוכנית שאלי יצר באס400
            DBS.executeInsertQuery(qry);
        }



        /// <summary>
        /// שולף את כל המק"טים והמק"טים הבנים שלהם מטבלת SQL
        /// </summary>
        private void BringAllCatalogNumber()
        {
            DbServiceSQL dbCatalog = new DbServiceSQL();
            DataTable CatalogNumTableSql = new DataTable();
            string qry = $@"SELECT *
                          FROM WasteTable";
            CatalogNumTableSql = dbCatalog.executeSelectQueryNoParam(qry);
            CatalogFilterTable = new DataView(CatalogNumTableSql);//להמשך יביא לסינונים בהתאם לטבלה שנבנית

            //הוספת אובייקטים לרשימה
            foreach (DataRow rowView in CatalogNumTableSql.Rows)
            {
                CatalogNumObj = new CatalogNum(rowView["CatalogNumber"].ToString(), rowView["CatalogNumberSon"].ToString(), rowView["Description"].ToString(),
                                               int.Parse(rowView["Department"].ToString().Trim()), rowView["Machine"].ToString(), int.Parse(rowView["Shift"].ToString()), int.Parse(rowView["WorkCenter"].ToString()), rowView["TableType"].ToString().Trim());//הוספה לרשימה את כל המקטים
                ListRowCatalogNum.Add(CatalogNumObj);
            }

            //לשאילתת שליפת נתונים
            ForQueryCatalogNum = "";
            ForQueryDepartment = "";
            DataTable distinctCatalogValues = CatalogFilterTable.ToTable(true, "CatalogNumberSon");
            foreach (DataRow row in distinctCatalogValues.Rows)
            {
                ForQueryCatalogNum += $@"'{row["CatalogNumberSon"].ToString().Trim()}',";//בשביל שאילתת השליפה
            }
            ForQueryCatalogNum = ForQueryCatalogNum.Remove(ForQueryCatalogNum.Length - 1);
            DataTable distinctDepartmentValues = CatalogFilterTable.ToTable(true, "Department");
            foreach (DataRow row in distinctDepartmentValues.Rows)
            {
                ForQueryDepartment += $@"{row["Department"].ToString().Trim()},";//בשביל שאילתת השליפה
            }
            ForQueryDepartment = ForQueryDepartment.Remove(ForQueryDepartment.Length - 1);

            //שליפה של איזה מקטים צריך עבור טבלת ריכוזי פסולות
            qry = $@"SELECT CatalogNumber
                   FROM WasteSum";
            DtTotalSql = dbCatalog.executeSelectQueryNoParam(qry);
        }

        /// <summary>
        /// שליפת נתוני פסולות מs400
        /// </summary>
        public void FillWasteTable()
        {
            dataTableWaste = new DataTable();
            string qry = $@"SELECT trim(t.tprod) as CatalogNumber, t.TTDTE as Date,t.TREF as Department,t.thmach as Machine,case when left(trim(t. TCOM),1) ='1' or left(trim(t. TCOM),1)='2' or left(trim(t. TCOM),1)='3'  then left(trim(t. TCOM),1)  else '0' end shift,sum(INTEGER(round(t.TQTY,0))) as Quantity
                          FROM  BPCSDV30.ITH t  
                          WHERE t.TTDTE between {DateByS400Start} and {DateByS400End} and t.TPROD in ({ForQueryCatalogNum}) and t.TREF in({ForQueryDepartment})
                          GROUP BY t.tprod,t.TTDTE,t.TREF,t.thmach,case when left(trim(t. TCOM),1) ='1' or left(trim(t. TCOM),1)='2' or left(trim(t. TCOM),1)='3'  then left(trim(t. TCOM),1)  else '0' end";
            dataTableWaste = DBS.executeSelectQueryNoParam(qry);
        }



        /// <summary>
        /// הכנסת ערכים לשדות של התאריכים
        /// </summary>
        public void InsertDatesValues(string MonthNumber, int Year, int Days)
        {

            this.DaysInMonth = Days;//כמה ימים בחודש העכשווי
            this.MonthNumber = MonthNumber;
            this.Year = Year;
            YearS400 = Year - 1928;//השנה המוזרה לפי s400
            DateByS400Start = YearS400.ToString() + MonthNumber + "01";//בניית תאריך התחלה
            DateByS400End = YearS400.ToString() + MonthNumber + Days.ToString();//בניית תאריך סיום
        }

        /// <summary>
        /// ממלא טבלה ריקה בתאריכים ריקים-יהיה בכל הטבלאות
        /// </summary>
        public void FillTableDates(Product WhichProduct)
        {
            var culture = new System.Globalization.CultureInfo("he-IL");

            //מילוי שמות ימים באותיות+מילוי תאריכים
            for (int i = 1; i <= DaysInMonth; i++)
            {
                DayNumber = i.ToString();
                if (int.Parse(DayNumber) < 10)
                    DayNumber = "0" + DayNumber;//הוספת 0 לחודש חד ספרתי
                var DayName = culture.DateTimeFormat.GetDayName(DateTime.Parse(DayNumber + "/" + MonthNumber + "/" + Year).DayOfWeek);//שם היום למטרת החלפה לאות -עמודה ראשונה
                GetDayLetter(DayName);
                switch (WhichProduct)
                {
                    case Product.Mixtures:
                        DtMixuresGeneral.Rows.Add(DayLetter);
                        DtMixuresGeneral.Rows[i - 1]["תאריך"] = DayNumber + "/" + MonthNumber + "/" + Year.ToString().Substring(2, 2);
                        break;

                    case Product.Fabric:
                        DtFabricGeneral.Rows.Add(DayLetter);
                        DtFabricGeneral.Rows[i - 1]["תאריך"] = DayNumber + "/" + MonthNumber + "/" + Year.ToString().Substring(2, 2);
                        break;

                    case Product.Steel:
                        DtSteelGeneral.Rows.Add(DayLetter);
                        DtSteelGeneral.Rows[i - 1]["תאריך"] = DayNumber + "/" + MonthNumber + "/" + Year.ToString().Substring(2, 2);
                        break;

                    case Product.Total:
                        DtTotal.Rows.Add(DayLetter);
                        DtTotal.Rows[i - 1]["תאריך"] = DayNumber + "/" + MonthNumber + "/" + Year.ToString().Substring(2, 2);
                        break;
                }
            }
        }


        /// <summary>
        /// טבלת תערובות
        /// </summary>
        public DataTable CreateTableMixures()
        {         
            //מילוי טבלת תערובות
            //יצירת עמודות עבור טבלת תערובות
            DtMixuresGeneral.Columns.Add("יום");
            DtMixuresGeneral.Columns.Add("תאריך");

            FillTableDates(Product.Mixtures);//ממלא תאריכים בטבלה
            CatalogFilterTable.RowFilter = "TableType ='תערובות'";
            foreach (DataRowView item in CatalogFilterTable)//הוספת עמודות
            {
                DtMixuresGeneral.Columns.Add(item["Description"].ToString());
            }
            CatalogSons = CatalogFilterTable.ToTable(true, "CatalogNumberSon");//עובר על כל המקטים הבנים של המקט תערובות ומכניס אותם לטבלת התערובות
            foreach (DataRow rowView in CatalogSons.Rows)
            {
                FillEveryTable(dataTableWaste, DtMixuresGeneral, rowView["CatalogNumberSon"].ToString());//תערובות//אחרי שיצרנו את הטבלה משלים אותה דרך נתוני הפסולות
            }
            CreateTotalRows(DtMixuresGeneral);//יצירת שורות טוטל
            return DtMixuresGeneral;
        }


        /// <summary>
        /// טבלת בדים
        /// </summary>
        public DataTable CreateTableFabric()
        {
            //יצירת עמודות עבור טבלת בדים
            DtFabricGeneral.Columns.Add("יום");
            DtFabricGeneral.Columns.Add("תאריך");

            FillTableDates(Product.Fabric);//ממלא תאריכים בטבלה
            CatalogFilterTable.RowFilter = "TableType ='בדים'";
            foreach (DataRowView item in CatalogFilterTable)//הוספת עמודות
            {
                DtFabricGeneral.Columns.Add(item["Description"].ToString());
            }
            CatalogSons = CatalogFilterTable.ToTable(true, "CatalogNumberSon");
            foreach (DataRow rowView in CatalogSons.Rows)
            {
                FillEveryTable(dataTableWaste, DtFabricGeneral, rowView["CatalogNumberSon"].ToString());//תערובות//אחרי שיצרנו את הטבלה משלים אותה דרך נתוני הפסולות
            }
            CreateTotalRows(DtFabricGeneral);//יצירת שורות טוטל
            return DtFabricGeneral;

        }

        /// <summary>
        /// טבלת פלדה
        /// </summary>
        public DataTable CreateTableSteel()
        {
            DtSteelGeneral.Columns.Add("יום");
            DtSteelGeneral.Columns.Add("תאריך");

            FillTableDates(Product.Steel);//ממלא תאריכים
            CatalogFilterTable.RowFilter = "TableType ='פלדה'";
            foreach (DataRowView item in CatalogFilterTable)//הוספת עמודות
            {
                DtSteelGeneral.Columns.Add(item["Description"].ToString());
            }
            CatalogSons = CatalogFilterTable.ToTable(true, "CatalogNumberSon");
            foreach (DataRow rowView in CatalogSons.Rows)
            {
                FillEveryTable(dataTableWaste, DtSteelGeneral, rowView["CatalogNumberSon"].ToString());//תערובות//אחרי שיצרנו את הטבלה משלים אותה דרך נתוני הפסולות
            }
            CreateTotalRows(DtSteelGeneral);//יצירת שורות טוטל
            return DtSteelGeneral;
        }

        /// <summary>
        /// טבלת ריכוזי פסולות
        /// </summary>
        public DataTable CreateTableTotal()
        {
            DtTotal.Columns.Add("יום");
            DtTotal.Columns.Add("תאריך");
            for (int i = 0; i < DtTotalSql.Rows.Count; i++)
            {
                DtTotal.Columns.Add(DtTotalSql.Rows[i]["CatalogNumber"].ToString());//תערובות
            }
    

            FillTableDates(Product.Total);//ממלא תאריכים
            GetSumOfProduct();
            return DtTotal;
        }





        /// <summary>
        /// המשך שיבוץ כל טבלה -המשך של פונקצית fillWasteTable
        /// </summary>
        public void FillEveryTable(DataTable dataTableWaste, DataTable SpecificTable, string CatalogNumber)//SpecificTable-mixures/fabric/steel,catalog number-מגדיר את הקוד פריט עליו אנחנו עובדים כרגע
        {

            int Quantity; //כמות ליום
            string WhichDay;// היום בו דווח 
            string ByWho="";//תחת מי הכמות
            DataColumnCollection columns = SpecificTable.Columns;//אוסף שמות העמודות
            DataView filterTable = new DataView(dataTableWaste);
            filterTable.RowFilter = "CatalogNumber = '" + CatalogNumber + "'";
            foreach (DataRowView row in filterTable)
            {
                Quantity = int.Parse(row["Quantity"].ToString());
                WhichDay = row["Date"].ToString().Substring(4, 2);
                catalogNumToFind = ListRowCatalogNum.Find(x => x.CatalogNumSon==row["CatalogNumber"].ToString() && x.Department == int.Parse(row["Department"].ToString()) && x.Machine==row["Machine"].ToString() && x.Shift==int.Parse(row["Shift"].ToString()));
                if(catalogNumToFind!=null)
                ByWho = catalogNumToFind.Description;

                //שיבוץ בטבלה בדים פלדה תערובות תלוי איזה טבלה ספציפית נשלחה לפונקציה
                for (int i = 0; i < SpecificTable.Rows.Count; i++)
                {

                    if (columns.Contains(ByWho))//בודק אם בכלל קיימת עמודה כזאת בטבלה
                    {
                        if (SpecificTable.Rows[i]["תאריך"].ToString().Split('/').First() == WhichDay)
                        {
                            SpecificTable.Rows[i][ByWho] = Quantity;//הכנסת הכמות לטור הרלוונטי
                        }
                    }
                }
            }

        }


        public void CreateTotalRows(DataTable SpecificTable)
        {
            //סיכומי אחוזים מכל טור
          
                SpecificTable.Rows.Add("TOTAL");
                SpecificTable.Rows.Add("PRECENT");

                int SumColumn;
                double SumTotal = 0;
                for (int i = 2; i < SpecificTable.Columns.Count; i++)//מתחיל אחרי תאריך ויום
                {
                    SumColumn = 0;
                    for (int j = 0; j < SpecificTable.Rows.Count - 2; j++)
                    {
                        if (!string.IsNullOrEmpty(SpecificTable.Rows[j][i].ToString()))
                            SumColumn += int.Parse(SpecificTable.Rows[j][i].ToString());
                    }
                    SpecificTable.Rows[SpecificTable.Rows.Count - 2][i] = SumColumn;
                    SumTotal += SumColumn;
                }

                //חישוב אחוזים
                for (int i = 2; i < SpecificTable.Columns.Count; i++)
                {
                    SpecificTable.Rows[SpecificTable.Rows.Count - 1][i] = ((double.Parse(SpecificTable.Rows[SpecificTable.Rows.Count - 2][i].ToString()) / SumTotal)).ToString("P2");
                }
            
        }

        /// <summary>
        /// שולף סיכומי כמויות מכל פריט
        /// </summary>
        public void GetSumOfProduct()
        {
            dataTableTotal = new DataTable();
            string qry = $@"SELECT trim(t.tprod) as CatalogNumber, t.TTDTE as Date,sum(INTEGER(round(t.TQTY,0))) as Quantity
                          FROM  BPCSDV30.ITH t  
                          WHERE t.TTDTE between {DateByS400Start} and {DateByS400End} and t.TPROD in ('FB-40002','SC-FC010','SC-FC050','SC-BE200','SC-SC100') 
                          GROUP BY t.tprod,t.TTDTE";
            dataTableTotal = DBS.executeSelectQueryNoParam(qry);
            int Quantity;
            string WhichDay,CatalogNumber;
            DataColumnCollection columns = DtTotal.Columns;//אוסף שמות העמודות
            foreach (DataRow row in dataTableTotal.Rows)
            {
                Quantity = int.Parse(row["Quantity"].ToString());
                WhichDay = row["Date"].ToString().Substring(4, 2);
                CatalogNumber = row["CatalogNumber"].ToString();

                for (int i = 0; i < DtTotal.Rows.Count; i++)
                {
                    if (DtTotal.Rows[i]["תאריך"].ToString().Split('/').First() == WhichDay)
                        DtTotal.Rows[i][CatalogNumber] = Quantity;
                }
            }

            //שורה אחרונה סכום כמויות חודשי
            qry = $@"SELECT trim(t.tprod) as CatalogNumber,sum(INTEGER(round(t.TQTY,0))) as Quantity
                          FROM  BPCSDV30.ITH t  
                          WHERE t.TTDTE between {DateByS400Start} and {DateByS400End} and t.TPROD in ('FB-40002','SC-FC010','SC-FC050','SC-BE200','SC-SC100')
                          GROUP BY t.tprod";
            dataTableTotal = DBS.executeSelectQueryNoParam(qry);
            DtTotal.Rows.Add("Total");
            DtTotal.Rows.Add("Precentage");
            double SumAll = 0;
            for (int i = 0; i < dataTableTotal.Rows.Count; i++)
            {
                CatalogNumber = dataTableTotal.Rows[i]["CatalogNumber"].ToString();
                DtTotal.Rows[DtTotal.Rows.Count - 2][CatalogNumber] = dataTableTotal.Rows[i]["Quantity"].ToString();
                SumAll+= double.Parse( dataTableTotal.Rows[i]["Quantity"].ToString());
            }
            for (int i = 2; i < DtTotal.Columns.Count; i++)
            {   
                if(!string.IsNullOrEmpty(DtTotal.Rows[DtTotal.Rows.Count - 2][i].ToString())&& DtTotal.Columns[i].ColumnName!="SC-SC100")
                DtTotal.Rows[DtTotal.Rows.Count - 1][i] = ((double.Parse(DtTotal.Rows[DtTotal.Rows.Count - 2][i].ToString()) / SumAll)).ToString("P2");
            }


        }

        /// <summary>
        /// מחזיר שם!!  מחלקה שאחראית על הפסולת או שם מכונה
        /// </summary>
        /// <param name="byWho"></param>
        public string GetStringDepartmentOrMachine(string ByWho,string CatalogNumber, string Machine)
        {
            switch (ByWho)//המרת מיקסר קלנדר וברומו שרוף למספרי מחלקות
            {
                case "52":
                    ByWho = "מיקסר";
                    break;

                case "55":
                    ByWho = "קלנדר";
                    break;

                case "9050":
                    ByWho = "ברומו שרוף";
                    break;

                case "9001":
                    ByWho = "הנצלה טכסטיל";
                    break;

                case "9020":
                    ByWho = "תקלות גימום טקסטיל";
                    break;

                case "9010":
                    ByWho = "הנצלה פוליאסטר";
                    break;

                case "9030":
                    ByWho = "עודפי ייצור ובניה";
                    break;

                case "9040":
                    ByWho = "בדים r&d";
                    break;

                case "0":
                    if (CatalogNumber == "SC-FC050" && Machine == "32")
                        ByWho = "ברסטוף";
                    else if (CatalogNumber == "SC-FC050" && Machine == "31")
                        ByWho = "ספדון";
                    else if (CatalogNumber == "SC-FC050" && Machine == "LOC" )
                        ByWho = "עודפי פלדה";
                    else if (CatalogNumber == "SC-SC100")
                        ByWho = "חוטי פלדה גולמיים";
                    else if (CatalogNumber == "SC-BE200")
                        ByWho = "חישוקים";
                    //else if (CatalogNumber == "SC-FC055")
                    //    ByWho = "עודפי פלדה";
                    break;

  
            }
            return ByWho;
        }

        /// <summary>
        /// מחזיר מספר!!  מחלקה שאחראית על הפסולת או שם מכונה
        /// </summary>
        /// <param name="byWho"></param>
        public string GetIntDepartmentOrMachine(string ByWho)
        {
            switch (ByWho)//המרת מיקסר קלנדר וברומו שרוף למספרי מחלקות
            {
                case "מיקסר":
                    ByWho = "52";
                    break;

                case "קלנדר":                    
                    ByWho = "55";
                    break;

                case "ברומו שרוף":
                    ByWho = "9050";
                    break;

                case "לילה חתכן 1":
                case "ערב חתכן 1":
                case "בוקר חתכן 1":
                case "ערב חתכן 3":
                case "בוקר חתכן 3":
                case "חתכן ברקר":
                case "לילה חתכן 5":
                case "ערב חתכן 5":
                case "בוקר חתכן 5":
                case "מכונת RJS חתכן 2":
                    ByWho = "57";
                    break;

                case "הנצלה טכסטיל":
                    ByWho = "9001";
                    break;

                case "הנצלה פוליאסטר":
                    ByWho = "9010";
                    break;

                case "תקלות גימום טקסטיל":
                    ByWho = "9020";
                    break;

                case "עודפי ייצור ובניה":
                    ByWho = "9030";
                    break;

                case "נסיונות":
                    ByWho = "9040";
                    break;

                case "ברסטוף":
                case "ספדון":
                case "עודפי פלדה":
                case "חישוקים":
                case "חוטי פלדה גולמיים":
                //case "נסיונות פלדה":
                    ByWho = "0";
                    break;

            
                        
            }
            return ByWho;
        }

        /// <summary>
        /// מחלקה 57 מאופיינת במכונות בטבלת בדים ,צריך לגלות לאיזה טור להכניס
        /// </summary>
        /// <param name="byWho"></param>
        public string GetColumnFabricName57(string Machine,string Shift)
        {
            string ColumnFabricName="";
            switch (Machine)
            {
                case "10":
                    if (Shift == "1")
                        ColumnFabricName = "לילה חתכן 1";
                    else if (Shift == "2")
                        ColumnFabricName = "בוקר חתכן 1";
                    else if(Shift=="3")
                        ColumnFabricName = "ערב חתכן 1";
                    break;

                case "30":
                    if (Shift == "2")
                        ColumnFabricName = "בוקר חתכן 3";
                    else if(Shift=="3")
                        ColumnFabricName = "ערב חתכן 3";
                    break;

                case "2":
                    ColumnFabricName = "חתכן ברקר";
                    break;

                case "1":
                    if (Shift == "1")
                        ColumnFabricName = "לילה חתכן 5";
                    else if (Shift == "2")
                        ColumnFabricName = "בוקר חתכן 5";
                    else if(Shift=="3")
                        ColumnFabricName = "ערב חתכן 5";
                    break;

                case "40":
                    ColumnFabricName = "מכונת RJS חתכן 2";
                    break;

            }
            return ColumnFabricName;
        }

        /// <summary>
        /// עבור טבלה של פלדה יש כמה קודי פריטים(חוטי פלדה חישוקים ועודפי פלדה
        /// </summary>
        internal string SteelCatalogNumber(string dataPropertyName)
        {
            switch(dataPropertyName)
            {
                //case "עודפי פלדה":
                //    dataPropertyName = "SC-FC055";
                //    break;

                case "חישוקים":
                    dataPropertyName = "SC-BE200";
                    break;

                case "חוטי פלדה גולמיים":
                    dataPropertyName = "SC-SC100";
                    break;

                default:
                    dataPropertyName = "SC-FC050";
                    break;
            }
            return dataPropertyName;
        }

        /// <summary>
        /// מחלקה 57 מאופיינת במכונות בטבלת בדים ,מקבל את מספר המכונה
        /// </summary>
        /// <param name="byWho"></param>
        public int[] GetColumnFabricNumber57(string ColumnName)
        {
            int [] NumberMachineAndShift=new int[3];//ו2 זה מרכז עבודה מקום 0 זה מכונה מקום 1 משמרת
            switch (ColumnName)
            {
                case "לילה חתכן 1":
                    NumberMachineAndShift[0] = 10;
                    NumberMachineAndShift[1] = 1;
                    NumberMachineAndShift[2] = 57002;
                    break;

                case "בוקר חתכן 1":
                    NumberMachineAndShift[0] = 10;
                    NumberMachineAndShift[1] = 2;
                    NumberMachineAndShift[2] = 57002;
                    break;

                case "ערב חתכן 1":
                    NumberMachineAndShift[0] = 10;
                    NumberMachineAndShift[1] = 3;
                    break;

                case "בוקר חתכן 3":
                    NumberMachineAndShift[0] = 30;
                    NumberMachineAndShift[1] = 2;
                    NumberMachineAndShift[2] = 57003;
                    break;

                case "ערב חתכן 3":
                    NumberMachineAndShift[0] = 30;
                    NumberMachineAndShift[1] = 3;
                    NumberMachineAndShift[2] = 57003;
                    break;

                case "חתכן ברקר":
                    NumberMachineAndShift[0] = 2;
                    NumberMachineAndShift[1] = 0;
                    NumberMachineAndShift[2] = 157002;
                    break;

                case "לילה חתכן 5":
                    NumberMachineAndShift[0] = 1;
                    NumberMachineAndShift[1] = 1;
                    NumberMachineAndShift[2] = 157001;
                    break;

                case "בוקר חתכן 5":
                    NumberMachineAndShift[0] = 1;
                    NumberMachineAndShift[1] = 2;
                    NumberMachineAndShift[2] = 157001;
                    break;

                case "ערב חתכן 5":
                    NumberMachineAndShift[0] = 1;
                    NumberMachineAndShift[1] = 3;
                    NumberMachineAndShift[2] = 157001;
                    break;

                case "מכונת RJS חתכן 2":
                    NumberMachineAndShift[0] = 40;
                    NumberMachineAndShift[1] = 0;
                    NumberMachineAndShift[2] = 570004;
                    break;

                //פלדה
                case "ברסטוף":
                    NumberMachineAndShift[0] = 32;
                    NumberMachineAndShift[1] = 0;
                    NumberMachineAndShift[2] = 720032;
                    break;

                case "ספדון":
                    NumberMachineAndShift[0] = 31;
                    NumberMachineAndShift[1] = 0;
                    NumberMachineAndShift[2] = 720031;
                    break;



            }
            return NumberMachineAndShift;
        }

        /// <summary>
        /// מביא קוד פריט במידה ומדובר בפלדה שלא SC-FC050
        /// </summary>
        internal string GetCatalogNumForsteel(string headerText)
        {
            switch (headerText)
            {
                case "חוטי פלדה גולמיים":
                    headerText = "SC-SC100";
                    break;

                case "חישוקים":
                    headerText = "SC-BE200";
                    break;

                //case "עודפי פלדה":
                //    headerText = "SC-FC055";
                //    break;

                default:
                    headerText = "SC-FC050";
                    break;
            }
            return headerText;
        }


        private void GetDayLetter(string DayName)
        {
            switch (DayName)
            {
                case "יום ראשון":
                    DayLetter = "א";
                    break;

                case "יום שני":
                    DayLetter = "ב";
                    break;

                case "יום שלישי":
                    DayLetter = "ג";
                    break;

                case "יום רביעי":
                    DayLetter = "ד";
                    break;

                case "יום חמישי":
                    DayLetter = "ה";
                    break;

                case "יום שישי":
                    DayLetter = "ו";
                    break;

                case "שבת":
                    DayLetter = "ש";
                    break;
            }
        }

        /// <summary>
        /// הכנסת נתונים לקובץ נין
        /// </summary>
        public bool InsertDataToNin(DataTable TableForUpdate)
        {
            if (TableForUpdate.Rows.Count != 0)
            {
                try
                {
                    string Date;
                    string ItemNumber, Quantity, Department, ReasonCode, TransactionType, UnitKg,Machine,ShiftAndComment,WorkCenter;
                    string StrSql = "";
                    try
                    {
                        if (!ExistMember)//אם אין קריאה לממבר
                        {
                            StrSql = $@"Create alias BPCSDV30.{NameUser} for BPCSDV30.NIN({NameUser})"; //יוצר טבלה וירטואלית שם מחשב
                            DBS.executeInsertQuery(StrSql);
                        }
                    }
                    catch(Exception ex)
                    {
                        ExistMember = true;//קיימת קריאה כבר לממבר הזה
                        InsertDataToNin(TableForUpdate);
                    }
                    for (int i = 0; i < TableForUpdate.Rows.Count; i++)
                    {
                        Date =ChangeDateToS400(DateTime.Parse(TableForUpdate.Rows[i]["תאריך"].ToString()).Year, DateTime.Parse(TableForUpdate.Rows[i]["תאריך"].ToString()).Month.ToString(), DateTime.Parse(TableForUpdate.Rows[i]["תאריך"].ToString()).Day.ToString());//משנה את התאריך לתאריך המוזר של S400              
                        ItemNumber = TableForUpdate.Rows[i]["קוד פריט"].ToString();//קוד פריט לדוגמא תערובות
                        Department = TableForUpdate.Rows[i]["מחלקה"].ToString();             
                        Quantity = TableForUpdate.Rows[i]["כמות"].ToString();
                        //TransactionType = TableForUpdate.Rows[i]["סוג תנועה"].ToString();
                        TransactionType = "PX";
                        UnitKg = TableForUpdate.Rows[i]["יחידת משקל"].ToString();
                        Machine = TableForUpdate.Rows[i]["מכונה"].ToString();
                        ShiftAndComment = TableForUpdate.Rows[i]["משמרת"].ToString()+ TableForUpdate.Rows[i]["הערות"].ToString();//שדה שמכיל משמרת ותו אחריו הערות
                        WorkCenter = TableForUpdate.Rows[i]["מרכז עבודה"].ToString();
                        ReasonCode= TableForUpdate.Rows[i]["קוד תקלה"].ToString();
                        StrSql = $@"INSERT INTO BPCSDV30.{NameUser} VALUES('IN',1,2,'{ItemNumber}','SC','LOT','{Machine}',{Quantity},{Department},99,cast('{ShiftAndComment}' as Character (100) CCSID 424)  ,{Date},'{UnitKg}','{TransactionType}','{ReasonCode}',{WorkCenter},' ')";
                        DBS.executeInsertQuery(StrSql);
                    }

                    StrSql = $@"drop alias BPCSDV30.{NameUser}";
                    DBS.executeInsertQuery(StrSql);
                    StrSql = $@"call SYS.SNDDTAQ('CIM600    BPCSDALI  ','00050','" + ("PX        " + NameUser).PadRight(50, ' ') + "')";
                    //StrSql = $@" SNDDTAQ DTAQ(BPCSDALI/CIM600) LEN(00050) TOHEN('PX        HSHIFTAN') ";//s400
                    DBS.executeInsertQuery(StrSql);
                    //StrSql = $@"call SYS.SNDDTAQ('CIM600    BPCSDALI  ','00050','" + ("P5        " + NameUser).PadRight(50, ' ') + "')";
                    //DBS.executeInsertQuery(StrSql);
                    MessageBox.Show("נתונים נשמרו בהצלחה");

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    //MessageBox.Show("שגיאה בשמירת נתונים", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }
            return true;

        }



        /// <summary>
        /// משיג פרטי קודי פסולת של החודש הנוכחי
        /// </summary>
        /// <returns></returns>
        public DataTable GetCellsData()
        {
            DataTable GetCellsDataTable = new DataTable();
            //string qry = $@"SELECT a.MSMKT as product,a.MSMCN as machine,a.MSMSM as shift,a.MSKOD as code,a.MSQTDT as date, a.MSMHL as department,a.MSQTY as quantity,b.MSDEC as description            //                FROM MSK.MSVQTP as a left join msk.MSSVGP as b on a.mskod=b.msskod
            //                WHERE MSQTDT between {Year.ToString().Substring(2, 2) + MonthNumber}01 and {Year.ToString().Substring(2,2) + MonthNumber + DaysInMonth}";
            string qry = $@"SELECT t.tprod as CatalogNumber, t.TTDTE as Date,t.TREF as Department,t.thmach as Machine,case when left(trim(t. TCOM),1) ='1' or left(trim(t. TCOM),1)='2' or left(trim(t. TCOM),1)='3'  then left(trim(t. TCOM),1)  else 0 end shift,sum(INTEGER(round(t.TQTY,0))) as Quantity,t.tres as ReasonCode,trim(z.data) as Description,
                            case when left(trim(t. TCOM),1) ='1' or left(trim(t. TCOM),1)='2' or left(trim(t. TCOM),1)='3'  then SUBSTR(t.tcom,2,15)  else t.tcom end Comment
                            FROM  BPCSDV30.ITH t left join 
                            (select distinct substring(PKEY,7,2) as Rcode, data                              from BPCSDV30.ZPAL01
                   			 where substring(PKEY,5,2) in ('PX')
		                    ) z on  t.tres=z.Rcode  
                          WHERE t.TTDTE between {DateByS400Start} and {DateByS400End} and t.TPROD in ({ForQueryCatalogNum}) and t.TREF in({ForQueryDepartment})
                          GROUP BY t.tprod,t.TTDTE,t.TREF,t.thmach,t.tcom,t.tres,z.data,SUBSTR(t.tcom,2,15),t.tcom
                          ORDER BY t.tprod,TTDTE";//תוספת של קוד תקלה-כל מה ששונה מ01  and t.tres<>01 and TTYPE='PX' , t.tres ||' ' ||trim(z.data) as ReasonCode
            GetCellsDataTable = DBS.executeSelectQueryNoParam(qry);
            return GetCellsDataTable;
        }

        /// <summary>
        /// משיג סכומי סיבות ספדון עבור אלחנן בדו"ח אקסל
        /// </summary>
        public DataTable GetSumSteelForExcel()
        {
            DataTable dataTable = new DataTable();
            string qry = $@" SELECT  trim(z.data) as Description,sum(INTEGER(round(t.TQTY,0))) as Quantity,t.tres as ReasonCode
                            FROM  BPCSDV30.ITH t left join 
                            (select distinct substring(PKEY,7,2) as Rcode, data 
                             from BPCSDV30.ZPAL01
                   			 where substring(PKEY,5,2) in ('PX')
		                    ) z on  t.tres=z.Rcode  
                          WHERE t.TTDTE between {DateByS400Start} and {DateByS400End} and t.TPROD='SC-FC050' and t.TREF in(0,52,54,55,57,71,160,9001,9010,9020,9030,9040,9050) and t.tres in(02,03,05,12)
                          GROUP BY t.tprod,t.TREF,t.thmach,t.tcom,t.tres,z.data,SUBSTR(t.tcom,2,15),t.tcom
                          ORDER BY t.tres";//המספרים של קודי תקלה הם שייכים לפירוט ספדון
            dataTable = DBS.executeSelectQueryNoParam(qry);
            return dataTable;
        }

        /// <summary>
        /// רשימת קודי פסולות עבור הפירוט-ימולא בקומבובוקס
        /// </summary>
        public DataTable GetWasteCodes()
        {
            DataTable GetWasteCodesTable = new DataTable();
            //string qry = $@"SELECT   MSSKOD as ReasonCode, MSDEC as desc         
            //              FROM msk.MSSVGP";
            string qry = $@"SELECT distinct substring(PKEY,7,2)  as ReasonCode, trim(DATA)||' '||substring(PKEY,7,2) as desc     
                            FROM BPCSDV30.ZPAL01 
                            WHERE substring(PKEY,5,2) in ('PX')";//למחוק פי 3 ופי4 זה לטסטים  
            GetWasteCodesTable = DBS.executeSelectQueryNoParam(qry);
            CodeAndDescWaste = new Dictionary<string, string>();//הוספה למילון קוד פסולת ותיאורו
            for (int i = 0; i < GetWasteCodesTable.Rows.Count; i++)
            {
                CodeAndDescWaste.Add(GetWasteCodesTable.Rows[i]["ReasonCode"].ToString(), GetWasteCodesTable.Rows[i]["desc"].ToString().Substring(3));
            }
            return GetWasteCodesTable;
        }

        /// <summary>
        /// מגדיר יחידות מידה של משקל עבור כל סוג פריט
        /// </summary>
        public void BringUnitMessaure()
        {
            UnitMessaure = new Dictionary<string, string>();
            DataTable dataTable = new DataTable();
            string qry = $@"SELECT trim(iprod) as prod,iums
                            FROM  BPCSDV30.IIML01   
                            WHERE iprod in ('FB-40002','SC-FC010','SC-FC050','SC-BE200','SC-SC100')";
            dataTable=DBS.executeSelectQueryNoParam(qry);
            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                UnitMessaure.Add(dataTable.Rows[i]["prod"].ToString(), dataTable.Rows[i]["iums"].ToString());
            }

        }


        public void DeleteCellUpdate()
        {
            string qry = $@"delete 
                          FROM MSK.MSVQTP
                          WHERE MSMKT='{CatalogNumber}' and MSQTDT={DateOfCellUpdate} and MSMHL={Department} and MSMCN='{Machine}' and  MSMSM='{Shift}'";
            DBS.executeInsertQuery(qry);
        }


        /// <summary>
        /// שינוי תאריך לפורמט של אס400 כולל שנה פחות 1928
        /// </summary>
        public string ChangeDateToS400(int Year,string MonthNumber,string DayNumber)
        {
            YearS400 = Year - 1928;
            if (int.Parse(MonthNumber) < 10)
                MonthNumber = "0" + MonthNumber;
            if (int.Parse(DayNumber) < 10)
                DayNumber = "0" + DayNumber;
            return YearS400 + MonthNumber + DayNumber;
        }

        public void NewDataTable()
        {
            DtMixuresGeneral = new DataTable();//טבלת תערובות
            DtFabricGeneral = new DataTable();//טבלת בדים
            DtSteelGeneral = new DataTable();//טבלת פלדה
            NewTotal();
        }

        public void NewTotal()
        {
            DtTotal = new DataTable();
        }

        /// <summary>
        /// בודק את קובץ נין,אם יש עדיין נתונים שם ,לא הכל עבר לith
        /// </summary>
        public bool CheckNinData()
        {
            DataTable NinTable = new DataTable();
            string StrSql;
            try
            {
                if (!ExistMember)//אם אין קריאה לממבר
                {
                    StrSql = $@"Create alias BPCSDV30.{NameUser} for BPCSDV30.NIN({NameUser})"; //יוצר טבלה וירטואלית שם מחשב
                    DBS.executeInsertQuery(StrSql);
                }
            }
            catch (Exception ex)
            {
                ExistMember = true;//קיימת קריאה כבר לממבר הזה
                CheckNinData();
            }
            StrSql = $@"SELECT * 
                       FROM BPCSDV30.{NameUser}";
            NinTable = DBS.executeSelectQueryNoParam(StrSql);
            StrSql = $@"drop alias BPCSDV30.{NameUser}";//זריקת ממבר אחרי 
            DBS.executeInsertQuery(StrSql);
            if (NinTable.Rows.Count > 0)//עדיין יש נתונים בנין שלא נכנסו לITH
                return true;
            else
            {
                return false;
            }
        }

    }

    }

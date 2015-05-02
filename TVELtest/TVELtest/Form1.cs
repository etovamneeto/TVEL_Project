using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using RiskCalculatorLib;
//Хреновины имеют значение, когда сравниваешь строки и строки, т.е. когда тебе надо ввести в текстовую ячейку текстовый параметр; В другие, где инты или флоаты, туда не нужны они
//"UPDATE [names] SET [userName]='" + nameBox.Text + "', [age]=" + Convert.ToInt32(ageBox.Text), но + " WHERE [id]=" + Convert.ToInt32(table.Rows[0]["id"]), connection);

namespace TVELtest
{
    public partial class Form1 : Form
    {

        
        /*-----Описание класса Человек, в котором хранится информация: id, пол, возраст при облучении, дозовая история-----*/
        public class Man
        {
            
            private int id = 0;
            private byte sex = 0;
            private short ageAtExp = 0;
            private RiskCalculator.DoseHistoryRecord[] doseHistory = null; 

            public Man(int id, byte sex, short ageAtExp, RiskCalculator.DoseHistoryRecord[] doseHistory)
            {
                this.id = id;
                this.sex = sex;
                this.ageAtExp = ageAtExp;
                this.doseHistory = doseHistory;
            }

            public int getID() { return this.id; }
            public byte getSex() { return this.sex; }
            public short getAgeAtExp() { return this.ageAtExp; }
            public RiskCalculator.DoseHistoryRecord[] getDoseHistory() { return this.doseHistory; }

            public void setID(int id) { this.id = id; }
            public void setSex(byte sex) { this.sex = sex; }
            public void setAgeAtExp(short ageAtExp) { this.ageAtExp = ageAtExp;}
            public void getDoseHistory(RiskCalculator.DoseHistoryRecord[] doseHistory) { this.doseHistory = doseHistory; }
        }

        /*-----Описание класса Объект, представляющий собой строку таблицы с параметрами: id, пол, доза суммарная, доза внутренняя, возраст при облучении-----*/
        public class dbObject
        {
            private int id = 0;
            private short ageAtExp = 0;
            private double dose = 0;
            private double doseInt = 0;
            private byte sex = 0;

            public dbObject(int id, byte sex, short ageAtExp, double dose, double doseInt)
            {
                this.id = id;
                this.sex = sex;
                this.ageAtExp = ageAtExp;
                this.dose = dose;
                this.doseInt = doseInt;
            }

            public void setId(int id) { this.id = id; }
            public void setAgeAtExp(short ageAtExp) { this.ageAtExp = ageAtExp; }
            public void setDose(double dose) { this.dose = dose; }
            public void setDoseInt(double doseInt) { this.doseInt = doseInt; }
            public void setSex(byte sex) { this.sex = sex; }

            public int getId() { return this.id; }
            public short getAgeAtExp() { return this.ageAtExp; }
            public double getDose() { return this.dose; }
            public double getDoseInt() { return this.doseInt; }
            public byte getSex() { return this.sex; }
        }

        /*-----Описание форм инициализации и инициализация библиотеки с рейтами 2012 года-----*/
        public Form1(String title)
        {
            InitializeComponent();
            this.Text = title;
        }

        String libPath = Path.GetDirectoryName(Application.ExecutablePath) + "\\DataRus2012";

        /*-----Функции для расчета LAR, необходимых для расчета ОРПО*-----*/
        public double getManExtLar(double meanAge)
        {
            double lar = 0;
            double secondPowerElement = (2 / Math.Pow(10, 6)) * Math.Pow(meanAge, 2);
            double firstPowerElement = (-13 / Math.Pow(10, 4)) * meanAge;
            double constant = 9.36 / Math.Pow(10, 2);
            return lar = secondPowerElement + firstPowerElement + constant;
        }

        public double getWomanExtLar(double meanAge)
        {
            double lar = 0;
            double secondPowerElement = (1 / Math.Pow(10, 5)) * Math.Pow(meanAge, 2);
            double firstPowerElement = (-31 / Math.Pow(10, 4)) * meanAge;
            double constant = 17.42 / Math.Pow(10, 2);
            return lar = secondPowerElement + firstPowerElement + constant;
        }

        public double getManIntLar(double meanAge)
        {
            double lar = 0;
            double secondPowerElement = (-3 / Math.Pow(10, 5)) * Math.Pow(meanAge, 2);
            double firstPowerElement = (23 / Math.Pow(10, 4)) * meanAge;
            double constant = 1.15 / Math.Pow(10, 2);
            return lar = secondPowerElement + firstPowerElement + constant;
        }

        public double getWomanIntLar(double meanAge)
        {
            double lar = 0;
            double secondPowerElement = (-4 / Math.Pow(10, 5)) * Math.Pow(meanAge, 2);
            double firstPowerElement = (27 / Math.Pow(10, 4)) * meanAge;
            double constant = 5.02 / Math.Pow(10, 2);
            return lar = secondPowerElement + firstPowerElement + constant;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            RiskCalculatorLib.RiskCalculator.FillData(ref libPath);
            this.Width = 300;
            this.Height = 300;
            this.CenterToScreen();
        }

        private void getOrpoButton_Click(object sender, EventArgs e)
        {
            /*-----Инициализация всяких входных параметров, подключения к БД, парсинга в таблицу нужных столбцов-----*/
            String dbPath = Path.GetDirectoryName(Application.ExecutablePath) + "\\dbTvel.mdb";
            String connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + dbPath;
            OleDbConnection connection = new OleDbConnection(connectionString);
            connection.Open();

            OleDbDataAdapter adapter = new OleDbDataAdapter("SELECT [ID], [Dose], [DoseInt], [Year], [Gender], [BirthYear], [AgeAtExp] FROM [Final] WHERE [Shop]='r3'", connectionString);//Выбор нужных столбцов из нужной таблицы
            DataSet dataSet = new DataSet();
            adapter.Fill(dataSet, "Final");
            DataTable table = dataSet.Tables[0];//Из Final в эту таблицу считываются поля, указанные в запросе; Выборка для МСК (shop = r3)
            /*-----Заполнения списка уникальных ID-----*/
            //List<int> uniqueIdList = new List<int>();
            //for (int i = 0; i < table.Rows.Count; i++)
            //{
            //    uniqueIdList.Add(Convert.ToInt32(table.Rows[i]["id"]));
            //}
            //uniqueIdList = uniqueIdList.Distinct().ToList();

            /*-----Список объектов; достаем все необходимое для расчетов: id, dose, doseInt, ageAtExp, gender-----*/
            List<dbObject> dbRecords = new List<dbObject>();
            for (int i = 0; i < table.Rows.Count; i++)
            {
                dbRecords.Add(new dbObject(Convert.ToInt32(table.Rows[i]["id"]), Convert.ToByte(table.Rows[i]["gender"]), Convert.ToInt16(table.Rows[i]["ageatexp"]), Convert.ToDouble(table.Rows[i]["dose"]), Convert.ToDouble(table.Rows[i]["doseint"])));
            }

            /*-----Список, в котором хранится пол-----*/
            List<byte> dbSex = new List<byte>();
            for (int i = 0; i < dbRecords.Count; i++)
                dbSex.Add(dbRecords[i].getSex());

            /*-----Определение пола; Меньшая цифра пола - М, большая - Ж-----*/
            byte sexMale = dbSex.Min();
            byte sexFemale = dbSex.Max();

            /*-----Список возрастов облучения из БД для М и Ж-----*/
            List<short> dbManAges = new List<short>();
            for (int i = 0; i < dbRecords.Count; i++)
                if(dbRecords[i].getSex() == sexMale)
                    dbManAges.Add(dbRecords[i].getAgeAtExp());

            List<short> dbWomanAges = new List<short>();
            for (int i = 0; i < dbRecords.Count; i++)
                if (dbRecords[i].getSex() == sexFemale)
                    dbWomanAges.Add(dbRecords[i].getAgeAtExp());

            /*-----Min и Max возраста облучения для М и Ж-----*/
            short manMinAge = dbManAges.Min();
            short manMaxAge = dbManAges.Max();

            short womanMinAge = dbWomanAges.Min();
            short womanMaxAge = dbWomanAges.Max();

            /*-----Массивы списков для мужчин и для женщин, в каждом из которых будут храниться объекты, сгруппированные по возрастам облучения. i = 0 - 18-летние, i = 1 - 19-летние и тд-----*/
            List<dbObject>[] manAgesGroupedArray = new List<dbObject>[manMaxAge - manMinAge + 1];
            for (int i = 0; i < manAgesGroupedArray.Length; i++)
                manAgesGroupedArray[i] = new List<dbObject>();

            List<dbObject>[] womanAgesGroupedArray = new List<dbObject>[womanMaxAge - womanMinAge + 1];
            for (int i = 0; i < womanAgesGroupedArray.Length; i++)
                womanAgesGroupedArray[i] = new List<dbObject>();

            /*-----Словари для мужчин и для женщин, где ключ - возраст облучения, а значение - список объектов с этим годом облучения-----*/
            Dictionary<int, List<dbObject>> manDict = new Dictionary<int, List<dbObject>>();      
            for (int i = manMinAge - manMinAge; i <= manMaxAge - manMinAge; i++)
            {
                for (int k = 0; k < dbRecords.Count; k++)
                {
                    if (dbRecords[k].getSex() == sexMale)
                        if (dbRecords[k].getAgeAtExp() == i + manMinAge)
                        {
                            manAgesGroupedArray[i].Add(dbRecords[k]);
                        }
                }
                manDict.Add(i + manMinAge, manAgesGroupedArray[i]);
            }

            Dictionary<int, List<dbObject>> womanDict = new Dictionary<int, List<dbObject>>();
            for (int i = womanMinAge - womanMinAge; i <= womanMaxAge - womanMinAge; i++)
            {
                for (int k = 0; k < dbRecords.Count; k++)
                {
                    if (dbRecords[k].getSex() == sexFemale)
                        if (dbRecords[k].getAgeAtExp() == i + womanMinAge)
                        {
                            womanAgesGroupedArray[i].Add(dbRecords[k]);
                        }
                }
                womanDict.Add(i + womanMinAge, womanAgesGroupedArray[i]);
            }




                    /* ПОЛНОСТЬЮ МЕНЯЕМ ЛОГИКУ ПРОГРАММЫ!
                     * ТО, ЧТО НАПИСАНО НИЖЕ, БУДЕТ ЗАКОММЕНТИРОВАННО,
                     * Т.К. ЭТО НЕПРАВИЛЬНАЯ ЛОГИКА, НО ТАМ МОЖНО ПОДСМОТРЕТЬ РЕШЕНИЯ.
                     * ВСЯ ПРАВИЛЬНАЯ ЛОГИКА ПРОГРАММЫ БУДЕТ НАПИСАНА НАД ЭТОЙ ЗАПИСЬЮ.
                     * РАБОТА ПРОДОЛЖАЕТСЯ . . .
                     */


                    ///*-----Создания массива списков, где каждый элемент массива - это список объектов, id которых совпадают с уникальными id;
                    // * ----например, если уникальный id = 1, то в элемент массива списков записываются все объекты с id = 1.
                    // */
                    //List<dbObject>[] manRecordsList = new List<dbObject>[uniqueIdList.Count];
                    //for (int i = 0; i < uniqueIdList.Count; i++)
                    //{
                    //    List<dbObject> buffer = new List<dbObject>();
                    //    for (int k = 0; k < dbObjectList.Count; k++)
                    //    {
                    //        if (Equals(uniqueIdList[i], dbObjectList[k].getId()))
                    //        {
                    //            buffer.Add(dbObjectList[k]);
                    //        }
                    //        manRecordsList[i] = buffer;
                    //    }
                    //}

                    ///*-----Создание пустого списка дозовых историй; для каждого уникального ID своя дозовая история (по сути, это ячейки, которые надо заполнить)-----*/
                    //List<RiskCalculator.DoseHistoryRecord[]> doseHistoryList = new List<RiskCalculator.DoseHistoryRecord[]>();
                    //for (int i = 0; i < uniqueIdList.Count; i++)
                    //{
                    //    doseHistoryList.Add(new RiskCalculator.DoseHistoryRecord[manRecordsList[i].Count]);
                    //}
                    //foreach (RiskCalculator.DoseHistoryRecord[] note in doseHistoryList)
                    //{
                    //    for (int i = 0; i < note.Length; i++)
                    //        note[i] = new RiskCalculator.DoseHistoryRecord();
                    //}

                    ///*-----Заполнение дозовых историй-----*/
                    //for (int i = 0; i < uniqueIdList.Count; i++)
                    //    for (int k = 0; k < manRecordsList[i].Count; k++)
                    //    {
                    //        doseHistoryList[i][k].AgeAtExposure = manRecordsList[i][k].getAgeAtExp();
                    //        doseHistoryList[i][k].AllSolidDoseInmGy = manRecordsList[i][k].getDose() - manRecordsList[i][k].getDoseInt();
                    //        doseHistoryList[i][k].LeukaemiaDoseInmGy = manRecordsList[i][k].getDose() - manRecordsList[i][k].getDoseInt();
                    //        doseHistoryList[i][k].LungDoseInmGy = manRecordsList[i][k].getDoseInt();
                    //    }
                    //RiskCalculator.DoseHistoryRecord[] rec = doseHistoryList[0];
                    //RiskCalculatorLib.RiskCalculator calc = new RiskCalculator(1, 1, ref rec, true);
                    //calc.getYLLBounds();

                    ///*-----Заполнение списка людей, где каждый определяется уникальным id, полом, у него есть дозовая история и тд и тп-----*/ 
                    //List<Man> manList = new List<Man>();
                    //for (int i = 0; i < uniqueIdList.Count; i++)
                    //{
                    //    manList.Add(new Man(uniqueIdList[i], manRecordsList[i][0].getSex(), manRecordsList[i][0].getAgeAtExp(), doseHistoryList[i]));
                    //}

                    ///*-----Разбиение списка людей на половозрастные группы
                    // * Группа, № Возраст, лет
                    //        1	18-24
                    //        2	25-29
                    //        3	30-34
                    //        4	35-39
                    //        5	40-44
                    //        6	45-49
                    //        7	50-54
                    //        8	55-59
                    //        9	60-64
                    //        10	65-69
                    //        11	70+
                    //-----*/
                    //List<String> ageGroups = new List<string>();//Строки, в которых указаны возростные группы. Это ключи для дальнейшей связи через словари.
                    //ageGroups.Add("18-24");
                    //ageGroups.Add("25-29");
                    //ageGroups.Add("30-34");
                    //ageGroups.Add("35-39");
                    //ageGroups.Add("40-44");
                    //ageGroups.Add("45-49");
                    //ageGroups.Add("50-54");
                    //ageGroups.Add("55-59");
                    //ageGroups.Add("60-64");
                    //ageGroups.Add("65-69");
                    //ageGroups.Add("70+");

                    ////List<double>[] manAmountAgeInGroup = new List<double>[ageGroups.Count-1];
                    ////for (int i = 0; i < manAmountAgeInGroup.Length; i++)
                    ////    manAmountAgeInGroup[i] = new List<double>();

                    ////List<double>[] womanAmountAgeInGroup = new List<double>[ageGroups.Count-1];
                    ////for (int i = 0; i < womanAmountAgeInGroup.Length; i++)
                    ////    womanAmountAgeInGroup[i] = new List<double>();

                    ////Переписать все это в адекватную функцию
                    ////for (int i = 0; i < manList.Count; i++)
                    ////{
                    ////    if (manList[i].getSex() == 1)
                    ////    {
                    ////        if (manList[i].getAgeAtExp() >= 18 && manList[i].getAgeAtExp() <= 24)
                    ////            manAmountAgeInGroup[0].Add(manList[i].getAgeAtExp());
                    ////        if (manList[i].getAgeAtExp() >= 25 && manList[i].getAgeAtExp() <= 29)
                    ////            manAmountAgeInGroup[1].Add(manList[i].getAgeAtExp());
                    ////        if (manList[i].getAgeAtExp() >= 30 && manList[i].getAgeAtExp() <= 34)
                    ////            manAmountAgeInGroup[2].Add(manList[i].getAgeAtExp());
                    ////        if (manList[i].getAgeAtExp() >= 35 && manList[i].getAgeAtExp() <= 39)
                    ////            manAmountAgeInGroup[3].Add(manList[i].getAgeAtExp());
                    ////        if (manList[i].getAgeAtExp() >= 40 && manList[i].getAgeAtExp() <= 44)
                    ////            manAmountAgeInGroup[4].Add(manList[i].getAgeAtExp());
                    ////        if (manList[i].getAgeAtExp() >= 45 && manList[i].getAgeAtExp() <= 49)
                    ////            manAmountAgeInGroup[5].Add(manList[i].getAgeAtExp());
                    ////        if (manList[i].getAgeAtExp() >= 50 && manList[i].getAgeAtExp() <= 54)
                    ////            manAmountAgeInGroup[6].Add(manList[i].getAgeAtExp());
                    ////        if (manList[i].getAgeAtExp() >= 55 && manList[i].getAgeAtExp() <= 59)
                    ////            manAmountAgeInGroup[7].Add(manList[i].getAgeAtExp());
                    ////        if (manList[i].getAgeAtExp() >= 60 && manList[i].getAgeAtExp() <= 64)
                    ////            manAmountAgeInGroup[8].Add(manList[i].getAgeAtExp());
                    ////        if (manList[i].getAgeAtExp() >= 65 && manList[i].getAgeAtExp() <= 69)
                    ////            manAmountAgeInGroup[9].Add(manList[i].getAgeAtExp());
                    ////        if (manList[i].getAgeAtExp() >= 70)
                    ////            manAmountAgeInGroup[10].Add(manList[i].getAgeAtExp());
                    ////    }
                    ////    if (manList[i].getSex() == 2)
                    ////    {
                    ////        if (manList[i].getAgeAtExp() >= 18 && manList[i].getAgeAtExp() <= 24)
                    ////            womanAmountAgeInGroup[0].Add(manList[i].getAgeAtExp());
                    ////        if (manList[i].getAgeAtExp() >= 25 && manList[i].getAgeAtExp() <= 29)
                    ////            womanAmountAgeInGroup[1].Add(manList[i].getAgeAtExp());
                    ////        if (manList[i].getAgeAtExp() >= 30 && manList[i].getAgeAtExp() <= 34)
                    ////            womanAmountAgeInGroup[2].Add(manList[i].getAgeAtExp());
                    ////        if (manList[i].getAgeAtExp() >= 35 && manList[i].getAgeAtExp() <= 39)
                    ////            womanAmountAgeInGroup[3].Add(manList[i].getAgeAtExp());
                    ////        if (manList[i].getAgeAtExp() >= 40 && manList[i].getAgeAtExp() <= 44)
                    ////            womanAmountAgeInGroup[4].Add(manList[i].getAgeAtExp());
                    ////        if (manList[i].getAgeAtExp() >= 45 && manList[i].getAgeAtExp() <= 49)
                    ////            womanAmountAgeInGroup[5].Add(manList[i].getAgeAtExp());
                    ////        if (manList[i].getAgeAtExp() >= 50 && manList[i].getAgeAtExp() <= 54)
                    ////            womanAmountAgeInGroup[6].Add(manList[i].getAgeAtExp());
                    ////        if (manList[i].getAgeAtExp() >= 55 && manList[i].getAgeAtExp() <= 59)
                    ////            womanAmountAgeInGroup[7].Add(manList[i].getAgeAtExp());
                    ////        if (manList[i].getAgeAtExp() >= 60 && manList[i].getAgeAtExp() <= 64)
                    ////            womanAmountAgeInGroup[8].Add(manList[i].getAgeAtExp());
                    ////        if (manList[i].getAgeAtExp() >= 65 && manList[i].getAgeAtExp() <= 69)
                    ////            womanAmountAgeInGroup[9].Add(manList[i].getAgeAtExp());
                    ////        if (manList[i].getAgeAtExp() >= 70)
                    ////            womanAmountAgeInGroup[10].Add(manList[i].getAgeAtExp());
                    ////    }
                    ////}
                    ///*-----Разбиение списка людей, по половозрастным группам-----*/
                    //List<Man>[] manAgeGroups = new List<Man>[ageGroups.Count - 1];
                    //for (int i = 0; i < manAgeGroups.Length; i++)
                    //    manAgeGroups[i] = new List<Man>();

                    //List<Man>[] womanAgeGroups = new List<Man>[ageGroups.Count - 1];
                    //for (int i = 0; i < womanAgeGroups.Length; i++)
                    //    womanAgeGroups[i] = new List<Man>();

                    //for (int i = 0; i < manList.Count; i++)
                    //{
                    //    if (manList[i].getSex() == 1)
                    //    {
                    //        if (manList[i].getAgeAtExp() >= 18 && manList[i].getAgeAtExp() <= 24)
                    //            manAgeGroups[0].Add(manList[i]);
                    //        if (manList[i].getAgeAtExp() >= 25 && manList[i].getAgeAtExp() <= 29)
                    //            manAgeGroups[1].Add(manList[i]);
                    //        if (manList[i].getAgeAtExp() >= 30 && manList[i].getAgeAtExp() <= 34)
                    //            manAgeGroups[2].Add(manList[i]);
                    //        if (manList[i].getAgeAtExp() >= 35 && manList[i].getAgeAtExp() <= 39)
                    //            manAgeGroups[3].Add(manList[i]);
                    //        if (manList[i].getAgeAtExp() >= 40 && manList[i].getAgeAtExp() <= 44)
                    //            manAgeGroups[4].Add(manList[i]);
                    //        if (manList[i].getAgeAtExp() >= 45 && manList[i].getAgeAtExp() <= 49)
                    //            manAgeGroups[5].Add(manList[i]);
                    //        if (manList[i].getAgeAtExp() >= 50 && manList[i].getAgeAtExp() <= 54)
                    //            manAgeGroups[6].Add(manList[i]);
                    //        if (manList[i].getAgeAtExp() >= 55 && manList[i].getAgeAtExp() <= 59)
                    //            manAgeGroups[7].Add(manList[i]);
                    //        if (manList[i].getAgeAtExp() >= 60 && manList[i].getAgeAtExp() <= 64)
                    //            manAgeGroups[8].Add(manList[i]);
                    //        if (manList[i].getAgeAtExp() >= 65 && manList[i].getAgeAtExp() <= 69)
                    //            manAgeGroups[9].Add(manList[i]);
                    //        if (manList[i].getAgeAtExp() >= 70)
                    //            manAgeGroups[10].Add(manList[i]);
                    //    }
                    //    if (manList[i].getSex() == 2)
                    //    {
                    //        if (manList[i].getAgeAtExp() >= 18 && manList[i].getAgeAtExp() <= 24)
                    //            womanAgeGroups[0].Add(manList[i]);
                    //        if (manList[i].getAgeAtExp() >= 25 && manList[i].getAgeAtExp() <= 29)
                    //            womanAgeGroups[1].Add(manList[i]);
                    //        if (manList[i].getAgeAtExp() >= 30 && manList[i].getAgeAtExp() <= 34)
                    //            womanAgeGroups[2].Add(manList[i]);
                    //        if (manList[i].getAgeAtExp() >= 35 && manList[i].getAgeAtExp() <= 39)
                    //            womanAgeGroups[3].Add(manList[i]);
                    //        if (manList[i].getAgeAtExp() >= 40 && manList[i].getAgeAtExp() <= 44)
                    //            womanAgeGroups[4].Add(manList[i]);
                    //        if (manList[i].getAgeAtExp() >= 45 && manList[i].getAgeAtExp() <= 49)
                    //            womanAgeGroups[5].Add(manList[i]);
                    //        if (manList[i].getAgeAtExp() >= 50 && manList[i].getAgeAtExp() <= 54)
                    //            womanAgeGroups[6].Add(manList[i]);
                    //        if (manList[i].getAgeAtExp() >= 55 && manList[i].getAgeAtExp() <= 59)
                    //            womanAgeGroups[7].Add(manList[i]);
                    //        if (manList[i].getAgeAtExp() >= 60 && manList[i].getAgeAtExp() <= 64)
                    //            womanAgeGroups[8].Add(manList[i]);
                    //        if (manList[i].getAgeAtExp() >= 65 && manList[i].getAgeAtExp() <= 69)
                    //            womanAgeGroups[9].Add(manList[i]);
                    //        if (manList[i].getAgeAtExp() >= 70)
                    //            womanAgeGroups[10].Add(manList[i]);
                    //    }
                    //}

                    ///*-----Заполнение массивов списков (для М и Ж раздельно), в которых хранятся года облучения для каждой возрастной группы (для М и Ж раздельно)-----*/
                    //List<double>[] manAgesInGroup = new List<double>[ageGroups.Count - 1];
                    //for (int i = 0; i < manAgesInGroup.Length; i++)
                    //    manAgesInGroup[i] = new List<double>();

                    //for (int i = 0; i < manAgeGroups.Length; i++)
                    //{
                    //    if (manAgeGroups[i].Count > 0)
                    //        for (int k = 0; k < manAgeGroups[i].Count; k++)
                    //        {
                    //            manAgesInGroup[i].Add(manAgeGroups[i][k].getAgeAtExp());
                    //        }
                    //}

                    //List<double>[] womanAgesInGroup = new List<double>[ageGroups.Count - 1];
                    //for (int i = 0; i < womanAgesInGroup.Length; i++)
                    //    womanAgesInGroup[i] = new List<double>();

                    //for (int i = 0; i < womanAgeGroups.Length; i++)
                    //{
                    //    if (womanAgeGroups[i].Count > 0)
                    //        for (int k = 0; k < womanAgeGroups[i].Count; k++)
                    //        {
                    //            womanAgesInGroup[i].Add(womanAgeGroups[i][k].getAgeAtExp());
                    //        }
                    //}

                    ///*-----Заполнение словарей (для М и Ж [внутреннее и внешнее] раздельно), в которых по ключам вида "18-24" хранятся LAR[внутренний и внешний] группы (для М и Ж раздельно)-----*/
                    //Dictionary<String, double> manIntLarInGroup = new Dictionary<String, double>();
                    //Dictionary<String, double> manExtLarInGroup = new Dictionary<String, double>();
                    //Dictionary<String, double> womanIntLarInGroup = new Dictionary<String, double>();
                    //Dictionary<String, double> womanExtLarInGroup = new Dictionary<String, double>();
                    ///*  Здесь должны быть также выделены словари с ключами такого же типа,
                    // *  но не для LAR, а для DET.
                    // *  Заполнение осуществить в этих же циклах.
                    // */
                    //for (int i = 0; i < manAgesInGroup.Length; i++)
                    //{
                    //    if (manAgesInGroup[i].Count > 0)
                    //    {
                    //        manIntLarInGroup.Add(ageGroups[i], getManIntLar(manAgesInGroup[i].Average()));
                    //        manExtLarInGroup.Add(ageGroups[i], getManExtLar(manAgesInGroup[i].Average()));
                    //    }
                    //}
                    //for (int i = 0; i < womanAgesInGroup.Length; i++)
                    //{
                    //    if (womanAgesInGroup[i].Count > 0)
                    //    {
                    //        womanIntLarInGroup.Add(ageGroups[i], getWomanIntLar(womanAgesInGroup[i].Average()));
                    //        womanExtLarInGroup.Add(ageGroups[i], getWomanExtLar(womanAgesInGroup[i].Average()));
                    //    }
                    //}

                    ///*-----Заполнение коллекций, в которых будут храниться дозы облучения по годам, чтобы вычислять среднюю дозу облучения за 5 лет в группе-----*/
                    ///*  Может быть, следует это сделать аналогично тому,
                    // *  как это было сделано при заполнении массивов списков,
                    // *  в которых хранятся года облучения:
                    // *  manAgesInGroup и womanAgesInGroup.
                    // */
                    //List<double> blabla = new List<double>();
                    //for (int i = 0; i < doseHistoryList.Count; i++)
                    //    for (int k = 0; k < doseHistoryList[i].Length; k++){
                    //        if (doseHistoryList[i][k].AgeAtExposure >= 18 && doseHistoryList[i][k].AgeAtExposure <= 24)
                    //        {
                    //            blabla.Add(doseHistoryList[i][k].AgeAtExposure);
                    //        }

                    //    }

            testTextBox.Text = manDict[manMaxAge][0].getId().ToString();
            resultTextBox.Text = womanDict[womanMaxAge][0].getId().ToString();        
        }

        private void testTextBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void resultTextBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void testLabel_Click(object sender, EventArgs e)
        {

        }

        private void resultLabel_Click(object sender, EventArgs e)
        {

        }
    }
}

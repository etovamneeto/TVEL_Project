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
using System.Diagnostics;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
//Хреновины имеют значение, когда сравниваешь строки и строки, т.е. когда тебе надо ввести в текстовую ячейку текстовый параметр; В другие, где инты или флоаты, туда не нужны они
//"UPDATE [names] SET [userName]='" + nameBox.Text + "', [age]=" + Convert.ToInt32(ageBox.Text), но + " WHERE [id]=" + Convert.ToInt32(table.Rows[0]["id"]), connection);

namespace TVELtest
{
    public partial class Form1 : Form
    {
        /*-----Список глобальных переменных-----*/
        /*-----Переменная, отвечающая за путь к exe-файлу-----*/
        String exePath = "";
        /*-----Переменная, отвечающая путь к папке с выводами-----*/
        String outPath = "";
        /*-----Переменная, отвечающая за путь к папке с рейтами-----*/
        String libPath = "";
        /*-----Переменная-буффер для задания имен вложенных папок внутри папок с выводами-----*/
        String bufferPath = "";
        /*-----Переменная для задания имени файла-----*/
        String saveAs = "";
        /*-----Переменная для хранения названия предприятия-----*/
        String shopName = "";
        /*-----Переменная для замера времени работы приложения-----*/
        Stopwatch stopWatch = new Stopwatch();
        /*-----Список, в котором хранятся строковые параметры, инентифицирующие возрастные группы-----*/
        List<String> ageGroups = null;
        /*-----Список, в котором хранятся нижние границы возрастов для возрастных групп-----*/
        List<int> ageLowerBound = null;
        /*-----Список, в котором хранятся верхние границы возрастов для возрастных групп-----*/
        List<int> ageUpperBound = null;
        /*-----Список объектов из базы Final; достаем все необходимое для расчетов-----*/
        List<dbObject> dbFinalRecords = null;
        /*-----Список объектов из базы Dose; достаем все необходимое для расчетов-----*/
        List<dbObject> dbDoseRecords = null;
        /*-----Строка подключения к выбранной базе данных-----*/
        String connectionString = "";
        /*-----Переменные, отвечающие за пол-----*/
        byte sexMale = 0;
        byte sexFemale = 0;
        /*-----Определение пути до базы данных-----*/
        String dbPath = "";
        String request = "";
        /*-----Таблицы данных для вывода на форму-----*/
        DataTable manOrpoTable = null;
        DataTable womanOrpoTable = null;
        DataTable manIbpoTable = null;
        DataTable womanIbpoTable = null;

        /*-----Массивы списков для мужчин и для женщин, в каждом из которых хранятся дозы для соответствующей половозрастной группы-----*/
        List<double>[] manSadExtArray = null;//SAD - SexAgeDose
        List<double>[] manSadIntArray = null;
        List<double>[] womanSadExtArray = null;
        List<double>[] womanSadIntArray = null;

        /*-----Массив списоков, через которые будут вычисляться средние возраста половозрастных групп-----*/
        List<int>[] manExtYearsArray = null;
        List<int>[] manIntYearsArray = null;
        List<int>[] manYearsArray = null;
        List<int>[] womanExtYearsArray = null;
        List<int>[] womanIntYearsArray = null;    
        List<int>[] womanYearsArray = null;

        /*-----Создание массивов, храящих LAR п/в групп-----*/
        List<double>[] manExtLarArray = null;
        List<double>[] manIntLarArray = null;
        List<double>[] womanExtLarArray = null;
        List<double>[] womanIntLarArray = null;

        /*-----Массивы, хранящие ОРПО для половозрастных групп-----*/
        double[] manExtOrpo = null;
        double[] manIntOrpo = null;
        double[] manSumOrpo = null;

        double[] womanExtOrpo = null;
        double[] womanIntOrpo = null;
        double[] womanSumOrpo = null;

        double[] manExtOrpo95 = null;
        double[] manIntOrpo95 = null;
        double[] manSumOrpo95 = null;

        double[] womanExtOrpo95 = null;
        double[] womanIntOrpo95 = null;
        double[] womanSumOrpo95 = null;

        /*-----Списки для вычисления взвешенных величин ОРПО-----*/
        List<double> manWeightedExtOrpo = null;
        List<double> manWeightedIntOrpo = null;
        List<double> womanWeightedExtOrpo = null;
        List<double> womanWeightedIntOrpo = null;

        List<double> manWeightedExtOrpo95 = null;
        List<double> manWeightedIntOrpo95 = null;
        List<double> womanWeightedExtOrpo95 = null;
        List<double> womanWeightedIntOrpo95 = null;

        bool orpoButtonAverAge = false;
        bool orpoButtonAverLar = false;
        bool shopTrigger = false;



        /*-----Описание класса Объект, представляющий собой строку таблицы с параметрами: id, пол, доза суммарная, доза внутренняя, возраст при облучении-----*/
        public class dbObject
        {
            private int id = 0;
            private short ageAtExp = 0;
            private double dose = 0;
            private double doseInt = 0;
            private byte sex = 0;
            private int year = 0;

            public dbObject() { }

            public dbObject(int id, byte sex, int year, short ageAtExp, double dose, double doseInt)
            {
                this.id = id;
                this.sex = sex;
                this.year = year;
                this.ageAtExp = ageAtExp;
                this.dose = dose;
                this.doseInt = doseInt;
            }

            public dbObject(int id, int year, double dose, double doseInt)
            {
                this.id = id;
                this.year = year;
                this.dose = dose;
                this.doseInt = doseInt;
            }

            public void setId(int id) { this.id = id; }
            public void setAgeAtExp(short ageAtExp) { this.ageAtExp = ageAtExp; }
            public void setYear(int year) { this.year = year; }
            public void setDose(double dose) { this.dose = dose; }
            public void setDoseInt(double doseInt) { this.doseInt = doseInt; }
            public void setSex(byte sex) { this.sex = sex; }

            public int getId() { return this.id; }
            public short getAgeAtExp() { return this.ageAtExp; }
            public int getYear() { return this.year; }
            public double getDose() { return this.dose; }
            public double getDoseInt() { return this.doseInt; }
            public byte getSex() { return this.sex; }
        }

        /*-----Функции для расчета LAR, необходимых для расчета ОРПО-----*/
        public double getManExtLar(double meanAge)
        {
            double lar = 0;
            double secondPowerElement = 4 * Math.Pow(10, -6) * Math.Pow(meanAge, 2);
            double firstPowerElement = -11 * Math.Pow(10, -4)  * meanAge;
            double constant = 6.63 * Math.Pow(10, -2);
            return lar = secondPowerElement + firstPowerElement + constant;
        }

        public double getWomanExtLar(double meanAge)
        {
            double lar = 0;
            double secondPowerElement = -1 * Math.Pow(10, -6) * Math.Pow(meanAge, 2);
            double firstPowerElement = -9 * Math.Pow(10, -4) * meanAge;
            double constant = 7.74 * Math.Pow(10, -2);
            return lar = secondPowerElement + firstPowerElement + constant;
        }

        public double getManIntLar(double meanAge)
        {
            double lar = 0;
            double secondPowerElement = -3 * Math.Pow(10, -5) * Math.Pow(meanAge, 2);
            double firstPowerElement = 22 * Math.Pow(10, -4) * meanAge;
            double constant = 85 * Math.Pow(10, -4);
            return lar = secondPowerElement + firstPowerElement + constant;
        }

        public double getWomanIntLar(double meanAge)
        {
            double lar = 0;
            double secondPowerElement = -3 * Math.Pow(10, -5) * Math.Pow(meanAge, 2);
            double firstPowerElement = 24 * Math.Pow(10, -4) * meanAge;
            double constant = 4.39 * Math.Pow(10, -2);
            return lar = secondPowerElement + firstPowerElement + constant;
        }

        /*-----Функции для расчета Det, необходимых для расчета ОРПО-----*/
        public double getManExtDet(double meanAge)
        {
            double det = 0;
            double secondPowerElement = (4 * Math.Pow(10, -6)) * (Math.Pow(meanAge, 2));
            double firstPowerElement = (-11 * Math.Pow(10, -4)) * meanAge;
            double constant = 6.63 * Math.Pow(10, -2);
            return det = secondPowerElement + firstPowerElement + constant;
        }

        public double getWomanExtDet(double meanAge)
        {
            double det = 0;
            double secondPowerElement = (-1 * Math.Pow(10, -6)) * (Math.Pow(meanAge, 2));
            double firstPowerElement = (-9 * Math.Pow(10, -4)) * meanAge;
            double constant = 7.74 * Math.Pow(10, -2);
            return det = secondPowerElement + firstPowerElement + constant;
        }

        public double getManIntDet(double meanAge)
        {
            double det = 0;
            double secondPowerElement = (-3 * Math.Pow(10, -5)) * (Math.Pow(meanAge, 2));
            double firstPowerElement = (26 * Math.Pow(10, -4)) * meanAge;
            double constant = -1.53 * Math.Pow(10, -2);
            return det = secondPowerElement + firstPowerElement + constant;
        }
       
        public double getWomanIntDet(double meanAge)
        {
            double det = 0;
            double secondPowerElement = (-4 * Math.Pow(10, -5)) * (Math.Pow(meanAge, 2));
            double firstPowerElement = (34 * Math.Pow(10, -4)) * meanAge;
            double constant = 19 * Math.Pow(10, -4);
            return det = secondPowerElement + firstPowerElement + constant;
        }

        /*-----Функции для расчета ОРПО-----*/
        public double getOrpo(double lar, double averageDose)
        {
            double orpo = 0;
            orpo = lar * averageDose;
            return orpo;
        }

        public double getOrpo_95(double lar, double averageDose, double deviation)
        {
            double orpo = 0;
            orpo = lar * (averageDose + 1.96 * deviation);
            return orpo;
        }

        public double getOrpo_95(double lar, double dose95)
        {
            double orpo = 0;
            orpo = lar * dose95;
            return orpo;
        }

        public double getDeviation(List<double> list)
        {
            double deviation = 0;
            double[] buffer = new double[list.Count];
            for (int i = 0; i < list.Count; i++)
            {
                buffer[i] = Math.Pow((list[i] - list.Average()), 2);
                deviation += buffer[i];
            }
            deviation = Math.Sqrt(deviation / buffer.Length);
            return deviation;
        }

        /*-----Функция для расчета ИБПО-----*/
        public double getIbpo(List<double> groupedLar, double orpo)
        {
            //womanIntIbpo[i] = 100 / (1 + womanIntOrpo[i] / (2 * Math.Pow(10, -4) * (1 - ((womanLarIntArray[i].Sum() / womanLarIntArray[i].Count) / (4.1 * Math.Pow(10, -2))))));
            double r = groupedLar.Average();
            double q = 1 - r / (4.1 * Math.Pow(10, -2));
            double denominator = 1 + orpo / (2 * Math.Pow(10, -4) * q);
            return 100 / denominator;
        }

        /*-----Описание форм инициализации и инициализация библиотеки с рейтами 2012 года-----*/
        public Form1(String title)
        {
            InitializeComponent();
            this.Text = title;
        }

        private void Form1_Load(object sender, EventArgs e)
        {           
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;

            larRB.Checked = true;
            detRB.Checked = false;

            aMethodRB.Checked = true;
            bMethodRB.Checked = false;

            exePath = Path.GetDirectoryName(Application.ExecutablePath);
            libPath = exePath + "\\DataRus2012";
            outPath = exePath + "\\Табличные выводы";
            Directory.CreateDirectory(outPath);

            RiskCalculatorLib.RiskCalculator.FillData(ref libPath);
        }

        private void openFileButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                dbPath = ofd.FileName;
                connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + dbPath;
            }

            /*-----Список, в котором хранятся строковые параметры, инентифицирующие возрастные группы-----*/
            ageGroups = new List<string>();//Строки, в которых указаны возростные группы. Это ключи для дальнейшей связи через словари.
            ageGroups.Add("18-24");
            ageGroups.Add("25-29");
            ageGroups.Add("30-34");
            ageGroups.Add("35-39");
            ageGroups.Add("40-44");
            ageGroups.Add("45-49");
            ageGroups.Add("50-54");
            ageGroups.Add("55-59");
            ageGroups.Add("60-64");
            ageGroups.Add("65-69");
            ageGroups.Add("70+");

            /*-----Список, в котором хранятся нижние границы возрастов для возрастных групп-----*/
            ageLowerBound = new List<int>();
            ageLowerBound.Add(18);
            ageLowerBound.Add(25);
            ageLowerBound.Add(30);
            ageLowerBound.Add(35);
            ageLowerBound.Add(40);
            ageLowerBound.Add(45);
            ageLowerBound.Add(50);
            ageLowerBound.Add(55);
            ageLowerBound.Add(60);
            ageLowerBound.Add(65);
            ageLowerBound.Add(70);

            /*-----Список, в котором хранятся верхние границы возрастов для возрастных групп-----*/
            ageUpperBound = new List<int>();
            ageUpperBound.Add(24);
            ageUpperBound.Add(29);
            ageUpperBound.Add(34);
            ageUpperBound.Add(39);
            ageUpperBound.Add(44);
            ageUpperBound.Add(49);
            ageUpperBound.Add(54);
            ageUpperBound.Add(59);
            ageUpperBound.Add(64);
            ageUpperBound.Add(69);
            ageUpperBound.Add(100);
        }

        private void getOrpoAverAgeButton_Click(object sender, EventArgs e)
        {
            /*-----Инициализация всяких входных параметров, подключения к БД, парсинга в таблицу нужных столбцов-----*/
            OleDbConnection connection = new OleDbConnection(connectionString);
            try
            {
                orpoButtonAverAge = true;
                orpoButtonAverLar = false;
                shopTrigger = true;

                if (shopComboBox.SelectedItem == "СХК")
                    shopName = "r1";
                else if (shopComboBox.SelectedItem == "АЭХК")
                    shopName = "r2";
                else if (shopComboBox.SelectedItem == "МСЗ")
                    shopName = "r3";
                else if (shopComboBox.SelectedItem == "УЭХК")
                    shopName = "r4";
                else if (shopComboBox.SelectedItem == "ПО ЭХЗ")
                    shopName = "r5";
                else if (shopComboBox.SelectedItem == "ЧМЗ")
                    shopName = "r6";

                request = "SELECT [ID], [Dose], [Year], [DoseInt], [Gender], [AgeAtExp] FROM [Final] WHERE [Shop]='" + shopName + "'";

                if (shopComboBox.SelectedItem == "ВСЕ ПРЕДПРИЯТИЯ")
                    request = "SELECT [ID], [Dose], [Year], [DoseInt], [Gender], [AgeAtExp] FROM [Final]";

                connection.Open();

                try
                {
                    /*-----Выбор нужных столбцов из нужной таблицы в таблицу table-----*/
                    OleDbDataAdapter adapter = new OleDbDataAdapter(request, connectionString);
                    DataSet dataSet = new DataSet();
                    adapter.Fill(dataSet, "Final");
                    DataTable table = dataSet.Tables[0];

                    /*-----Список объектов; достаем все необходимое для расчетов: id, dose, doseInt, ageAtExp, gender-----*/
                    dbFinalRecords = new List<dbObject>();
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        if(aMethodRB.Checked)
                            dbFinalRecords.Add(new dbObject(Convert.ToInt32(table.Rows[i]["id"]), Convert.ToByte(table.Rows[i]["gender"]), Convert.ToInt32(table.Rows[i]["year"]), Convert.ToInt16(table.Rows[i]["ageatexp"]), Convert.ToDouble(table.Rows[i]["dose"]) / 1000, Convert.ToDouble(table.Rows[i]["doseint"]) / 1000));
                        if(bMethodRB.Checked)
                            dbFinalRecords.Add(new dbObject(Convert.ToInt32(table.Rows[i]["id"]), Convert.ToByte(table.Rows[i]["gender"]), Convert.ToInt32(table.Rows[i]["year"]), Convert.ToInt16(table.Rows[i]["ageatexp"]), Convert.ToDouble(table.Rows[i]["dose"]), Convert.ToDouble(table.Rows[i]["doseint"])));
                    }

                    /*-----Список, в котором хранится пол-----*/
                    List<byte> dbSex = new List<byte>();
                    for (int i = 0; i < dbFinalRecords.Count; i++)
                        dbSex.Add(dbFinalRecords[i].getSex());

                    /*-----Определение пола; Меньшая цифра пола - М, большая - Ж-----*/
                    sexMale = dbSex.Min();
                    sexFemale = dbSex.Max();

                    /*-----Счетчики, определяющие количество мужских и женских записей-----*/
                    double dbMan = 0;
                    for (int i = 0; i < dbFinalRecords.Count; i++)
                        if (dbFinalRecords[i].getSex() == sexMale)
                            dbMan++;

                    double dbWoman = 0;
                    for (int i = 0; i < dbFinalRecords.Count; i++)
                        if (dbFinalRecords[i].getSex() == sexFemale)
                            dbWoman++;

                    /*-----Массивы списков для мужчин и для женщин, в каждом из которых хранятся дозы (внешние и внутренние) для соответствующий половозрастной группы-----*/
                    manSadExtArray = new List<double>[ageGroups.Count];
                    manSadIntArray = new List<double>[ageGroups.Count];
                    womanSadExtArray = new List<double>[ageGroups.Count];
                    womanSadIntArray = new List<double>[ageGroups.Count];

                    /*-----Массивы списоков, через которые будут вычесляться средние возраста половозрастных групп-----*/
                    if (aMethodRB.Checked)
                    {
                        manExtYearsArray = new List<int>[ageGroups.Count];
                        manIntYearsArray = new List<int>[ageGroups.Count];
                        womanExtYearsArray = new List<int>[ageGroups.Count];
                        womanIntYearsArray = new List<int>[ageGroups.Count];
                    }
                    if (bMethodRB.Checked)
                    {
                        manYearsArray = new List<int>[ageGroups.Count];
                        womanYearsArray = new List<int>[ageGroups.Count];
                    }

                    for (int i = 0; i < ageGroups.Count; i++)
                    {
                        manSadExtArray[i] = new List<double>();
                        manSadIntArray[i] = new List<double>();
                        womanSadExtArray[i] = new List<double>();
                        womanSadIntArray[i] = new List<double>();

                        if (aMethodRB.Checked)
                        {
                            manExtYearsArray[i] = new List<int>();
                            womanExtYearsArray[i] = new List<int>();
                            manIntYearsArray[i] = new List<int>();
                            womanIntYearsArray[i] = new List<int>();
                        }
                        if (bMethodRB.Checked)
                        {
                            manYearsArray[i] = new List<int>();
                            womanYearsArray[i] = new List<int>();
                        }
                    }

                    if (aMethodRB.Checked)
                    {
                        for (int i = 0; i < ageGroups.Count; i++)
                            for (int k = 0; k < dbFinalRecords.Count; k++)
                            {
                                if (dbFinalRecords[k].getSex() == sexMale)
                                    if (dbFinalRecords[k].getAgeAtExp() >= ageLowerBound[i] && dbFinalRecords[k].getAgeAtExp() <= ageUpperBound[i])
                                    {
                                        manSadExtArray[i].Add(dbFinalRecords[k].getDose() - dbFinalRecords[k].getDoseInt());
                                        manExtYearsArray[i].Add(dbFinalRecords[k].getAgeAtExp());
                                        if (dbFinalRecords[k].getDoseInt() > 0)
                                        {
                                            manSadIntArray[i].Add(dbFinalRecords[k].getDoseInt());
                                            manIntYearsArray[i].Add(dbFinalRecords[k].getAgeAtExp());
                                        }
                                    }
                                if (dbFinalRecords[k].getSex() == sexFemale)
                                    if (dbFinalRecords[k].getAgeAtExp() >= ageLowerBound[i] && dbFinalRecords[k].getAgeAtExp() <= ageUpperBound[i])
                                    {
                                        womanSadExtArray[i].Add(dbFinalRecords[k].getDose() - dbFinalRecords[k].getDoseInt());
                                        womanExtYearsArray[i].Add(dbFinalRecords[k].getAgeAtExp());
                                        if (dbFinalRecords[k].getDoseInt() > 0)
                                        {
                                            womanSadIntArray[i].Add(dbFinalRecords[k].getDoseInt());
                                            womanIntYearsArray[i].Add(dbFinalRecords[k].getAgeAtExp());
                                        }
                                    }
                            }
                    }


                    if (bMethodRB.Checked)
                    {
                        /*-----Заполнение массива списков доз-----*/
                        for (int i = 0; i < ageGroups.Count; i++)
                            for (int k = 0; k < dbFinalRecords.Count; k++)
                            {
                                if (dbFinalRecords[k].getSex() == sexMale)
                                    if (dbFinalRecords[k].getAgeAtExp() >= ageLowerBound[i] && dbFinalRecords[k].getAgeAtExp() <= ageUpperBound[i])
                                    {
                                        manSadExtArray[i].Add(dbFinalRecords[k].getDose() - dbFinalRecords[k].getDoseInt());
                                        manSadIntArray[i].Add(dbFinalRecords[k].getDoseInt());
                                        manYearsArray[i].Add(dbFinalRecords[k].getAgeAtExp());
                                    }
                                if (dbFinalRecords[k].getSex() == sexFemale)
                                    if (dbFinalRecords[k].getAgeAtExp() >= ageLowerBound[i] && dbFinalRecords[k].getAgeAtExp() <= ageUpperBound[i])
                                    {
                                        womanSadExtArray[i].Add(dbFinalRecords[k].getDose() - dbFinalRecords[k].getDoseInt());
                                        womanSadIntArray[i].Add(dbFinalRecords[k].getDoseInt());
                                        womanYearsArray[i].Add(dbFinalRecords[k].getAgeAtExp());
                                    }
                            }

                        //Задание весовых коэффициентов для тканей (в нашем случае учитывается только влияние на лёгкие)
                        double wLung = 0.12;

                        /*-----Создание дозовых историй-----*/
                        List<RiskCalculator.DoseHistoryRecord[]>[] manDoseHistoryList = new List<RiskCalculator.DoseHistoryRecord[]>[ageGroups.Count];
                        List<RiskCalculator.DoseHistoryRecord[]>[] womanDoseHistoryList = new List<RiskCalculator.DoseHistoryRecord[]>[ageGroups.Count];
                        for (int i = 0; i < ageGroups.Count; i++)
                        {
                            manDoseHistoryList[i] = new List<RiskCalculator.DoseHistoryRecord[]>();
                            womanDoseHistoryList[i] = new List<RiskCalculator.DoseHistoryRecord[]>();
                        }
                        for (int i = 0; i < ageGroups.Count; i++)
                            for (int k = 0; k < manSadExtArray[i].Count; k++)
                            {
                                manDoseHistoryList[i].Add(new RiskCalculator.DoseHistoryRecord[1]);
                            }
                        for (int i = 0; i < ageGroups.Count; i++)
                            for (int k = 0; k < manSadExtArray[i].Count; k++)
                            {
                                womanDoseHistoryList[i].Add(new RiskCalculator.DoseHistoryRecord[1]);
                            }

                        /*-----Создание массивов, храящих LAR п/в групп-----*/
                        manExtLarArray = new List<double>[ageGroups.Count];
                        manIntLarArray = new List<double>[ageGroups.Count];
                        for (int i = 0; i < ageGroups.Count; i++)
                        {
                            manExtLarArray[i] = new List<double>();
                            manIntLarArray[i] = new List<double>();
                        }

                        womanExtLarArray = new List<double>[ageGroups.Count];
                        womanIntLarArray = new List<double>[ageGroups.Count];
                        for (int i = 0; i < ageGroups.Count; i++)
                        {
                            womanExtLarArray[i] = new List<double>();
                            womanIntLarArray[i] = new List<double>();
                        }

                        /*-----Заполнение ДИ-----*/
                        RiskCalculator.DoseHistoryRecord[] record = null;
                        RiskCalculatorLib.RiskCalculator calculator = null;
                        bool isIncidence = false;
                        for (int i = 0; i < ageGroups.Count; i++)
                            for (int k = 0; k < manSadExtArray[i].Count; k++)
                            {
                                manDoseHistoryList[i][k][0] = new RiskCalculator.DoseHistoryRecord();
                                manDoseHistoryList[i][k][0].AgeAtExposure = (short)manYearsArray[i][k];
                                manDoseHistoryList[i][k][0].AllSolidDoseInmGy = manSadExtArray[i][k];
                                manDoseHistoryList[i][k][0].LeukaemiaDoseInmGy = manSadExtArray[i][k];
                                manDoseHistoryList[i][k][0].LungDoseInmGy = manSadIntArray[i][k] / wLung;

                                record = manDoseHistoryList[i][k];
                                calculator = new RiskCalculatorLib.RiskCalculator(RiskCalculator.SEX_MALE, manDoseHistoryList[i][k][0].AgeAtExposure, ref record, true);
                                if (larRB.Checked)
                                {
                                    manExtLarArray[i].Add(calculator.getLAR(false, isIncidence).AllCancers);
                                    if (manDoseHistoryList[i][k][0].LungDoseInmGy > 0)
                                        manIntLarArray[i].Add(calculator.getLAR(false, isIncidence).Lung);
                                }
                                if (detRB.Checked)
                                {
                                    calculator.createEARSamples(0, ref isIncidence);
                                    manExtLarArray[i].Add(calculator.getDetriment().Value.AllCancers);
                                    if (manDoseHistoryList[i][k][0].LungDoseInmGy > 0)
                                        manIntLarArray[i].Add(calculator.getDetriment().Value.Lung);
                                }
                            }
                        for (int i = 0; i < ageGroups.Count; i++)
                            for (int k = 0; k < womanSadExtArray[i].Count; k++)
                            {
                                womanDoseHistoryList[i][k][0] = new RiskCalculator.DoseHistoryRecord();
                                womanDoseHistoryList[i][k][0].AgeAtExposure = (short)womanYearsArray[i][k];
                                womanDoseHistoryList[i][k][0].AllSolidDoseInmGy = womanSadExtArray[i][k];
                                womanDoseHistoryList[i][k][0].LeukaemiaDoseInmGy = womanSadExtArray[i][k];
                                womanDoseHistoryList[i][k][0].LungDoseInmGy = womanSadIntArray[i][k] / wLung;

                                record = womanDoseHistoryList[i][k];
                                calculator = new RiskCalculatorLib.RiskCalculator(RiskCalculator.SEX_FEMALE, womanDoseHistoryList[i][k][0].AgeAtExposure, ref record, true);
                                if (larRB.Checked)
                                {
                                    womanExtLarArray[i].Add(calculator.getLAR(false, isIncidence).AllCancers);
                                    if (womanDoseHistoryList[i][k][0].LungDoseInmGy != 0)
                                        womanIntLarArray[i].Add(calculator.getLAR(false, isIncidence).Lung);
                                }
                                if (detRB.Checked)
                                {
                                    calculator.createEARSamples(0, ref isIncidence);
                                    womanExtLarArray[i].Add(calculator.getDetriment().Value.AllCancers);
                                    if (womanDoseHistoryList[i][k][0].LungDoseInmGy != 0)
                                        womanIntLarArray[i].Add(calculator.getDetriment().Value.Lung);
                                }
                            }
                    }

///*---------------------------------------------------------------------------------------------------------------------------------------*/
                    /*-----Инициализация массивов, хранящих ОРПО для половозрастных групп-----*/
                    manExtOrpo = new double[ageGroups.Count];
                    manIntOrpo = new double[ageGroups.Count];
                    manSumOrpo = new double[ageGroups.Count];

                    womanExtOrpo = new double[ageGroups.Count];
                    womanIntOrpo = new double[ageGroups.Count];
                    womanSumOrpo = new double[ageGroups.Count];

                    manExtOrpo95 = new double[ageGroups.Count];
                    manIntOrpo95 = new double[ageGroups.Count];
                    manSumOrpo95 = new double[ageGroups.Count];

                    womanExtOrpo95 = new double[ageGroups.Count];
                    womanIntOrpo95 = new double[ageGroups.Count];
                    womanSumOrpo95 = new double[ageGroups.Count];

                    for (int i = 0; i < ageGroups.Count; i++)
                    {
                        if (aMethodRB.Checked)
                        {
                            if (manSadExtArray[i].Count > 0)
                            {
                                if (larRB.Checked)
                                {
                                    manExtOrpo[i] = getOrpo(getManExtLar(manExtYearsArray[i].Average()), manSadExtArray[i].Average());
                                    manSadExtArray[i].Sort();
                                    if (manSadExtArray[i].Count == 1)
                                        manExtOrpo95[i] = getOrpo_95(getManExtLar(manExtYearsArray[i].Average()), manSadExtArray[i][0]);
                                    if (manSadExtArray[i].Count > 1)
                                        manExtOrpo95[i] = getOrpo_95(getManExtLar(manExtYearsArray[i].Average()), manSadExtArray[i][manSadExtArray[i].Count * 95 / 100 - 1]);
                                }
                                if (detRB.Checked)
                                {
                                    manExtOrpo[i] = getOrpo(getManExtDet(manExtYearsArray[i].Average()), manSadExtArray[i].Average());
                                    manSadExtArray[i].Sort();
                                    if (manSadExtArray[i].Count == 1)
                                        manExtOrpo95[i] = getOrpo_95(getManExtDet(manExtYearsArray[i].Average()), manSadExtArray[i][0]);
                                    if (manSadExtArray[i].Count > 1)
                                        manExtOrpo95[i] = getOrpo_95(getManExtDet(manExtYearsArray[i].Average()), manSadExtArray[i][manSadExtArray[i].Count * 95 / 100 - 1]);
                                }
                            }

                            if (manSadIntArray[i].Count > 0)
                            {
                                if (larRB.Checked)
                                {
                                    manIntOrpo[i] = getOrpo(getManIntLar(manIntYearsArray[i].Average()), manSadIntArray[i].Average());
                                    manSadIntArray[i].Sort();
                                    if (manSadIntArray[i].Count == 1)
                                        manIntOrpo95[i] = getOrpo_95(getManIntLar(manIntYearsArray[i].Average()), manSadIntArray[i][0]);
                                    if (manSadIntArray[i].Count > 1)
                                        manIntOrpo95[i] = getOrpo_95(getManIntLar(manIntYearsArray[i].Average()), manSadIntArray[i][manSadIntArray[i].Count * 95 / 100 - 1]);
                                }
                                if (detRB.Checked)
                                {
                                    manIntOrpo[i] = getOrpo(getManIntDet(manIntYearsArray[i].Average()), manSadIntArray[i].Average());
                                    manSadIntArray[i].Sort();
                                    if (manSadIntArray[i].Count == 1)
                                        manIntOrpo95[i] = getOrpo_95(getManIntDet(manIntYearsArray[i].Average()), manSadIntArray[i][0]);
                                    if (manSadIntArray[i].Count > 1)
                                        manIntOrpo95[i] = getOrpo_95(getManIntDet(manIntYearsArray[i].Average()), manSadIntArray[i][manSadIntArray[i].Count * 95 / 100 - 1]);
                                }
                            }

                            manSumOrpo[i] = manExtOrpo[i] + manIntOrpo[i];
                            manSumOrpo95[i] = manExtOrpo95[i] + manIntOrpo95[i];

                            if (womanSadExtArray[i].Count > 0)
                            {
                                if (larRB.Checked)
                                {
                                    womanExtOrpo[i] = getOrpo(getWomanExtLar(womanExtYearsArray[i].Average()), womanSadExtArray[i].Average());
                                    womanSadExtArray[i].Sort();
                                    if (womanSadExtArray[i].Count == 1)
                                        womanExtOrpo95[i] = getOrpo_95(getWomanExtLar(womanExtYearsArray[i].Average()), womanSadExtArray[i][0]);
                                    if (womanSadExtArray[i].Count > 1)
                                        womanExtOrpo95[i] = getOrpo_95(getWomanExtLar(womanExtYearsArray[i].Average()), womanSadExtArray[i][womanSadExtArray[i].Count * 95 / 100 - 1]);
                                }
                                if (detRB.Checked)
                                {
                                    womanExtOrpo[i] = getOrpo(getWomanExtDet(womanExtYearsArray[i].Average()), womanSadExtArray[i].Average());
                                    womanSadExtArray[i].Sort();
                                    if (womanSadExtArray[i].Count == 1)
                                        womanExtOrpo95[i] = getOrpo_95(getWomanExtDet(womanExtYearsArray[i].Average()), womanSadExtArray[i][0]);
                                    if (womanSadExtArray[i].Count > 1)
                                        womanExtOrpo95[i] = getOrpo_95(getWomanExtDet(womanExtYearsArray[i].Average()), womanSadExtArray[i][womanSadExtArray[i].Count * 95 / 100 - 1]);
                                }
                            }

                            if (womanSadIntArray[i].Count > 0)
                            {
                                if (larRB.Checked)
                                {
                                    womanIntOrpo[i] = getOrpo(getWomanIntLar(womanIntYearsArray[i].Average()), womanSadIntArray[i].Average());
                                    womanSadIntArray[i].Sort();
                                    if (womanSadIntArray[i].Count == 1)
                                        womanIntOrpo95[i] = getOrpo_95(getWomanIntLar(womanIntYearsArray[i].Average()), womanSadIntArray[i][0]);
                                    if (womanSadIntArray[i].Count > 1)
                                        womanIntOrpo95[i] = getOrpo_95(getWomanIntLar(womanIntYearsArray[i].Average()), womanSadIntArray[i][womanSadIntArray[i].Count * 95 / 100 - 1]);
                                }
                                if (detRB.Checked)
                                {
                                    womanIntOrpo[i] = getOrpo(getWomanIntDet(womanIntYearsArray[i].Average()), womanSadIntArray[i].Average());
                                    womanSadIntArray[i].Sort();
                                    if (womanSadIntArray[i].Count == 1)
                                        womanIntOrpo95[i] = getOrpo_95(getWomanIntDet(womanIntYearsArray[i].Average()), womanSadIntArray[i][0]);
                                    if (womanSadIntArray[i].Count > 1)
                                        womanIntOrpo95[i] = getOrpo_95(getWomanIntDet(womanIntYearsArray[i].Average()), womanSadIntArray[i][womanSadIntArray[i].Count * 95 / 100 - 1]);
                                }
                            }

                            womanSumOrpo[i] = womanExtOrpo[i] + womanIntOrpo[i];
                            womanSumOrpo95[i] = womanExtOrpo95[i] + womanIntOrpo95[i];
                        }

                        if (bMethodRB.Checked)
                        {
                            if (manExtLarArray[i].Count > 0)
                            {
                                manExtOrpo[i] = manExtLarArray[i].Average();
                                manExtLarArray[i].Sort();
                                if (manExtLarArray[i].Count == 1)
                                    manExtOrpo95[i] = manExtLarArray[i][0];
                                if (manExtLarArray[i].Count > 1)
                                    manExtOrpo95[i] = manExtLarArray[i][manExtLarArray[i].Count * 95 / 100 - 1];
                            }

                            if (manIntLarArray[i].Count > 0)
                            {
                                manIntOrpo[i] = manIntLarArray[i].Average();
                                manIntLarArray[i].Sort();
                                if (manIntLarArray[i].Count == 1)
                                    manIntOrpo95[i] = manIntLarArray[i][0];
                                if (manIntLarArray[i].Count > 1)
                                    manIntOrpo95[i] = manIntLarArray[i][manIntLarArray[i].Count * 95 / 100 - 1];
                            }

                            manSumOrpo[i] = manExtOrpo[i] + manIntOrpo[i];
                            manSumOrpo95[i] = manExtOrpo95[i] + manIntOrpo95[i];

                            if (womanExtLarArray[i].Count > 0)
                            {
                                womanExtOrpo[i] = womanExtLarArray[i].Average();
                                womanExtLarArray[i].Sort();
                                if (womanExtLarArray[i].Count == 1)
                                    womanExtOrpo95[i] = womanExtLarArray[i][0];
                                if (womanExtLarArray[i].Count > 1)
                                    womanExtOrpo95[i] = womanExtLarArray[i][womanExtLarArray[i].Count * 95 / 100 - 1];
                            }

                            if (womanIntLarArray[i].Count > 0)
                            {
                                womanIntOrpo[i] = womanIntLarArray[i].Average();
                                womanIntLarArray[i].Sort();
                                if (womanIntLarArray[i].Count == 1)
                                    womanIntOrpo95[i] = womanIntLarArray[i][0];
                                if (womanIntLarArray[i].Count > 1)
                                    womanIntOrpo95[i] = womanIntLarArray[i][womanIntLarArray[i].Count * 95 / 100 - 1];
                            }

                            womanSumOrpo[i] = womanExtOrpo[i] + womanIntOrpo[i];
                            womanSumOrpo95[i] = womanExtOrpo95[i] + womanIntOrpo95[i];
                        }
                    }

                    manWeightedExtOrpo = new List<double>();
                    manWeightedIntOrpo = new List<double>();
                    womanWeightedExtOrpo = new List<double>();
                    womanWeightedIntOrpo = new List<double>();
                    manWeightedExtOrpo95 = new List<double>();
                    manWeightedIntOrpo95 = new List<double>();
                    womanWeightedExtOrpo95 = new List<double>();
                    womanWeightedIntOrpo95 = new List<double>();

                    for (int i = 0; i < ageGroups.Count; i++)
                    {
                        if (manSadExtArray[i].Count > 0)
                            manWeightedExtOrpo.Add(manExtOrpo[i] * manSadExtArray[i].Count);

                        if (manSadIntArray[i].Count > 0)
                            manWeightedIntOrpo.Add(manIntOrpo[i] * manSadIntArray[i].Count);

                        if (womanSadExtArray[i].Count > 0)
                            womanWeightedExtOrpo.Add(womanExtOrpo[i] * womanSadExtArray[i].Count);

                        if (womanSadIntArray[i].Count > 0)
                            womanWeightedIntOrpo.Add(womanIntOrpo[i] * womanSadIntArray[i].Count);

                        if (manSadExtArray[i].Count > 0)
                            manWeightedExtOrpo95.Add(manExtOrpo95[i] * manSadExtArray[i].Count);

                        if (manSadIntArray[i].Count > 0)
                            manWeightedIntOrpo95.Add(manIntOrpo95[i] * manSadIntArray[i].Count);

                        if (womanSadExtArray[i].Count > 0)
                            womanWeightedExtOrpo95.Add(womanExtOrpo95[i] * womanSadExtArray[i].Count);

                        if (womanSadIntArray[i].Count > 0)
                            womanWeightedIntOrpo95.Add(womanIntOrpo95[i] * womanSadIntArray[i].Count);

                        //if (bMethodRB.Checked)
                        //{
                        //    manWeightedExtOrpo.Add(manExtOrpo[i] * manSadExtArray[i].Count);
                        //    manWeightedIntOrpo.Add(manIntOrpo[i] * manSadIntArray[i].Count);
                        //    womanWeightedExtOrpo.Add(womanExtOrpo[i] * womanSadExtArray[i].Count);
                        //    womanWeightedIntOrpo.Add(womanIntOrpo[i] * womanSadIntArray[i].Count);

                        //    manWeightedExtOrpo95.Add(manExtOrpo95[i] * manSadExtArray[i].Count);
                        //    manWeightedIntOrpo95.Add(manIntOrpo95[i] * manSadIntArray[i].Count);
                        //    womanWeightedExtOrpo95.Add(womanExtOrpo95[i] * womanSadExtArray[i].Count);
                        //    womanWeightedIntOrpo95.Add(womanIntOrpo95[i] * womanSadIntArray[i].Count);
                        //}
                    }

                    manOrpoTable = new DataTable();
                    manOrpoTable.Columns.Add("Возрастные группы");
                    manOrpoTable.Columns.Add("Внешнее облучение");
                    manOrpoTable.Columns.Add("Внутреннее облучение");
                    manOrpoTable.Columns.Add("Сумма");
                    manOrpoTable.Columns.Add("Внешнее облучение (95%)");
                    manOrpoTable.Columns.Add("Внутреннее облучение (95%)");
                    manOrpoTable.Columns.Add("Сумма (95%)");

                    for (int i = 0; i < ageGroups.Count; i++)
                    {
                        DataRow row = manOrpoTable.NewRow();
                        row["Возрастные группы"] = ageGroups[i];
                        row["Внешнее облучение"] = Math.Round(manExtOrpo[i], 8);
                        row["Внутреннее облучение"] = Math.Round(manIntOrpo[i], 8);
                        row["Сумма"] = Math.Round(manSumOrpo[i], 8);
                        row["Внешнее облучение (95%)"] = Math.Round(manExtOrpo95[i], 8);
                        row["Внутреннее облучение (95%)"] = Math.Round(manIntOrpo95[i], 8);
                        row["Сумма (95%)"] = Math.Round(manSumOrpo95[i], 8);
                        manOrpoTable.Rows.Add(row);
                    }

                    DataRow manOrpoRow = manOrpoTable.NewRow();
                    manOrpoRow["Возрастные группы"] = "Взвешенные величины";
                    if (manWeightedExtOrpo.Sum() > 0)
                        manOrpoRow["Внешнее облучение"] = Math.Round(manWeightedExtOrpo.Sum() / dbMan, 8);
                    else
                        manOrpoRow["Внешнее облучение"] = "Облучения нет";

                    if (manWeightedIntOrpo.Sum() > 0)
                        manOrpoRow["Внутреннее облучение"] = Math.Round(manWeightedIntOrpo.Sum() / dbMan, 8);
                    else
                        manOrpoRow["Внутреннее облучение"] = "Облучения нет!";

                    manOrpoRow["Сумма"] = Math.Round((manWeightedExtOrpo.Sum() / dbMan) + (manWeightedIntOrpo.Sum() / dbMan), 8);

                    if (manWeightedExtOrpo95.Sum() > 0)
                        manOrpoRow["Внешнее облучение (95%)"] = Math.Round(manWeightedExtOrpo95.Sum() / dbMan, 8);
                    else
                        manOrpoRow["Внешнее облучение (95%)"] = "Облучения нет!";

                    if (manWeightedIntOrpo95.Sum() > 0)
                        manOrpoRow["Внутреннее облучение (95%)"] = Math.Round(manWeightedIntOrpo95.Sum() / dbMan, 8);
                    else
                        manOrpoRow["Внутреннее облучение (95%)"] = "Облучения нет!";

                    manOrpoRow["Сумма (95%)"] = Math.Round((manWeightedExtOrpo95.Sum() / dbMan) + (manWeightedIntOrpo95.Sum() / dbMan), 8);
                    manOrpoTable.Rows.Add(manOrpoRow);

                    manOrpoGridView.DataSource = manOrpoTable;
                    manOrpoGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                    manOrpoGridView.AllowUserToAddRows = false;

                    womanOrpoTable = new DataTable();
                    womanOrpoTable.Columns.Add("Возрастные группы");
                    womanOrpoTable.Columns.Add("Внешнее облучение");
                    womanOrpoTable.Columns.Add("Внутреннее облучение");
                    womanOrpoTable.Columns.Add("Сумма");
                    womanOrpoTable.Columns.Add("Внешнее облучение (95%)");
                    womanOrpoTable.Columns.Add("Внутреннее облучение (95%)");
                    womanOrpoTable.Columns.Add("Сумма (95%)");

                    for (int i = 0; i < ageGroups.Count; i++)
                    {
                        DataRow row = womanOrpoTable.NewRow();
                        row["Возрастные группы"] = ageGroups[i];
                        row["Внешнее облучение"] = Math.Round(womanExtOrpo[i], 8);
                        row["Внутреннее облучение"] = Math.Round(womanIntOrpo[i], 8);
                        row["Сумма"] = Math.Round(womanSumOrpo[i], 8);
                        row["Внешнее облучение (95%)"] = Math.Round(womanExtOrpo95[i], 8);
                        row["Внутреннее облучение (95%)"] = Math.Round(womanIntOrpo95[i], 8);
                        row["Сумма (95%)"] = Math.Round(womanSumOrpo95[i], 8);
                        womanOrpoTable.Rows.Add(row);
                    }
                    DataRow womanOrpoRow = womanOrpoTable.NewRow();
                    womanOrpoRow["Возрастные группы"] = "Взвешенные величины";
                    if (womanWeightedExtOrpo.Sum() > 0)
                        womanOrpoRow["Внешнее облучение"] = Math.Round(womanWeightedExtOrpo.Sum() / dbWoman, 8);
                    else
                        womanOrpoRow["Внешнее облучение"] = "Облучения нет";

                    if (womanWeightedIntOrpo.Sum() > 0)
                        womanOrpoRow["Внутреннее облучение"] = Math.Round(womanWeightedIntOrpo.Sum() / dbWoman, 8);
                    else
                        womanOrpoRow["Внутреннее облучение"] = "Облучения нет!";

                    womanOrpoRow["Сумма"] = Math.Round((womanWeightedExtOrpo.Sum() / dbWoman) + (womanWeightedIntOrpo.Sum() / dbWoman), 8);

                    if (womanWeightedExtOrpo95.Sum() > 0)
                        womanOrpoRow["Внешнее облучение (95%)"] = Math.Round(womanWeightedExtOrpo95.Sum() / dbWoman, 8);
                    else
                        womanOrpoRow["Внешнее облучение (95%)"] = "Облучения нет!";

                    if (womanWeightedIntOrpo95.Sum() > 0)
                        womanOrpoRow["Внутреннее облучение (95%)"] = Math.Round(womanWeightedIntOrpo95.Sum() / dbWoman, 8);
                    else
                        womanOrpoRow["Внутреннее облучение (95%)"] = "Облучения нет!";

                    womanOrpoRow["Сумма (95%)"] = Math.Round((womanWeightedExtOrpo95.Sum() / dbWoman) + (womanWeightedIntOrpo95.Sum() / dbWoman), 8);
                    womanOrpoTable.Rows.Add(womanOrpoRow);

                    womanOrpoGridView.DataSource = womanOrpoTable;
                    womanOrpoGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                    womanOrpoGridView.AllowUserToAddRows = false;
/*------------------------------------------------Вывод в эксель-файл-----------------------------------------------------------------------------*/
                    //if (aMethodRB.Checked)
                    //    textBox1.Text = "ок А";
                    //if (bMethodRB.Checked)
                    //    textBox1.Text = "ок Б";
                }
                catch
                {
                    MessageBox.Show("Не выбрано предприятие!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
                    Application.DoEvents();
                }
            }

            catch/*(OleDbException ex)*/
            {
                MessageBox.Show("Не выбрана база данных!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
                Application.DoEvents();
            }
        }

        private void getOrpoAverLarButton_Click(object sender, EventArgs e)
        {
            /*-----Инициализация всяких входных параметров, подключения к БД, парсинга в таблицу нужных столбцов-----*/
            OleDbConnection connection = new OleDbConnection(connectionString);
            try
            {
                shopTrigger = true;
                orpoButtonAverAge = false;
                orpoButtonAverLar = true;

                if (shopComboBox.SelectedItem == "СХК")
                    shopName = "r1";
                else if (shopComboBox.SelectedItem == "АЭХК")
                    shopName = "r2";
                else if (shopComboBox.SelectedItem == "МСЗ")
                    shopName = "r3";
                else if (shopComboBox.SelectedItem == "УЭХК")
                    shopName = "r4";
                else if (shopComboBox.SelectedItem == "ПО ЭХЗ")
                    shopName = "r5";
                else if (shopComboBox.SelectedItem == "ЧМЗ")
                    shopName = "r6";

                request = "SELECT [ID], [Dose], [Year], [DoseInt], [Gender], [AgeAtExp] FROM [Final] WHERE [Shop]='" + shopName + "'";

                if (shopComboBox.SelectedItem == "ВСЕ ПРЕДПРИЯТИЯ")
                    request = "SELECT [ID], [Dose], [Year], [DoseInt], [Gender], [AgeAtExp] FROM [Final]";

                connection.Open();

                try
                {
                    OleDbDataAdapter adapter = new OleDbDataAdapter(request, connectionString);//Выбор нужных столбцов из нужной таблицы
                    //OleDbDataAdapter adapter = new OleDbDataAdapter("SELECT [ID], [Dose], [Year], [DoseInt], [Gender], [AgeAtExp] FROM [Final] WHERE [Shop]='r3'", connectionString);//Выбор нужных столбцов из нужной таблицы
                    DataSet dataSet = new DataSet();
                    adapter.Fill(dataSet, "Final");
                    DataTable table = dataSet.Tables[0];//Из Final в эту таблицу считываются поля, указанные в запросе; Выборка для МСК (shop = r3)

                    /*-----Список объектов; достаем все необходимое для расчетов: id, dose, doseInt, ageAtExp, gender-----*/
                    dbFinalRecords = new List<dbObject>();
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        dbFinalRecords.Add(new dbObject(Convert.ToInt32(table.Rows[i]["id"]), Convert.ToByte(table.Rows[i]["gender"]), Convert.ToInt32(table.Rows[i]["year"]), Convert.ToInt16(table.Rows[i]["ageatexp"]), Convert.ToDouble(table.Rows[i]["dose"]), Convert.ToDouble(table.Rows[i]["doseint"])));
                    }

                    /*-----Список, в котором хранится пол-----*/
                    List<byte> dbSex = new List<byte>();
                    for (int i = 0; i < dbFinalRecords.Count; i++)
                        dbSex.Add(dbFinalRecords[i].getSex());

                    /*-----Определение пола; Меньшая цифра пола - М, большая - Ж-----*/
                    sexMale = dbSex.Min();
                    sexFemale = dbSex.Max();

                    /*-----Счетчики, определяющие количество мужских и женских записей-----*/
                    double dbMan = 0;
                    for (int i = 0; i < dbFinalRecords.Count; i++)
                        if (dbFinalRecords[i].getSex() == sexMale)
                            dbMan++;

                    double dbWoman = 0;
                    for (int i = 0; i < dbFinalRecords.Count; i++)
                        if (dbFinalRecords[i].getSex() == sexFemale)
                            dbWoman++;

                    /*-----Массивы списков для мужчин и для женщин, в каждом из которых хранятся дозы (внешние и внутренние) для соответствующий половозрастной группы-----*/
                    List<double>[] manSadExtArray = new List<double>[ageGroups.Count];//SAD - SexAgeDose
                    List<double>[] manSadIntArray = new List<double>[ageGroups.Count];
                    List<double>[] womanSadExtArray = new List<double>[ageGroups.Count];
                    List<double>[] womanSadIntArray = new List<double>[ageGroups.Count];

                    /*-----Массив списоков, через которые будут вычесляться средние возроста половозрастных групп-----*/
                    List<int>[] manYearsArray = new List<int>[ageGroups.Count];
                    List<int>[] womanYearsArray = new List<int>[ageGroups.Count];

                    for (int i = 0; i < ageGroups.Count; i++)
                    {
                        manSadExtArray[i] = new List<double>();
                        manSadIntArray[i] = new List<double>();
                        womanSadExtArray[i] = new List<double>();
                        womanSadIntArray[i] = new List<double>();

                        manYearsArray[i] = new List<int>();
                        womanYearsArray[i] = new List<int>();
                    }

                    /*-----Заполнение массива списков доз-----*/
                    for (int i = 0; i < ageGroups.Count; i++)
                        for (int k = 0; k < dbFinalRecords.Count; k++)
                        {
                            if (dbFinalRecords[k].getSex() == sexMale)
                                if (dbFinalRecords[k].getAgeAtExp() >= ageLowerBound[i] && dbFinalRecords[k].getAgeAtExp() <= ageUpperBound[i])
                                {
                                    manSadExtArray[i].Add(dbFinalRecords[k].getDose() - dbFinalRecords[k].getDoseInt());
                                    manSadIntArray[i].Add(dbFinalRecords[k].getDoseInt());
                                    manYearsArray[i].Add(dbFinalRecords[k].getAgeAtExp());
                                }
                            if (dbFinalRecords[k].getSex() == sexFemale)
                                if (dbFinalRecords[k].getAgeAtExp() >= ageLowerBound[i] && dbFinalRecords[k].getAgeAtExp() <= ageUpperBound[i])
                                {
                                    womanSadExtArray[i].Add(dbFinalRecords[k].getDose() - dbFinalRecords[k].getDoseInt());
                                    womanSadIntArray[i].Add(dbFinalRecords[k].getDoseInt());
                                    womanYearsArray[i].Add(dbFinalRecords[k].getAgeAtExp());
                                }
                        }

                    //Задание весовых коэффициентов для тканей (в нашем случае учитывается только влияние на лёгкие)
                    double wLung = 0.12;

                    /*-----Создание дозовых историй-----*/
                    List<RiskCalculator.DoseHistoryRecord[]>[] manDoseHistoryList = new List<RiskCalculator.DoseHistoryRecord[]>[ageGroups.Count];
                    List<RiskCalculator.DoseHistoryRecord[]>[] womanDoseHistoryList = new List<RiskCalculator.DoseHistoryRecord[]>[ageGroups.Count];
                    for (int i = 0; i < ageGroups.Count; i++)
                    {
                        manDoseHistoryList[i] = new List<RiskCalculator.DoseHistoryRecord[]>();
                        womanDoseHistoryList[i] = new List<RiskCalculator.DoseHistoryRecord[]>();
                    }
                    for (int i = 0; i < ageGroups.Count; i++)
                        for (int k = 0; k < manSadExtArray[i].Count; k++)
                        {
                            manDoseHistoryList[i].Add(new RiskCalculator.DoseHistoryRecord[1]);
                        }
                    for (int i = 0; i < ageGroups.Count; i++)
                        for (int k = 0; k < manSadExtArray[i].Count; k++)
                        {
                            womanDoseHistoryList[i].Add(new RiskCalculator.DoseHistoryRecord[1]);
                        }

                    /*-----Создание массивов, храящих LAR п/в групп-----*/
                    List<double>[] manExtLarArray = new List<double>[ageGroups.Count];
                    List<double>[] manIntLarArray = new List<double>[ageGroups.Count];
                    for (int i = 0; i < ageGroups.Count; i++)
                    {
                        manExtLarArray[i] = new List<double>();
                        manIntLarArray[i] = new List<double>();
                    }

                    List<double>[] womanExtLarArray = new List<double>[ageGroups.Count];
                    List<double>[] womanIntLarArray = new List<double>[ageGroups.Count];
                    for (int i = 0; i < ageGroups.Count; i++)
                    {
                        womanExtLarArray[i] = new List<double>();
                        womanIntLarArray[i] = new List<double>();
                    }

                    /*-----Заполнение ДИ-----*/
                    RiskCalculator.DoseHistoryRecord[] record = null;
                    RiskCalculatorLib.RiskCalculator calculator = null;
                    bool isIncidence = false;
                    for (int i = 0; i < ageGroups.Count; i++)
                        for (int k = 0; k < manSadExtArray[i].Count; k++)
                        {
                            manDoseHistoryList[i][k][0] = new RiskCalculator.DoseHistoryRecord();
                            manDoseHistoryList[i][k][0].AgeAtExposure = (short)manYearsArray[i][k];
                            manDoseHistoryList[i][k][0].AllSolidDoseInmGy = manSadExtArray[i][k];
                            manDoseHistoryList[i][k][0].LeukaemiaDoseInmGy = manSadExtArray[i][k];
                            manDoseHistoryList[i][k][0].LungDoseInmGy = manSadIntArray[i][k] / wLung;

                            record = manDoseHistoryList[i][k];
                            calculator = new RiskCalculatorLib.RiskCalculator(RiskCalculator.SEX_MALE, manDoseHistoryList[i][k][0].AgeAtExposure, ref record, true);
                            if (larRB.Checked)
                            {
                                manExtLarArray[i].Add(calculator.getLAR(false, isIncidence).AllCancers);
                                if (manDoseHistoryList[i][k][0].LungDoseInmGy > 0)
                                    manIntLarArray[i].Add(calculator.getLAR(false, isIncidence).Lung);
                            }
                            if (detRB.Checked)
                            {
                                calculator.createEARSamples(0, ref isIncidence);
                                manExtLarArray[i].Add(calculator.getDetriment().Value.AllCancers);
                                if (manDoseHistoryList[i][k][0].LungDoseInmGy > 0)
                                    manIntLarArray[i].Add(calculator.getDetriment().Value.Lung);
                            }
                        }
                    for (int i = 0; i < ageGroups.Count; i++)
                        for (int k = 0; k < womanSadExtArray[i].Count; k++)
                        {
                            womanDoseHistoryList[i][k][0] = new RiskCalculator.DoseHistoryRecord();
                            womanDoseHistoryList[i][k][0].AgeAtExposure = (short)womanYearsArray[i][k];
                            womanDoseHistoryList[i][k][0].AllSolidDoseInmGy = womanSadExtArray[i][k];
                            womanDoseHistoryList[i][k][0].LeukaemiaDoseInmGy = womanSadExtArray[i][k];
                            womanDoseHistoryList[i][k][0].LungDoseInmGy = womanSadIntArray[i][k] / wLung;

                            record = womanDoseHistoryList[i][k];
                            calculator = new RiskCalculatorLib.RiskCalculator(RiskCalculator.SEX_FEMALE, womanDoseHistoryList[i][k][0].AgeAtExposure, ref record, true);
                            if (larRB.Checked)
                            {
                                womanExtLarArray[i].Add(calculator.getLAR(false, isIncidence).AllCancers);
                                if (womanDoseHistoryList[i][k][0].LungDoseInmGy != 0)
                                    womanIntLarArray[i].Add(calculator.getLAR(false, isIncidence).Lung);
                            }
                            if (detRB.Checked)
                            {
                                calculator.createEARSamples(0, ref isIncidence);
                                womanExtLarArray[i].Add(calculator.getDetriment().Value.AllCancers);
                                if (womanDoseHistoryList[i][k][0].LungDoseInmGy != 0)
                                    womanIntLarArray[i].Add(calculator.getDetriment().Value.Lung);
                            }
                        }

                    /*-----Инициализация массивов, хранящих ОРПО для половозрастных групп-----*/
                    manExtOrpo = new double[ageGroups.Count];
                    manIntOrpo = new double[ageGroups.Count];
                    manSumOrpo = new double[ageGroups.Count];

                    womanExtOrpo = new double[ageGroups.Count];
                    womanIntOrpo = new double[ageGroups.Count];
                    womanSumOrpo = new double[ageGroups.Count];

                    manExtOrpo95 = new double[ageGroups.Count];
                    manIntOrpo95 = new double[ageGroups.Count];
                    manSumOrpo95 = new double[ageGroups.Count];

                    womanExtOrpo95 = new double[ageGroups.Count];
                    womanIntOrpo95 = new double[ageGroups.Count];
                    womanSumOrpo95 = new double[ageGroups.Count];

                    for (int i = 0; i < ageGroups.Count; i++)
                    {
                        if (manExtLarArray[i].Count > 0)
                        {

                            //manExtOrpo[i] = getOrpo(getManExtLar(manYearsArray[i].Average()), manSadExtArray[i].Average());
                            //manExtOrpo_95[i] = getOrpo_95(getManExtLar(manYearsArray[i].Average()), manSadExtArray[i].Average(), getDeviation(manSadExtArray[i]));
                            manExtOrpo[i] = manExtLarArray[i].Average();//manSadExtArray[i].Average());
                            manExtLarArray[i].Sort();
                            if (manExtLarArray[i].Count == 1)
                                manExtOrpo95[i] = manExtLarArray[i][0];
                            if (manExtLarArray[i].Count > 1)
                                manExtOrpo95[i] = manExtLarArray[i][manExtLarArray[i].Count * 95 / 100 - 1];
                        }

                        if (manIntLarArray[i].Count > 0)
                        {
                            //manIntOrpo[i] = getOrpo(getManIntLar(manYearsArray[i].Average()), manSadIntArray[i].Average());
                            // manIntOrpo_95[i] = getOrpo_95(getManIntLar(manYearsArray[i].Average()), manSadIntArray[i].Average(), getDeviation(manSadIntArray[i]));
                            manIntOrpo[i] = manIntLarArray[i].Average();//manSadIntArray[i].Average());
                            manIntLarArray[i].Sort();
                            if (manIntLarArray[i].Count == 1)
                                manIntOrpo95[i] = manIntLarArray[i][0];
                            if (manIntLarArray[i].Count > 1)
                                manIntOrpo95[i] = manIntLarArray[i][manIntLarArray[i].Count * 95 / 100 - 1];
                            //manSadIntArray[i].Sort();
                            //manIntOrpo_95[i] = getOrpo_95(manIntLarArray[i].Average(), manSadIntArray[i][manSadIntArray[i].Count * 95 / 100 - 1]);
                        }

                        manSumOrpo[i] = manExtOrpo[i] + manIntOrpo[i];
                        manSumOrpo95[i] = manExtOrpo95[i] + manIntOrpo95[i];

                        if (womanExtLarArray[i].Count > 0)
                        {
                            //womanExtOrpo[i] = getOrpo(getWomanExtLar(womanYearsArray[i].Average()), womanSadExtArray[i].Average());
                            //womanExtOrpo_95[i] = getOrpo_95(getWomanExtLar(womanYearsArray[i].Average()), womanSadExtArray[i].Average(), getDeviation(womanSadExtArray[i]));
                            womanExtOrpo[i] = womanExtLarArray[i].Average();//womanSadExtArray[i].Average());
                            womanExtLarArray[i].Sort();
                            
                            if (womanExtLarArray[i].Count == 1)
                                womanExtOrpo95[i] = womanExtLarArray[i][0];
                            if (womanExtLarArray[i].Count > 1)
                                womanExtOrpo95[i] = womanExtLarArray[i][womanExtLarArray[i].Count * 95 / 100 - 1];
                            //womanSadExtArray[i].Sort();
                            //womanExtOrpo_95[i] = getOrpo_95(womanExtLarArray[i].Average(), womanSadExtArray[i][womanSadExtArray[i].Count * 95 / 100 - 1]);
                        }

                        if (womanIntLarArray[i].Count > 0)
                        {
                            //womanIntOrpo[i] = getOrpo(getWomanIntLar(womanYearsArray[i].Average()), womanSadIntArray[i].Average());
                            //womanIntOrpo_95[i] = getOrpo_95(getWomanIntLar(womanYearsArray[i].Average()), womanSadIntArray[i].Average(), getDeviation(womanSadIntArray[i]));
                            womanIntOrpo[i] = womanIntLarArray[i].Average();//womanSadIntArray[i].Average());
                            womanIntLarArray[i].Sort();
                            
                            if (womanIntLarArray[i].Count == 1)
                                womanIntOrpo95[i] = womanIntLarArray[i][0];
                            if (womanIntLarArray[i].Count > 1)
                                womanIntOrpo95[i] = womanIntLarArray[i][womanIntLarArray[i].Count * 95 / 100 - 1];
                            //womanSadIntArray[i].Sort();
                            //womanIntOrpo_95[i] = getOrpo_95(womanIntLarArray[i].Average(), womanSadIntArray[i][womanSadIntArray[i].Count * 95 / 100 - 1]);
                        }

                        womanSumOrpo[i] = womanExtOrpo[i] + womanIntOrpo[i];
                        womanSumOrpo95[i] = womanExtOrpo95[i] + womanIntOrpo95[i];
                    }

                    List<double> manWeightedExtOrpo = new List<double>();
                    List<double> manWeightedIntOrpo = new List<double>();
                    List<double> womanWeightedExtOrpo = new List<double>();
                    List<double> womanWeightedIntOrpo = new List<double>();

                    List<double> manWeightedExtOrpo95 = new List<double>();
                    List<double> manWeightedIntOrpo95 = new List<double>();
                    List<double> womanWeightedExtOrpo95 = new List<double>();
                    List<double> womanWeightedIntOrpo95 = new List<double>();

                    for (int i = 0; i < ageGroups.Count; i++)
                    {
                        manWeightedExtOrpo.Add(manExtOrpo[i] * manSadExtArray[i].Count);
                        manWeightedIntOrpo.Add(manIntOrpo[i] * manSadIntArray[i].Count);
                        womanWeightedExtOrpo.Add(womanExtOrpo[i] * womanSadExtArray[i].Count);
                        womanWeightedIntOrpo.Add(womanIntOrpo[i] * womanSadIntArray[i].Count);

                        manWeightedExtOrpo95.Add(manExtOrpo95[i] * manSadExtArray[i].Count);
                        manWeightedIntOrpo95.Add(manIntOrpo95[i] * manSadIntArray[i].Count);
                        womanWeightedExtOrpo95.Add(womanExtOrpo95[i] * womanSadExtArray[i].Count);
                        womanWeightedIntOrpo95.Add(womanIntOrpo95[i] * womanSadIntArray[i].Count);
                    }

                    //manExtOrpoBox.Text = "2-б) " + Math.Round(manWeightedExtOrpo.Sum() / dbMan, 7).ToString();
                    //manIntOrpoBox.Text = "2-б) " + /*manWeightedIntOrpo.Sum() / dbMan;//*/Math.Round(manWeightedIntOrpo.Sum() / dbMan, 7).ToString();
                    //womanExtOrpoBox.Text = "2-б) " + Math.Round(womanWeightedExtOrpo.Sum() / dbWoman, 7).ToString();
                    //womanIntOrpoBox.Text = "2-б) " + /*womanWeightedIntOrpo.Sum() / dbWoman;//*/Math.Round(womanWeightedIntOrpo.Sum() / dbWoman, 7).ToString();

                    //manExtOrpoBox95.Text = "2-б) " + Math.Round(manWeightedExtOrpo95.Sum() / dbMan, 7).ToString();
                    //manIntOrpoBox95.Text = "2-б) " + /*manWeightedIntOrpo_95.Sum() / dbMan;//*/Math.Round(manWeightedIntOrpo95.Sum() / dbMan, 7).ToString();
                    //womanExtOrpoBox95.Text = "2-б) " + Math.Round(womanWeightedExtOrpo95.Sum() / dbWoman, 7).ToString();
                    //womanIntOrpoBox95.Text = "2-б) " + /*womanWeightedIntOrpo_95.Sum() / dbWoman;//*/Math.Round(womanWeightedIntOrpo95.Sum() / dbWoman, 7).ToString();

                    //if (manWeightedExtOrpo.Sum() > 0)
                    //    manExtOrpoBox.Text = "2-б) " + Math.Round(manWeightedExtOrpo.Sum() / dbMan, 8);
                    //else
                    //    manIntOrpoBox.Text = "Внешнего облучения нет!";

                    //if (manWeightedIntOrpo.Sum() > 0)
                    //    manIntOrpoBox.Text = "2-б) " + Math.Round(manWeightedIntOrpo.Sum() / dbMan, 8);
                    //else
                    //    manIntOrpoBox.Text = "Внутреннего облучения нет!";

                    //manSumOrpoBox.Text = "2-б) " + Math.Round((manWeightedExtOrpo.Sum() / dbMan) + (manWeightedIntOrpo.Sum() / dbMan), 8);

                    //if (womanWeightedExtOrpo.Sum() > 0)
                    //    womanExtOrpoBox.Text = "2-б) " + Math.Round(womanWeightedExtOrpo.Sum() / dbWoman, 8);
                    //else
                    //    womanExtOrpoBox.Text = "Внешнего облучения нет!";

                    //if (womanWeightedIntOrpo.Sum() > 0)
                    //    womanIntOrpoBox.Text = "2-б) " + Math.Round(womanWeightedIntOrpo.Sum() / dbWoman, 8);
                    //else
                    //    womanIntOrpoBox.Text = "Внутреннего облучения нет!";

                    //womanSumOrpoBox.Text = "2-б) " + Math.Round((womanWeightedExtOrpo.Sum() / dbWoman) + (womanWeightedIntOrpo.Sum() / dbWoman), 8);

                    //if (manWeightedExtOrpo95.Sum() > 0)
                    //    manExtOrpoBox95.Text = "2-б) " + Math.Round(manWeightedExtOrpo95.Sum() / dbMan, 8);
                    //else
                    //    manIntOrpoBox95.Text = "Внешнего облучения нет!";

                    //if (manWeightedIntOrpo95.Sum() > 0)
                    //    manIntOrpoBox95.Text = "2-б) " + Math.Round(manWeightedIntOrpo95.Sum() / dbMan, 8);
                    //else
                    //    manIntOrpoBox95.Text = "Внутреннего облучения нет!";

                    //manSumOrpo95Box.Text = "2-а) " + Math.Round((manWeightedExtOrpo95.Sum() / dbMan) + (manWeightedIntOrpo95.Sum() / dbMan), 8);

                    //if (womanWeightedExtOrpo95.Sum() > 0)
                    //    womanExtOrpoBox95.Text = "2-б) " + Math.Round(womanWeightedExtOrpo95.Sum() / dbWoman, 8);
                    //else
                    //    womanExtOrpoBox95.Text = "Внешнего облучения нет!";

                    //if (womanWeightedIntOrpo95.Sum() > 0)
                    //    womanIntOrpoBox95.Text = "2-б) " + Math.Round(womanWeightedIntOrpo95.Sum() / dbWoman, 8);
                    //else
                    //    womanIntOrpoBox95.Text = "Внутреннего облучения нет!";

                    //womanSumOrpo95Box.Text = "2-б) " + Math.Round((womanWeightedExtOrpo95.Sum() / dbWoman) + (womanWeightedIntOrpo95.Sum() / dbWoman), 8);

                    //manExtIbpoBox.Text = "" + manWeightedExtOrpo.Count;
                    //manIntIbpoBox.Text = "" + manWeightedIntOrpo.Count;
                    //womanExtIbpoBox.Text = "" + womanWeightedExtOrpo.Count;
                    //womanIntIbpoBox.Text = "" + womanWeightedIntOrpo.Count;

                    //manExtIbpoBox95.Text = "" + manWeightedExtOrpo95.Count;
                    //manIntIbpoBox95.Text = "" + manWeightedIntOrpo95.Count;
                    //womanExtIbpoBox95.Text = "" + womanWeightedExtOrpo95.Count;
                    //womanIntIbpoBox95.Text = "" + womanWeightedIntOrpo95.Count;
                }
                catch
                {
                    MessageBox.Show("Не выбрано предприятие!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
                    Application.DoEvents();
                }
            }

            catch/*(OleDbException ex)*/
            {
                MessageBox.Show("Не выбрана база данных!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
                Application.DoEvents();
            }
        }

        private void getIbpoButton_Click(object sender, EventArgs e)
        {
            //stopWatch.Start();
            OleDbConnection connection = new OleDbConnection(connectionString);
            try
            {
                if (!shopTrigger)
                {
                    MessageBox.Show("Смените предприятие и пересчитайте ОРПО!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
                    Application.DoEvents();
                }
                else
                {
                    connection.Open();
                    try
                    {
                        OleDbDataAdapter adapter = new OleDbDataAdapter("SELECT [ID], [Year], [Dose], [DoseInt] FROM [Dose]", connectionString);
                        DataSet dataSet = new DataSet();
                        adapter.Fill(dataSet, "Dose");
                        DataTable table = dataSet.Tables[0];

                        /*-----Список объектов, хранящих данные из таблицы Dose-----*/
                        dbDoseRecords = new List<dbObject>();
                        for (int i = 0; i < table.Rows.Count; i++)
                        {
                            dbDoseRecords.Add(new dbObject(Convert.ToInt32(table.Rows[i]["id"]), Convert.ToInt32(table.Rows[i]["year"]), Convert.ToDouble(table.Rows[i]["dose"]) / 1000, Convert.ToDouble(table.Rows[i]["doseint"]) / 1000));
                        }

                        /*-----Списки id, у которых есть записи в 2012 году-----*/
                        List<dbObject> manRecordsList = new List<dbObject>();
                        List<dbObject> womanRecordsList = new List<dbObject>();
                        for (int i = 0; i < dbFinalRecords.Count; i++)
                        {
                            if (dbFinalRecords[i].getSex() == sexMale && dbFinalRecords[i].getYear() == 2012)
                                manRecordsList.Add(dbFinalRecords[i]);
                            if (dbFinalRecords[i].getSex() == sexFemale && dbFinalRecords[i].getYear() == 2012)
                                womanRecordsList.Add(dbFinalRecords[i]);
                        }

                        if (manRecordsList.Count == 0 || womanRecordsList.Count == 0)
                        {
                            //manExtIbpoBox.Text = "Нет данных за 2012 год!";
                            //manIntIbpoBox.Text = "Нет данных за 2012 год!";
                            //manExtIbpoBox95.Text = "Нет данных за 2012 год!";
                            //manIntIbpoBox95.Text = "Нет данных за 2012 год!";
                            //womanExtIbpoBox.Text = "Нет данных за 2012 год!";
                            //womanIntIbpoBox.Text = "Нет данных за 2012 год!";
                            //womanExtIbpoBox95.Text = "Нет данных за 2012 год!";
                            //womanIntIbpoBox95.Text = "Нет данных за 2012 год!";
                        }

                        else
                        {
                            /*-----Массивы списков для мужчин и женщин.
                             * Каждый элемент массива - список, содержащий элементы,
                             * у которых одинаковые id (это записи дозовой истории конкретного человека)-----*/
                            List<dbObject>[] manGroupedRecordsArray = new List<dbObject>[manRecordsList.Count];
                            for (int i = 0; i < manGroupedRecordsArray.Length; i++)
                                manGroupedRecordsArray[i] = new List<dbObject>();
                            List<dbObject>[] womanGroupedRecordsArray = new List<dbObject>[womanRecordsList.Count];
                            for (int i = 0; i < womanGroupedRecordsArray.Length; i++)
                                womanGroupedRecordsArray[i] = new List<dbObject>();

                            /*-----Заполнение этих массивов-----*/
                            List<dbObject> buffer = null;
                            for (int i = 0; i < manGroupedRecordsArray.Length; i++)
                            {
                                buffer = new List<dbObject>();
                                for (int k = 0; k < dbDoseRecords.Count; k++)
                                {
                                    if (manRecordsList[i].getId() == dbDoseRecords[k].getId())
                                    {
                                        dbDoseRecords[k].setSex(sexMale);
                                        dbDoseRecords[k].setAgeAtExp(Convert.ToInt16(manRecordsList[i].getAgeAtExp() - (manRecordsList[i].getYear() - dbDoseRecords[k].getYear())));
                                        buffer.Add(dbDoseRecords[k]);
                                    }
                                }
                                manGroupedRecordsArray[i] = buffer;
                            }
                            for (int i = 0; i < womanGroupedRecordsArray.Length; i++)
                            {
                                buffer = new List<dbObject>();
                                for (int k = 0; k < dbDoseRecords.Count; k++)
                                {
                                    if (womanRecordsList[i].getId() == dbDoseRecords[k].getId())
                                    {
                                        dbDoseRecords[k].setSex(sexFemale);
                                        dbDoseRecords[k].setAgeAtExp(Convert.ToInt16(womanRecordsList[i].getAgeAtExp() - (womanRecordsList[i].getYear() - dbDoseRecords[k].getYear())));
                                        buffer.Add(dbDoseRecords[k]);
                                    }
                                }
                                womanGroupedRecordsArray[i] = buffer;
                            }

                            /*-----Задание весовых коэффициентов для тканей (в нашем случае учитывается только влияние на лёгкие)-----*/
                            double wLung = 0.12;

                            /*-----Создание пустого списка дозовых историй мужчин; для каждого уникального ID своя дозовая история (по сути, это ячейки, которые надо заполнить)-----*/
                            List<RiskCalculator.DoseHistoryRecord[]> manDoseHistoryList = new List<RiskCalculator.DoseHistoryRecord[]>();
                            for (int i = 0; i < manGroupedRecordsArray.Length; i++)
                            {
                                manDoseHistoryList.Add(new RiskCalculator.DoseHistoryRecord[manGroupedRecordsArray[i].Count]);
                            }
                            foreach (RiskCalculator.DoseHistoryRecord[] note in manDoseHistoryList)
                            {
                                for (int i = 0; i < note.Length; i++)
                                    note[i] = new RiskCalculator.DoseHistoryRecord();
                            }

                            /*-----Создание аналогичного списка дозовых историй для женщин-----*/
                            List<RiskCalculator.DoseHistoryRecord[]> womanDoseHistoryList = new List<RiskCalculator.DoseHistoryRecord[]>();
                            for (int i = 0; i < womanGroupedRecordsArray.Length; i++)
                            {
                                womanDoseHistoryList.Add(new RiskCalculator.DoseHistoryRecord[womanGroupedRecordsArray[i].Count]);
                            }
                            foreach (RiskCalculator.DoseHistoryRecord[] note in womanDoseHistoryList)
                            {
                                for (int i = 0; i < note.Length; i++)
                                    note[i] = new RiskCalculator.DoseHistoryRecord();
                            }

                            /*-----Заполнение дозовых историй мужчин-----*/
                            for (int i = 0; i < manGroupedRecordsArray.Length; i++)
                                for (int k = 0; k < manGroupedRecordsArray[i].Count; k++)
                                {
                                    manDoseHistoryList[i][k].AgeAtExposure = manGroupedRecordsArray[i][k].getAgeAtExp();
                                    manDoseHistoryList[i][k].AllSolidDoseInmGy = manGroupedRecordsArray[i][k].getDose() - manGroupedRecordsArray[i][k].getDoseInt();
                                    manDoseHistoryList[i][k].LeukaemiaDoseInmGy = manGroupedRecordsArray[i][k].getDose() - manGroupedRecordsArray[i][k].getDoseInt();
                                    manDoseHistoryList[i][k].LungDoseInmGy = manGroupedRecordsArray[i][k].getDoseInt() / wLung;
                                }
                            /*-----Заполнение дозовых историй женщин-----*/
                            for (int i = 0; i < womanGroupedRecordsArray.Length; i++)
                                for (int k = 0; k < womanGroupedRecordsArray[i].Count; k++)
                                {
                                    womanDoseHistoryList[i][k].AgeAtExposure = womanGroupedRecordsArray[i][k].getAgeAtExp();
                                    womanDoseHistoryList[i][k].AllSolidDoseInmGy = womanGroupedRecordsArray[i][k].getDose() - womanGroupedRecordsArray[i][k].getDoseInt();
                                    womanDoseHistoryList[i][k].LeukaemiaDoseInmGy = womanGroupedRecordsArray[i][k].getDose() - womanGroupedRecordsArray[i][k].getDoseInt();
                                    womanDoseHistoryList[i][k].LungDoseInmGy = womanGroupedRecordsArray[i][k].getDoseInt() / wLung;
                                }

                            /*-----Создание массива списков для п/в групп мужчин, хранящих LAR п/в группы;
                             * каждый элемент массива - список LAR-ов п/в группы.
                             * Это массивы LAR от внешнего облучения.
                             * Для внутреннего отдельно надо-----*/
                            List<double>[] manExtLarArray = new List<double>[ageGroups.Count];
                            List<double>[] manIntLarArray = new List<double>[ageGroups.Count];
                            List<double>[] manSumLarArray = new List<double>[ageGroups.Count];
                            for (int i = 0; i < ageGroups.Count; i++)
                            {
                                manExtLarArray[i] = new List<double>();
                                manIntLarArray[i] = new List<double>();
                                manSumLarArray[i] = new List<double>();
                            }

                            /*-----Создание аналогичного массива для женщин-----*/
                            List<double>[] womanExtLarArray = new List<double>[ageGroups.Count];
                            List<double>[] womanIntLarArray = new List<double>[ageGroups.Count];
                            List<double>[] womanSumLarArray = new List<double>[ageGroups.Count];
                            for (int i = 0; i < ageGroups.Count; i++)
                            {
                                womanExtLarArray[i] = new List<double>();
                                womanIntLarArray[i] = new List<double>();
                                womanSumLarArray[i] = new List<double>();
                            }

                            /*-----Заполнение этих массивов-----*/
                            RiskCalculator.DoseHistoryRecord[] record = null;
                            RiskCalculatorLib.RiskCalculator calculator = null;
                            for (int i = 0; i < ageGroups.Count; i++)
                                for (int k = 0; k < manDoseHistoryList.Count; k++)
                                {
                                    if (manRecordsList[k].getAgeAtExp() == manDoseHistoryList[k][manDoseHistoryList[k].Length - 1].AgeAtExposure)
                                        if (manRecordsList[k].getAgeAtExp() >= ageLowerBound[i] && manRecordsList[k].getAgeAtExp() <= ageUpperBound[i])
                                        {
                                            record = manDoseHistoryList[k];
                                            calculator = new RiskCalculatorLib.RiskCalculator(RiskCalculator.SEX_MALE, manDoseHistoryList[k][0].AgeAtExposure, ref record, true);
                                            manExtLarArray[i].Add(calculator.getLAR(false, true).AllCancers);//Кажется, здесь считается LAR...
                                            manIntLarArray[i].Add(calculator.getLAR(false, true).Lung);
                                            manSumLarArray[i].Add(calculator.getLAR(false, true).AllCancers + calculator.getLAR(false, true).Lung);
                                        }
                                }

                            for (int i = 0; i < ageGroups.Count; i++)
                                for (int k = 0; k < womanDoseHistoryList.Count; k++)
                                {
                                    if (womanRecordsList[k].getAgeAtExp() == womanDoseHistoryList[k][womanDoseHistoryList[k].Length - 1].AgeAtExposure)
                                        if (womanRecordsList[k].getAgeAtExp() >= ageLowerBound[i] && womanRecordsList[k].getAgeAtExp() <= ageUpperBound[i])
                                        {
                                            record = womanDoseHistoryList[k];
                                            calculator = new RiskCalculatorLib.RiskCalculator(RiskCalculator.SEX_FEMALE, womanDoseHistoryList[k][0].AgeAtExposure, ref record, true);
                                            womanExtLarArray[i].Add(calculator.getLAR(false, true).AllCancers);//в этой строчке должен вычисляться LAR и записываться в этот список
                                            womanIntLarArray[i].Add(calculator.getLAR(false, true).Lung);
                                            womanSumLarArray[i].Add(calculator.getLAR(false, true).AllCancers + calculator.getLAR(false, true).Lung);
                                        }
                                }

                            double[] manExtIbpo = new double[ageGroups.Count];
                            double[] manExtIbpo95 = new double[ageGroups.Count];

                            double[] manIntIbpo = new double[ageGroups.Count];
                            double[] manIntIbpo95 = new double[ageGroups.Count];

                            double[] manSumIbpo = new double[ageGroups.Count];
                            double[] manSumIbpo95 = new double[ageGroups.Count];

                            double[] womanExtIbpo = new double[ageGroups.Count];
                            double[] womanExtIbpo95 = new double[ageGroups.Count];

                            double[] womanIntIbpo = new double[ageGroups.Count];
                            double[] womanIntIbpo95 = new double[ageGroups.Count];

                            double[] womanSumIbpo = new double[ageGroups.Count];
                            double[] womanSumIbpo95 = new double[ageGroups.Count];


                            for (int i = 0; i < ageGroups.Count; i++)
                            {
                                if (manExtLarArray[i].Count > 0)
                                {
                                    manExtIbpo[i] = getIbpo(manExtLarArray[i], manExtOrpo[i]);
                                    manExtIbpo95[i] = getIbpo(manExtLarArray[i], manExtOrpo95[i]);
                                }

                                if (manIntLarArray[i].Count > 0)
                                {
                                    manIntIbpo[i] = getIbpo(manIntLarArray[i], manIntOrpo[i]);
                                    manIntIbpo95[i] = getIbpo(manIntLarArray[i], manIntOrpo95[i]);
                                }

                                if (manSumLarArray[i].Count > 0)
                                {
                                    manSumIbpo[i] = getIbpo(manSumLarArray[i], manSumOrpo[i]);
                                    manSumIbpo95[i] = getIbpo(manSumLarArray[i], manSumOrpo95[i]);
                                }

                                if (womanExtLarArray[i].Count > 0)
                                {
                                    womanExtIbpo[i] = getIbpo(womanExtLarArray[i], womanExtOrpo[i]);
                                    womanExtIbpo95[i] = getIbpo(womanExtLarArray[i], womanExtOrpo95[i]);
                                }

                                if (womanIntLarArray[i].Count > 0)
                                {
                                    womanIntIbpo[i] = getIbpo(womanIntLarArray[i], womanIntOrpo[i]);
                                    womanIntIbpo95[i] = getIbpo(womanIntLarArray[i], womanIntOrpo95[i]);
                                }

                                if (womanSumLarArray[i].Count > 0)
                                {
                                    womanSumIbpo[i] = getIbpo(womanSumLarArray[i], womanSumOrpo[i]);
                                    womanSumIbpo95[i] = getIbpo(womanSumLarArray[i], womanSumOrpo95[i]);
                                }
                            }

                            /*-----Создание списков для подсчета взвешенных величин ИБПО-----*/
                            List<double> manWeightedExtIbpo = new List<double>();
                            List<double> manWeightedIntIbpo = new List<double>();
                            List<double> manWeightedSumIbpo = new List<double>();
                            List<double> womanWeightedExtIbpo = new List<double>();
                            List<double> womanWeightedIntIbpo = new List<double>();
                            List<double> womanWeightedSumIbpo = new List<double>();

                            List<double> manWeightedExtIbpo95 = new List<double>();
                            List<double> manWeightedIntIbpo95 = new List<double>();
                            List<double> manWeightedSumIbpo95 = new List<double>();
                            List<double> womanWeightedExtIbpo95 = new List<double>();
                            List<double> womanWeightedIntIbpo95 = new List<double>();
                            List<double> womanWeightedSumIbpo95 = new List<double>();

                            for (int i = 0; i < ageGroups.Count; i++)
                            {
                                manWeightedExtIbpo.Add(manExtIbpo[i] * manExtLarArray[i].Count);
                                manWeightedIntIbpo.Add(manIntIbpo[i] * manIntLarArray[i].Count);
                                manWeightedSumIbpo.Add(manSumIbpo[i] * manSumLarArray[i].Count);
                                womanWeightedExtIbpo.Add(womanExtIbpo[i] * womanExtLarArray[i].Count);
                                womanWeightedIntIbpo.Add(womanIntIbpo[i] * womanIntLarArray[i].Count);
                                womanWeightedSumIbpo.Add(womanSumIbpo[i] * womanSumLarArray[i].Count);

                                manWeightedExtIbpo95.Add(manExtIbpo95[i] * manExtLarArray[i].Count);
                                manWeightedIntIbpo95.Add(manIntIbpo95[i] * manIntLarArray[i].Count);
                                manWeightedSumIbpo95.Add(manSumIbpo95[i] * manSumLarArray[i].Count);
                                womanWeightedExtIbpo95.Add(womanExtIbpo95[i] * womanExtLarArray[i].Count);
                                womanWeightedIntIbpo95.Add(womanIntIbpo95[i] * womanIntLarArray[i].Count);
                                womanWeightedSumIbpo95.Add(womanSumIbpo95[i] * womanSumLarArray[i].Count);
                            }

                            //if (manWeightedExtIbpo.Sum() / manRecordsList.Count < 100)
                            //    manExtIbpoBox.Text = Math.Round(manWeightedExtIbpo.Sum() / manRecordsList.Count, 2).ToString();
                            //else
                            //    manExtIbpoBox.Text = "Нет внешнего облучения";

                            //if (manWeightedIntIbpo.Sum() / manRecordsList.Count < 100)
                            //    manIntIbpoBox.Text = Math.Round(manWeightedIntIbpo.Sum() / manRecordsList.Count, 2).ToString();
                            //else
                            //    manIntIbpoBox.Text = "Нет внутреннего облучения";

                            //if (manWeightedSumIbpo.Sum() / manRecordsList.Count < 100)
                            //    manSumIbpoBox.Text = Math.Round(manWeightedSumIbpo.Sum() / manRecordsList.Count, 2).ToString();
                            //else
                            //    manSumIbpoBox.Text = "Нет внутреннего облучения";

                            //if (manWeightedExtIbpo95.Sum() / manRecordsList.Count < 100)
                            //    manExtIbpoBox95.Text = Math.Round(manWeightedExtIbpo95.Sum() / manRecordsList.Count, 2).ToString();
                            //else
                            //    manExtIbpoBox95.Text = "Нет внешнего облучения";

                            //if (manWeightedIntIbpo95.Sum() / manRecordsList.Count < 100)
                            //    manIntIbpoBox95.Text = Math.Round(manWeightedIntIbpo95.Sum() / manRecordsList.Count, 2).ToString();
                            //else
                            //    manIntIbpoBox95.Text = "Нет внутреннего облучения";

                            //if (manWeightedSumIbpo95.Sum() / manRecordsList.Count < 100)
                            //    manSumIbpo95Box.Text = Math.Round(manWeightedSumIbpo95.Sum() / manRecordsList.Count, 2).ToString();
                            //else
                            //    manSumIbpoBox.Text = "Нет внутреннего облучения";

                            //if (womanWeightedExtIbpo.Sum() / womanRecordsList.Count < 100)
                            //    womanExtIbpoBox.Text = Math.Round(womanWeightedExtIbpo.Sum() / womanRecordsList.Count, 2).ToString();
                            //else
                            //    womanExtIbpoBox.Text = "Нет внешнего облучения";

                            //if (womanWeightedIntIbpo.Sum() / womanRecordsList.Count < 100)
                            //    womanIntIbpoBox.Text = Math.Round(womanWeightedIntIbpo.Sum() / womanRecordsList.Count, 2).ToString();
                            //else
                            //    womanIntIbpoBox.Text = "Нет внутреннего облучения";

                            //if (womanWeightedSumIbpo.Sum() / womanRecordsList.Count < 100)
                            //    womanSumIbpoBox.Text = Math.Round(womanWeightedSumIbpo.Sum() / womanRecordsList.Count, 2).ToString();
                            //else
                            //    womanSumIbpoBox.Text = "Нет внутреннего облучения";

                            //if (womanWeightedExtIbpo95.Sum() / womanRecordsList.Count < 100)
                            //    womanExtIbpoBox95.Text = Math.Round(womanWeightedExtIbpo95.Sum() / womanRecordsList.Count, 2).ToString();
                            //else
                            //    womanExtIbpoBox95.Text = "Нет внешнего облучения";

                            //if (womanWeightedIntIbpo95.Sum() / womanRecordsList.Count < 100)
                            //    womanIntIbpoBox95.Text = Math.Round(womanWeightedIntIbpo95.Sum() / womanRecordsList.Count, 2).ToString();
                            //else
                            //    womanIntIbpoBox95.Text = "Нет внутреннего облучения";

                            //if (womanWeightedSumIbpo95.Sum() / womanRecordsList.Count < 100)
                            //    womanSumIbpo95Box.Text = Math.Round(womanWeightedSumIbpo95.Sum() / womanRecordsList.Count, 2).ToString();
                            //else
                            //    womanSumIbpo95Box.Text = "Нет внутреннего облучения";

                            /*-----Вывод в Excel-файл-----*/
                            /*-----Инициализация Excel-файла-----*/
                            Excel.Application excelApp = new Excel.Application();
                            //excelApp.Visible = true;
                            //excelApp.DisplayAlerts = true;
                            excelApp.StandardFont = "Times-New-Roman";
                            excelApp.StandardFontSize = 12;

                            /*-----Создание рабочей книги с 4 страницами, в которые будет выводиться информация-----*/
                            excelApp.Workbooks.Add(Type.Missing);
                            Excel.Workbook excelWorkbook = excelApp.Workbooks[1];
                            excelApp.SheetsInNewWorkbook = 2;
                            Excel.Worksheet excelWorksheet = null;
                            Excel.Range excelCells = null;

                            /*-----Вывод в столбцы-----*/
                            excelWorksheet = (Excel.Worksheet)excelWorkbook.Worksheets.get_Item(1);
                            excelWorksheet.Name = "Мужчины";

                            /*-----Описываем ячейку А1 на странице-----*/
                            excelCells = excelWorksheet.get_Range("A1");
                            excelCells.VerticalAlignment = Excel.Constants.xlCenter;
                            excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
                            excelCells.Borders.Weight = Excel.XlBorderWeight.xlThick;
                            excelCells.Value2 = "Возрастные группы";

                            /*-----Описываем ячейку B1 на странице-----*/
                            excelCells = excelWorksheet.get_Range("B1");
                            excelCells.VerticalAlignment = Excel.Constants.xlCenter;
                            excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
                            excelCells.Borders.Weight = Excel.XlBorderWeight.xlThick;
                            excelCells.Value2 = "Внешнее облучение";

                            /*-----Описываем ячейку C1 на странице-----*/
                            excelCells = excelWorksheet.get_Range("C1");
                            excelCells.VerticalAlignment = Excel.Constants.xlCenter;
                            excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
                            excelCells.Borders.Weight = Excel.XlBorderWeight.xlThick;
                            excelCells.Value2 = "Внутреннее облучение";

                            /*-----Описываем ячейку B1 на странице-----*/
                            excelCells = excelWorksheet.get_Range("D1");
                            excelCells.VerticalAlignment = Excel.Constants.xlCenter;
                            excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
                            excelCells.Borders.Weight = Excel.XlBorderWeight.xlThick;
                            excelCells.Value2 = "Сумма";

                            /*-----Описываем ячейку C1 на странице-----*/
                            excelCells = excelWorksheet.get_Range("E1");
                            excelCells.VerticalAlignment = Excel.Constants.xlCenter;
                            excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
                            excelCells.Borders.Weight = Excel.XlBorderWeight.xlThick;
                            excelCells.Value2 = "Внешнее облучение (95%)";

                            /*-----Описываем ячейку C1 на странице-----*/
                            excelCells = excelWorksheet.get_Range("F1");
                            excelCells.VerticalAlignment = Excel.Constants.xlCenter;
                            excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
                            excelCells.Borders.Weight = Excel.XlBorderWeight.xlThick;
                            excelCells.Value2 = "Внутреннее облучение (95%)";

                            /*-----Описываем ячейку C1 на странице-----*/
                            excelCells = excelWorksheet.get_Range("G1");
                            excelCells.VerticalAlignment = Excel.Constants.xlCenter;
                            excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
                            excelCells.Borders.Weight = Excel.XlBorderWeight.xlThick;
                            excelCells.Value2 = "Сумма (95%)";

                            for (int i = 2; i <= manExtIbpo.Length + 1; i++)
                            {
                                excelCells = (Excel.Range)excelWorksheet.Cells[i, "A"];
                                excelCells.Value2 = ageGroups[i - 2];
                                excelCells.Borders.ColorIndex = 1;

                                excelCells = (Excel.Range)excelWorksheet.Cells[i, "B"];
                                if (manExtIbpo[i - 2] < 100)
                                    excelCells.Value2 = manExtIbpo[i - 2];
                                else
                                    excelCells.Value2 = "Нет облучения";
                                excelCells.Borders.ColorIndex = 1;

                                excelCells = (Excel.Range)excelWorksheet.Cells[i, "C"];
                                if (manIntIbpo[i - 2] < 100)
                                    excelCells.Value2 = manIntIbpo[i - 2];
                                else
                                    excelCells.Value2 = "Нет облучения";
                                excelCells.Borders.ColorIndex = 1;

                                excelCells = (Excel.Range)excelWorksheet.Cells[i, "D"];
                                if (manSumIbpo[i - 2] < 100)
                                    excelCells.Value2 = manSumIbpo[i - 2];
                                else
                                    excelCells.Value2 = "Нет облучения";
                                excelCells.Borders.ColorIndex = 1;

                                excelCells = (Excel.Range)excelWorksheet.Cells[i, "E"];
                                if (manExtIbpo95[i - 2] < 100)
                                    excelCells.Value2 = manExtIbpo95[i - 2];
                                else
                                    excelCells.Value2 = "Нет облучения";
                                excelCells.Borders.ColorIndex = 1;

                                excelCells = (Excel.Range)excelWorksheet.Cells[i, "F"];
                                if (manIntIbpo95[i - 2] < 100)
                                    excelCells.Value2 = manIntIbpo95[i - 2];
                                else
                                    excelCells.Value2 = "Нет облучения";
                                excelCells.Borders.ColorIndex = 1;

                                excelCells = (Excel.Range)excelWorksheet.Cells[i, "G"];
                                if (manSumIbpo95[i - 2] < 100)
                                    excelCells.Value2 = manSumIbpo95[i - 2];
                                else
                                    excelCells.Value2 = "Нет облучения";
                                excelCells.Borders.ColorIndex = 1;
                            }

                            /*-----Описываем ячейку А1 на странице-----*/
                            excelCells = excelWorksheet.get_Range("A" + 13);
                            excelCells.VerticalAlignment = Excel.Constants.xlCenter;
                            excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
                            excelCells.Borders.Weight = Excel.XlBorderWeight.xlMedium;
                            excelCells.Value2 = "Взвешенные величины";

                            /*-----Описываем ячейку А1 на странице-----*/
                            excelCells = excelWorksheet.get_Range("B" + 13);
                            excelCells.VerticalAlignment = Excel.Constants.xlCenter;
                            excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
                            excelCells.Borders.Weight = Excel.XlBorderWeight.xlMedium;
                            if (manWeightedExtIbpo.Sum() / manRecordsList.Count < 100)
                                excelCells.Value2 = Math.Round(manWeightedExtIbpo.Sum() / manRecordsList.Count, 8);
                            else
                                excelCells.Value2 = "Нет облучения";

                            /*-----Описываем ячейку А1 на странице-----*/
                            excelCells = excelWorksheet.get_Range("C" + 13);
                            excelCells.VerticalAlignment = Excel.Constants.xlCenter;
                            excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
                            excelCells.Borders.Weight = Excel.XlBorderWeight.xlMedium;
                            if (manWeightedIntIbpo.Sum() / manRecordsList.Count < 100)
                                excelCells.Value2 = Math.Round(manWeightedIntIbpo.Sum() / manRecordsList.Count, 8);
                            else
                                excelCells.Value2 = "Нет облучения";

                            /*-----Описываем ячейку А1 на странице-----*/
                            excelCells = excelWorksheet.get_Range("D" + 13);
                            excelCells.VerticalAlignment = Excel.Constants.xlCenter;
                            excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
                            excelCells.Borders.Weight = Excel.XlBorderWeight.xlMedium;
                            if (manWeightedSumIbpo.Sum() / manRecordsList.Count < 100)
                                excelCells.Value2 = Math.Round(manWeightedSumIbpo.Sum() / manRecordsList.Count, 8);
                            else
                                excelCells.Value2 = "Нет облучения";

                            /*-----Описываем ячейку А1 на странице-----*/
                            excelCells = excelWorksheet.get_Range("E" + 13);
                            excelCells.VerticalAlignment = Excel.Constants.xlCenter;
                            excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
                            excelCells.Borders.Weight = Excel.XlBorderWeight.xlMedium;
                            if (manWeightedExtIbpo95.Sum() / manRecordsList.Count < 100)
                                excelCells.Value2 = Math.Round(manWeightedExtIbpo95.Sum() / manRecordsList.Count, 8);
                            else
                                excelCells.Value2 = "Нет облучения";

                            /*-----Описываем ячейку А1 на странице-----*/
                            excelCells = excelWorksheet.get_Range("F" + 13);
                            excelCells.VerticalAlignment = Excel.Constants.xlCenter;
                            excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
                            excelCells.Borders.Weight = Excel.XlBorderWeight.xlMedium;
                            if (manWeightedIntIbpo95.Sum() / manRecordsList.Count < 100)
                                excelCells.Value2 = Math.Round(manWeightedIntIbpo95.Sum() / manRecordsList.Count, 8);
                            else
                                excelCells.Value2 = "Нет облучения";

                            /*-----Описываем ячейку А1 на странице-----*/
                            excelCells = excelWorksheet.get_Range("G" + 13);
                            excelCells.VerticalAlignment = Excel.Constants.xlCenter;
                            excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
                            excelCells.Borders.Weight = Excel.XlBorderWeight.xlMedium;
                            if (manWeightedSumIbpo95.Sum() / manRecordsList.Count < 100)
                                excelCells.Value2 = Math.Round(manWeightedSumIbpo95.Sum() / manRecordsList.Count, 8);
                            else
                                excelCells.Value2 = "Нет облучения";

                            excelWorksheet = (Excel.Worksheet)excelWorkbook.Worksheets.get_Item(2);
                            excelWorksheet.Name = "Женщины";

                            /*-----Описываем ячейку А1 на странице-----*/
                            excelCells = excelWorksheet.get_Range("A1");
                            excelCells.VerticalAlignment = Excel.Constants.xlCenter;
                            excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
                            excelCells.Borders.Weight = Excel.XlBorderWeight.xlThick;
                            excelCells.Value2 = "Возрастные группы";

                            /*-----Описываем ячейку B1 на странице-----*/
                            excelCells = excelWorksheet.get_Range("B1");
                            excelCells.VerticalAlignment = Excel.Constants.xlCenter;
                            excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
                            excelCells.Borders.Weight = Excel.XlBorderWeight.xlThick;
                            excelCells.Value2 = "Внешнее облучение";

                            /*-----Описываем ячейку C1 на странице-----*/
                            excelCells = excelWorksheet.get_Range("C1");
                            excelCells.VerticalAlignment = Excel.Constants.xlCenter;
                            excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
                            excelCells.Borders.Weight = Excel.XlBorderWeight.xlThick;
                            excelCells.Value2 = "Внутреннее облучение";

                            /*-----Описываем ячейку B1 на странице-----*/
                            excelCells = excelWorksheet.get_Range("D1");
                            excelCells.VerticalAlignment = Excel.Constants.xlCenter;
                            excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
                            excelCells.Borders.Weight = Excel.XlBorderWeight.xlThick;
                            excelCells.Value2 = "Сумма";

                            /*-----Описываем ячейку C1 на странице-----*/
                            excelCells = excelWorksheet.get_Range("E1");
                            excelCells.VerticalAlignment = Excel.Constants.xlCenter;
                            excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
                            excelCells.Borders.Weight = Excel.XlBorderWeight.xlThick;
                            excelCells.Value2 = "Внешнее облучение (95%)";

                            /*-----Описываем ячейку C1 на странице-----*/
                            excelCells = excelWorksheet.get_Range("F1");
                            excelCells.VerticalAlignment = Excel.Constants.xlCenter;
                            excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
                            excelCells.Borders.Weight = Excel.XlBorderWeight.xlThick;
                            excelCells.Value2 = "Внутреннее облучение (95%)";

                            /*-----Описываем ячейку C1 на странице-----*/
                            excelCells = excelWorksheet.get_Range("G1");
                            excelCells.VerticalAlignment = Excel.Constants.xlCenter;
                            excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
                            excelCells.Borders.Weight = Excel.XlBorderWeight.xlThick;
                            excelCells.Value2 = "Сумма (95%)";

                            for (int i = 2; i <= womanExtIbpo.Length + 1; i++)
                            {
                                excelCells = (Excel.Range)excelWorksheet.Cells[i, "A"];
                                excelCells.Value2 = ageGroups[i - 2];
                                excelCells.Borders.ColorIndex = 1;

                                excelCells = (Excel.Range)excelWorksheet.Cells[i, "B"];
                                if (womanExtIbpo[i - 2] < 100)
                                    excelCells.Value2 = womanExtIbpo[i - 2];
                                else
                                    excelCells.Value2 = "Нет облучения";
                                excelCells.Borders.ColorIndex = 1;

                                excelCells = (Excel.Range)excelWorksheet.Cells[i, "C"];
                                if (womanIntIbpo[i - 2] < 100)
                                    excelCells.Value2 = womanIntIbpo[i - 2];
                                else
                                    excelCells.Value2 = "Нет облучения";
                                excelCells.Borders.ColorIndex = 1;

                                excelCells = (Excel.Range)excelWorksheet.Cells[i, "D"];
                                if (womanSumIbpo[i - 2] < 100)
                                    excelCells.Value2 = womanSumIbpo[i - 2];
                                else
                                    excelCells.Value2 = "Нет облучения";
                                excelCells.Borders.ColorIndex = 1;

                                excelCells = (Excel.Range)excelWorksheet.Cells[i, "E"];
                                if (womanExtIbpo95[i - 2] < 100)
                                    excelCells.Value2 = womanExtIbpo95[i - 2];
                                else
                                    excelCells.Value2 = "Нет облучения";
                                excelCells.Borders.ColorIndex = 1;

                                excelCells = (Excel.Range)excelWorksheet.Cells[i, "F"];
                                if (womanIntIbpo95[i - 2] < 100)
                                    excelCells.Value2 = womanIntIbpo95[i - 2];
                                else
                                    excelCells.Value2 = "Нет облучения";
                                excelCells.Borders.ColorIndex = 1;

                                excelCells = (Excel.Range)excelWorksheet.Cells[i, "G"];
                                if (womanSumIbpo95[i - 2] < 100) 
                                    excelCells.Value2 = womanSumIbpo95[i - 2];
                                else
                                    excelCells.Value2 = "Нет облучения";
                                excelCells.Borders.ColorIndex = 1;
                            }

                            /*-----Описываем ячейку А1 на странице-----*/
                            excelCells = excelWorksheet.get_Range("A" + 13);
                            excelCells.VerticalAlignment = Excel.Constants.xlCenter;
                            excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
                            excelCells.Borders.Weight = Excel.XlBorderWeight.xlMedium;
                            excelCells.Value2 = "Взвешенные величины";

                            /*-----Описываем ячейку А1 на странице-----*/
                            excelCells = excelWorksheet.get_Range("B" + 13);
                            excelCells.VerticalAlignment = Excel.Constants.xlCenter;
                            excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
                            excelCells.Borders.Weight = Excel.XlBorderWeight.xlMedium;
                            if (womanWeightedExtIbpo.Sum() / womanRecordsList.Count < 100)
                                excelCells.Value2 = Math.Round(womanWeightedExtIbpo.Sum() / womanRecordsList.Count, 8);
                            else
                                excelCells.Value2 = "Нет облучения";

                            /*-----Описываем ячейку А1 на странице-----*/
                            excelCells = excelWorksheet.get_Range("C" + 13);
                            excelCells.VerticalAlignment = Excel.Constants.xlCenter;
                            excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
                            excelCells.Borders.Weight = Excel.XlBorderWeight.xlMedium;
                            if (womanWeightedIntIbpo.Sum() / womanRecordsList.Count < 100)
                                excelCells.Value2 = Math.Round(womanWeightedIntIbpo.Sum() / womanRecordsList.Count, 8);
                            else
                                excelCells.Value2 = "Нет облучения";

                            /*-----Описываем ячейку А1 на странице-----*/
                            excelCells = excelWorksheet.get_Range("D" + 13);
                            excelCells.VerticalAlignment = Excel.Constants.xlCenter;
                            excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
                            excelCells.Borders.Weight = Excel.XlBorderWeight.xlMedium;
                            if (womanWeightedSumIbpo.Sum() / womanRecordsList.Count < 100)
                                excelCells.Value2 = Math.Round(womanWeightedSumIbpo.Sum() / womanRecordsList.Count, 8);
                            else
                                excelCells.Value2 = "Нет облучения";

                            /*-----Описываем ячейку А1 на странице-----*/
                            excelCells = excelWorksheet.get_Range("E" + 13);
                            excelCells.VerticalAlignment = Excel.Constants.xlCenter;
                            excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
                            excelCells.Borders.Weight = Excel.XlBorderWeight.xlMedium;
                            if (womanWeightedExtIbpo95.Sum() / womanRecordsList.Count < 100)
                                excelCells.Value2 = Math.Round(womanWeightedExtIbpo95.Sum() / womanRecordsList.Count, 8);
                            else
                                excelCells.Value2 = "Нет облучения";

                            /*-----Описываем ячейку А1 на странице-----*/
                            excelCells = excelWorksheet.get_Range("F" + 13);
                            excelCells.VerticalAlignment = Excel.Constants.xlCenter;
                            excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
                            excelCells.Borders.Weight = Excel.XlBorderWeight.xlMedium;
                            if (womanWeightedIntIbpo95.Sum() / womanRecordsList.Count < 100)
                                excelCells.Value2 = Math.Round(womanWeightedIntIbpo95.Sum() / womanRecordsList.Count, 8);
                            else
                                excelCells.Value2 = "Нет облучения";

                            /*-----Описываем ячейку А1 на странице-----*/
                            excelCells = excelWorksheet.get_Range("G" + 13);
                            excelCells.VerticalAlignment = Excel.Constants.xlCenter;
                            excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
                            excelCells.Borders.Weight = Excel.XlBorderWeight.xlMedium;
                            if (womanWeightedSumIbpo95.Sum() / womanRecordsList.Count < 100)
                                excelCells.Value2 = Math.Round(womanWeightedSumIbpo95.Sum() / womanRecordsList.Count, 8);
                            else
                                excelCells.Value2 = "Нет облучения";

                            char[] timeNameBuffer = DateTime.Now.ToString().ToCharArray();
                            for (int i = 0; i < timeNameBuffer.Length; i++)
                            {
                                if (timeNameBuffer[i] == ':')
                                    timeNameBuffer[i] = '-';
                            }

                            if (larRB.Checked)
                                saveAs = " ИБПО_LAR";
                            if (detRB.Checked)
                                saveAs = " ИБПО_Det";

                            if (orpoButtonAverAge)
                            {
                                bufferPath = outPath + "\\ИБПО (Средний возраст)";
                                Directory.CreateDirectory(bufferPath);
                                excelWorkbook.SaveAs(@bufferPath + "\\" + shopComboBox.SelectedItem + saveAs + " (Средний возраст)" + "(" + new string(timeNameBuffer) + ").xlsx",  //object Filename
                                        Excel.XlFileFormat.xlOpenXMLWorkbook,                       //object FileFormat
                                        Type.Missing,                       //object Password 
                                        Type.Missing,                       //object WriteResPassword  
                                        Type.Missing,                       //object ReadOnlyRecommended
                                        Type.Missing,                       //object CreateBackup
                                        Excel.XlSaveAsAccessMode.xlNoChange,//XlSaveAsAccessMode AccessMode
                                        Type.Missing,                       //object ConflictResolution
                                        Type.Missing,                       //object AddToMru 
                                        Type.Missing,                       //object TextCodepage
                                        Type.Missing,                       //object TextVisualLayout
                                        Type.Missing);                      //object Local
                            }
                            if (orpoButtonAverLar)
                            {
                                bufferPath = outPath + "\\ИБПО (Средний LAR(Det))";
                                Directory.CreateDirectory(bufferPath);
                                excelWorkbook.SaveAs(@bufferPath + "\\" + shopComboBox.SelectedItem + saveAs + " (Средний LAR(Det))" + "(" + new string(timeNameBuffer) + ").xlsx",  //object Filename
                                        Excel.XlFileFormat.xlOpenXMLWorkbook,                       //object FileFormat
                                        Type.Missing,                       //object Password 
                                        Type.Missing,                       //object WriteResPassword  
                                        Type.Missing,                       //object ReadOnlyRecommended
                                        Type.Missing,                       //object CreateBackup
                                        Excel.XlSaveAsAccessMode.xlNoChange,//XlSaveAsAccessMode AccessMode
                                        Type.Missing,                       //object ConflictResolution
                                        Type.Missing,                       //object AddToMru 
                                        Type.Missing,                       //object TextCodepage
                                        Type.Missing,                       //object TextVisualLayout
                                        Type.Missing);                      //object Local
                            }
                            excelApp.Quit();
                        }

                        connection.Close();
                        shopTrigger = false;
                        ///*-----Замер времени работы кнопки-----*/
                        //stopWatch.Stop();
                        //TimeSpan ts = stopWatch.Elapsed;
                        //string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
                        //ts.Hours, ts.Minutes, ts.Seconds,
                        //ts.Milliseconds / 10);
                        //womanExtIbpoBox95.Text = "Время " + elapsedTime;
                    }
                    catch
                    {
                        MessageBox.Show("ОРПО не посчитано!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
                        Application.DoEvents();
                    }
                }
            }
            catch
            {
                MessageBox.Show("Нет связи с базой данных! Подключите базу!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
                Application.DoEvents();
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}

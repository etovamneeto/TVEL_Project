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
using Excel = Microsoft.Office.Interop.Excel;
//Хреновины имеют значение, когда сравниваешь строки и строки, т.е. когда тебе надо ввести в текстовую ячейку текстовый параметр; В другие, где инты или флоаты, туда не нужны они
//"UPDATE [names] SET [userName]='" + nameBox.Text + "', [age]=" + Convert.ToInt32(ageBox.Text), но + " WHERE [id]=" + Convert.ToInt32(table.Rows[0]["id"]), connection);

namespace TVELtest
{
    public partial class Form1 : Form
    {
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

        /*-----Описание форм инициализации и инициализация библиотеки с рейтами 2012 года-----*/
        public Form1(String title)
        {
            InitializeComponent();
            this.Text = title;

            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
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

        public double getIbpo(List<double> groupedLar, double orpo)
        {
            //womanIntIbpo[i] = 100 / (1 + womanIntOrpo[i] / (2 * Math.Pow(10, -4) * (1 - ((womanLarIntArray[i].Sum() / womanLarIntArray[i].Count) / (4.1 * Math.Pow(10, -2))))));
            double r = groupedLar.Average();
            double q = 1 - r / (4.1 * Math.Pow(10, -2));
            double denominator = 1 + orpo / (2 * Math.Pow(10, -4) * q);
            return 100 / denominator;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            RiskCalculatorLib.RiskCalculator.FillData(ref libPath);
            manExtIbpoBox.Text = "0";
            manIntIbpoBox.Text = "0";
        }

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
        /*-----Массивы, хранящие ОРПО для половозрастных групп-----*/
        double[] manExtOrpo = null;
        double[] manIntOrpo = null;
        double[] womanExtOrpo = null;
        double[] womanIntOrpo = null;

        double[] manExtOrpo_95 = null;
        double[] manIntOrpo_95 = null;
        double[] womanExtOrpo_95 = null;
        double[] womanIntOrpo_95 = null;

        private void openFileButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                dbPath = ofd.FileName;
                connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + dbPath;
            }
        }

        private void getOrpoButton_Click(object sender, EventArgs e)
        {
            /*-----Инициализация всяких входных параметров, подключения к БД, парсинга в таблицу нужных столбцов-----*/
            OleDbConnection connection = new OleDbConnection(connectionString);
            try
            {
                connection.Open();
                OleDbDataAdapter adapter = new OleDbDataAdapter("SELECT [ID], [Dose], [Year], [DoseInt], [Gender], [AgeAtExp] FROM [Final] WHERE [Shop]='r3'", connectionString);//Выбор нужных столбцов из нужной таблицы
                DataSet dataSet = new DataSet();
                adapter.Fill(dataSet, "Final");
                DataTable table = dataSet.Tables[0];//Из Final в эту таблицу считываются поля, указанные в запросе; Выборка для МСК (shop = r3)

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

                /*-----Список объектов; достаем все необходимое для расчетов: id, dose, doseInt, ageAtExp, gender-----*/
                dbFinalRecords = new List<dbObject>();
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    dbFinalRecords.Add(new dbObject(Convert.ToInt32(table.Rows[i]["id"]), Convert.ToByte(table.Rows[i]["gender"]), Convert.ToInt32(table.Rows[i]["year"]), Convert.ToInt16(table.Rows[i]["ageatexp"]), Convert.ToDouble(table.Rows[i]["dose"]) / 1000, Convert.ToDouble(table.Rows[i]["doseint"]) / 1000));
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

                /*-----Инициализация массиво, хранящих ОРПО для половозрастных групп-----*/
                manExtOrpo = new double[ageGroups.Count];
                manIntOrpo = new double[ageGroups.Count];
                womanExtOrpo = new double[ageGroups.Count];
                womanIntOrpo = new double[ageGroups.Count];

                manExtOrpo_95 = new double[ageGroups.Count];
                manIntOrpo_95 = new double[ageGroups.Count];
                womanExtOrpo_95 = new double[ageGroups.Count];
                womanIntOrpo_95 = new double[ageGroups.Count];

                for (int i = 0; i < ageGroups.Count; i++)
                {
                    if (manSadExtArray[i].Count > 0)
                    {
                        manExtOrpo[i] = getOrpo(getManExtLar(manYearsArray[i].Average()), manSadExtArray[i].Average());
                        manExtOrpo_95[i] = getOrpo_95(getManExtLar(manYearsArray[i].Average()), manSadExtArray[i].Average(), getDeviation(manSadExtArray[i]));
                    }

                    if (manSadIntArray[i].Count > 0)
                    {
                        manIntOrpo[i] = getOrpo(getManIntLar(manYearsArray[i].Average()), manSadIntArray[i].Average());
                        manIntOrpo_95[i] = getOrpo_95(getManIntLar(manYearsArray[i].Average()), manSadIntArray[i].Average(), getDeviation(manSadIntArray[i]));
                    }

                    if (womanSadExtArray[i].Count > 0)
                    {
                        womanExtOrpo[i] = getOrpo(getWomanExtLar(womanYearsArray[i].Average()), womanSadExtArray[i].Average());
                        womanExtOrpo_95[i] = getOrpo_95(getWomanExtLar(womanYearsArray[i].Average()), womanSadExtArray[i].Average(), getDeviation(womanSadExtArray[i]));
                    }

                    if (womanSadIntArray[i].Count > 0)
                    {
                        womanIntOrpo[i] = getOrpo(getWomanIntLar(womanYearsArray[i].Average()), womanSadIntArray[i].Average());
                        womanIntOrpo_95[i] = getOrpo_95(getWomanIntLar(womanYearsArray[i].Average()), womanSadIntArray[i].Average(), getDeviation(womanSadIntArray[i]));
                    }
                }

                List<double> manWeightedExtOrpo = new List<double>();
                List<double> manWeightedIntOrpo = new List<double>();
                List<double> womanWeightedExtOrpo = new List<double>();
                List<double> womanWeightedIntOrpo = new List<double>();

                List<double> manWeightedExtOrpo_95 = new List<double>();
                List<double> manWeightedIntOrpo_95 = new List<double>();
                List<double> womanWeightedExtOrpo_95 = new List<double>();
                List<double> womanWeightedIntOrpo_95 = new List<double>();
                for (int i = 0; i < ageGroups.Count; i++)
                {
                    manWeightedExtOrpo.Add(manExtOrpo[i] * manSadExtArray[i].Count);
                    manWeightedIntOrpo.Add(manIntOrpo[i] * manSadIntArray[i].Count);
                    womanWeightedExtOrpo.Add(womanExtOrpo[i] * womanSadExtArray[i].Count);
                    womanWeightedIntOrpo.Add(womanIntOrpo[i] * womanSadIntArray[i].Count);

                    manWeightedExtOrpo_95.Add(manExtOrpo_95[i] * manSadExtArray[i].Count);
                    manWeightedIntOrpo_95.Add(manIntOrpo_95[i] * manSadIntArray[i].Count);
                    womanWeightedExtOrpo_95.Add(womanExtOrpo_95[i] * womanSadExtArray[i].Count);
                    womanWeightedIntOrpo_95.Add(womanIntOrpo_95[i] * womanSadIntArray[i].Count);
                }

                manExtOrpoBox.Text = manSadIntArray.Length.ToString();
                manIntOrpoBox.Text = ageGroups.Count.ToString();

                //manExtOrpoBox.Text = (manWeightedExtOrpo.Sum() / dbMan).ToString();
                //manIntOrpoBox.Text = (manWeightedIntOrpo.Sum() / dbMan).ToString();
                //womanExtOrpoBox.Text = (womanWeightedExtOrpo.Sum() / dbWoman).ToString();
                //womanIntOrpoBox.Text = (womanWeightedIntOrpo.Sum() / dbWoman).ToString();

                //manExtOrpoBox95.Text = (manWeightedExtOrpo_95.Sum() / dbMan).ToString();
                //manIntOrpoBox95.Text = (manWeightedIntOrpo_95.Sum() / dbMan).ToString();
                //womanExtOrpoBox95.Text = (womanWeightedExtOrpo_95.Sum() / dbWoman).ToString();
                //womanIntOrpoBox95.Text = (womanWeightedIntOrpo_95.Sum() / dbWoman).ToString();


                ///*-----Вывод в Excel-файл-----*/
                ///*-----Инициализация Excel-файла-----*/
                //Excel.Application excelApp = new Excel.Application();
                ////excelApp.Visible = true;
                ////excelApp.DisplayAlerts = true;
                //excelApp.StandardFont = "Times-New-Roman";
                //excelApp.StandardFontSize = 12;

                ///*-----Создание рабочей книги с 4 страницами, в которые будет выводиться информация-----*/
                //excelApp.Workbooks.Add(Type.Missing);
                //Excel.Workbook excelWorkbook = excelApp.Workbooks[1];
                //excelApp.SheetsInNewWorkbook = 4;
                //Excel.Worksheet excelWorksheet = null;
                //Excel.Range excelCells = null;

                ///*-----Вывод в столбцы-----*/
                //excelWorksheet = (Excel.Worksheet)excelWorkbook.Worksheets.get_Item(1);
                //excelWorksheet.Name = "Мужчины, ОРПО внеш.";

                ///*-----Описываем ячейку А1 на странице-----*/
                //excelCells = excelWorksheet.get_Range("A1");
                //excelCells.VerticalAlignment = Excel.Constants.xlCenter;
                //excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
                //excelCells.Borders.Weight = Excel.XlBorderWeight.xlThick;
                //excelCells.Value2 = "Возрастные группы";

                ///*-----Описываем ячейку B1 на странице-----*/
                //excelCells = excelWorksheet.get_Range("B1");
                //excelCells.VerticalAlignment = Excel.Constants.xlCenter;
                //excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
                //excelCells.Borders.Weight = Excel.XlBorderWeight.xlThick;
                //excelCells.Value2 = "ОРПО";

                ///*-----Описываем ячейку C1 на странице-----*/
                //excelCells = excelWorksheet.get_Range("C1");
                //excelCells.VerticalAlignment = Excel.Constants.xlCenter;
                //excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
                //excelCells.Borders.Weight = Excel.XlBorderWeight.xlThick;
                //excelCells.Value2 = "ОРПО_95";

                //for (int i = 2; i <= manExtOrpo.Length + 1; i++)
                //{
                //    excelCells = (Excel.Range)excelWorksheet.Cells[i, "A"];
                //    excelCells.Value2 = ageGroups[i - 2];
                //    excelCells.Borders.ColorIndex = 1;
                //    excelCells = (Excel.Range)excelWorksheet.Cells[i, "B"];
                //    excelCells.Value2 = manExtOrpo[i - 2];
                //    excelCells.Borders.ColorIndex = 1;
                //    excelCells = (Excel.Range)excelWorksheet.Cells[i, "C"];
                //    excelCells.Value2 = manExtOrpo_95[i - 2];
                //    excelCells.Borders.ColorIndex = 1;
                //}

                //excelWorksheet = (Excel.Worksheet)excelWorkbook.Worksheets.get_Item(2);
                //excelWorksheet.Name = "Мужчины, ОРПО внут.";

                ///*-----Описываем ячейку А1 на странице-----*/
                //excelCells = excelWorksheet.get_Range("A1");
                //excelCells.VerticalAlignment = Excel.Constants.xlCenter;
                //excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
                //excelCells.Borders.Weight = Excel.XlBorderWeight.xlThick;
                //excelCells.Value2 = "Возрастные группы";

                ///*-----Описываем ячейку B1 на странице-----*/
                //excelCells = excelWorksheet.get_Range("B1");
                //excelCells.VerticalAlignment = Excel.Constants.xlCenter;
                //excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
                //excelCells.Borders.Weight = Excel.XlBorderWeight.xlThick;
                //excelCells.Value2 = "ОРПО";

                ///*-----Описываем ячейку C1 на странице-----*/
                //excelCells = excelWorksheet.get_Range("C1");
                //excelCells.VerticalAlignment = Excel.Constants.xlCenter;
                //excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
                //excelCells.Borders.Weight = Excel.XlBorderWeight.xlThick;
                //excelCells.Value2 = "ОРПО_95";

                //for (int i = 2; i <= manIntOrpo.Length + 1; i++)
                //{
                //    excelCells = (Excel.Range)excelWorksheet.Cells[i, "A"];
                //    excelCells.Value2 = ageGroups[i - 2];
                //    excelCells.Borders.ColorIndex = 1;
                //    excelCells = (Excel.Range)excelWorksheet.Cells[i, "B"];
                //    excelCells.Value2 = manIntOrpo[i - 2];
                //    excelCells.Borders.ColorIndex = 1;
                //    excelCells = (Excel.Range)excelWorksheet.Cells[i, "C"];
                //    excelCells.Value2 = manIntOrpo_95[i - 2];
                //    excelCells.Borders.ColorIndex = 1;

                //}

                //excelWorksheet = (Excel.Worksheet)excelWorkbook.Worksheets.get_Item(3);
                //excelWorksheet.Name = "Женщины, ОРПО внеш.";

                ///*-----Описываем ячейку А1 на странице-----*/
                //excelCells = excelWorksheet.get_Range("A1");
                //excelCells.VerticalAlignment = Excel.Constants.xlCenter;
                //excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
                //excelCells.Borders.Weight = Excel.XlBorderWeight.xlThick;
                //excelCells.Value2 = "Возрастные группы";

                ///*-----Описываем ячейку B1 на странице-----*/
                //excelCells = excelWorksheet.get_Range("B1");
                //excelCells.VerticalAlignment = Excel.Constants.xlCenter;
                //excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
                //excelCells.Borders.Weight = Excel.XlBorderWeight.xlThick;
                //excelCells.Value2 = "ОРПО";

                ///*-----Описываем ячейку C1 на странице-----*/
                //excelCells = excelWorksheet.get_Range("C1");
                //excelCells.VerticalAlignment = Excel.Constants.xlCenter;
                //excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
                //excelCells.Borders.Weight = Excel.XlBorderWeight.xlThick;
                //excelCells.Value2 = "ОРПО_95";

                //for (int i = 2; i <= womanExtOrpo.Length + 1; i++)
                //{
                //    excelCells = (Excel.Range)excelWorksheet.Cells[i, "A"];
                //    excelCells.Value2 = ageGroups[i - 2];
                //    excelCells.Borders.ColorIndex = 1;
                //    excelCells = (Excel.Range)excelWorksheet.Cells[i, "B"];
                //    excelCells.Value2 = womanExtOrpo[i - 2];
                //    excelCells.Borders.ColorIndex = 1;
                //    excelCells = (Excel.Range)excelWorksheet.Cells[i, "C"];
                //    excelCells.Value2 = womanExtOrpo_95[i - 2];
                //    excelCells.Borders.ColorIndex = 1;
                //}

                //excelWorksheet = (Excel.Worksheet)excelWorkbook.Worksheets.get_Item(4);
                //excelWorksheet.Name = "Женщины, ОРПО внут.";

                ///*-----Описываем ячейку А1 на странице-----*/
                //excelCells = excelWorksheet.get_Range("A1");
                //excelCells.VerticalAlignment = Excel.Constants.xlCenter;
                //excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
                //excelCells.Borders.Weight = Excel.XlBorderWeight.xlThick;
                //excelCells.Value2 = "Возрастные группы";

                ///*-----Описываем ячейку B1 на странице-----*/
                //excelCells = excelWorksheet.get_Range("B1");
                //excelCells.VerticalAlignment = Excel.Constants.xlCenter;
                //excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
                //excelCells.Borders.Weight = Excel.XlBorderWeight.xlThick;
                //excelCells.Value2 = "ОРПО";

                ///*-----Описываем ячейку C1 на странице-----*/
                //excelCells = excelWorksheet.get_Range("C1");
                //excelCells.VerticalAlignment = Excel.Constants.xlCenter;
                //excelCells.HorizontalAlignment = Excel.Constants.xlCenter;
                //excelCells.Borders.Weight = Excel.XlBorderWeight.xlThick;
                //excelCells.Value2 = "ОРПО_95";

                //for (int i = 2; i <= womanIntOrpo.Length + 1; i++)
                //{
                //    excelCells = (Excel.Range)excelWorksheet.Cells[i, "A"];
                //    excelCells.Value2 = ageGroups[i - 2];
                //    excelCells.Borders.ColorIndex = 1;
                //    excelCells = (Excel.Range)excelWorksheet.Cells[i, "B"];
                //    excelCells.Value2 = womanIntOrpo[i - 2];
                //    excelCells.Borders.ColorIndex = 1;
                //    excelCells = (Excel.Range)excelWorksheet.Cells[i, "C"];
                //    excelCells.Value2 = womanIntOrpo_95[i - 2];
                //    excelCells.Borders.ColorIndex = 1;
                //}

                //char[] timeNameBuffer = DateTime.Now.ToString().ToCharArray();
                //for (int i = 0; i < timeNameBuffer.Length; i++)
                //{
                //    if (timeNameBuffer[i] == ':')
                //        timeNameBuffer[i] = '-';
                //}

                //excelWorkbook.SaveAs(@Path.GetDirectoryName(Application.ExecutablePath) + "\\ОРПО" + "(" + new string(timeNameBuffer) + ").xlsx",  //object Filename
                //        Excel.XlFileFormat.xlOpenXMLWorkbook,                       //object FileFormat
                //        Type.Missing,                       //object Password 
                //        Type.Missing,                       //object WriteResPassword  
                //        Type.Missing,                       //object ReadOnlyRecommended
                //        Type.Missing,                       //object CreateBackup
                //        Excel.XlSaveAsAccessMode.xlNoChange,//XlSaveAsAccessMode AccessMode
                //        Type.Missing,                       //object ConflictResolution
                //        Type.Missing,                       //object AddToMru 
                //        Type.Missing,                       //object TextCodepage
                //        Type.Missing,                       //object TextVisualLayout
                //        Type.Missing);                      //object Local
                //excelApp.Quit();
                connection.Close();
            }

            catch/*(OleDbException ex)*/
            {
                MessageBox.Show("Не выбрана база данных!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
                Application.DoEvents();
            }
        }

        private void getIbpoButton_Click(object sender, EventArgs e)
        {
            OleDbConnection connection = new OleDbConnection(connectionString);
            //try
            {
                connection.Open();
                OleDbDataAdapter adapter = new OleDbDataAdapter("SELECT [ID], [Year], [Dose], [DoseInt] FROM [Dose]", connectionString);
                DataSet dataSet = new DataSet();
                adapter.Fill(dataSet, "Dose");
                DataTable table = dataSet.Tables[0];

                //try
                {
                    /*-----Список объектов, хранящих данные из таблицы Dose-----*/
                    dbDoseRecords = new List<dbObject>();
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        dbDoseRecords.Add(new dbObject(Convert.ToInt32(table.Rows[i]["id"]), Convert.ToInt32(table.Rows[i]["year"]), Convert.ToDouble(table.Rows[i]["dose"]) / 1000, Convert.ToDouble(table.Rows[i]["doseint"]) / 1000));
                    }

                    /*-----Списки id, у которых есть записи в 2012 году-----*/
                    List<dbObject> manIbpoIdList = new List<dbObject>();
                    List<dbObject> womanIbpoIdList = new List<dbObject>();
                    for (int i = 0; i < dbFinalRecords.Count; i++)
                    {
                        if (dbFinalRecords[i].getSex() == sexMale && dbFinalRecords[i].getYear() == 2012)
                            manIbpoIdList.Add(dbFinalRecords[i]);
                        if (dbFinalRecords[i].getSex() == sexFemale && dbFinalRecords[i].getYear() == 2012)
                            womanIbpoIdList.Add(dbFinalRecords[i]);
                    }

                    List<dbObject>[] manIbpoArray = new List<dbObject>[manIbpoIdList.Count];
                    for (int i = 0; i < manIbpoArray.Length; i++)
                        manIbpoArray[i] = new List<dbObject>();
                    List<dbObject>[] womanIbpoArray = new List<dbObject>[womanIbpoIdList.Count];
                    for (int i = 0; i < womanIbpoArray.Length; i++)
                        womanIbpoArray[i] = new List<dbObject>();

                    List<dbObject> buffer = null;
                    for (int i = 0; i < manIbpoArray.Length; i++)
                    {
                        buffer = new List<dbObject>();
                        for (int k = 0; k < dbDoseRecords.Count; k++)
                        {
                            if (manIbpoIdList[i].getId() == dbDoseRecords[k].getId())
                            {
                                dbDoseRecords[k].setSex(sexMale);
                                dbDoseRecords[k].setAgeAtExp(Convert.ToInt16(manIbpoIdList[i].getAgeAtExp() - (manIbpoIdList[i].getYear() - dbDoseRecords[k].getYear())));
                                buffer.Add(dbDoseRecords[k]);
                            }
                        }
                        manIbpoArray[i] = buffer;
                    }
                    for (int i = 0; i < womanIbpoArray.Length; i++)
                    {
                        buffer = new List<dbObject>();
                        for (int k = 0; k < dbDoseRecords.Count; k++)
                        {
                            if (womanIbpoIdList[i].getId() == dbDoseRecords[k].getId())
                            {
                                dbDoseRecords[k].setSex(sexFemale);
                                dbDoseRecords[k].setAgeAtExp(Convert.ToInt16(womanIbpoIdList[i].getAgeAtExp() - (womanIbpoIdList[i].getYear() - dbDoseRecords[k].getYear())));
                                buffer.Add(dbDoseRecords[k]);
                            }
                        }
                        womanIbpoArray[i] = buffer;
                    }

                    /*-----Задание весовых коэффициентов для тканей (в нашем случае учитывается только влияние на лёгкие)-----*/
                    double wLung = 0.12;

                    /*-----Создание пустого списка дозовых историй мужчин; для каждого уникального ID своя дозовая история (по сути, это ячейки, которые надо заполнить)-----*/
                    List<RiskCalculator.DoseHistoryRecord[]> manDoseHistoryList = new List<RiskCalculator.DoseHistoryRecord[]>();
                    for (int i = 0; i < manIbpoArray.Length; i++)
                    {
                        manDoseHistoryList.Add(new RiskCalculator.DoseHistoryRecord[manIbpoArray[i].Count]);
                    }
                    foreach (RiskCalculator.DoseHistoryRecord[] note in manDoseHistoryList)
                    {
                        for (int i = 0; i < note.Length; i++)
                            note[i] = new RiskCalculator.DoseHistoryRecord();
                    }

                    /*-----Создание аналогичного списка дозовых историй для женщин-----*/
                    List<RiskCalculator.DoseHistoryRecord[]> womanDoseHistoryList = new List<RiskCalculator.DoseHistoryRecord[]>();
                    for (int i = 0; i < womanIbpoArray.Length; i++)
                    {
                        womanDoseHistoryList.Add(new RiskCalculator.DoseHistoryRecord[womanIbpoArray[i].Count]);
                    }
                    foreach (RiskCalculator.DoseHistoryRecord[] note in womanDoseHistoryList)
                    {
                        for (int i = 0; i < note.Length; i++)
                            note[i] = new RiskCalculator.DoseHistoryRecord();
                    }
                    
                    /*-----Заполнение дозовых историй мужчин-----*/
                    for (int i = 0; i < manIbpoArray.Length; i++)
                        for (int k = 0; k < manIbpoArray[i].Count; k++)
                        {
                            manDoseHistoryList[i][k].AgeAtExposure = manIbpoArray[i][k].getAgeAtExp();
                            manDoseHistoryList[i][k].AllSolidDoseInmGy = manIbpoArray[i][k].getDose() - manIbpoArray[i][k].getDoseInt();
                            manDoseHistoryList[i][k].LeukaemiaDoseInmGy = manIbpoArray[i][k].getDose() - manIbpoArray[i][k].getDoseInt();
                            manDoseHistoryList[i][k].LungDoseInmGy = manIbpoArray[i][k].getDoseInt() / wLung;
                        }
                    /*-----Заполнение дозовых историй женщин-----*/
                    for (int i = 0; i < womanIbpoArray.Length; i++)
                        for (int k = 0; k < womanIbpoArray[i].Count; k++)
                        {
                            womanDoseHistoryList[i][k].AgeAtExposure = womanIbpoArray[i][k].getAgeAtExp();
                            womanDoseHistoryList[i][k].AllSolidDoseInmGy = womanIbpoArray[i][k].getDose() - womanIbpoArray[i][k].getDoseInt();
                            womanDoseHistoryList[i][k].LeukaemiaDoseInmGy = womanIbpoArray[i][k].getDose() - womanIbpoArray[i][k].getDoseInt();
                            womanDoseHistoryList[i][k].LungDoseInmGy = womanIbpoArray[i][k].getDoseInt() / wLung;
                        }

                    /*-----Здесь пример использования калькулятора-----*/
                    ////Создание словаря, где ключ - возраст, а значение - LAR
                    //Dictionary<short, double> ageLar = new Dictionary<short, double>();
                    //for (int i = 0; i <= ages; i++)
                    //{
                    //    RiskCalculator.DoseHistoryRecord[] record = listOfDoseHistories[i];
                    //    if (externalRB.Checked)
                    //    {
                    //        RiskCalculatorLib.RiskCalculator calculator = new RiskCalculatorLib.RiskCalculator(sex, listOfDoseHistories[i][0].AgeAtExposure, ref record, true);
                    //        ageLar.Add(listOfDoseHistories[i][0].AgeAtExposure, calculator.getLAR(false, true).AllCancers);
                    //        sheetName = sexName + " Внешнее";
                    //    }
                    //    else if (internalRB.Checked)
                    //    {
                    //        RiskCalculatorLib.RiskCalculator calculator = new RiskCalculatorLib.RiskCalculator(sex, listOfDoseHistories[i][0].AgeAtExposure, ref record, true);
                    //        ageLar.Add(listOfDoseHistories[i][0].AgeAtExposure, calculator.getLAR(false, true).Lung);
                    //        sheetName = sexName + " Внутреннее";
                    //    }
                    //}

                    

                    manExtIbpoBox95.Text = "Элементов " + manIbpoArray[Convert.ToInt32(manExtIbpoBox.Text)].Count.ToString();
                    manIntIbpoBox95.Text = "Элементы " + manIbpoArray[Convert.ToInt32(manExtIbpoBox.Text)][Convert.ToInt32(manIntIbpoBox.Text)].getDose().ToString();
                    womanExtIbpoBox.Text = "id " + manIbpoArray[Convert.ToInt32(manExtIbpoBox.Text)][Convert.ToInt32(manIntIbpoBox.Text)].getId().ToString();
                    womanIntIbpoBox.Text = "Длина " + manIbpoArray.Length.ToString();
                    womanExtIbpoBox95.Text = "Пол " + manIbpoArray[Convert.ToInt32(manExtIbpoBox.Text)][Convert.ToInt32(manIntIbpoBox.Text)].getSex().ToString();
                    womanIntIbpoBox95.Text = "ВозПриОб " + manIbpoArray[Convert.ToInt32(manExtIbpoBox.Text)][Convert.ToInt32(manIntIbpoBox.Text)].getAgeAtExp().ToString();

                    connection.Close();




                            /*-----Это годный вариант заполнения и разбиения на п/в группы, но как заполнять дозовые истории?!-----*/
                            ///*-----Массивы списков, в которых хранятся записи из базы Final для п/в групп в 2012 году-----*/
                            //List<dbObject>[] manIbpoArray = new List<dbObject>[ageGroups.Count];
                            //List<dbObject>[] womanIbpoArray = new List<dbObject>[ageGroups.Count];

                            //for (int i = 0; i < ageGroups.Count; i++)
                            //{
                            //    manIbpoArray[i] = new List<dbObject>();
                            //    womanIbpoArray[i] = new List<dbObject>();
                            //}

                            //for (int i = 0; i < ageGroups.Count; i++)
                            //    for (int k = 0; k < dbFinalRecords.Count; k++)
                            //    {
                            //        if (dbFinalRecords[k].getSex() == sexMale && dbFinalRecords[k].getYear() == 2012)
                            //            if (dbFinalRecords[k].getAgeAtExp() >= ageLowerBound[i] && dbFinalRecords[k].getAgeAtExp() <= ageUpperBound[i])
                            //            {
                            //                manIbpoArray[i].Add(dbFinalRecords[k]);
                            //            }
                            //        if (dbFinalRecords[k].getSex() == sexFemale && dbFinalRecords[k].getYear() == 2012)
                            //            if (dbFinalRecords[k].getAgeAtExp() >= ageLowerBound[i] && dbFinalRecords[k].getAgeAtExp() <= ageUpperBound[i])
                            //            {
                            //                womanIbpoArray[i].Add(dbFinalRecords[k]);
                            //            }
                            //    }

                            /*-----Сложный скелет для заполнения дозовых историй-----*/
                            //List<int> array = new List<int>();//Живучая идея для заполнения индивидуальных доз, но потом надо их по группам распихать
                            //int buffer = 0;
                            //for (int i = 0; i < ageGroups.Count; i++)
                            //{
                            //    for (int k = 0; k < manIbpoArray[i].Count; k++)
                            //    {
                            //        buffer = 0;
                            //        for (int m = 0; m < dbDoseRecords.Count; m++)
                            //        {
                            //            if (manIbpoArray[i][k].getId() == dbDoseRecords[i].getId())
                            //            {
                            //                buffer++;
                            //                //Здесь как-то должно быть столько массивов дозовых историй, сколько счетчик К пробегает
                            //                //А всего 11 групп, как всегда
                            //            }
                            //            array.Add(buffer);
                            //        }
                            //    }
                            //}
                            //int[] array = new int[manIbpoIdList.Count];//Живучая идея для заполнения индивидуальных доз, но потом надо их по группам распихать
                            //int buffer = 0;
                            //for (int i = 0; i < manIbpoIdList.Count; i++)
                            //{
                            //    buffer = 0;
                            //    for (int k = 0; k < dbDoseRecords.Count; k++)
                            //    {

                            //        if (manIbpoIdList[i] == dbDoseRecords[k].getId())
                            //        {
                            //            buffer++;
                            //        }
                            //        array[i] = buffer;
                            //    }
                            //}
                            //List<dbObject>[] manIbpoArray = new List<dbObject>[ageGroups.Count];
                            //List<dbObject>[] womanIbpoArray = new List<dbObject>[ageGroups.Count];

                            //for (int i = 0; i < ageGroups.Count; i++)
                            //{
                            //    manIbpoArray[i] = new List<dbObject>();
                            //    womanIbpoArray[i] = new List<dbObject>();
                            //}

                            //for (int i = 0; i < ageGroups.Count; i++)
                            //    for (int k = 0; k < dbRecords.Count; k++)
                            //    {
                            //        if (dbRecords[k].getSex() == sexMale && dbRecords[k].getYear() == 2012)
                            //            if (dbRecords[k].getAgeAtExp() >= ageLowerBound[i] && dbRecords[k].getAgeAtExp() <= ageUpperBound[i])
                            //            {
                            //                manIbpoArray[i].Add(dbRecords[k]);
                            //            }
                            //        if (dbRecords[k].getSex() == sexFemale && dbRecords[k].getYear() == 2012)
                            //            if (dbRecords[k].getAgeAtExp() >= ageLowerBound[i] && dbRecords[k].getAgeAtExp() <= ageUpperBound[i])
                            //            {
                            //                womanIbpoArray[i].Add(dbRecords[k]);
                            //            }
                            //    }




                    ///*-----Создание пустого списка дозовых историй мужчин; для каждого уникального ID своя дозовая история (по сути, это ячейки, которые надо заполнить)-----*/
                    //List<RiskCalculator.DoseHistoryRecord[]> manDoseHistoryList = new List<RiskCalculator.DoseHistoryRecord[]>();
                    //for (int i = 0; i < manIdRecordsArray.Length; i++)
                    //{
                    //    manDoseHistoryList.Add(new RiskCalculator.DoseHistoryRecord[manIdRecordsArray[i].Count]);
                    //}
                    //foreach (RiskCalculator.DoseHistoryRecord[] note in manDoseHistoryList)
                    //{
                    //    for (int i = 0; i < note.Length; i++)
                    //        note[i] = new RiskCalculator.DoseHistoryRecord();
                    //}

                    ///*-----Задание весовых коэффициентов для тканей (в нашем случае учитывается только влияние на лёгкие)-----*/
                    //double wLung = 0.12;

                    ///*-----Заполнение дозовых историй мужчин-----*/
                    //for (int i = 0; i < manIdRecordsArray.Length; i++)
                    //    for (int k = 0; k < manIdRecordsArray[i].Count; k++)
                    //    {
                    //        manDoseHistoryList[i][k].AgeAtExposure = manIdRecordsArray[i][k].getAgeAtExp();
                    //        manDoseHistoryList[i][k].AllSolidDoseInmGy = manIdRecordsArray[i][k].getDose() - manIdRecordsArray[i][k].getDoseInt();
                    //        manDoseHistoryList[i][k].LeukaemiaDoseInmGy = manIdRecordsArray[i][k].getDose() - manIdRecordsArray[i][k].getDoseInt();
                    //        manDoseHistoryList[i][k].LungDoseInmGy = manIdRecordsArray[i][k].getDoseInt() / wLung;
                    //    }

                    ///*-----Создание аналогичного списка дозовых историй для женщин-----*/
                    //List<RiskCalculator.DoseHistoryRecord[]> womanDoseHistoryList = new List<RiskCalculator.DoseHistoryRecord[]>();
                    //for (int i = 0; i < womanIdRecordsArray.Length; i++)
                    //{
                    //    womanDoseHistoryList.Add(new RiskCalculator.DoseHistoryRecord[womanIdRecordsArray[i].Count]);
                    //}
                    //foreach (RiskCalculator.DoseHistoryRecord[] note in womanDoseHistoryList)
                    //{
                    //    for (int i = 0; i < note.Length; i++)
                    //        note[i] = new RiskCalculator.DoseHistoryRecord();
                    //}

                    ///*-----Заполнение дозовых историй женщин-----*/
                    //for (int i = 0; i < womanIdRecordsArray.Length; i++)
                    //    for (int k = 0; k < womanIdRecordsArray[i].Count; k++)
                    //    {
                    //        womanDoseHistoryList[i][k].AgeAtExposure = womanIdRecordsArray[i][k].getAgeAtExp();
                    //        womanDoseHistoryList[i][k].AllSolidDoseInmGy = womanIdRecordsArray[i][k].getDose() - womanIdRecordsArray[i][k].getDoseInt();
                    //        womanDoseHistoryList[i][k].LeukaemiaDoseInmGy = womanIdRecordsArray[i][k].getDose() - womanIdRecordsArray[i][k].getDoseInt();
                    //        womanDoseHistoryList[i][k].LungDoseInmGy = womanIdRecordsArray[i][k].getDoseInt() / wLung;
                    //    }
                }
                //catch
                //{
                //    MessageBox.Show("ОРПО не посчитано!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
                //    Application.DoEvents();
                //}
            }
            //catch/*(OleDbException ex)*/
            //    {
            //        MessageBox.Show("Нет связи с базой данных! Подключите базу!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            //        Application.DoEvents();
            //    }
        }
    }
}

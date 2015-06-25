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

            public dbObject(int id, byte sex, int year, short ageAtExp, double dose, double doseInt)
            {
                this.id = id;
                this.sex = sex;
                this.year = year;
                this.ageAtExp = ageAtExp;
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
        }

        /*-----Список, в котором хранятся строковые параметры, инентифицирующие возрастные группы-----*/
        List<String> ageGroups = null;
        /*-----Список, в котором хранятся нижние границы возрастов для возрастных групп-----*/
        List<int> ageLowerBound = null;
        /*-----Список, в котором хранятся верхние границы возрастов для возрастных групп-----*/
        List<int> ageUpperBound = null;
        /*-----Список объектов; достаем все необходимое для расчетов: id, dose, doseInt, ageAtExp, gender-----*/
        List<dbObject> dbRecords = null;
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
                dbRecords = new List<dbObject>();
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    dbRecords.Add(new dbObject(Convert.ToInt32(table.Rows[i]["id"]), Convert.ToByte(table.Rows[i]["gender"]), Convert.ToInt32(table.Rows[i]["year"]), Convert.ToInt16(table.Rows[i]["ageatexp"]), Convert.ToDouble(table.Rows[i]["dose"]) / 1000, Convert.ToDouble(table.Rows[i]["doseint"]) / 1000));
                }

                /*-----Список, в котором хранится пол-----*/
                List<byte> dbSex = new List<byte>();
                for (int i = 0; i < dbRecords.Count; i++)
                    dbSex.Add(dbRecords[i].getSex());

                /*-----Определение пола; Меньшая цифра пола - М, большая - Ж-----*/
                sexMale = dbSex.Min();
                sexFemale = dbSex.Max();

                /*-----Счетчики, определяющие количество мужских и женских записей-----*/
                double dbMan = 0;
                for (int i = 0; i < dbRecords.Count; i++)
                    if (dbRecords[i].getSex() == sexMale)
                        dbMan++;

                double dbWoman = 0;
                for (int i = 0; i < dbRecords.Count; i++)
                    if (dbRecords[i].getSex() == sexFemale)
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
                    for (int k = 0; k < dbRecords.Count; k++)
                    {
                        if (dbRecords[k].getSex() == sexMale)
                            if (dbRecords[k].getAgeAtExp() >= ageLowerBound[i] && dbRecords[k].getAgeAtExp() <= ageUpperBound[i])
                            {
                                manSadExtArray[i].Add(dbRecords[k].getDose() - dbRecords[k].getDoseInt());
                                manSadIntArray[i].Add(dbRecords[k].getDoseInt());
                                manYearsArray[i].Add(dbRecords[k].getAgeAtExp());
                            }
                        if (dbRecords[k].getSex() == sexFemale)
                            if (dbRecords[k].getAgeAtExp() >= ageLowerBound[i] && dbRecords[k].getAgeAtExp() <= ageUpperBound[i])
                            {
                                womanSadExtArray[i].Add(dbRecords[k].getDose() - dbRecords[k].getDoseInt());
                                womanSadIntArray[i].Add(dbRecords[k].getDoseInt());
                                womanYearsArray[i].Add(dbRecords[k].getAgeAtExp());
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
            try
            {
                connection.Open();
                OleDbDataAdapter adapter = new OleDbDataAdapter("SELECT [ID], [Year], [Dose], [DoseInt] FROM [Dose]", connectionString);
                DataSet dataSet = new DataSet();
                adapter.Fill(dataSet, "Dose");
                DataTable table = dataSet.Tables[0];

                try
                {
                    /*-----Списки ID людей, у которых есть записи в 2012 году-----*/
                    List<dbObject> manIbpoId = new List<dbObject>();
                    List<dbObject> womanIbpoId = new List<dbObject>();

                    for (int i = 0; i < dbRecords.Count; i++)
                    {
                        if (dbRecords[i].getYear() == 2012)
                        {
                            if (dbRecords[i].getSex() == sexMale)
                                manIbpoId.Add(dbRecords[i]);
                            if (dbRecords[i].getSex() == sexFemale)
                                womanIbpoId.Add(dbRecords[i]);
                        }
                    }



                    manExtIbpoBox.Text = "Мужчинки " + manIbpoId.Count;
                    manIntIbpoBox.Text = "Тетьки " + womanIbpoId.Count;
                }
                catch
                {
                    MessageBox.Show("ОРПО не посчитано!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
                    Application.DoEvents();
                }
            }
            catch/*(OleDbException ex)*/
                {
                    MessageBox.Show("Нет связи с базой данных! Подключите базу!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
                    Application.DoEvents();
                }

        }

           

            ///*-----Список, в котором хранится пол-----*/
            //List<byte> dbSex = new List<byte>();
            //for (int i = 0; i < dbRecord.Count; i++)
            //    dbSex.Add(dbRecord[i].getSex());

            ///*-----Определение пола; Меньшая цифра пола - М, большая - Ж-----*/
            //byte sexMale = dbSex.Min();
            //byte sexFemale = dbSex.Max();

            ///*-----Уникальные ID мужчин-----*/
            //List<int> manUniqueIdList = new List<int>();
            //for (int i = 0; i < dbRecord.Count; i++)
            //{
            //    if (dbRecord[i].getSex() == sexMale && dbRecord[i].getYear() == 2012)
            //    //if (dbRecord[i].getSex() == sexMale)
            //    {
            //        manUniqueIdList.Add(dbRecord[i].getId());
            //    }
            //}
            //manUniqueIdList = manUniqueIdList.Distinct().ToList();

            ///*-----Уникальные ID женщин-----*/
            //List<int> womanUniqueIdList = new List<int>();
            //for (int i = 0; i < dbRecord.Count; i++)
            //{
            //    if (dbRecord[i].getSex() == sexFemale && dbRecord[i].getYear() == 2012)
            //    //if (dbRecord[i].getSex() == sexFemale)
            //    {
            //        womanUniqueIdList.Add(dbRecord[i].getId());
            //    }
            //}
            //womanUniqueIdList = womanUniqueIdList.Distinct().ToList();

            ///*-----Разделение записей БД на мужские и женские-----*/
            //List<dbObject> manList = new List<dbObject>();
            //for (int i = 0; i < dbRecord.Count; i++)
            //    if (dbRecord[i].getSex() == sexMale)
            //        manList.Add(dbRecord[i]);

            //List<dbObject> womanList = new List<dbObject>();
            //for (int i = 0; i < dbRecord.Count; i++)
            //    if (dbRecord[i].getSex() == sexFemale)
            //        womanList.Add(dbRecord[i]);

            ///*
            // * -----
            // * Создания массива списков, где каждый элемент
            // * массива - это список объектов, id которых
            // * совпадают с уникальными id; например, если уникальный id = 1,
            // * то в элемент массива списков записываются все объекты с id = 1.
            // * -----
            // */
            //List<dbObject>[] manIdRecordsArray = new List<dbObject>[manUniqueIdList.Count];
            //for (int i = 0; i < manIdRecordsArray.Length; i++)
            //    manIdRecordsArray[i] = new List<dbObject>();

            //for (int i = 0; i < manIdRecordsArray.Length; i++)
            //    for (int k = 0; k < manList.Count; k++)
            //    {
            //        if (Equals(manUniqueIdList[i], manList[k].getId()))
            //        {
            //            manIdRecordsArray[i].Add(manList[k]);
            //        }
            //    }

            ///*-----Создание аналогичного массива списков для женщин-----*/
            //List<dbObject>[] womanIdRecordsArray = new List<dbObject>[womanUniqueIdList.Count];
            //for (int i = 0; i < womanIdRecordsArray.Length; i++)
            //    womanIdRecordsArray[i] = new List<dbObject>();

            //for (int i = 0; i < womanIdRecordsArray.Length; i++)
            //    for (int k = 0; k < womanList.Count; k++)
            //    {
            //        if (Equals(womanUniqueIdList[i], womanList[k].getId()))
            //        {
            //            womanIdRecordsArray[i].Add(womanList[k]);
            //        }
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

            ///*-----Вычленение только тех членов персонала, что наблюдались включительно по 2012 год-----*/
            //List<double>[] manLarExtArray = new List<double>[ageGroups.Count];
            //List<double>[] manLarIntArray = new List<double>[ageGroups.Count];

            ///*-----Создание аналогичного массива списков LAR для возрастных групп женщин-----*/
            //List<double>[] womanLarExtArray = new List<double>[ageGroups.Count];
            //List<double>[] womanLarIntArray = new List<double>[ageGroups.Count];

            ///*-----Инициализация всех элементов массивов-----*/
            //for (int i = 0; i < ageGroups.Count; i++)
            //{
            //    manLarExtArray[i] = new List<double>();
            //    manLarIntArray[i] = new List<double>();
            //    womanLarExtArray[i] = new List<double>();
            //    womanLarIntArray[i] = new List<double>();
            //}

            //for (int i = 0; i < manIdRecordsArray.Length; i++)
            //    for (int k = 0; k < ageGroups.Count; k++)
            //        if (manIdRecordsArray[i][0].getAgeAtExp() >= ageLowerBound[k] && manIdRecordsArray[i][0].getAgeAtExp() <= ageUpperBound[k])
            //        {
            //            RiskCalculator.DoseHistoryRecord[] record = manDoseHistoryList[i];
            //            RiskCalculatorLib.RiskCalculator calculator = new RiskCalculatorLib.RiskCalculator(RiskCalculator.SEX_MALE, manIdRecordsArray[i][0].getAgeAtExp(), ref record, true);
            //            manLarExtArray[k].Add(calculator.getLAR(false, true).AllCancers);
            //            manLarIntArray[k].Add(calculator.getLAR(false, true).Lung);
            //        }

            //for (int i = 0; i < womanIdRecordsArray.Length; i++)
            //    for (int k = 0; k < ageGroups.Count; k++)
            //        if (womanIdRecordsArray[i][0].getAgeAtExp() >= ageLowerBound[k] && womanIdRecordsArray[i][0].getAgeAtExp() <= ageUpperBound[k])
            //        {
            //            RiskCalculator.DoseHistoryRecord[] record = womanDoseHistoryList[i];
            //            RiskCalculatorLib.RiskCalculator calculator = new RiskCalculatorLib.RiskCalculator(RiskCalculator.SEX_FEMALE, womanIdRecordsArray[i][0].getAgeAtExp(), ref record, true);
            //            womanLarExtArray[k].Add(calculator.getLAR(false, true).AllCancers);
            //            womanLarIntArray[k].Add(calculator.getLAR(false, true).Lung);
            //        }

            ///*-----Вычисление знаменателя R, используемого при расчете q, используемого в ИБПО-----*/
            //double[] manExtIbpo = new double[ageGroups.Count];
            //double[] manIntIbpo = new double[ageGroups.Count];
            //double[] womanExtIbpo = new double[ageGroups.Count];
            //double[] womanIntIbpo = new double[ageGroups.Count];

            //double[] manExtOrpo = new double[ageGroups.Count];//Чем заполняются эти ОРПО?
            //double[] manIntOrpo = new double[ageGroups.Count];
            //double[] womanExtOrpo = new double[ageGroups.Count];
            //double[] womanIntOrpo = new double[ageGroups.Count];

            //for (int i = 0; i < ageGroups.Count; i++)
            //{
            //    manExtIbpo[i] = getIbpo(manLarExtArray[i], manExtOrpo[i]);
            //    manIntIbpo[i] = getIbpo(manLarIntArray[i], manIntOrpo[i]);
            //    womanExtIbpo[i] = getIbpo(womanLarExtArray[i], womanExtOrpo[i]);
            //    womanIntIbpo[i] = getIbpo(womanLarIntArray[i], womanIntOrpo[i]);
            //}
       

    }
}

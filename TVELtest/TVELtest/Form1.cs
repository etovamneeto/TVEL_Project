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

        private void Form1_Load(object sender, EventArgs e)
        {
            RiskCalculatorLib.RiskCalculator.FillData(ref libPath);
            this.Width = 300;
            this.Height = 300;
            this.CenterToScreen();
        }

        /*-----Функции для расчета LAR, необходимых для расчета ОРПО*-----*/
        public double getManExtLar(double meanAge)
        {
            double lar = 0;
            double secondPowerElement = (2 / Math.Pow(10, 6)) * Math.Pow(meanAge, 2);
            double firstPowerElement = (-13 / Math.Pow(10,4)) * meanAge;
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
            List<int> uniqueIdList = new List<int>();
            for (int i = 0; i < table.Rows.Count; i++)
            {
                uniqueIdList.Add(Convert.ToInt32(table.Rows[i]["id"]));
            }
            uniqueIdList = uniqueIdList.Distinct().ToList();

            /*-----Список объектов; достаем все необходимое для расчетов: id, dose, doseInt, ageAtExp, gender-----*/
            List<dbObject> dbObjectList = new List<dbObject>();
            for (int i = 0; i < table.Rows.Count; i++)
            {
                dbObjectList.Add(new dbObject(Convert.ToInt32(table.Rows[i]["id"]), Convert.ToByte(table.Rows[i]["gender"]), Convert.ToInt16(table.Rows[i]["ageatexp"]), Convert.ToDouble(table.Rows[i]["dose"]), Convert.ToDouble(table.Rows[i]["doseint"])));
            }

            /*-----Создания массива списков, где каждый элемент массива - это список объектов, id которых совпадают с уникальными id;
             * ----например, если уникальный id = 1, то в элемент массива списков записываются все объекты с id = 1.
             */
            List<dbObject>[] manRecordsList = new List<dbObject>[uniqueIdList.Count];
            for (int i = 0; i < uniqueIdList.Count; i++)
            {
                List<dbObject> buffer = new List<dbObject>();
                for (int k = 0; k < dbObjectList.Count; k++)
                {
                    if (Equals(uniqueIdList[i], dbObjectList[k].getId()))
                    {
                        buffer.Add(dbObjectList[k]);
                    }
                    manRecordsList[i] = buffer;
                }
            }

            /*-----Создание пустого списка дозовых историй; для каждого уникального ID своя дозовая история (по сути, это ячейки, которые надо заполнить)-----*/
            List<RiskCalculator.DoseHistoryRecord[]> doseHistoryList = new List<RiskCalculator.DoseHistoryRecord[]>();
            for (int i = 0; i < uniqueIdList.Count; i++)
            {
                doseHistoryList.Add(new RiskCalculator.DoseHistoryRecord[manRecordsList[i].Count]);
            }
            foreach (RiskCalculator.DoseHistoryRecord[] note in doseHistoryList)
            {
                for (int i = 0; i < note.Length; i++)
                    note[i] = new RiskCalculator.DoseHistoryRecord();
            }

            /*-----Заполнение дозовых историй-----*/
            for (int i = 0; i < uniqueIdList.Count; i++)
                for (int k = 0; k < manRecordsList[i].Count; k++)
                {
                    doseHistoryList[i][k].AgeAtExposure = manRecordsList[i][k].getAgeAtExp();
                    doseHistoryList[i][k].AllSolidDoseInmGy = manRecordsList[i][k].getDose() - manRecordsList[i][k].getDoseInt();
                    doseHistoryList[i][k].LungDoseInmGy = manRecordsList[i][k].getDoseInt();
                }

            /*-----Заполнение списка людей, каждый определяется уникальным id, у него есть дозовая история и тд и тп-----*/ 
            List<Man> manList = new List<Man>();
            for (int i = 0; i < uniqueIdList.Count; i++)
            {
                manList.Add(new Man(uniqueIdList[i], manRecordsList[i][0].getSex(), manRecordsList[i][0].getAgeAtExp(), doseHistoryList[i]));
            }
            
            /*-----Закладываем формулы для расчета LAR-----*/
            double meanDoseInGroup = 0;//средняя доза за 5 лет в группе
            double meanAgeInGroup = 0;//средний возраст группы

            resultTextBox.Text = getWomanExtLar(2).ToString();
        }

        private void resultTextBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void testTextBox_TextChanged(object sender, EventArgs e)
        {

        }
    }
}

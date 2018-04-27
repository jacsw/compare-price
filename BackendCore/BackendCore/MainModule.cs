using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Xml;
using ExcelTest;

namespace ExcelTest
{
    class CapitalData
    {
        public string Name { get; set; }
        public string FileName { get; set; }
        public string Sheet { get; set; }
        public int GTotRow { get; set; }
        public int GTotCol { get; set; }
        public int RentRow { get; set; }
        public int RentCol { get; set; }
    }

    class MainModule
    {
        ExcelParser[] excels;
        int CompanyCount = 3;

        void init()
        {
            // TODO : Set by JSON File
            CapitalData[] CompanyInfo = new CapitalData[CompanyCount];

            CompanyInfo[0] = new CapitalData();
            CompanyInfo[0].Name = @"하나캐피탈";
            CompanyInfo[0].FileName = @"D:\Rent\Excel\HANA\HANA_20180310.xlsx";
            CompanyInfo[0].Sheet = @"(국산차)견적";
            CompanyInfo[0].GTotRow = 14;
            CompanyInfo[0].GTotCol = 59;
            CompanyInfo[0].RentRow = 30;
            CompanyInfo[0].RentCol = 8;

            CompanyInfo[1] = new CapitalData();
            CompanyInfo[1].Name = @"JB우리캐피탈";
            CompanyInfo[1].FileName = @"D:\Rent\Excel\JBWOORI\JBWOORI_20180421.xlsm";
            CompanyInfo[1].Sheet = @"국산";
            CompanyInfo[1].GTotRow = 16;
            CompanyInfo[1].GTotCol = 50;
            CompanyInfo[1].RentRow = 27;
            CompanyInfo[1].RentCol = 9;

            CompanyInfo[2] = new CapitalData();
            CompanyInfo[2].Name = @"효성캐피탈";
            CompanyInfo[2].FileName = @"D:\Rent\Excel\HYOSUNG\HYOSUNG_20180418-2.xlsx";
            CompanyInfo[2].Sheet = @"견적서및입력시트";
            CompanyInfo[2].GTotRow = 18;
            CompanyInfo[2].GTotCol = 11;
            CompanyInfo[2].RentRow = 20;
            CompanyInfo[2].RentCol = 46;

            excels = new ExcelParser[CompanyCount];
            for (int i = 0; i < CompanyCount; i++)
            {
                excels[i] = new ExcelParser(CompanyInfo[i].Name, CompanyInfo[i].FileName, CompanyInfo[i].Sheet);
                excels[i].SetGTotalPos(CompanyInfo[i].GTotRow, CompanyInfo[i].GTotCol);
                excels[i].SetRentCostPos(CompanyInfo[i].RentRow, CompanyInfo[i].RentCol);
                excels[i].Init();
            }
        }

        void run()
        {
            while (true)
            {
                System.Console.Write("자동차 비용 ( Exit = 0 ) :  ");
                int carcost = Convert.ToInt32(Console.ReadLine());
                if (carcost == 0)
                {
                    break;
                }

                for (int i = 0; i < CompanyCount; i++)
                {
                    System.Console.WriteLine("{0} : {1}", excels[i].getCompanyName(), excels[i].GetRentCost(carcost));
                }
            }
            System.Console.Write("아무내용이나 입력후 엔터 : ");
            System.Console.ReadLine();
        }

        void deinit()
        {
            for (int i = 0; i < CompanyCount; i++)
            {
                excels[i].Deinit();
            }
        }

        static void Main(string[] args)
        {
            ExcelParser.ModuleInit();
            MainModule main = new MainModule();
            main.init();
            main.run();
            main.deinit();

            ExcelParser.ModuleDeInit();
        }
    }
}

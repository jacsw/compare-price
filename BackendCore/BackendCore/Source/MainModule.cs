using BackendCore.Source;
using System;

namespace BackendCore
{
    class MainModule
    {
        static void Main(string[] args)
        {
            ExcelParser.ModuleInit();
            MainModule main = new MainModule();
            main.Init();
            main.Run();
            main.Deinit();

            ExcelParser.ModuleDeInit();
        }

        public void Init()
        {
            // TODO : Set by JSON File
            mConfig = new CapitalConfig(@"D:\Rent\Excel\CapitalList.json");
            mConfig.LoadConfig();

            mExcels = new ExcelParser[mConfig.GetCount()];
            for (int i = 0; i < mConfig.GetCount(); i++)
            {
                var info = mConfig.GetCapitalData(i);

                mExcels[i] = new ExcelParser(info.Com, info.File, info.Worksheet);
                mExcels[i].SetGTotalPos(info.Price.Row, info.Price.Col);
                mExcels[i].SetRentCostPos(info.Fee.M36.Row, info.Fee.M36.Col);
                mExcels[i].Init();
            }

            System.Console.WriteLine("--------------------------------------------------");
        }

        public void Run()
        {
            Console.WriteLine("Running sync server.");
            HttpService service = new HttpService(mExcels);

            System.Console.Write("\nPress Enter Key!!");
            System.Console.ReadKey();
        }

        public void Deinit()
        {
            for (int i = 0; i < mConfig.GetCount(); i++)
            {
                mExcels[i].Deinit();
            }
        }

        private ExcelParser[] mExcels;
        private CapitalConfig mConfig;
    }
}

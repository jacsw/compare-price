using System;

namespace BackendCore
{
    class MainModule
    {
        static void Main(string[] args)
        {
            ExcelParser.ModuleInit();
            MainModule main = new MainModule();
            main.init();
            main.run();
            main.deinit();

            ExcelParser.ModuleDeInit();
        }

        public void init()
        {
            // TODO : Set by JSON File
            mConfig = new CapitalConfig(@"D:\Rent\Excel\CapitalList.json");
            mConfig.LoadConfig();

            mExcels = new ExcelParser[mConfig.GetCount()];
            for (int i = 0; i < mConfig.GetCount(); i++)
            {
                var info = mConfig.getCapitalData(i);

                mExcels[i] = new ExcelParser(info.Com, info.File, info.Worksheet);
                mExcels[i].SetGTotalPos(info.Price.Row, info.Price.Col);
                mExcels[i].SetRentCostPos(info.Fee.M36.Row, info.Fee.M36.Col);
                mExcels[i].Init();
            }

            System.Console.WriteLine("--------------------------------------------------");
        }

        public void run()
        {
            while (true)
            {
                System.Console.Write("자동차 비용 ( Exit = 0 ) :  ");
                int carcost = Convert.ToInt32(Console.ReadLine());

                if (carcost == 0)
                {
                    break;
                }

                for (int i = 0; i < mConfig.GetCount(); i++)
                {
                    System.Console.WriteLine("{0} : {1}", mExcels[i].GetCompanyName(), mExcels[i].GetRentCost(carcost));
                }
            }

            System.Console.Write("\nPress Enter Key!!");
            System.Console.ReadKey();
        }

        public void deinit()
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

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
                if (info.Price != null)
                {
                    mExcels[i].SetPricePos(info.Price.Row, info.Price.Col);
                }
                if (info.Rate != null)
                {
                    mExcels[i].SetRatePos(info.Rate.Row, info.Rate.Col);
                }
                if (info.Fee.M36 != null)
                {
                    mExcels[i].SetFeeM36Pos(info.Fee.M36.Row, info.Fee.M36.Col);
                }
                if (info.Fee.M48 != null)
                {
                    mExcels[i].SetFeeM48Pos(info.Fee.M48.Row, info.Fee.M48.Col);
                }
                if (info.Fee.M60 != null)
                {
                    mExcels[i].SetFeeM60Pos(info.Fee.M60.Row, info.Fee.M60.Col);
                }

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

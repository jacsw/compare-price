using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace BackendCore.Source.ExcelParser
{
    class CarName
    {
        public string Hyosung;
        public string KB;
        public string Meriz;
        public string Woori;
        public string JBWoori;
        public string Hana;
    }

    class CarList
    {   
        static public CarList getInstance() { return sCarList; }
        static private CarList sCarList = new CarList();

        private Dictionary<string, CarName>  mMap = new Dictionary<string, CarName>();

        public CarName AddCarList(string pName)
        {
            mMap.Add(pName, new CarName());
            return mMap[pName];
        }

        public CarName GetCarName(string pName)
        {
            System.Console.WriteLine("GetCarName... : {0}", pName);
            return mMap[pName];
        }
    }

    class ExcelCarList
    {
        //
        // ExcelBase Common 

        public ExcelCarList(string pPath, string[] pSheet)
        {
            mExcelFileName = pPath;
            mSheetName = pSheet;
        }

        private void loadCarList()
        {
            var list = CarList.getInstance();

            //Get the used Range
            Excel.Range usedRange = mWorksheet.UsedRange;

            //Iterate the rows in the used range
            foreach (Excel.Range row in usedRange.Rows)
            {
                var company = row.Cells[1, 1].Value2;
                var model = row.Cells[1, 2].Value2;
                var trim = row.Cells[1, 3].Value2;

                if (company != null && model != null && trim != null)
                {
                    if (company == "제조사") continue;
                    string carinfo = company.ToString() + "/" + model.ToString() + "/" + trim.ToString();
                    System.Console.Write("Company : {0}", carinfo);

                    var item = list.AddCarList(carinfo);
                    item.Hyosung = row.Cells[1, 5].Value2.ToString();
                    item.KB      = row.Cells[1, 6].Value2.ToString();
                    item.Meriz   = row.Cells[1, 7].Value2.ToString();
                    item.Woori   = row.Cells[1, 8].Value2.ToString();
                    item.JBWoori = row.Cells[1, 9].Value2.ToString();
                    item.Hana    = row.Cells[1,10].Value2.ToString();

                    System.Console.WriteLine(" / JB = {0}", item.JBWoori);
                }
            }
        }

        public void Init()
        {
            try
            {
                System.Console.Write("Data CarList...File : {0}", mExcelFileName);
                mWorkbook = ExcelInstance.getInstance().Workbooks.Open(mExcelFileName);

                foreach(string name in mSheetName) {
                    mWorksheet = mWorkbook.Sheets[name];
                    loadCarList();
                }

                System.Console.WriteLine("Completed");

            }
            finally
            {
            }
        }

        public void Deinit()
        {
            try
            {
                if (mWorkbook != null) mWorkbook.Close(false);
            }
            finally
            {
                ExcelInstance.ReleaseExcelObject(mWorksheet);
                ExcelInstance.ReleaseExcelObject(mWorkbook);
            }
        }

        private Excel.Workbook mWorkbook = null;
        private Excel.Worksheet mWorksheet = null;

        private string mExcelFileName;
        private string[] mSheetName;
    }
}

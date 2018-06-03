using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace BackendCore.Source.ExcelParser
{
    class ExcelInstance
    {
        public static void ModuleInit()
        {
            // sExcelApp = new Excel.Application();
        }

        public static void ModuleDeInit()
        {
            if (sExcelApp != null)
            {
                sExcelApp.Quit();
                ReleaseExcelObject(sExcelApp);
            }
        }

        public static Excel.Application getInstance()
        {
            return sExcelApp;
        }

        protected static Excel.Application sExcelApp = new Excel.Application();

        public static void ReleaseExcelObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    Marshal.ReleaseComObject(obj);
                    obj = null;
                }
            }
            catch (Exception ex)
            {
                obj = null;
                throw ex;
            }
            finally
            {
                GC.Collect();
            }
        }

    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BackendCore.Source.Interface;

namespace BackendCore.Source.ExcelParser
{
    class ExcelHana : ExcelBase
    {
        public ExcelHana(Config.CapitalData pConfig) : base (pConfig)
        {
        }

        public override JsonResponseType GetResonseInfo()
        {
            return null;
            //throw new NotImplementedException();
        }

        public override void SetRequestInfo(JsonRequest request)
        {
            //throw new NotImplementedException();
        }

        protected override void SetPositionfromConfigEach()
        {
            // throw new NotImplementedException();
        }
    }
}

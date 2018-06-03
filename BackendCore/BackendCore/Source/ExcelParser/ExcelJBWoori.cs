using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BackendCore.Source.Interface;

namespace BackendCore.Source.ExcelParser
{
    class ExcelJBWoori : ExcelBase
    {
        public ExcelJBWoori(Config.CapitalData pConfig) : base(pConfig)
        {
        }

        public override void GetResonseInfo()
        {
            throw new NotImplementedException();
        }

        public override void SetRequestInfo(JsonRequest request)
        {
            throw new NotImplementedException();
        }

        protected override void SetPositionfromConfigEach()
        {
            throw new NotImplementedException();
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;

namespace 将.xls转换为.xlsx
{
    [ExcelImporter(IsLabelingError = true, IsDisableAllFilter = true)]
    public  class Class1
    {

        [ImporterHeader(Name = "ED_PDNo")]
        public string a{get;set;}
        [ImporterHeader(Name = "型号")]
        public string b{get;set;}
        [ImporterHeader(Name = "包装型态")]
        public string g{get;set;}
        [ImporterHeader(Name = "材料批号")]
        public string c{get;set;}
        [ImporterHeader(Name = "加工数量")]
        public string d{get;set;}
        [ImporterHeader(Name = "订单编码")]
        public string f{get;set;}
      
    }
}

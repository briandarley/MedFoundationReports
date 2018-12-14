using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SplitReport
{
    public class MasterDataRecord
    {
        //Department Name	PeopleSoft Source #	Fund Title	 Cash 	 Graystone 	 Total Available Expendable Bal 
        public string DepartmentName { get; set; }
        public string PeopleSoftSource { get; set; }
        public string FundTitle { get; set; }
        public double Cash { get; set; }
        public double Graystone { get; set; }
        public double TotalAvailable { get; set; }
        
    }
}

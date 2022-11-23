using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WForRestGet
{
    public class SetOutOfOffice
    {
        public string InternalContent { get; set; }
        public string ExternalContent { get; set; }
        public DateTime DateStart { get; set; }
        public DateTime DateEnd { get; set; }
        public string RequestNumber = "";
        public string Explanation = "Automatic status processing";
        public string Uid { get; set; }
        public string Account { get; set; }
        public string Domain = "IE.CORP";
        public string Entity = "ADAccount";
        public bool Disable { get; set; }

    }

    public class SetOutOffResult
    {
        public bool Status { get; set; }
        public string Message { get; set; }
        public string Data { get; set; }
    }

    public class GetOutOfOffice
    {
        public string Domain = "IE.CORP";
        public string Account { get; set; }

    }

}

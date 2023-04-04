using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WForRestGet.DataModel
{
    public class Leave
    {
        public int Id { get; set; }
        public string LeaveType { get; set; }
        public string LeaveDescription { get; set; }
        public List<Datauser> Datausers { get; set; }

    }
}

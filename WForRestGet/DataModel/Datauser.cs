using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WForRestGet.DataModel
{
    public class Datauser
    {
        public string FimSyncKey { get; set; }
        public string AccountId { get; set; }
        public string AccountName { get; set; }
        public string LastName { get; set; }
        public string FirstName { get; set; }
        public string MiddleName { get; set; }
        public string EmployeeNumber { get; set; }
        public DateTime? Birthday { get; set; }
        public string CompanyName { get; set; }
        public string DepartmentName { get; set; }
        public string JobTitle { get; set; }
        public DateTime? DateIn { get; set; }
        public int? LeaveId { get; set; }
        public Leave Leave { get; set; }
        public DateTime? LeaveStart { get; set; }
        public DateTime? LeaveEnd { get; set; }
        public string City { get; set; }
        public string Phone { get; set; }
        public string Email { get; set; }
        public bool Disabled { get; set; }
    }
}

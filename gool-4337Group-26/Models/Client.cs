
using System;

namespace gool_4337Group_26.Models
{
    public class Client
    {
        public int Id { get; set; }
        public string FullName { get; set; }
        public string Email { get; set; }
        public DateTime BirthDate { get; set; }

        public int Age
        {
            get
            {
                var age = DateTime.Now.Year - BirthDate.Year;
                if (BirthDate > DateTime.Now.AddYears(-age)) age--;
                return age;
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ErrorEmp
{
    public class ErrorEmployee
    {
        //this is for collect data which id containing error(empty field)//
        //not used//
        public int id { get; set; }
        public string name { get; set; }
        public string email { get; set; }
        public string mobile { get; set; }
        public int age { get; set; }
        public string error { get; set; }
    }
}

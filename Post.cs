using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApplication4
{
    class Post
    {
        public DateTime? period { get; set; } 
        public string currency_id { get; set; }
       // public string currency_name_th { get; set; }
      //  public string currency_name_eng { get; set; }
        public double? buying_sight { get; set; }
        public double? buying_transfer { get; set; }
        public double? selling { get; set; }

        public DateTime? CallDate { get; set; }
        public double? mid_rate { get; set; }

    }
}

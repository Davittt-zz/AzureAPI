using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HAA.BusinessEntities.Input
{
    public class CrossProductModel
    {
        public Vector V1 { get; set; }

        public Vector V2 { get; set; }

        public Vector V3 { get; set; }
    }

    public class Vector
    {
        public float X { get; set; }
        
        public float Y { get; set; }

        public float Z { get; set; }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Pipeline
{
    class Defect
    {
        public int Id { get; set; }

        public string Name { get; set; }

        public double LengthCoordinate { get; set; }

        public DateTime HeightCoordinate { get; set; }

        public int Diameter { get; set; }
 
        public double SectionLength { get; set; }
    }
}

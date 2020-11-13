using _8.gyakorlat.Abstarctions;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _8.gyakorlat.Entities
{
    class PresentFactory : IToyFactory
    {
        public Color RibbonColor { get; set; }
        public Color BoxColor { get; set; }

        public Toy CreateNew()
        {
            return new Present(BoxColor,RibbonColor);
            
            
        }
    }
}

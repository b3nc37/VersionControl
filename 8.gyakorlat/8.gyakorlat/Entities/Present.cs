using _8.gyakorlat.Abstarctions;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _8.gyakorlat.Entities
{
    class Present : Toy
    {
        public SolidBrush RibbonColor { get; private set; }
        public SolidBrush BoxColor { get; private set; }

        public Present(Color box,Color ribbon)
        {
            RibbonColor = new SolidBrush(ribbon);
            BoxColor = new SolidBrush(box);
        }
        

        protected override void DrawImage(Graphics g)
        {
            g.FillRectangle(BoxColor,0,0,Width,Height);
            g.FillRectangle(RibbonColor, 20, 0, Width / 5, Height);
            g.FillRectangle(RibbonColor, 0, 20, Width,Height/5);
        }
    }
}

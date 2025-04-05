using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Sd = System.Drawing;

using Rg = Rhino.Geometry;

using XL = Microsoft.Office.Interop.Excel;
using MC = Microsoft.Office.Core;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;

namespace Bumblebee
{
    public class ExShape
    {
        #region members

        public ExWorksheet sheet = null;
        public XL.Shape ComObj = null;

        public enum ShapeTypes { Illustration, Control, SmartArt, Line };
        protected ShapeTypes shapeType = ShapeTypes.Illustration;

        #endregion

        #region constructors

        public ExShape(ExShape exSmart)
        {
            this.sheet = new ExWorksheet(exSmart.sheet);
            this.ComObj = exSmart.ComObj;
            this.shapeType = exSmart.shapeType;
        }

        public ExShape(XL.Shape comObj, ExWorksheet sheet, ShapeTypes shapeType)
        {
            this.sheet = new ExWorksheet(sheet);
            this.ComObj = comObj;
            this.shapeType = shapeType;
        }

        public ExShape(ExWorksheet sheet, string name, Rg.Rectangle3d boundary, List<string> values, List<int> levels)
        {
            this.sheet = sheet;
            this.shapeType = ShapeTypes.SmartArt;
            int countV = values.Count;
            int countL = levels.Count;

            for (int i = countL; i < countV; i++)
            {
                levels.Add(countL - 1);
            }

            foreach (XL.Shape obj in sheet.ComObj.Shapes)
            {
                if (name == obj.Name)
                {
                    obj.Delete();
                    break;
                }
            }

            this.ComObj = sheet.ComObj.Shapes.AddSmartArt(sheet.ParentApp.ComObj.SmartArtLayouts[1], boundary.Corner(0).Y, boundary.Width, boundary.Height);

            int count = this.ComObj.SmartArt.AllNodes.Count;
            for (int i = 0; i < count; i++)
            {
                this.ComObj.SmartArt.AllNodes[1].Delete();
            }

            MC.SmartArtNode node = this.ComObj.SmartArt.AllNodes.Add();
            node.TextFrame2.TextRange.Text = values[0];

            for (int i = 1; i < values.Count; i++)
            {
                MC.SmartArtNode nodeA = this.ComObj.SmartArt.AllNodes.Add();
                nodeA.TextFrame2.TextRange.Text = values[i];
                for(int j = 1;j<levels[i]-1;j++)
                {
                    nodeA.Demote();
                }
            }

            this.ComObj.Name = name;
        }

        #endregion

        #region properties

        public virtual ShapeTypes ShapeType
        {
            get { return shapeType; }
        }

        public virtual string FontFamily
        {
            get { return this.ComObj.TextFrame.Characters().Font.Name; }
            set { this.ComObj.TextFrame.Characters().Font.Name = value; }
        }

        public virtual double FontSize
        {
            get { return this.ComObj.TextFrame.Characters().Font.Size; }
            set { this.ComObj.TextFrame.Characters().Font.Size = value; }
        }

        public virtual Sd.Color FontColor
        {
            get { return Sd.ColorTranslator.FromOle((int)this.ComObj.TextFrame.Characters().Font.Color); }
            set { this.ComObj.TextFrame.Characters().Font.Color = value; }
        }

        public virtual Justification FontJustification
        {
            set
            {
                this.ComObj.TextFrame.HorizontalAlignment = value.ToExcelHalign();
                this.ComObj.TextFrame.VerticalAlignment = value.ToExcelValign();
            }
            get
            {
                int align = 0;
                switch (this.ComObj.TextFrame.VerticalAlignment)
                {
                    case XL.XlVAlign.xlVAlignCenter:
                        align = 3;
                        break;
                    case XL.XlVAlign.xlVAlignTop:
                        align = 6;
                        break;
                }

                switch (this.ComObj.TextFrame.HorizontalAlignment)
                {
                    case XL.XlHAlign.xlHAlignLeft:
                    case XL.XlHAlign.xlHAlignGeneral:
                        break;
                    case XL.XlHAlign.xlHAlignRight:
                        align += 2;
                        break;
                    default:
                        align += 1;
                        break;
                }
                return (Justification)align;
            }
        }

        public bool Bold
        {
            get { return this.ComObj.TextFrame.Characters().Font.Bold; }
            set { this.ComObj.TextFrame.Characters().Font.Bold = value; }
        }

        public bool Italic
        {
            get { return this.ComObj.TextFrame.Characters().Font.Italic; }
            set { this.ComObj.TextFrame.Characters().Font.Italic = value; }
        }


        #endregion

        #region methods

        public string GetText()
        {
            string output = string.Empty;
            if (this.shapeType != ShapeTypes.Line)
            {
                output = this.ComObj.TextFrame.Characters().Text;
            }
            return output;
        }

        public void SetText(string text)
        {
            if (this.shapeType != ShapeTypes.Line)
            {
                this.ComObj.TextFrame.Characters().Text = text;
            }
        }

        public void SetFillColor(Sd.Color color)
        {
            if (this.shapeType != ShapeTypes.Line)
            {
                this.ComObj.Fill.ForeColor.RGB = Sd.ColorTranslator.ToOle(color);
                this.ComObj.Fill.BackColor.RGB = Sd.ColorTranslator.ToOle(color);
            }
        }

        public void SetStrokeColor(Sd.Color color)
        {
            this.ComObj.Line.ForeColor.RGB = Sd.ColorTranslator.ToOle(color);
            this.ComObj.Line.BackColor.RGB = Sd.ColorTranslator.ToOle(color);
        }

        public void SetStrokeWeight(double weight)
        {
            this.ComObj.Line.Weight = (float)weight;
        }

        public void SetList(int type)
        {
            this.ComObj.SmartArt.Layout = this.sheet.ParentApp.ComObj.SmartArtLayouts[type];
        }

        #endregion

        #region overrides

        public override string ToString()
        {
            return "Smart | " + ShapeType.ToString();
        }

        #endregion

    }
}

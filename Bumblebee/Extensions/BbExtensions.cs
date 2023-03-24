using Grasshopper.Kernel.Types;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Sd = System.Drawing;

using Rg = Rhino.Geometry;

using XL = Microsoft.Office.Interop.Excel;
using Grasshopper.Kernel;

namespace Bumblebee
{
    public static class BbExtensions
    {

        #region controls

        public static ExShape AddButton(this ExWorksheet sheet, string name, string title, string macro, Rg.Rectangle3d boundary)
        {
            name = sheet.PrepControl(name);

            XL.Shape shape = sheet.ComObj.Shapes.AddFormControl(XL.XlFormControl.xlButtonControl, (int)boundary.Corner(0).X, (int)boundary.Corner(0).Y, (int)boundary.Width, (int)boundary.Height);
            shape.Name = name;
            shape.OLEFormat.Object.Text = title;
            try
            {
                shape.OnAction = macro;
            }
            catch
            {

            }

            return new ExShape(shape, sheet, ExShape.ShapeTypes.Control);
        }

        public static ExShape AddCheckBox(this ExWorksheet sheet, string name, string title, ExCell cell, Rg.Rectangle3d boundary)
        {
            name = sheet.PrepControl(name);

            XL.Shape shape = sheet.ComObj.Shapes.AddFormControl(XL.XlFormControl.xlCheckBox, (int)boundary.Corner(0).X, (int)boundary.Corner(0).Y, (int)boundary.Width, (int)boundary.Height);
            shape.Name = name;
            shape.OLEFormat.Object.Text = title;
            shape.OLEFormat.Object.LinkedCell = cell.ToString();

            return new ExShape(shape, sheet, ExShape.ShapeTypes.Control);
        }

        public static ExShape AddLabel(this ExWorksheet sheet, string name, string title, Rg.Rectangle3d boundary)
        {
            name = sheet.PrepControl(name);

            XL.Shape shape = sheet.ComObj.Shapes.AddFormControl(XL.XlFormControl.xlLabel, (int)boundary.Corner(0).X, (int)boundary.Corner(0).Y, (int)boundary.Width, (int)boundary.Height);
            shape.Name = name;
            shape.OLEFormat.Object.Text = title;

            return new ExShape(shape, sheet, ExShape.ShapeTypes.Control);
        }

        public static ExShape AddDropDown(this ExWorksheet sheet, string name, string title, ExCell cell, List<string> data, Rg.Rectangle3d boundary)
        {
            name = sheet.PrepControl(name);

            XL.Shape shape = sheet.ComObj.Shapes.AddFormControl(XL.XlFormControl.xlDropDown, (int)boundary.Corner(0).X, (int)boundary.Corner(0).Y, (int)boundary.Width, (int)boundary.Height);
            shape.Name = name;
            shape.OLEFormat.Object.Text = title;

            foreach (string d in data)
            {
                shape.OLEFormat.Object.AddItem(d);
            }
            shape.OLEFormat.Object.LinkedCell = cell.ToString();

            return new ExShape(shape, sheet, ExShape.ShapeTypes.Control);
        }

        public static ExShape AddListBox(this ExWorksheet sheet, string name, ExCell cell, List<string> data, Rg.Rectangle3d boundary)
        {
            name = sheet.PrepControl(name);

            XL.Shape shape = sheet.ComObj.Shapes.AddFormControl(XL.XlFormControl.xlListBox, (int)boundary.Corner(0).X, (int)boundary.Corner(0).Y, (int)boundary.Width, (int)boundary.Height);
            shape.Name = name;

            foreach (string d in data)
            {
                shape.OLEFormat.Object.AddItem(d);
            }
            shape.OLEFormat.Object.LinkedCell = cell.ToString();

            return new ExShape(shape, sheet, ExShape.ShapeTypes.Control);
        }

        public static ExShape AddScrollBar(this ExWorksheet sheet, string name, ExCell cell, Rg.Interval domain, double increment, Rg.Rectangle3d boundary)
        {
            name = sheet.PrepControl(name);

            XL.Shape shape = sheet.ComObj.Shapes.AddFormControl(XL.XlFormControl.xlScrollBar, (int)boundary.Corner(0).X, (int)boundary.Corner(0).Y, (int)boundary.Width, (int)boundary.Height);
            shape.Name = name;

            shape.OLEFormat.Object.Min = domain.Min;
            shape.OLEFormat.Object.Max = domain.Max;

            shape.OLEFormat.Object.SmallChange = increment;
            shape.OLEFormat.Object.LargeChange = increment;

            shape.OLEFormat.Object.LinkedCell = cell.ToString();

            return new ExShape(shape, sheet, ExShape.ShapeTypes.Control);
        }

        public static ExShape AddSpinner(this ExWorksheet sheet, string name, ExCell cell, Rg.Interval domain, double increment, Rg.Rectangle3d boundary)
        {
            name = sheet.PrepControl(name);

            XL.Shape shape = sheet.ComObj.Shapes.AddFormControl(XL.XlFormControl.xlSpinner, (int)boundary.Corner(0).X, (int)boundary.Corner(0).Y, (int)boundary.Width, (int)boundary.Height);
            shape.Name = name;

            shape.OLEFormat.Object.Min = domain.Min;
            shape.OLEFormat.Object.Max = domain.Max;

            shape.OLEFormat.Object.SmallChange = increment;

            shape.OLEFormat.Object.LinkedCell = cell.ToString();

            return new ExShape(shape, sheet, ExShape.ShapeTypes.Control);
        }

        public static ExShape AddEditBox(this ExWorksheet sheet, string name, ExCell cell, string data, Rg.Rectangle3d boundary)
        {
            name = sheet.PrepControl(name);

            XL.Shape shape = sheet.ComObj.Shapes.AddFormControl(XL.XlFormControl.xlEditBox, (int)boundary.Corner(0).X, (int)boundary.Corner(0).Y, (int)boundary.Width, (int)boundary.Height);
            shape.Name = name;
            shape.OLEFormat.Object.Text = data;

            shape.OLEFormat.Object.LinkedCell = cell.ToString();

            return new ExShape(shape, sheet, ExShape.ShapeTypes.Control);
        }

        private static string PrepControl(this ExWorksheet sheet, string name)
        {
            string[] subNames = name.Split('-');
            Array.Reverse(subNames);
            string n = subNames[2] + subNames[1] + subNames[0];
            sheet.RemoveControl(n);

            return n;
        }

        private static void RemoveControl(this ExWorksheet sheet, string name)
        {
            foreach (XL.Shape shp in sheet.ComObj.Shapes)
            {
                if (shp.Name == name) shp.Delete();
            }
        }

        #endregion

        #region drawing

        public static ExShape AddLine(this ExWorksheet sheet,string name,ArrowStyle startArrow, ArrowStyle endArrow, Rg.Line line)
        {
            name = sheet.PrepControl(name);

            XL.Shape shape = sheet.ComObj.Shapes.AddLine((float)line.From.X, -(float)line.From.Y, (float)line.To.X, -(float)line.To.Y);
            shape.Name = name;
            shape.Line.BeginArrowheadStyle = startArrow.ToExcel();
            shape.Line.EndArrowheadStyle = endArrow.ToExcel();
            return new ExShape(shape, sheet, ExShape.ShapeTypes.Line);
        }

        public static ExShape AddShape(this ExWorksheet sheet, string name, ShapeArrow type, Rg.Rectangle3d boundary)
        {
            name = sheet.PrepControl(name);

            XL.Shape shape = sheet.ComObj.Shapes.AddShape(type.ToExcel(), (float)boundary.Corner(0).X, (float)boundary.Corner(0).Y, (float)boundary.Width, (float)boundary.Height);
            shape.Name = name;

            return new ExShape(shape, sheet, ExShape.ShapeTypes.Illustration);
        }

        public static ExShape AddShape(this ExWorksheet sheet, string name, ShapeStar type, Rg.Rectangle3d boundary)
        {
            name = sheet.PrepControl(name);

            XL.Shape shape = sheet.ComObj.Shapes.AddShape(type.ToExcel(), (float)boundary.Corner(0).X, (float)boundary.Corner(0).Y, (float)boundary.Width, (float)boundary.Height);
            shape.Name = name;

            return new ExShape(shape, sheet, ExShape.ShapeTypes.Illustration);
        }

        public static ExShape AddShape(this ExWorksheet sheet, string name, ShapeFlowChart type, Rg.Rectangle3d boundary)
        {
            name = sheet.PrepControl(name);

            XL.Shape shape = sheet.ComObj.Shapes.AddShape(type.ToExcel(), (float)boundary.Corner(0).X, (float)boundary.Corner(0).Y, (float)boundary.Width, (float)boundary.Height);
            shape.Name = name;

            return new ExShape(shape, sheet, ExShape.ShapeTypes.Illustration);
        }
        
        public static ExShape AddShape(this ExWorksheet sheet, string name, ShapeSymbol type, Rg.Rectangle3d boundary)
        {
            name = sheet.PrepControl(name);

            XL.Shape shape = sheet.ComObj.Shapes.AddShape(type.ToExcel(), (float)boundary.Corner(0).X, (float)boundary.Corner(0).Y, (float)boundary.Width, (float)boundary.Height);
            shape.Name = name;

            return new ExShape(shape, sheet, ExShape.ShapeTypes.Illustration);
        }
        
        public static ExShape AddShape(this ExWorksheet sheet, string name, ShapeGeometry type, Rg.Rectangle3d boundary)
        {
            name = sheet.PrepControl(name);

            XL.Shape shape = sheet.ComObj.Shapes.AddShape(type.ToExcel(), (float)boundary.Corner(0).X, (float)boundary.Corner(0).Y, (float)boundary.Width, (float)boundary.Height);
            shape.Name = name;

            return new ExShape(shape, sheet, ExShape.ShapeTypes.Illustration);
        }
        public static ExShape AddShape(this ExWorksheet sheet, string name, ShapeFigure type, Rg.Rectangle3d boundary)
        {
            name = sheet.PrepControl(name);

            XL.Shape shape = sheet.ComObj.Shapes.AddShape(type.ToExcel(), (float)boundary.Corner(0).X, (float)boundary.Corner(0).Y, (float)boundary.Width, (float)boundary.Height);
            shape.Name = name;

            return new ExShape(shape, sheet, ExShape.ShapeTypes.Illustration);
        }

        #endregion
    }
}

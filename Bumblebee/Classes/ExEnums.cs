using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Bumblebee
{

    public enum HorizontalBorder { None, Bottom, Top, Both };
    public enum VerticalBorder { None, Left, Right, Both };
    public enum LineType { None, Continuous, Dash, DashDot, DashDotDot, Dot, Double, SlantDashDot };
    public enum BorderWeight { Hairline, Thin, Medium, Thick, None };
    public enum Justification { BottomLeft, BottomMiddle, BottomRight, CenterLeft, CenterMiddle, CenterRight, TopLeft, TopMiddle, TopRight };
    public enum ValueCondition { Greater, GreaterEqual, Less, LessEqual, Equal, NotEqual };
    public enum AverageCondition { AboveAverage, AboveEqualAverage,AboveDeviation, BelowAverage, BelowEqualAverage, BelowDeviation };
    public enum VbModuleType { ClassModule, Document, MSForm, StdModule, ActiveX};
    public enum ChartFill { Cluster, Stack, Fill};
    public enum RadialChartType { Pie, Pie3D, Donut, Radar,RadarFilled};
    public enum BarChartType { Basic, Box, Pyramid, Cylinder, Cone };
    public enum LineChartType { Line, LineMarkers, Area, Area3d };
    public enum ScatterChartType { Scatter, ScatterLines, ScatterSmooth, Bubble, Bubble3D };
    public enum SurfaceChartType { Surface, SurfaceWireframe, SurfaceTop, SurfaceWireframeTop };
    public enum LegendLocations { None, Left,Right,Top, Bottom};
    public enum LabelType { None, Value, Category};
    public enum GridType { None,Primary,All };
    public enum ShapeList { AlternatingFlow, AlternatingHexagons, BasicBlockList, CircleAccentTimeline, ConvergingArrows, DivergingArrows, Grouped, HorizontalBullet, LinearVenn, Lined, MultidirectionalCycle, NondirectionalCycle, Process, SquareAccent, Stacked, Trapezoid, VerticalAccent, VerticalArrow, VerticalBlock, VerticalBox, VerticalBullet, VerticalCircle };
    public enum ArrowStyle { None,Open,Oval, Diamond, Stealth, Triangle};
    public enum ShapeArrow { Right,Left,Up,Down,LeftRight,UpDown,Quad,LeftRightUp,Bent,UTurn,LeftUp,BentUp,CurvedRight,CurvedLeft,CurvedUp,CurvedDown,StripedRight,NotchedRight, Circular,Swoosh,LeftCircular, LeftRightCircular };
    public enum ShapeStar { Pt4,Pt5,Pt6,Pt7,Pt8,Pt10, Pt12,Pt16,Pt24,Pt32};
    public enum ShapeFlowChart { Process, AlternateProcess, Decision, Data, PredefinedProcess, InternalStorage, Document, Multidocument, Terminator, Preparation, ManualInput, ManualOperation, Connector, OfflineStorage, OffpageConnector, Card, PunchedTape, SummingJunction, Or, Collate, Sort, Extract, Merge, StoredData, Delay, SequentialAccessStorage, MagneticDisk, DirectAccessStorage, Display };
    public enum ShapeSymbol { Plus,Minus,Multiply,Divide,Equal,NotEqual,LeftBracket,RightBracket,DoubleBracket,LeftBrace,RightBrace,DoubleBrace };
    public enum ShapeGeometry { Rectangle, Parallelogram, Trapezoid, Diamond, RoundedRectangle, Octagon, IsoscelesTriangle, RightTriangle, Oval, Hexagon, Cross, RegularPentagon, Pentagon, Donut, BlockArc, NonIsoscelesTrapezoid, Decagon, Heptagon, Dodecagon, Round1Rectangle, Round2SameRectangle, SnipRoundRectangle, Snip1Rectangle, Snip2SameRectangle, Round2DiagRectangle, Snip2DiagRectangle };
    public enum ShapeFigure { Can, Cube, Bevel, FoldedCorner, SmileyFace, NoSymbol, Heart, LightningBolt, Sun, Moon, Arc, Plaque, Cloud, Gear6, Gear9, Funnel, Chevron, Explosion1, Balloon, Explosion2, Wave, DoubleWave, DiagonalStripe, Pie, Frame, HalfFrame, Tear, Chord, Corner, PieWedge };
    public enum Extensions { xlxs,xlxm};
}
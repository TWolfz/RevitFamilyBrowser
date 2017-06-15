﻿using RevitFamilyBrowser.Revit_Classes;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using System.Windows.Shapes;
using Point = System.Drawing.Point;

namespace RevitFamilyBrowser.WPF_Classes
{
    public class WpfCoordinates
    {
        public double X { get; set; }
        public double Y { get; set; }


        public double GetLength(Line line)
        {
            return Math.Sqrt(Math.Pow((line.X1 - line.X2), 2) + Math.Pow((line.Y1 - line.Y2), 2));
        }

        public double GetSlope(Line line)
        {
            return (line.Y2 - line.Y1) / (line.X2 - line.X1);
        }

        public Point GetCenter(Line line)
        {
            Point point = new Point();
            point.X = (int)(line.X2 + line.X1) / 2;
            point.Y = (int)(line.Y2 + line.Y1) / 2;
            return point;
        }

        //-----Return coefficients of Line equation ax+by+c=0
        public List<double> LineEquation(Line line)
        {
            List<double> result = new List<double>();
            double a = line.Y2 - line.Y1;
            double b = line.X1 - line.X2;
            double c = line.Y1 * (line.X1 - line.X2) - line.X1 * (line.Y1 - line.Y2);
            result.Add(a);
            result.Add(b);
            result.Add(-c);
            return result;
        }

        //-----Get two lines and return intersection point
        public Point GetIntersection(Line box, Line wall)
        {
            Line normalBox = OrtoNormalization(box);
            Line normalWall = OrtoNormalization(wall);

            List<double> wallCoefs = LineEquation(normalWall);
            double a1 = wallCoefs[0];
            double b1 = wallCoefs[1];
            double c1 = wallCoefs[2];

            List<double> boxCoefs = LineEquation(normalBox);
            double a2 = boxCoefs[0];
            double b2 = boxCoefs[1];
            double c2 = boxCoefs[2];
            Point intersection = new Point();

            {
                double x = (c1 * b2 - c2 * b1) / (a2 * b1 - a1 * b2);
                double y = 0;
                if (b1.Equals(0))
                {
                    y = (int)(-c2 - a2 * x) / b2;
                }
                else if (b2.Equals(0))
                    y = (-c1 - a1 * x) / b1;

                else
                    y = (-c2 - (a2 * x)) / b2;

                intersection.X = (int)x;
                intersection.Y = (int)y;
            }
            return intersection;
        }

        public PointF GetIntersectionD(Line box, Line wall)
        {
            List<double> wallCoefs = LineEquation(wall);
            double a1 = wallCoefs[0];
            double b1 = wallCoefs[1];
            double c1 = wallCoefs[2];

            List<double> boxCoefs = LineEquation(box);
            double a2 = boxCoefs[0];
            double b2 = boxCoefs[1];
            double c2 = boxCoefs[2];

            PointF intersection = new PointF();
            {
                double x = (c1 * b2 - c2 * b1) / (a2 * b1 - a1 * b2);
                double y = 0;
                if (b1.Equals(0))
                {
                    y = (int)(-c2 - a2 * x) / b2;
                }
                else if (b2.Equals(0))
                    y = (-c1 - a1 * x) / b1;

                else
                    y = (-c2 - (a2 * x)) / b2;

                if (x.Equals(Double.NaN))
                {
                    x = float.PositiveInfinity;
                }
                if (y.Equals(Double.NaN))
                {
                    y = float.PositiveInfinity;
                }
                intersection.X = (float)x;
                intersection.Y = (float)y;
            }
            return intersection;
        }

        //-----Check if point belongs to line
        public bool IntersectionPositionCheck(Line line, Point point)
        {
            int lineMaxX = (int)(line.X1 > line.X2 ? line.X1 : line.X2);
            int lineMinX = (int)(line.X1 < line.X2 ? line.X1 : line.X2);
            int lineMaxY = (int)(line.Y1 > line.Y2 ? line.Y1 : line.Y2);
            int lineMinY = (int)(line.Y1 < line.Y2 ? line.Y1 : line.Y2);

            if (point.X <= lineMaxX && point.X >= lineMinX && point.Y <= lineMaxY && point.Y >= lineMinY)
            {
                return true;
            }
            return false;
        }

        public bool IntersectionPositionCheckF(Line line, PointF point)
        {
            double lineMaxX = (line.X1 > line.X2 ? line.X1 : line.X2);
            double lineMinX = (line.X1 < line.X2 ? line.X1 : line.X2);
            double lineMaxY = (line.Y1 > line.Y2 ? line.Y1 : line.Y2);
            double lineMinY = (line.Y1 < line.Y2 ? line.Y1 : line.Y2);

            if (point.X <= lineMaxX && point.X >= lineMinX && point.Y <= lineMaxY && point.Y >= lineMinY)
            {
                return true;
            }
            return false;
        }


        //-----If line is paralle to one of axis replace infinity coord to defined
        public Line OrtoNormalization(Line perpend)
        {
            Line normal = new Line();
            int xA = 0; int xB = 0;
            int yA = 0; int yB = 0;

            if (perpend.X1 != int.MaxValue || perpend.X1 != int.MinValue)
            {
                xA = (int)perpend.X1;
            }
            if (perpend.Y1 != int.MaxValue || perpend.Y1 != int.MinValue)
            {
                yA = (int)perpend.Y1;
            }
            if (perpend.X2 != int.MaxValue || perpend.X2 != int.MinValue)
            {
                xB = (int)perpend.X2;
            }
            if (perpend.Y2 != int.MaxValue || perpend.Y2 != int.MinValue)
            {
                yB = (int)perpend.Y2;
            }

            if ((xA == int.MaxValue || xA == int.MinValue) && yA == 0)
            {
                yA = 0; xA = xB;
            }

            if ((xB == int.MaxValue || xB == int.MinValue) && yB == 0)
            {
                xB = 0; yB = yA;
            }

            if ((yA == int.MaxValue || yA == int.MinValue) && xA == 0)
            {
                yA = 0; xA = xB;
            }

            if ((yB == int.MaxValue || yB == int.MinValue) && xB == 0)
            {
                yB = 0; xB = xA;
            }
            normal.X1 = xA;
            normal.X2 = xB;
            normal.Y1 = yA;
            normal.Y2 = yB;

            return normal;
        }

        //-----Build dashed line in limits of box around the room
        public Line BuildInstallAxis(List<Line> boundingBox, Line perpend)
        {
            Line gridLine = new Line();
            Line normal = OrtoNormalization(perpend);
            List<Point> allIntersections = new List<Point>();
            foreach (var side in boundingBox)
            {
                Point intersection = GetIntersection(side, normal);
                if (IntersectionPositionCheck(side, intersection))
                {
                    allIntersections.Add(intersection);
                }
            }

            int count = 0;
            foreach (var item in allIntersections)
            {
                count++;
                if (count == 1)
                {
                    gridLine.X1 = item.X;
                    gridLine.Y1 = item.Y;
                }
                else if (count == 2)
                {
                    gridLine.X2 = item.X;
                    gridLine.Y2 = item.Y;
                }
            }
            return DrawDashedLine(gridLine);
        }

        public Line BuildInstallAxisF(List<Line> boundingBox, Line perpend)
        {
            Line gridLine = new Line();
           
            List<PointF> allIntersections = new List<PointF>();
            foreach (var side in boundingBox)
            {
                PointF intersection = GetIntersection(side, perpend);
                if (IntersectionPositionCheckF(side, intersection))
                {
                    allIntersections.Add(intersection);
                }
            }

            int count = 0;
            foreach (var item in allIntersections)
            {
                count++;
                if (count == 1)
                {
                    gridLine.X1 = item.X;
                    gridLine.Y1 = item.Y;
                }
                else if (count == 2)
                {
                    gridLine.X2 = item.X;
                    gridLine.Y2 = item.Y;
                }
            }
            return DrawDashedLine(gridLine);
        }


        public List<PointF> SplitLineProportional(Line line, int lineNumber)
        {
            List<PointF> points = new List<PointF>();
            int partNumber = lineNumber * 2;
            for (int i = 1; i < partNumber; i = i + 2)
            {
                PointF point = new PointF();
                double top = i;
                double bottom = (partNumber - i);
                if ((partNumber - i) == 0)
                {
                    bottom = 1;
                }
                var proportion = top / bottom;
                point.X = (float)((line.X1 + (line.X2 * proportion)) / (1 + proportion));
                point.Y = (float)((line.Y1 + (line.Y2 * proportion)) / (1 + proportion));
                points.Add(point);
            }
            return points;
        }

        public List<PointF> SplitLineEqual(Line line, int lineNumber)
        {
            List<PointF> points = new List<PointF>();

            int partNumber = lineNumber + 1;
            for (int i = 1; i < partNumber; i++)
            {
                PointF point = new PointF();
                double top = i;
                double bottom = (partNumber - i);
                if ((partNumber - i) == 0)
                {
                    bottom = 1;
                }
                var proportion = top / bottom;
                point.X = (float)((line.X1 + (line.X2 * proportion)) / (1 + proportion));
                point.Y = (float)((line.Y1 + (line.Y2 * proportion)) / (1 + proportion));
                points.Add(point);
            }
            return points;
        }

        #region Split line on parts by segment lenght

        private int GetLinePartsNumber(Line line, double distance)
        {
            int parts;
            double wallLength = this.GetLength(line);
            if ((wallLength % distance).Equals(0) && wallLength / distance > 2)
            {
                parts = (int)(wallLength / distance) - 1;
            }
            else if (wallLength / distance <= 2 && wallLength > distance)
            {
                parts = 2;
            }
            else
            {
                parts = (int)Math.Truncate(wallLength / distance);
            }
            return parts;
        }

        private List<double> GetPartsSizes(Line line, double distance)
        {
            List<double> partLenghts = new List<double>();
            int parts = GetLinePartsNumber(line, distance);
            double firstPart = (GetLength(line) - ((parts - 1) * distance)) / 2;
            partLenghts.Add(firstPart);
            for (int i = 1; i <= parts - 1; i++)
            {
                double part = firstPart + i * distance;
                partLenghts.Add(part);
            }
            return partLenghts;
        }

        public double GetAngle(Line line)
        {
            List<double> lineCoefs = LineEquation(line);
            double angle = -Math.Atan(lineCoefs[1] / lineCoefs[0]);
            return angle;
        }

        public PointF GetSecondCoord(Line line, double distance)
        {
            PointF point = new Point();
            double angle = GetAngle(line);
            point.X = (float)(line.X1 + distance * Math.Sin(angle));
            point.Y = (float)(line.Y1 + distance * Math.Cos(angle));
            return point;
        }
        public List<PointF> SplitLineDistance(Line line, double distance)
        {
            List<PointF> points = new List<PointF>();
            List<double> partSizes = GetPartsSizes(line, distance);
            foreach (var part in partSizes)
            {
                points.Add(GetSecondCoord(line, part));
            }
            return points;
        }
        #endregion 
        public List<Line> GetBoundingBox(PointF min, PointF max, GridSetup grid)
        // public List<Line> GetBoundingBox(ConversionPoint min, ConversionPoint max, GridSetup grid)
        {
            int scale = grid.Scale;
            float derX = grid.Derrivation.X;
            float derY = grid.Derrivation.Y;

            List<Line> boxSides = new List<Line>();
            int offset = 800 / scale;

            Line SideA = new Line();
            SideA.X1 = min.X / scale + derX - offset;
            SideA.Y1 = -min.Y / scale + derY + offset;
            SideA.X2 = min.X / scale + derX - offset;
            SideA.Y2 = -max.Y / scale + derY - offset;
            boxSides.Add(SideA);

            Line SideB = new Line();
            SideB.X1 = min.X / scale + derX - offset;
            SideB.Y1 = -max.Y / scale + derY - offset;
            SideB.X2 = max.X / scale + derX + offset;
            SideB.Y2 = -max.Y / scale + derY - offset;
            boxSides.Add(SideB);

            Line SideC = new Line();
            SideC.X1 = max.X / scale + derX + offset;
            SideC.Y1 = -max.Y / scale + derY - offset;
            SideC.X2 = max.X / scale + derX + offset;
            SideC.Y2 = -min.Y / scale + derY + offset;
            boxSides.Add(SideC);

            Line SideD = new Line();
            SideD.X1 = max.X / scale + derX + offset;
            SideD.Y1 = -min.Y / scale + derY + offset;
            SideD.X2 = min.X / scale + derX - offset;
            SideD.Y2 = -min.Y / scale + derY + offset;
            boxSides.Add(SideD);
            foreach (var item in boxSides)
            {
                item.Stroke = System.Windows.Media.Brushes.Transparent;
            }
            return boxSides;
        }

        //-----Create perpendiculars to given line in given points
        public List<Line> DrawPerp(Line baseWall, List<PointF> points)
        {
            List<Line> lines = new List<Line>();
            double slope = -1 / (GetSlope(baseWall));

            foreach (PointF point in points)
            {
                Point target = new Point();
                target.X = 0;
                target.Y = (int)(point.Y - (slope * point.X));

                Line perpendicular = new Line();
                perpendicular.X1 = target.X;
                perpendicular.Y1 = target.Y;
                perpendicular.X2 = point.X;
                perpendicular.Y2 = point.Y;
                // perpendicular.Stroke = System.Windows.Media.Brushes.Red;
                lines.Add(perpendicular);
            }
            return lines;
        }

        public List<Line> GetPerpendicularsF(Line baseWall, List<PointF> points)
        {
            List<Line> perpList = new List<Line>();

            foreach (PointF point in points)
            {
                PointF target = new PointF();

                double tolerance = 0.000000001;
                if (Math.Abs(GetSlope(baseWall)) < tolerance)
                {
                    target.X = point.X;
                    target.Y = 0;
                }
                else if (double.IsInfinity(GetSlope(baseWall)))
                {
                    target.X = 0;
                    target.Y = point.Y;
                }
                else
                {
                    double slope = -1 / GetSlope(baseWall);
                    target.X = 0;
                    target.Y = (float)(point.Y - (slope * point.X));
                }

                Line perpendicular = new Line();
                perpendicular.X1 = target.X;
                perpendicular.Y1 = target.Y;
                perpendicular.X2 = point.X;
                perpendicular.Y2 = point.Y;
                perpList.Add(perpendicular);
            }
            return perpList;
        }

        //-----DrawWalls Center Line on Canvas by given Start and End points
        public Line DrawCenterLine(Line line)
        {
            line.Stroke = System.Windows.Media.Brushes.Red;
            DoubleCollection dash = new DoubleCollection() { 130, 8, 15, 8 };
            line.StrokeDashArray = dash;
            return line;
        }

        //-----DrawWalls Dashed Line on Canvas by given Start and End points
        public Line DrawDashedLine(System.Windows.Shapes.Line line)
        {
            line.Stroke = System.Windows.Media.Brushes.SteelBlue;
            line.SnapsToDevicePixels = true;
            line.SetValue(RenderOptions.EdgeModeProperty, EdgeMode.Aliased);
            DoubleCollection dash = new DoubleCollection() { 20, 5, 20, 5 };
            line.StrokeDashArray = dash;
            return line;
        }

        
        public List<Point> GetGridPoints(List<Line> listPerpendiculars, List<Line> wallNormals)
        {
            wallNormals.AddRange(listPerpendiculars);
            List<Point> temp = new List<Point>();

            foreach (var normalA in wallNormals)
            {
                foreach (var normalB in wallNormals)
                {
                    if (!normalA.Equals(normalB))
                        temp.Add(GetIntersection(normalA, normalB));
                }
            }

            IEnumerable<Point> distinctPoints = temp.Distinct();
            List<Point> filteredPoints = new List<Point>();
            foreach (var item in distinctPoints)
            {
                if (item.Y != int.MaxValue && item.Y != int.MinValue)
                {
                    filteredPoints.Add(item);
                }
            }
            return filteredPoints;
        }

        public List<PointF> GetGridPointsF(List<Line> listPerpendiculars, List<Line> wallNormals)
        {
            wallNormals.AddRange(listPerpendiculars);
            List<PointF> temp = new List<PointF>();

            foreach (var normalA in wallNormals)
            {
                foreach (var normalB in wallNormals)
                {
                    if (!normalA.Equals(normalB))
                        temp.Add(GetIntersection(normalA, normalB));
                }
            }

            IEnumerable<PointF> distinctPoints = temp.Distinct();
            List<PointF> filteredPoints = new List<PointF>();
            foreach (var item in distinctPoints)
            {
                {
                    filteredPoints.Add(item);
                }
            }
            return filteredPoints;
        }
    }
}

using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Revit4_4
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                Document doc = commandData.Application.ActiveUIDocument.Document;
                List<Level> levelList = GetLevels(doc);

                var walls = CreateWalls(doc, levelList, 10000, 5000);
                var door = CreateDoor(doc, levelList[0], walls[0]);
                List<FamilyInstance> windows = new List<FamilyInstance>();
                for (int i = 1; i < 4; i++)
                {
                    windows.Add(CreateWindow(doc, levelList[0], walls[i]));
                }
                var roof = CreateRoof(doc, levelList[1], walls);
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }

            return Result.Succeeded;
        }

        private RoofBase CreateRoof(Document doc, Level level, List<Wall> walls)
        {
            var roofType = new FilteredElementCollector(doc)
                                .OfClass(typeof(RoofType))
                                .OfType<RoofType>()
                                .Where(x => x.Name.Equals("Типовой - 400мм"))
                                .Where(x => x.FamilyName.Equals("Базовая крыша"))
                                .FirstOrDefault();
            List<XYZ> points = new List<XYZ>();
            foreach (var wall in walls)
            {
                LocationCurve wallCurve = wall.Location as LocationCurve;
                points.Add(wallCurve.Curve.GetEndPoint(1));
            }

            Application app = doc.Application;
            //CurveArray curveArray = app.Create.NewCurveArray();
            //for (int i = 0; i < 4; i++)
            //{
            //    LocationCurve curve = walls[i].Location as LocationCurve;
            //    Line line = Line.CreateBound(curve.Curve.GetEndPoint(0) + points[i], curve.Curve.GetEndPoint(1) + points[i + 1]);
            //    curveArray.Append(line);
            //}

            CurveArray curveArray = new CurveArray();
            double wallHeight = walls[0].get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).AsDouble();
            double dt = walls[0].Width / 2;
            XYZ point1 = new XYZ(points[0].X + dt, points[0].Y - dt, wallHeight);
            XYZ point2 = new XYZ(points[1].X - dt, points[1].Y + dt, wallHeight);
            XYZ highPoint = (points[0] + points[1]) / 2 + new XYZ(dt, 0, points[0].DistanceTo(points[1]) / 2 + wallHeight);
            curveArray.Append(Line.CreateBound(point1, highPoint));
            curveArray.Append(Line.CreateBound(highPoint, point2));

            ExtrusionRoof roof = null;

            using (Transaction tr = new Transaction(doc, "Create ExtrusionRoof"))
            {
                tr.Start();
                ReferencePlane plane = doc.Create.NewReferencePlane2(point1, new XYZ(point1.X, point1.Y, highPoint.Z), new XYZ(highPoint.X, highPoint.Y, point1.Z), doc.ActiveView);
                roof = doc.Create.NewExtrusionRoof(curveArray, plane, level, roofType, 0, points[0].DistanceTo(points[3]) + dt * 2);
                tr.Commit();
            }

            return roof;

        }

        public FamilyInstance CreateDoor(Document doc, Level level, Wall wall)
        {
            var doorType = new FilteredElementCollector(doc)
                                .OfClass(typeof(FamilySymbol))
                                .OfCategory(BuiltInCategory.OST_Doors)
                                .OfType<FamilySymbol>()
                                .FirstOrDefault();

            var wallCurve = wall.Location as LocationCurve;
            XYZ point = (wallCurve.Curve.GetEndPoint(0) + wallCurve.Curve.GetEndPoint(1)) / 2;
            FamilyInstance door = null;
            try
            {
                using (var ts = new Transaction(doc, "Creating of door"))
                {
                    ts.Start();
                    if (!doorType.IsActive)
                        doorType.Activate();

                    door = doc.Create.NewFamilyInstance(point, doorType, wall, level, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                    ts.Commit();
                }
            }
            catch (Exception)
            {

                throw;
            }

            return door;
        }

        public FamilyInstance CreateWindow(Document doc, Level level, Wall wall)
        {
            var windowType = new FilteredElementCollector(doc)
                                .OfClass(typeof(FamilySymbol))
                                .OfCategory(BuiltInCategory.OST_Windows)
                                .OfType<FamilySymbol>()
                                .FirstOrDefault();


            var wallCurve = wall.Location as LocationCurve;
            XYZ point = (wallCurve.Curve.GetEndPoint(0) + wallCurve.Curve.GetEndPoint(1)) / 2;
            point = new XYZ(point.X, point.Y, UnitUtils.ConvertToInternalUnits(1500, DisplayUnitType.DUT_MILLIMETERS));
            FamilyInstance window = null;
            try
            {
                using (var ts = new Transaction(doc, "Creating of window"))
                {
                    ts.Start();
                    if (!windowType.IsActive)
                        windowType.Activate();

                    window = doc.Create.NewFamilyInstance(point, windowType, wall, level, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                    ts.Commit();
                }
            }
            catch (Exception)
            {

                throw;
            }

            return window;
        }

        private List<Wall> CreateWalls(Document doc, List<Level> levelList, double width, double depth)
        {
            width = UnitUtils.ConvertToInternalUnits(width, DisplayUnitType.DUT_MILLIMETERS);
            depth = UnitUtils.ConvertToInternalUnits(depth, DisplayUnitType.DUT_MILLIMETERS);

            double dx = width / 2;
            double dy = depth / 2;

            List<XYZ> points = new List<XYZ>()
            {
                new XYZ(-dx, -dy, 0),
                new XYZ(dx, -dy, 0),
                new XYZ(dx, dy, 0),
                new XYZ(-dx, dy, 0),
                new XYZ(-dx, -dy, 0)
            };

            var walls = new List<Wall>();

            try
            {
                using (var ts = new Transaction(doc, "Creating Walls"))
                {
                    ts.Start();

                    for (int i = 0; i < 4; i++)
                    {
                        Line line = Line.CreateBound(points[i], points[i + 1]);
                        Wall wall = Wall.Create(doc, line, levelList[0].Id, false);
                        walls.Add(wall);
                        wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).Set(levelList[1].Id);
                    }

                    ts.Commit();
                }
            }
            catch (Exception)
            {
                throw;
            }

            return walls;
        }

        private List<Level> GetLevels(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .OfType<Level>()
                .OrderBy(x => x.Name)
                .ToList();
        }
    }
}

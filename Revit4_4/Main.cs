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
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }

            return Result.Succeeded;
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

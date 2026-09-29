using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.GraphicsInterface;

namespace AsdRcSlab
{
    internal class ImrInsertionJig : DrawJig
    {
        private Point3d _currentPoint;
        private readonly ImrCommand.PlotMapInfo _plot;

        public ImrInsertionJig(ImrCommand.PlotMapInfo plot)
        {
            _plot = plot;
            _currentPoint = Point3d.Origin;
        }

        public Point3d CurrentPoint => _currentPoint;

        protected override SamplerStatus Sampler(JigPrompts prompts)
        {
            var opts = new JigPromptPointOptions("\nSpecify insertion point (top-left of T1 map): ");
            opts.UserInputControls = UserInputControls.Accept3dCoordinates;

            var res = prompts.AcquirePoint(opts);
            if (res.Status == PromptStatus.Cancel || res.Status == PromptStatus.Error)
                return SamplerStatus.Cancel;

            if (_currentPoint.DistanceTo(res.Value) < 1e-6)
                return SamplerStatus.NoChange;

            _currentPoint = res.Value;
            return SamplerStatus.OK;
        }

        protected override bool WorldDraw(WorldDraw draw)
        {
            Vector3d disp = _currentPoint - _plot.ReferencePoint;

            // Biały kolor obrysów (index 7 = paper color, biały na czarnym tle AutoCAD)
            draw.SubEntityTraits.Color = 7;

            var frames = new[] { _plot.T1, _plot.T2, _plot.B1, _plot.B2 };
            foreach (var f in frames)
            {
                Point3d p1 = new Point3d(f.Xmin + disp.X, f.Ymin + disp.Y, 0);
                Point3d p2 = new Point3d(f.Xmax + disp.X, f.Ymin + disp.Y, 0);
                Point3d p3 = new Point3d(f.Xmax + disp.X, f.Ymax + disp.Y, 0);
                Point3d p4 = new Point3d(f.Xmin + disp.X, f.Ymax + disp.Y, 0);

                draw.Geometry.WorldLine(p1, p2);
                draw.Geometry.WorldLine(p2, p3);
                draw.Geometry.WorldLine(p3, p4);
                draw.Geometry.WorldLine(p4, p1);
            }

            return true;
        }
    }
}

using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.GraphicsInterface;

namespace AsdRcSlab
{
    internal class ManualTextInsertionJig : DrawJig
    {
        private readonly string   _contents;
        private readonly double   _textHeight;
        private readonly ObjectId _textStyleId;
        private readonly Database _db;
        private Point3d           _currentPoint;

        public ManualTextInsertionJig(string contents, double textHeight, ObjectId textStyleId, Database db)
        {
            _contents    = contents;
            _textHeight  = textHeight;
            _textStyleId = textStyleId;
            _db          = db;
            _currentPoint = Point3d.Origin;
        }

        public Point3d Point => _currentPoint;

        protected override SamplerStatus Sampler(JigPrompts prompts)
        {
            var opts = new JigPromptPointOptions("\nSpecify insertion point for MANUAL text: ");
            opts.UserInputControls = UserInputControls.Accept3dCoordinates
                                   | UserInputControls.NoNegativeResponseAccepted;
            var res = prompts.AcquirePoint(opts);
            if (res.Status == PromptStatus.Cancel || res.Status == PromptStatus.None)
                return SamplerStatus.Cancel;
            if (res.Value.IsEqualTo(_currentPoint))
                return SamplerStatus.NoChange;
            _currentPoint = res.Value;
            return SamplerStatus.OK;
        }

        protected override bool WorldDraw(WorldDraw draw)
        {
            using (var mt = new MText())
            {
                mt.SetDatabaseDefaults(_db);
                mt.Contents    = _contents;
                mt.TextHeight  = _textHeight;
                mt.TextStyleId = _textStyleId;
                mt.Location    = _currentPoint;
                mt.Attachment  = AttachmentPoint.MiddleLeft;
                mt.Width       = 0;
                mt.Color       = Color.FromColorIndex(ColorMethod.ByAci, 1);
                draw.Geometry.Draw(mt);
            }
            return true;
        }
    }
}

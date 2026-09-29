using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;

namespace AsdRcSlab
{
    /// <summary>
    /// Czyta layouty z aktualnego drawingu i wyciąga atrybuty title
    /// block A1-BL z każdego. Bazuje na wzorcu z ExtractA1BLAttributes
    /// (Commands.cs, private — tutaj własna kopia publiczna).
    /// </summary>
    public static class DrawingTitleBlockReader
    {
        private const string TitleBlockName = "A1-BL";

        /// <summary>
        /// Zwraca wszystkie layouty (bez Model) z aktywnego drawingu,
        /// dla których znaleziono blok tytułowy A1-BL.
        /// Posortowane alfabetycznie po LayoutName.
        /// </summary>
        public static List<BbsLayoutInfo> ReadAllLayouts(Document doc)
        {
            var result = new List<BbsLayoutInfo>();
            var db = doc.Database;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var layoutDict = (DBDictionary)tr.GetObject(
                    db.LayoutDictionaryId, OpenMode.ForRead);

                foreach (DBDictionaryEntry entry in layoutDict)
                {
                    var layout = tr.GetObject(entry.Value, OpenMode.ForRead)
                        as Layout;
                    if (layout == null) continue;
                    if (string.Equals(layout.LayoutName, "Model",
                        StringComparison.OrdinalIgnoreCase)) continue;

                    var attrs = ReadLayoutAttributes(tr, layout);
                    if (attrs == null) continue;  // brak A1-BL w tym layout

                    string drawingNumber = null;
                    attrs.TryGetValue("DRAWING_NUMBER", out drawingNumber);

                    result.Add(new BbsLayoutInfo
                    {
                        LayoutName    = layout.LayoutName,
                        DrawingNumber = drawingNumber,
                        Attributes    = attrs
                    });
                }

                tr.Commit();
            }

            // Sort po LayoutName dla deterministycznego porządku w dialogu
            result.Sort((a, b) =>
                string.Compare(a.LayoutName, b.LayoutName,
                    StringComparison.OrdinalIgnoreCase));
            return result;
        }

        private static Dictionary<string, string> ReadLayoutAttributes(
            Transaction tr, Layout layout)
        {
            var btr = (BlockTableRecord)tr.GetObject(
                layout.BlockTableRecordId, OpenMode.ForRead);

            foreach (ObjectId id in btr)
            {
                var br = tr.GetObject(id, OpenMode.ForRead) as BlockReference;
                if (br == null) continue;
                if (!string.Equals(br.Name, TitleBlockName,
                    StringComparison.OrdinalIgnoreCase)) continue;

                var attrs = new Dictionary<string, string>(
                    StringComparer.OrdinalIgnoreCase);
                foreach (ObjectId attId in br.AttributeCollection)
                {
                    var att = tr.GetObject(attId, OpenMode.ForRead)
                        as AttributeReference;
                    if (att == null) continue;
                    attrs[att.Tag] = att.TextString;
                }
                return attrs;
            }
            return null;  // brak A1-BL w tym layoutcie
        }
    }
}

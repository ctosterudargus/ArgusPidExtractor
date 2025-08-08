using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

[assembly: CommandClass(typeof(ArgusPidExtractor.Commands))]

namespace ArgusPidExtractor
{
    public class BlockInstance
    {
        public string Name { get; set; }
        public string SourceDrawing { get; set; }
        public bool FromXref { get; set; }
        public string XrefPath { get; set; }
        public string Layer { get; set; }
        public double[] PositionWcs { get; set; }
        public Dictionary<string, string> Attributes { get; set; }
        public List<string> ShownOnSheets { get; set; } = new();
    }

    public class Commands : IExtensionApplication
    {
        public void Initialize() { }
        public void Terminate() { }

        [CommandMethod("ARGUS_EXPORT_BLOCKS")]
        public void ExportBlocks()
        {
            try
            {
                string inDwg = Environment.GetEnvironmentVariable("ARGUS_IN_DWG");
                string outPath = Environment.GetEnvironmentVariable("ARGUS_OUT_JSON");
                if (string.IsNullOrWhiteSpace(inDwg) || string.IsNullOrWhiteSpace(outPath)) return;

                string filtersEnv = Environment.GetEnvironmentVariable("ARGUS_FILTERS");
                var filters = BuildFilters(filtersEnv);
                var results = new List<BlockInstance>();
                string rootPath = Path.GetFullPath(inDwg);

                using (var db = new Database(false, true))
                {
                    db.ReadDwgFile(rootPath, FileOpenMode.OpenForReadAndAllShare, true, "");
                    db.CloseInput(true);
                    HostApplicationServices.WorkingDatabase = db;

                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                        var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                        TraverseBtr(db, ms, tr, Matrix3d.Identity, false, null, rootPath, results, filters, new HashSet<string>(StringComparer.OrdinalIgnoreCase), new HashSet<ObjectId>());
                        tr.Commit();
                    }
                }

                var dir = Path.GetDirectoryName(outPath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir)) Directory.CreateDirectory(dir);
                var json = JsonSerializer.Serialize(results, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(outPath, json);
            }
            catch
            {
                // fail silently
            }
        }

        private static List<Regex> BuildFilters(string env)
        {
            if (string.IsNullOrWhiteSpace(env)) return null;
            var parts = env.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            var list = new List<Regex>();
            foreach (var p in parts)
            {
                var pattern = "^" + Regex.Escape(p).Replace("\\*", ".*") + "$";
                list.Add(new Regex(pattern, RegexOptions.IgnoreCase));
            }
            return list;
        }

        private static void TraverseBtr(Database db, BlockTableRecord btr, Transaction tr, Matrix3d transform, bool fromXref, string xrefPath, string sourceDrawing, List<BlockInstance> results, List<Regex> filters, HashSet<string> visitedXrefs, HashSet<ObjectId> btrStack)
        {
            if (!btrStack.Add(btr.ObjectId)) return;

            foreach (ObjectId id in btr)
            {
                if (!id.ObjectClass.IsDerivedFrom(RXClass.GetClass(typeof(BlockReference)))) continue;

                var br = (BlockReference)tr.GetObject(id, OpenMode.ForRead);
                var def = (BlockTableRecord)tr.GetObject(br.BlockTableRecord, OpenMode.ForRead);
                if (def.IsLayout || def.IsAnonymous || def.Name.StartsWith("*")) continue;

                var instanceTransform = br.BlockTransform * transform;
                var pos = Point3d.Origin.TransformBy(instanceTransform);

                var attrs = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                if (br.AttributeCollection != null)
                {
                    foreach (ObjectId attId in br.AttributeCollection)
                    {
                        var ar = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                        if (ar != null) attrs[ar.Tag] = ar.TextString;
                    }
                }

                bool isXref = def.IsFromExternalReference;
                bool nextFromXref = fromXref || isXref;
                string nextXrefPath = nextFromXref ? (fromXref ? xrefPath : Path.GetFullPath(def.PathName)) : null;

                if (filters == null || filters.Any(r => r.IsMatch(def.Name)))
                {
                    results.Add(new BlockInstance
                    {
                        Name = def.Name,
                        SourceDrawing = sourceDrawing,
                        FromXref = nextFromXref,
                        XrefPath = nextFromXref ? nextXrefPath : null,
                        Layer = br.Layer,
                        PositionWcs = new[] { pos.X, pos.Y, pos.Z },
                        Attributes = attrs,
                        ShownOnSheets = new List<string>()
                    });
                }

                if (isXref)
                {
                    string path = Path.GetFullPath(def.PathName);
                    if (!visitedXrefs.Contains(path))
                    {
                        visitedXrefs.Add(path);
                        try
                        {
                            using (var xdb = new Database(false, true))
                            {
                                xdb.ReadDwgFile(path, FileOpenMode.OpenForReadAndAllShare, true, "");
                                xdb.CloseInput(true);
                                using (var xtr = xdb.TransactionManager.StartTransaction())
                                {
                                    var xbt = (BlockTable)xtr.GetObject(xdb.BlockTableId, OpenMode.ForRead);
                                    var xms = (BlockTableRecord)xtr.GetObject(xbt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
                                    TraverseBtr(xdb, xms, xtr, instanceTransform, true, path, sourceDrawing, results, filters, visitedXrefs, new HashSet<ObjectId>());
                                    xtr.Commit();
                                }
                            }
                        }
                        catch { }
                    }
                }
                else
                {
                    TraverseBtr(db, def, tr, instanceTransform, nextFromXref, nextXrefPath, sourceDrawing, results, filters, visitedXrefs, btrStack);
                }
            }

            btrStack.Remove(btr.ObjectId);
        }
    }
}

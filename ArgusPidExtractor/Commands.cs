using System;
using System.Collections.Generic;
using System.IO;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;

[assembly: CommandClass(typeof(ArgusPidExtractor.Commands))]

namespace ArgusPidExtractor
{
    public class Commands : IExtensionApplication
    {
        public void Initialize() { }
        public void Terminate() { }

        [CommandMethod("ARGUS_EXPORT_BLOCKS")]
        public void ExportBlocks()
        {
            string inDwg = Environment.GetEnvironmentVariable("ARGUS_IN_DWG");
            string outPath = Environment.GetEnvironmentVariable("ARGUS_OUT_JSON");
            if (string.IsNullOrWhiteSpace(inDwg) || string.IsNullOrWhiteSpace(outPath)) return;

            var lines = new List<string>();
            using (var db = new Database(false, true))
            {
                db.ReadDwgFile(inDwg, FileOpenMode.OpenForReadAndAllShare, true, "");
                db.CloseInput(true);
                HostApplicationServices.WorkingDatabase = db;

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                    foreach (ObjectId id in ms)
                    {
                        if (!id.ObjectClass.IsDerivedFrom(RXClass.GetClass(typeof(BlockReference)))) continue;
                        var br = (BlockReference)tr.GetObject(id, OpenMode.ForRead);
                        var btr = (BlockTableRecord)tr.GetObject(br.BlockTableRecord, OpenMode.ForRead);
                        if (btr.IsLayout) continue;
                        lines.Add($"{btr.Name} @ ({br.Position.X},{br.Position.Y},{br.Position.Z})");
                    }
                    tr.Commit();
                }
            }

            var dir = Path.GetDirectoryName(outPath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir)) Directory.CreateDirectory(dir);
            var txt = Path.ChangeExtension(outPath, ".txt");
            File.WriteAllLines(txt, lines);
        }
    }
}

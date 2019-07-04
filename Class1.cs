using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Colors;
using CAD = Autodesk.AutoCAD.ApplicationServices.Core;

[assembly: CommandClass(typeof(ConvertTxt.Class1))]

namespace ConvertTxt
{
    public class Class1
    {
        [CommandMethod("ATTDEFTOTEXT")]
        public void AttDefToText()
        {
            Document CurrentDoc = CAD.Application.DocumentManager.MdiActiveDocument; // Obtém o documento actual
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor; // define ed como o Editor

            TypedValue[] filterlist = new TypedValue[1];
            filterlist[0] = new TypedValue(100);

            SelectionFilter filter = new SelectionFilter(filterlist);

            PromptSelectionOptions PSOptions = new PromptSelectionOptions();
            PromptSelectionResult acSSPrompt = ed.GetSelection(PSOptions, filter);

            SelectionSet selSet = acSSPrompt.Value;

            ObjectId[] idArray = selSet.GetObjectIds();

            Database db = CurrentDoc.Database;

            Point3d pos = new Point3d(0, 0, 0);
            double wf = 0;
            AttributeDefinition att = null;
            string tstg = null;
            Color col = null;
            double hei = 0;
            ObjectId sty;
            string lay = null;

            foreach (ObjectId id in idArray)
            {
                RXClass objclass = id.ObjectClass;

                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    att = trans.GetObject(id, OpenMode.ForWrite) as AttributeDefinition;

                    // save the properties of the AttributeDefinition
                    tstg = att.TextString;
                    pos = att.Position;
                    wf = att.WidthFactor;
                    col = att.Color;
                    hei = att.Height;
                    sty = att.TextStyleId;
                    lay = att.Layer;

                    // Open the Block table for read
                    BlockTable acBlkTbl;
                    acBlkTbl = trans.GetObject(db.BlockTableId,
                                                    OpenMode.ForRead) as BlockTable;

                    // Open the Block table record Model space for write
                    BlockTableRecord acBlkTblRec;
                    acBlkTblRec = trans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                                        OpenMode.ForWrite) as BlockTableRecord;

                    // Erase the object
                    att.Erase(true);

                    // Commit the changes and dispose of the transaction
                    trans.Commit();
                }

                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    DBText dbt = new DBText();

                    // Define properties of the DBText
                    dbt.TextString = tstg;
                    dbt.Position = pos;
                    dbt.WidthFactor = wf;
                    dbt.Color = col;
                    dbt.Height = hei;
                    dbt.TextStyleId = sty;
                    dbt.Layer = lay;

                    // Open the Block table for read
                    BlockTable acBlkTbl;
                    acBlkTbl = trans.GetObject(db.BlockTableId,
                                                    OpenMode.ForRead) as BlockTable;

                    // Open the Block table record Model space for write
                    BlockTableRecord acBlkTblRec;
                    acBlkTblRec = trans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                                        OpenMode.ForWrite) as BlockTableRecord;

                    // Add the new object to Model space and the transaction
                    acBlkTblRec.AppendEntity(dbt);
                    trans.AddNewlyCreatedDBObject(dbt, true);

                    // Draw the new object
                    dbt.Draw();

                    // Commit the changes and dispose of the transaction
                    trans.Commit();
                }
            }
        }
    }
}

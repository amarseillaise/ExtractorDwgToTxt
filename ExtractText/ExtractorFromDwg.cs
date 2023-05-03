using System;
using System.IO;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Application = Autodesk.AutoCAD.ApplicationServices.Core.Application;

namespace ExtractText
{
    public class Class1
    {
        private const double TOLERANCE = 0.1;
        
        public string GetFolderPath()
        {
            FolderBrowserDialog folder = new FolderBrowserDialog();
            folder.ShowDialog();
            return folder.SelectedPath;
        }
        
        [CommandMethod("ExtractText")]
        public void ExtractText()
        {
            string selectedFolder = GetFolderPath();
            string preResultTxtPath = selectedFolder + "\\PreResult\\";
            var allFiles = Directory.GetFiles(selectedFolder);
            
            foreach (string fileName in allFiles)
            {
                Document doc = Application.DocumentManager.Open(fileName, true);

                Database db = doc.Database;
                Transaction tr = db.TransactionManager.StartTransaction();

                try
                {
                    // Get the block table and model space
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

                    // Iterate through the entities in model space
                    
                    StreamWriter f = new StreamWriter(preResultTxtPath + Path.GetFileName(fileName) + ".txt", false);
                    foreach (ObjectId id in btr)
                    {
                        Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;

                        // Check if the entity is a text object
                        if (ent is DBText txt)
                        {
                            // Get the text string, position, and other properties
                            string text = txt.TextString;
                            Point3d position = txt.Position;
                            if (!text.Trim().Equals(""))
                            {
                                double x = position.X;
                                double y = position.Y;
                                double height = txt.Height;
                                string result = $"{text}||{x}||{y}||{height}";
                                f.WriteLine(result);
                            }
                        }
                        // Check if the entity is a Line object
                        if (ent is Line line && (Math.Abs(line.Length - 210.000) < TOLERANCE || 
                                                 Math.Abs(line.Length - 197.2500) < TOLERANCE || 
                                                 Math.Abs(line.Length - 297.000) < TOLERANCE ||
                                                 Math.Abs(line.Length - 74.0259) < TOLERANCE ||
                                                 Math.Abs(line.Length - 218.2594) < 1))
                        {
                            Point3d startPoint = line.StartPoint;
                            Point3d endPoint = line.EndPoint;
                            double yEndPoint = line.EndPoint.Y;
                            double yStartPoint = line.StartPoint.Y;
                            
                            if (Math.Abs(line.Length - 197.25000) < TOLERANCE)
                            {
                                if (line.Delta.Y > 0) yStartPoint -= 12.75000;
                                else yEndPoint -= 12.75000;
                            }
                            
                            if (Math.Abs(line.Length - 74.0259) < TOLERANCE)
                            {
                                if (line.Delta.Y > 0) yStartPoint -= 135.9741;
                                else yEndPoint -= 135.9741;
                            }
                            
                            string result = $"$ListLine||{startPoint.X}||{endPoint.X}||{yStartPoint}||{yEndPoint}||{line.Delta.X}||{line.Delta.Y}";
                            f.WriteLine(result);
                        }
                        // Check if the entity is a Polyline object
                        if (ent is Polyline pl && Math.Abs(pl.Length - 1014.00000) < TOLERANCE)
                        {
                            Point3d startPoint = pl.StartPoint;

                            string result =
                                $"$ListLine||{startPoint.X}||{startPoint.X}||{startPoint.Y}||{startPoint.Y + 210.00000}||{0}||{1}" +
                                "\n" +
                                $"$ListLine||{startPoint.X}||{startPoint.X + 297.00000}||{startPoint.Y + 210.00000}||{startPoint.Y + 210.00000}||{1}||{0}" +
                                "\n" +
                                $"$ListLine||{startPoint.X + 297.00000}||{startPoint.X + 297.00000}||{startPoint.Y}||{startPoint.Y + 210.00000}||{0}||{1}" +
                                "\n" +
                                $"$ListLine||{startPoint.X}||{startPoint.X + 297.00000}||{startPoint.Y}||{startPoint.Y}||{1}||{0}";
                            
                            f.WriteLine(result);
                        }
                    }
                    f.Close();

                    // Commit the transaction
                    tr.Commit();
                }
                catch (SystemException ex)
                {
                    // Handle any exceptions
                    Console.WriteLine(ex.Message);
                }
                finally
                {
                    // Dispose of the transaction and document
                    tr.Dispose();
                    doc.CloseAndDiscard();
                }
            }
            MessageBox.Show("Done");
        }
    }
}
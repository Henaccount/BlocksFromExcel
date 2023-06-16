using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Excel = Microsoft.Office.Interop.Excel;

namespace BlocksFromExcel
{
    public class BlocksFromExcel
    {
        [CommandMethod("BlocksFromExcel")]
        public static void ImportFromExcel()
        {
            // Get the active AutoCAD document and database
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor editor = doc.Editor;

            // Prompt user to select an Excel file
            PromptOpenFileOptions options = new PromptOpenFileOptions("Select Excel file");
            options.Filter = "Excel Files (*.xlsx)|*.xlsx";
            PromptFileNameResult result = editor.GetFileNameForOpen(options);
            if (result.Status != PromptStatus.OK)
            {
                editor.WriteMessage("\nInvalid file selection.");
                return;
            }

            // Read Excel data and insert blocks into the drawing
            string excelFilePath = result.StringResult;
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(excelFilePath);
            Excel.Worksheet worksheet = workbook.ActiveSheet;
            Excel.Range usedRange = worksheet.UsedRange;

            try
            {
                int rowCount = usedRange.Rows.Count;
                int columnCount = usedRange.Columns.Count;

                // Iterate over each row in the used range
                for (int row = 2; row <= rowCount; row++) // Assuming header row is skipped
                {
                    // Read blockname and tag from the Excel columns
                    string blockName = usedRange.Cells[row, 1].Value.ToString();
                    string tag = usedRange.Cells[row, 2].Value.ToString();
                    ObjectId curblock = new ObjectId();
                    // Insert the block into the drawing
                    using (Transaction transaction = db.TransactionManager.StartTransaction())
                    {
                        BlockTable blockTable = transaction.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                        BlockTableRecord modelSpace = transaction.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                        //editor.WriteMessage("\nblockName: " + blockName);
                        //editor.WriteMessage("\ntag: " + tag);


                        using (BlockReference blockRef = new BlockReference(new Point3d(0, row * -5000, 0), blockTable[blockName]))
                        {
                            curblock = modelSpace.AppendEntity(blockRef);
                            transaction.AddNewlyCreatedDBObject(blockRef, true);

                        }

                        transaction.Commit();


                    }

                    editor.Command("attsync", "S", curblock, "Y");

                    using (Transaction transaction = db.TransactionManager.StartTransaction())
                    {
                        BlockReference blockinstRef = transaction.GetObject(curblock, OpenMode.ForWrite) as BlockReference;
                        //AttributeReference attributeRef = transaction.GetObject(nestedResult.ObjectId, OpenMode.ForWrite) as AttributeReference;

                        foreach (ObjectId attributeId in blockinstRef.AttributeCollection)
                        {
                            AttributeReference attribute = transaction.GetObject(attributeId, OpenMode.ForWrite) as AttributeReference;
                            editor.WriteMessage("\nattribute.Tag: " + attribute.Tag);
                            if (attribute.Tag == "FDSA")
                            {
                                editor.WriteMessage("\nfound attribute and want to update with: " + tag);
                                attribute.TextString = tag;
                            }
                        }



                        transaction.Commit();
                    }
                }

                editor.WriteMessage($"\n{rowCount - 1} blocks inserted into the drawing.");
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage($"\nAn error occurred: {ex.Message + "\n" + ex.StackTrace}");
            }
            finally
            {
                workbook.Close();
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }
        }
    }
}

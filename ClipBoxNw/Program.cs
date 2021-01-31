using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using System.Collections.Generic;
using Autodesk.ProcessPower.ProjectManager;
using System;
using System.IO;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.ApplicationServices;
using System.Reflection;
using System.Linq;
using System.Diagnostics;

[assembly: CommandClass(typeof(ClipBoxNw.Program))]

namespace ClipBoxNw
{
    public class Program
    {

        [CommandMethod("ClipBoxNwWriteBatchOnly", CommandFlags.UsePickSet)]
        public static void ClipBoxNwWriteBatchOnly()
        { doclipboxnw("ClipBoxNwWriteBatchOnly"); }

        [CommandMethod("ClipBoxNwReset", CommandFlags.UsePickSet)]
        public static void ClipBoxNwReset()
        { doclipboxnw("ClipBoxNwReset"); }

        [CommandMethod("ClipBoxNw", CommandFlags.UsePickSet)]
        public static void ClipBoxNw()
        { doclipboxnw("ClipBoxNw"); }

        public static void doclipboxnw(string command)
        {
            Helper.Initialize();

            try

            {

                PromptSelectionResult selectionRes = Helper.oEditor.SelectImplied();

                // If there's no pickfirst set available...

                if (selectionRes.Status == PromptStatus.Error)

                {

                    // ... ask the user to select entities
                    if (!command.Equals("ClipBoxNwReset"))
                    {
                        PromptSelectionOptions selectionOpts = new PromptSelectionOptions();

                        selectionOpts.SingleOnly = true;
                        selectionOpts.MessageForAdding = "\n(Unload NWD first, press esc now if still loaded) Select object to define area of interest..";

                        if (command.Equals("ClipBoxNwWriteBatchOnly"))
                        {
                            selectionOpts.SingleOnly = false;
                            selectionOpts.MessageForAdding = "\nSelect all objects defining areas and then press enter..";

                        }
                        selectionRes = Helper.oEditor.GetSelection(selectionOpts);
                    }
                }

                else

                {

                    // If there was a pickfirst set, clear it

                    Helper.oEditor.SetImpliedSelection(new ObjectId[0]);

                }

                // If the user has not cancellHelper.oEditor...

                if (selectionRes.Status == PromptStatus.OK || command.Equals("ClipBoxNwReset"))

                {

                    // ... take the selected objects one by one




                    using (Transaction tr = Helper.ActiveDocument.Database.TransactionManager.StartTransaction())
                    {

                        string documentspath = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                        string batpath = documentspath + "\\" + "clipboxnw.bat";
                        string commandline = "";
                        Point3d min = new Point3d(0, 0, 0);
                        Point3d max = new Point3d(0, 0, 0);
                        ObjectId[] objIds = new ObjectId[] { };

                        if (!command.Equals("ClipBoxNwReset")) { objIds = selectionRes.Value.GetObjectIds(); }
                        else objIds = new ObjectId[] { new ObjectId() };

                        int counter = 0;

                        foreach (ObjectId objId in objIds)
                        {
                            string suffix = "";
                            if (counter > 0) suffix = "_" + counter.ToString();
                            ++counter;                           

                            if (!command.Equals("ClipBoxNwReset"))
                            {
                                Entity ent = (Entity)tr.GetObject(objId, OpenMode.ForRead);

                                min = ent.Bounds.Value.MinPoint;
                                max = ent.Bounds.Value.MaxPoint;

                                ent.Dispose();
                            }

                            double conversionfactor = UnitsConverter.GetConversionFactor(Helper.oDatabase.Insunits, UnitsValue.Meters);

                            commandline += "\"D:\\Program Files\\Autodesk\\Navisworks Manage 2020\\Roamer.exe\" -NoGui -OpenFile";
                            commandline += " \"" + documentspath + "\\" + "CoordinationModel.nwd\"";
                            commandline += " -ExecuteAddInPlugin \"ClipBox.AClipBox.ADSK\"";
                            commandline += " \"" + documentspath + "\\" + "CoordinationModel" + suffix + ".nwd\"";
                            commandline += " " + Math.Round(min.X * conversionfactor, 3).ToString().Replace("-", "neg");
                            commandline += " " + Math.Round(min.Y * conversionfactor, 3).ToString().Replace("-", "neg");
                            commandline += " " + Math.Round(min.Z * conversionfactor, 3).ToString().Replace("-", "neg");
                            commandline += " " + Math.Round(max.X * conversionfactor, 3).ToString().Replace("-", "neg");
                            commandline += " " + Math.Round(max.Y * conversionfactor, 3).ToString().Replace("-", "neg");
                            commandline += " " + Math.Round(max.Z * conversionfactor, 3).ToString().Replace("-", "neg");
                            commandline += "\r\n";
                        }

                        File.WriteAllText(batpath, commandline);


                        if (!command.Equals("ClipBoxNwWriteBatchOnly"))
                        {
                            Process ExternalProcess = new Process();
                            ExternalProcess.StartInfo.FileName = batpath;
                            ExternalProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                            ExternalProcess.Start();
                            ExternalProcess.WaitForExit();
                        }
                        /*var btr = (BlockTableRecord)tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(Helper.oDatabase), OpenMode.ForRead);

                        foreach (ObjectId id in btr)
                        {
                            //Entity ente = (Entity)tr.GetObject(id, OpenMode.ForRead);
                            DBObject ente = (DBObject)tr.GetObject(id, OpenMode.ForRead);
                            Helper.oEditor.WriteMessage("\nmessage: " + ente.GetType().ToString());
                            if (ente.GetType().ToString().EndsWith("ImpEntity"))
                            {
                                ((AcBIMUnderlayDbx)ente).Unload(false);
                                //ShowType(id);
                            }

                        }*/
                        string logtext = "";
                        string logfilepath = documentspath + "\\" + "CoordinationModel.nwd.txt";
                        if (File.Exists(logfilepath))
                        {
                            logtext = File.ReadAllText(logfilepath);
                        }

                        if (command.Equals("ClipBoxNwWriteBatchOnly"))
                            Helper.oEditor.WriteMessage("\nclipboxnw.bat was created in the documents folder. Edit clipboxnw.bat for further use.");
                        else {
                            if (!logtext.Equals(""))
                                Helper.oEditor.WriteMessage("\n" + logtext);
                            else
                                Helper.oEditor.WriteMessage("\nCoordinationModel.nwd was updated. Now reload CoordinationModel.nwd");
                        }
                        tr.Commit();
                    }

                }

            }

            catch (Autodesk.AutoCAD.Runtime.Exception ex)

            {

                Helper.oEditor.WriteMessage(ex.Message);

            }

            Helper.Terminate();
        }

        //https://adndevblog.typepad.com/autocad/2016/07/unload-all-the-pdf-underlay.html
        //https://adndevblog.typepad.com/autocad/2016/11/how-to-retrieve-naviswork-model-nwd-or-bim-underlay-information-in-autocad-through-programming.html
        public static void ShowType(ObjectId id) //unrelated method (test)
        {
            //Helper.Initialize();
            Document dwg = Helper.ActiveDocument;

            using (Transaction t = dwg.TransactionManager.StartTransaction())
            {
                DBObject x = t.GetObject(id, OpenMode.ForRead);
                Type type = x.GetType();

                Helper.oEditor.WriteMessage(string.Format("\nType: '{0}'", type.ToString()));
                Helper.oEditor.WriteMessage(string.Format("\nIsPublic: '{0}'", type.IsPublic));
                Helper.oEditor.WriteMessage(string.Format("\nAssembly: '{0}'\n", type.Assembly.FullName));
                Helper.oEditor.WriteMessage(new string('*', 30));
                Helper.oEditor.WriteMessage("\n");
                PropertyInfo[] props = type.GetProperties().OrderBy(n => n.Name).ToArray();
                Dictionary<string, object> dict = new Dictionary<string, object>();
                foreach (PropertyInfo item in props)
                {
                    object value;
                    try
                    {
                        value = item.GetValue(x, null);
                    }
                    catch (System.Exception e)
                    {
                        value = string.Format("Exception: '{0}'", e.Message);
                    }
                    dict.Add(string.Format("{0} [{1}]", item.Name, item.PropertyType), value);

                    Helper.oEditor.WriteMessage(string.Format("\n\tProperty: '{0}';\tType: '{1}';\tValue.ToString: '{2}'", item.Name, item.PropertyType, value));
                }
            }

            //Helper.Terminate();
        }
    }
}

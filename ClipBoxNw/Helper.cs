//
// (C) Copyright 2013 by Autodesk, Inc.
//
// Permission to use, copy, modify, and distribute this software in
// object code form for any purpose and without fee is hereby granted,
// provided that the above copyright notice appears in all copies and
// that both that copyright notice and the limited warranty and
// restricted rights notice below appear in all supporting
// documentation.
//
// AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS.
// AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  AUTODESK, INC.
// DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
// UNINTERRUPTED OR ERROR FREE.
//
// Use, duplication, or disclosure by the U.S. Government is subject to
// restrictions set forth in FAR 52.227-19 (Commercial Computer
// Software - Restricted Rights) and DFAR 252.227-7013(c)(1)(ii)
// (Rights in Technical Data and Computer Software), as applicable.
//

using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.ApplicationServices;



using Autodesk.ProcessPower.ProjectManager;
using PlantApp = Autodesk.ProcessPower.PlantInstance.PlantApplication;
using Autodesk.AutoCAD.EditorInput;



namespace ClipBoxNw
{


    public class Helper
    {
        //public static PnIdProject PnIdProject { get; set; }
        public static Project PlantProject { get; set; }
        public static Document ActiveDocument { get; set; }
        public static Database oDatabase { get; set; }
        public static Editor oEditor { get; set; }






        public static bool Initialize()
        {
            if (PlantApp.CurrentProject == null)
                return false;

            Helper.PlantProject = PlantApp.CurrentProject.ProjectParts["Piping"];
            //Helper.PnIdProject = (PnIdProject)PlantApp.CurrentProject.ProjectParts["PnId"];
            //Helper.ActiveDataLinksManager = Helper.PnIdProject.DataLinksManager;

            Helper.ActiveDocument = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Helper.oDatabase = Helper.ActiveDocument.Database;
            Helper.oEditor = Helper.ActiveDocument.Editor;
            return true;
        }

        public static void Terminate()
        {
            Helper.PlantProject = null;
            //Helper.PnIdProject = null;

            Helper.ActiveDocument = null;
            Helper.oDatabase = null;
            Helper.oEditor = null;
        }



    }
}

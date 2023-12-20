using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace PackAndGoCSharp.csproj
{
    partial class SolidWorksMacro
    {
        public void Main()
        {
            ModelDoc2 swModelDoc = default(ModelDoc2);
            ModelDocExtension swModelDocExt = default(ModelDocExtension);
            PackAndGo swPackAndGo = default(PackAndGo);
            string openFile = null;
            bool status = false;
            int warnings = 0;
            int errors = 0;
            int i = 0;
            int namesCount = 0;
            string myPath = null;
            int[] statuses = null;

            // Open assembly
            openFile = "C:\\Users\\fs136650\\OneDrive - First Solar\\Desktop\\PackNGor Test\\Assem1.SLDASM";
            swModelDoc = (ModelDoc2)swApp.OpenDoc6(openFile, (int)swDocumentTypes_e.swDocASSEMBLY, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);
            swModelDocExt = (ModelDocExtension)swModelDoc.Extension;

            // Get Pack and Go object
            Debug.Print("Pack and Go");
            swPackAndGo = (PackAndGo)swModelDocExt.GetPackAndGo();

            // Get number of documents in assembly
            namesCount = swPackAndGo.GetDocumentNamesCount();
            Debug.Print("  Number of model documents: " + namesCount);


            // Include any drawings, SOLIDWORKS Simulation results, and SOLIDWORKS Toolbox components
            swPackAndGo.IncludeDrawings = true;
            Debug.Print(" Include drawings: " + swPackAndGo.IncludeDrawings);
            swPackAndGo.IncludeSimulationResults = true;
            Debug.Print(" Include SOLIDWORKS Simulation results: " + swPackAndGo.IncludeSimulationResults);
            swPackAndGo.IncludeToolboxComponents = true;
            Debug.Print(" Include SOLIDWORKS Toolbox components: " + swPackAndGo.IncludeToolboxComponents);

            // Get current paths and filenames of the assembly's documents
            object fileNames;
            object[] pgFileNames = new object[namesCount - 1];
            status = swPackAndGo.GetDocumentNames(out fileNames);
            pgFileNames = (object[])fileNames;

            Debug.Print("");
            Debug.Print("  Current path and filenames: ");
            if ((pgFileNames != null))
            {
                for (i = 0; i <= pgFileNames.GetUpperBound(0); i++)
                {
                    Debug.Print("    The path and filename is: " + pgFileNames[i]);
                }
            }

            // Get current save-to paths and filenames of the assembly's documents
            object pgFileStatus;
            status = swPackAndGo.GetDocumentSaveToNames(out fileNames, out pgFileStatus);
            pgFileNames = (object[])fileNames;
            Debug.Print("");
            Debug.Print("  Current default save-to filenames: ");
            if ((pgFileNames != null))
            {
                for (i = 0; i <= pgFileNames.GetUpperBound(0); i++)
                {
                    Debug.Print("   The path and filename is: " + pgFileNames[i]);
                }
            }

            // Set folder where to save the files
            myPath = "C:\\temp\\";
            status = swPackAndGo.SetSaveToName(true, myPath);

            // Flatten the Pack and Go folder structure; save all files to the root directory
            swPackAndGo.FlattenToSingleFolder = true;

            // Add a prefix and suffix to the filenames
            swPackAndGo.AddPrefix = "SW_";
            swPackAndGo.AddSuffix = "_PackAndGo";

            // Verify document paths and filenames after adding prefix and suffix
            object getFileNames;
            object getDocumentStatus;
            string[] pgGetFileNames = new string[namesCount - 1];

            status = swPackAndGo.GetDocumentSaveToNames(out getFileNames, out getDocumentStatus);
            pgGetFileNames = (string[])getFileNames;
            Debug.Print("");
            Debug.Print("  My Pack and Go path and filenames after adding prefix and suffix: ");
            for (i = 0; i <= namesCount - 1; i++)
            {
                Debug.Print("    My path and filename is: " + pgGetFileNames[i]);
            }

            // Pack and Go
            statuses = (int[])swModelDocExt.SavePackAndGo(swPackAndGo);

        }


        /// <summary>
        /// The SldWorks swApp variable is pre-assigned for you.
        /// </summary>

        public SldWorks swApp;

    }
}
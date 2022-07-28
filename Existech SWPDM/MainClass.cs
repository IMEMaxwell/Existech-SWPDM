using System;
using EPDM.Interop.epdm;
using System.Windows.Forms;
using System.IO;
using System.Collections.Generic;

namespace Existech_SWPDM
{
    public class MainClass : IEdmAddIn5
    {
        IEdmVault5 PoVault = new EdmVault5();

        public void GetAddInInfo(ref EdmAddInInfo poInfo, IEdmVault5 poVault, IEdmCmdMgr5 poCmdMgr)
        {
            //Specify information to display in the add-in's Properties dialog box
            poInfo.mbsAddInName = "Existech SWPDM Add-in";
            poInfo.mbsCompany = "IME Technology";
            poInfo.mbsDescription = "Customize Add-in";
            poInfo.mlAddInVersion = 1;

            //Specify the minimum required version of SolidWorks PDM Professional
            poInfo.mlRequiredVersionMajor = 6;
            poInfo.mlRequiredVersionMinor = 4;

            // Notify the add-in when state change
            poCmdMgr.AddHook(EdmCmdType.EdmCmd_PreState);
            poCmdMgr.AddHook(EdmCmdType.EdmCmd_PostState);
            //poCmdMgr.AddHook(EdmCmdType.EdmCmd_CardInput);

            // Customize command to assign multiple file
            // poCmdMgr.AddCmd(1000, "Assign File Type", (int)EdmMenuFlags.EdmMenu_OnlyFiles, "Assign variable to file(s)", "Assign File Type", 0, 99);

            // Past vault variable to public
            PoVault = poVault;
        }

        //readonly Microsoft.VisualBasic.CompilerServices.StaticLocalInitFlag static_OnCmd_VariableChangeInProgress_Init = new Microsoft.VisualBasic.CompilerServices.StaticLocalInitFlag();

        bool static_OnCmd_VariableChangeInProgress;
        public void OnCmd(ref EdmCmd poCmd, ref EdmCmdData[] ppoData)
        {
            try
            {
                #region Declare Variables
                // Declare variables here
                string AffectedFileNames = "";
                string FileName = "";
                string ext = "";
                string curConfig = "";
                string FileTypeString = null;
                string SerialNbrValueValue;
                string destSubFolder = "";
                string selectedFileType = "";
                object FileTypeObj = null;
                bool GetVarSuccess = false;
                int fileCount = 0;
                IEdmFolder9 Folder = default(IEdmFolder9);
                IEdmFolder5 dest = default(IEdmFolder5);
                IEdmSerNoGen7 SerialNbrs;
                IEdmSerNoValue SerialNbrValue = default(IEdmSerNoValue);
                IEdmFile10 aFile = default(IEdmFile10);
                IEdmVariableMgr5 variableMgr = default(IEdmVariableMgr5);
                IEdmEnumeratorVariable8 EnumVarObj = default(IEdmEnumeratorVariable8);
                IEdmStrLst5 configList = default(IEdmStrLst5);
                IEdmPos5 pos = default(IEdmPos5);
                IEdmBatchChangeState3 batchChanger = default(IEdmBatchChangeState3);
                #endregion

                // Get root folder
                Folder = (IEdmFolder9)PoVault.RootFolder;

                // Instantiate variable manager
                variableMgr = (IEdmVariableMgr5)PoVault;

                switch (poCmd.meCmdType)
                {
                    // hook this add in to after transition
                    case EdmCmdType.EdmCmd_PreState:

                        #region Pre State

                        foreach (EdmCmdData AffectedFile in ppoData)
                        {
                            // If the state is "Released"
                            if (AffectedFile.mbsStrData2 == "Released")
                            {
                                #region Get File information
                                // Get affected full file name
                                AffectedFileNames = AffectedFile.mbsStrData1;

                                // Get the name with extension
                                FileName = Path.GetFileName(AffectedFile.mbsStrData1);

                                // Get the extension ".sldprt"
                                ext = Path.GetExtension(AffectedFileNames);
                                #endregion

                                #region Process Files
                                // If affectedfile != null
                                if (AffectedFileNames.Length > 0)
                                {
                                    #region Get File and Variables

                                    // Get the affected file object
                                    aFile = (IEdmFile10)PoVault.GetObject(EdmObjectType.EdmObject_File, AffectedFile.mlObjectID1);

                                    // get the variable for this file
                                    EnumVarObj = (IEdmEnumeratorVariable8)aFile.GetEnumeratorVariable();

                                    // get "File Type" variable for this file and pass to FileTypeObj
                                    GetVarSuccess = EnumVarObj.GetVar("File Type", "@", out FileTypeObj);

                                    // convert the variable from object to string
                                    FileTypeString = (string)FileTypeObj;
                                    #endregion

                                    #region Get all variables

                                    #region Reference
                                    /* Index Reference
                                    0 - Description;
                                    1 - Description2;
                                    2 - Description3;
                                    3 - MaterialNo;
                                    4 - Treatment1;
                                    5 - Treatment2;
                                    6 - Treatment3;
                                    7 - Reference;
                                    8 - Customer;
                                    9 - Size;
                                    10 - BrandOEM;
                                    11 - Supplier;
                                    12 - ExtReference;
                                    13 - Package;
                                    14 - Width;
                                    15 - Length;
                                    16 - Thickness;
                                    17 - IQC;
                                    18 - SparePart;
                                    19 - Category;
                                    20 - SubCategory;
                                    21 - Lifespan;
                                    22 - Template;
                                    23 - RnD;
                                    24 - Standard;
                                    25 - Version;
                                    */
                                    #endregion

                                    object[] varObj = new object[25];
                                    string[] varStr = { "","","","","", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""};
                                    string[] variableName = {"Description", "Description2", "Description3", "MaterialNo",
                                                             "Treatment1", "Treatment2", "Treatment3", "Reference",
                                                             "Customer", "Size", "Brand", "Supplier", "ExtReference",
                                                             "Package", "Width", "Length", "Thickness", "IQC", "SparePart",
                                                             "Category", "SubCategory", "LifeSpan", "Template", "RnD",
                                                             "Standard", "Version"};
                                    for (int startInt = 0; startInt < 25; startInt++)
                                    {
                                        GetVarSuccess = EnumVarObj.GetVar(variableName[startInt], "@", out varObj[startInt]);
                                        if (varObj[startInt] != null)
                                        {
                                            varStr[startInt] = varObj[startInt].ToString();
                                        }
                                    }
                                    #endregion

                                    #region Parameter Check
                                    switch (FileTypeString)
                                    {
                                        // template
                                        /*if (varStr[0] == "" | varStr[1] == "" | varStr[2] == "" |
                                              varStr[3] == "" | varStr[4] == "" | varStr[5] == "" |
                                              varStr[6] == "" | varStr[7] == "" | varStr[8] == "" |
                                              varStr[9] == "" | varStr[10] == "" | varStr[11] == "" |
                                              varStr[12] == "" | varStr[13] == "" | varStr[14] == "" |
                                              varStr[15] == "" | varStr[16] == "" | varStr[17] == "" |
                                              varStr[18] == "" | varStr[19] == "" | varStr[20] == "" |
                                              varStr[21] == "" | varStr[22] == "" | varStr[23] == "" |
                                              varStr[24] == "" | varStr[25] == "")
                                              poCmd.mbCancel = 1*/

                                        case "66":
                                        case "86":
                                            if ((varStr[0] == "") || (varStr[1] == "") || (varStr[2] == "") ||
                                                (varStr[7] == "") || (varStr[8] == "") || (varStr[18] == "") ||
                                                (varObj[0] == null) || (varObj[1] == null) || (varObj[2] == null) ||
                                                (varObj[7] == null) || (varObj[8] == null) || (varObj[18] == null) )
                                                {
                                                    poCmd.mbCancel = 1;
                                                    MessageBox.Show("Please fill in all required field");
                                                    return;
                                                }
                                            break;

                                        case "71":
                                            if ((varStr[0] == "") || (varStr[1] == "") || (varStr[2] == "") ||
                                                (varStr[7] == "") || (varStr[8] == "") || (varStr[13] == "") ||
                                                (varStr[14] == "") || (varStr[15] == "") || (varStr[16] == "") ||
                                                (varObj[0] == null) || (varObj[1] == null) || (varObj[2] == null) ||
                                                (varObj[7] == null) || (varObj[8] == null) || (varObj[13] == null) ||
                                                (varObj[14] == null) || (varObj[15] == null) || (varObj[16] == null) )
                                                {
                                                    poCmd.mbCancel = 1;
                                                    MessageBox.Show("Please fill in all required field");
                                                    return;
                                                }
                                            break;

                                        case "72":
                                            if ((varStr[0] == "") || (varStr[13] == "") ||
                                                (varObj[0] == null) || (varObj[13] == null) )
                                                {
                                                    poCmd.mbCancel = 1;
                                                    MessageBox.Show("Please fill in all required field");
                                                    return;
                                                }
                                            break;

                                        case "89":
                                            if ((varStr[0] == "") || (varStr[1] == "") || (varStr[2] == "") ||
                                                (varStr[7] == "") || (varStr[8] == "") || (varStr[19] == "") ||
                                                (varObj[0] == null) || (varObj[1] == null) || (varObj[2] == null) ||
                                                (varObj[7] == null) || (varObj[8] == null) || (varObj[19] == null))
                                                {
                                                    poCmd.mbCancel = 1;
                                                    MessageBox.Show("Please fill in all required field");
                                                    return;
                                                }
                                            break;

                                        case "89 - HW System":
                                        case "90":
                                        case "91 - CK":
                                        case "91 - Tooling":
                                            if ((varStr[0] == "") || (varStr[1] == "") || (varStr[2] == "") ||
                                                (varStr[7] == "") || (varStr[8] == "") || (varStr[13] == "") ||
                                                (varStr[14] == "") || (varStr[15] == "") || (varStr[16] == "") || (varStr[19] == "") ||
                                                (varObj[0] == null) || (varObj[1] == null) || (varObj[2] == null) ||
                                                (varObj[7] == null) || (varObj[8] == null) || (varObj[13] == null) ||
                                                (varObj[14] == null) || (varObj[15] == null) || (varObj[16] == null) || (varObj[19] == null))
                                                {
                                                poCmd.mbCancel = 1;
                                                    MessageBox.Show("Please fill in all required field");
                                                    return;
                                                }
                                            break;
                                        
                                        case "92":
                                            if ((varStr[0] == "") || (varStr[1] == "") || (varStr[2] == "") ||
                                                (varStr[7] == "") || (varStr[8] == "") || (varStr[18] == "") ||
                                                (varStr[19] == "") || (varStr[20] == "") || (varStr[21] == "") ||
                                                (varObj[0] == null) || (varObj[1] == null) || (varObj[2] == null) ||
                                                (varObj[7] == null) || (varObj[8] == null) || (varObj[18] == null) ||
                                                (varObj[19] == null) || (varObj[20] == null) || (varObj[21] == null) )
                                                {
                                                    poCmd.mbCancel = 1;
                                                    MessageBox.Show("Please fill in all required field");
                                                    return;
                                                }
                                            break;

                                        case "94":
                                            if ((varStr[0] == "") || (varStr[1] == "") || (varStr[2] == "") ||
                                                (varStr[3] == "") || (varStr[4] == "") || (varStr[5] == "") ||
                                                (varStr[6] == "") || (varStr[7] == "") || (varStr[8] == "") ||
                                                (varStr[18] == "") || (varStr[19] == "") || (varStr[20] == "") ||
                                                (varStr[21] == "") || (varStr[23] == "") || (varStr[24] == "") || (varStr[25] == "") ||
                                                (varObj[0] == null) || (varObj[1] == null) || (varObj[2] == null) ||
                                                (varObj[3] == null) || (varObj[4] == null) || (varObj[5] == null) ||
                                                (varObj[6] == null) || (varObj[7] == null) || (varObj[8] == null) ||
                                                (varObj[18] == null) || (varObj[19] == null) || (varObj[20] == null) ||
                                                (varObj[21] == null) || (varObj[23] == null) || (varObj[24] == null) || (varObj[25] == null))
                                                {
                                                    poCmd.mbCancel = 1;
                                                    MessageBox.Show("Please fill in all required field");
                                                    return;
                                                }
                                            break;

                                        case "94 - Fabrication":
                                            if ((varStr[0] == "") || (varStr[1] == "") || (varStr[2] == "") ||
                                                (varStr[3] == "") || (varStr[4] == "") || (varStr[5] == "") ||
                                                (varStr[6] == "") || (varStr[7] == "") || (varStr[8] == "") || (varStr[17] == "") ||
                                                (varStr[18] == "") || (varStr[19] == "") || (varStr[20] == "") ||
                                                (varStr[21] == "") || (varStr[23] == "") || (varStr[24] == "") || (varStr[25] == "") ||
                                                (varObj[0] == null) || (varObj[1] == null) || (varObj[2] == null) ||
                                                (varObj[3] == null) || (varObj[4] == null) || (varObj[5] == null) ||
                                                (varObj[6] == null) || (varObj[7] == null) || (varObj[8] == null) || (varObj[17] == null) ||
                                                (varObj[18] == null) || (varObj[19] == null) || (varObj[20] == null) ||
                                                (varObj[21] == null) || (varObj[23] == null) || (varObj[24] == null) || (varObj[25] == null))
                                                {
                                                    poCmd.mbCancel = 1;
                                                    MessageBox.Show("Please fill in all required field");
                                                    return;
                                                }
                                            break;

                                        case "94 - Tooling":
                                            if ((varStr[0] == "") || (varStr[1] == "") || (varStr[2] == "") ||
                                                (varStr[3] == "") || (varStr[4] == "") || (varStr[5] == "") ||
                                                (varStr[6] == "") || (varStr[7] == "") || (varStr[8] == "") || 
                                                (varStr[13] == "") || (varStr[14] == "") || (varStr[15] == "") ||
                                                (varStr[16] == "") || (varStr[17] == "") ||
                                                (varStr[18] == "") || (varStr[19] == "") || (varStr[20] == "") ||
                                                (varStr[21] == "") || (varStr[22] == "") || (varStr[23] == "") || (varStr[25] == "") ||
                                                (varObj[0] == null) || (varObj[1] == null) || (varObj[2] == null) ||
                                                (varObj[3] == null) || (varObj[4] == null) || (varObj[5] == null) ||
                                                (varObj[6] == null) || (varObj[7] == null) || (varObj[8] == null) ||
                                                (varObj[13] == null) || (varObj[14] == null) || (varObj[15] == null) ||
                                                (varObj[16] == null) || (varObj[17] == null) ||
                                                (varObj[18] == null) || (varObj[19] == null) || (varObj[20] == null) ||
                                                (varObj[21] == null) || (varObj[22] == null) || (varObj[23] == null) || (varObj[25] == null))
                                                {
                                                    poCmd.mbCancel = 1;
                                                    MessageBox.Show("Please fill in all required field");
                                                    return;
                                                }
                                            break;

                                        //case "95":

                                        case "96":
                                            if ((varStr[0] == "") || (varStr[1] == "") || (varStr[2] == "") ||
                                                (varStr[3] == "") || (varStr[4] == "") || (varStr[5] == "") ||
                                                (varStr[6] == "") || (varStr[7] == "") || (varStr[8] == "") ||
                                                (varStr[18] == "") || (varStr[19] == "") || (varStr[20] == "") ||
                                                (varStr[21] == "") || (varStr[23] == "") || (varStr[24] == "") ||
                                                (varObj[0] == null) || (varObj[1] == null) || (varObj[2] == null) ||
                                                (varObj[3] == null) || (varObj[4] == null) || (varObj[5] == null) ||
                                                (varObj[6] == null) || (varObj[7] == null) || (varObj[8] == null) ||
                                                (varObj[18] == null) || (varObj[19] == null) || (varObj[20] == null) ||
                                                (varObj[21] == null) || (varObj[23] == null) || (varObj[24] == null))
                                                {
                                                    poCmd.mbCancel = 1;
                                                    MessageBox.Show("Please fill in all required field");
                                                    return;
                                                }
                                            break;

                                        case "96 - Fabrication":
                                            if ((varStr[0] == "") || (varStr[1] == "") || (varStr[2] == "") ||
                                                (varStr[3] == "") || (varStr[4] == "") || (varStr[5] == "") ||
                                                (varStr[6] == "") || (varStr[7] == "") || (varStr[8] == "") || (varStr[17] == "") ||
                                                (varStr[18] == "") || (varStr[19] == "") || (varStr[20] == "") ||
                                                (varStr[21] == "") || (varStr[23] == "") || (varStr[24] == "") ||
                                                (varObj[0] == null) || (varObj[1] == null) || (varObj[2] == null) ||
                                                (varObj[3] == null) || (varObj[4] == null) || (varObj[5] == null) ||
                                                (varObj[6] == null) || (varObj[7] == null) || (varObj[8] == null) || (varObj[17] == null) ||
                                                (varObj[18] == null) || (varObj[19] == null) || (varObj[20] == null) ||
                                                (varObj[21] == null) || (varObj[23] == null) || (varObj[24] == null))
                                                {
                                                    poCmd.mbCancel = 1;
                                                    MessageBox.Show("Please fill in all required field");
                                                    return;
                                                }
                                            break;

                                        case "96 - Tooling":
                                            if ((varStr[0] == "") || (varStr[1] == "") || (varStr[2] == "") ||
                                                (varStr[3] == "") || (varStr[4] == "") || (varStr[5] == "") ||
                                                (varStr[6] == "") || (varStr[7] == "") || (varStr[8] == "") ||
                                                (varStr[13] == "") || (varStr[14] == "") || (varStr[15] == "") ||
                                                (varStr[16] == "") || (varStr[17] == "") ||
                                                (varStr[18] == "") || (varStr[19] == "") || (varStr[20] == "") ||
                                                (varStr[21] == "") || (varStr[22] == "") || (varStr[23] == "") || (varStr[25] == "") ||
                                                (varObj[0] == null) || (varObj[1] == null) || (varObj[2] == null) ||
                                                (varObj[3] == null) || (varObj[4] == null) || (varObj[5] == null) ||
                                                (varObj[6] == null) || (varObj[7] == null) || (varObj[8] == null) ||
                                                (varObj[13] == null) || (varObj[14] == null) || (varObj[15] == null) ||
                                                (varObj[16] == null) || (varObj[17] == null) ||
                                                (varObj[18] == null) || (varObj[19] == null) || (varObj[20] == null) ||
                                                (varObj[21] == null) || (varObj[22] == null) || (varObj[23] == null) || (varObj[25] == null))
                                            {
                                                poCmd.mbCancel = 1;
                                                MessageBox.Show("Please fill in all required field");
                                                return;
                                            }
                                            break;

                                        case "97":
                                            if ((varStr[0] == "") || (varStr[1] == "") || (varStr[2] == "") ||
                                                (varStr[3] == "") || (varStr[4] == "") || (varStr[5] == "") ||
                                                (varStr[6] == "") || (varStr[7] == "") || (varStr[8] == "") ||
                                                (varStr[9] == "") || (varStr[10] == "") || (varStr[11] == "") ||
                                                (varStr[12] == "") ||
                                                (varObj[0] == null) || (varObj[1] == null) || (varObj[2] == null) ||
                                                (varObj[3] == null) || (varObj[4] == null) || (varObj[5] == null) ||
                                                (varObj[6] == null) || (varObj[7] == null) || (varObj[8] == null) ||
                                                (varObj[9] == null) || (varObj[10] == null) || (varObj[11] == null) ||
                                                (varObj[12] == null))
                                                {
                                                    poCmd.mbCancel = 1;
                                                    MessageBox.Show("Please fill in all required field");
                                                    return;
                                                }
                                            break;
                                        //case "98":
                                        //case "98 - Fabrication":
                                        //case "98 - Tooling":
                                        case "99":
                                            if ((varStr[0] == "") || (varStr[1] == "") || (varStr[2] == "") ||
                                                (varStr[7] == "") || (varStr[8] == "") || (varStr[13] == "") ||
                                                (varStr[14] == "") || (varStr[15] == "") || (varStr[16] == "") ||
                                                (varStr[19] == "") || (varStr[20] == "") ||
                                                (varObj[0] == null) || (varObj[1] == null) || (varObj[2] == null) ||
                                                (varObj[7] == null) || (varObj[8] == null) || (varObj[13] == null) ||
                                                (varObj[14] == null) || (varObj[15] == null) || (varObj[16] == null) ||
                                                (varObj[19] == null) || (varObj[20] == null))
                                                {
                                                    poCmd.mbCancel = 1;
                                                    MessageBox.Show("Please fill in all required field");
                                                    return;
                                                }
                                            break;
                                        //case "Standard Module":
                                    }
                                    #endregion
                                }
                                #endregion
                            }
                        }
                        #endregion
                        break;

                    case EdmCmdType.EdmCmd_PostState:
                        #region Post State
                        // Create serial number generation tility
                        SerialNbrs = (IEdmSerNoGen7)((IEdmVault11)PoVault).CreateUtility(EdmUtility.EdmUtil_SerNoGen);
                        foreach (EdmCmdData AffectedFile in ppoData)
                        {
                            // If the state is "Released"
                            if (AffectedFile.mbsStrData2 == "Released")
                            {
                                #region Get File information
                                // Get affected full file name
                                AffectedFileNames = AffectedFile.mbsStrData1;

                                // Get the name with extension
                                FileName = Path.GetFileName(AffectedFile.mbsStrData1);

                                // Get the extension ".sldprt"
                                ext = Path.GetExtension(AffectedFileNames);
                                #endregion

                                #region Process Files
                                // If affectedfile != null
                                if (AffectedFileNames.Length > 0)
                                {
                                    #region Get File and Variables

                                    // Get the affected file object
                                    aFile = (IEdmFile10)PoVault.GetObject(EdmObjectType.EdmObject_File, AffectedFile.mlObjectID1);

                                    // get the variable for this file
                                    EnumVarObj = (IEdmEnumeratorVariable8)aFile.GetEnumeratorVariable();

                                    // get "File Type" variable for this file and pass to FileTypeObj
                                    GetVarSuccess = EnumVarObj.GetVar("File Type", "@", out FileTypeObj);

                                    // convert the variable from object to string
                                    FileTypeString = (string)FileTypeObj;
                                    #endregion

                                    #region Handle Study Folder
                                    if (AffectedFileNames.Contains("\\Study Folder\\"))
                                    {
                                        switch (FileTypeString)
                                        {
                                            case "66":
                                            case "71":
                                            case "72":
                                            case "86":
                                            case "89":
                                            case "89 - HW System":
                                            case "90":
                                            case "91 - CK":
                                            case "91 - Tooling":
                                            case "92":
                                            case "94":
                                            case "94 - Fabrication":
                                            case "94 - Tooling":
                                            case "95":
                                            case "96":
                                            case "96 - Fabrication":
                                            case "96 - Tooling":
                                            case "97":
                                            case "98":
                                            case "98 - Fabrication":
                                            case "98 - Tooling":
                                            case "99":
                                                #region Handle file type with serial number
                                                // allocate a serial number for this file
                                                SerialNbrValue = SerialNbrs.AllocSerNoValue(FileTypeString.Substring(0, 2), (int)IntPtr.Zero, "", 0, 0, 0, 0);

                                                // get serial number value
                                                SerialNbrValueValue = SerialNbrValue.Value;

                                                // get sub folder path
                                                destSubFolder = SerialNbrValueValue.Substring(0, 5);

                                                // Check if the directory is exist
                                                if (!(Directory.Exists(Folder.LocalPath + "\\" + FileTypeString.Substring(0, 2) + "XXXXX\\" + destSubFolder + "XX")))
                                                {
                                                    IEdmFolder5 parentFolder = PoVault.GetFolderFromPath(Folder.LocalPath + "\\" + FileTypeString.Substring(0, 2) + "XXXXX\\");
                                                    parentFolder.AddFolder((int)IntPtr.Zero, destSubFolder + "XX");
                                                };

                                                // destination
                                                dest = PoVault.GetFolderFromPath(Folder.LocalPath + "\\" + FileTypeString.Substring(0, 2) + "XXXXX\\" + destSubFolder + "XX");
                                                #endregion

                                                #region Rename and move file

                                                // Check out file so that can change variable
                                                aFile.LockFile(AffectedFile.mlObjectID2, (int)IntPtr.Zero);

                                                // Assign serial number to variable
                                                EnumVarObj = (IEdmEnumeratorVariable8)aFile.GetEnumeratorVariable();

                                                // Get all configurations for this file
                                                configList = aFile.GetConfigurations();
                                                pos = configList.GetHeadPosition();
                                                while (!pos.IsNull)
                                                {
                                                    // set variable for each configuration in this file
                                                    curConfig = configList.GetNext(pos);
                                                    EnumVarObj.SetVar("PartNo", curConfig, SerialNbrValueValue);
                                                }

                                                // good practice to close it after done
                                                EnumVarObj.CloseFile(true);

                                                // Check in file
                                                aFile.UnlockFile((int)IntPtr.Zero, "");

                                                // Rename file
                                                aFile.Rename((int)IntPtr.Zero, SerialNbrValueValue + ext);

                                                // move file
                                                aFile.Move((int)IntPtr.Zero, AffectedFile.mlObjectID2, dest.ID, 0);
                                                #endregion
                                                break;

                                            case "Standard Module":
                                                #region Handle file type without serial number
                                                // destination
                                                dest = PoVault.GetFolderFromPath(Folder.LocalPath + "\\" + FileTypeString);
                                                #endregion

                                                #region Rename and move file
                                                // move file
                                                aFile.Move((int)IntPtr.Zero, AffectedFile.mlObjectID2, dest.ID, 0);
                                                #endregion
                                                break;
                                        }
                                    }
                                    #endregion
                                }
                                #endregion
                            }
                        }
                        #endregion
                        break;

                    case EdmCmdType.EdmCmd_Menu:

                        #region Assign variable command
                        if (poCmd.mlCmdID == 1000)
                        {
                            // create new form to choose variable and show it
                            using(assignVarForm newForm = new assignVarForm())
                            {
                                if(newForm.ShowDialog() == DialogResult.OK)
                                {
                                    selectedFileType = newForm.SelectedVariable;
                                }
                            }

                            // if confirm is clicked on the form
                            if (selectedFileType != "")
                            {
                                foreach (EdmCmdData AffectedFile in ppoData)
                                {
                                    // get affected file
                                    aFile = (IEdmFile10)PoVault.GetObject(EdmObjectType.EdmObject_File, AffectedFile.mlObjectID1);

                                    // Check if file is checked out
                                    if (aFile.IsLocked)
                                    {
                                        // Assign serial number to variable
                                        EnumVarObj = (IEdmEnumeratorVariable8)aFile.GetEnumeratorVariable();

                                        // Get all configurations for this file
                                        configList = aFile.GetConfigurations();
                                        pos = configList.GetHeadPosition();
                                        while (!pos.IsNull)
                                        {
                                            // set variable for each configuration in this file
                                            curConfig = configList.GetNext(pos);
                                            EnumVarObj.SetVar("File Type", curConfig, selectedFileType);
                                        }

                                        // good practice
                                        EnumVarObj.CloseFile(true);

                                        // count for processed file(s)
                                        fileCount++;
                                    }
                                    else
                                    {
                                        MessageBox.Show("Some files are not checked out, please check out first");
                                        return;
                                    }
                                }
                            };
                        }

                        MessageBox.Show("Finished set variables for " + fileCount + " files.");
                        #endregion
                        break;

                    case EdmCmdType.EdmCmd_CardInput:
                        #region When variable changed
                        /*
                        lock (static_OnCmd_VariableChangeInProgress_Init)
                        {
                            try
                            {
                                if (InitStaticVariableHelper(static_OnCmd_VariableChangeInProgress_Init))
                                {
                                    static_OnCmd_VariableChangeInProgress = false;
                                }
                            }
                            finally
                            {
                                static_OnCmd_VariableChangeInProgress_Init.State = 1;
                            }
                        }

                    */
                        #endregion
                        break;
                }
            }
            catch (System.Runtime.InteropServices.COMException ex) {
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }

        /*static bool InitStaticVariableHelper(Microsoft.VisualBasic.CompilerServices.StaticLocalInitFlag flag)
        {
            if (flag.State == 0)
            {
                flag.State = 2;
                return true;
            }
            else if (flag.State == 2)
            {
                throw new Microsoft.VisualBasic.CompilerServices.IncompleteInitialization();
            }
            else
            {
                return false;
            }
        }*/
    }
}

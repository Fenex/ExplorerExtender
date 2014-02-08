using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Drawing;
using System.Reflection;
using System.Linq;

namespace ExplorerExtender
{
    [ClassInterface(ClassInterfaceType.None)]
    [ComVisible(true)]
    [Guid("23a8b362-4cce-454a-b9d9-7490bfde4d83")]
    public class Main : IContextMenu, IShellExtInit
    {
        #region Fields

        private Dictionary<uint, string> Command;
        private List<string> SelectedItems;
        private bool isClickedOnEmptyArea;

        private IntPtr hIconAdd = IntPtr.Zero;
        private IntPtr hIconDelete = IntPtr.Zero;
        private IntPtr hIconTrash = IntPtr.Zero;
        private IntPtr hIconRename = IntPtr.Zero;
        private IntPtr hIconMove = IntPtr.Zero;
        private IntPtr hIconGroup = IntPtr.Zero;
        private IntPtr hIconUngroup = IntPtr.Zero;
        private IntPtr hIconReplace = IntPtr.Zero;
        private IntPtr hIconUrl = IntPtr.Zero;
        private IntPtr hIconHtml = IntPtr.Zero;

        #endregion Fields

        public Main()
        {
            this.hIconAdd = Helper.GetIcon("ExplorerExtender.Icons.Add.png");
            this.hIconDelete = Helper.GetIcon("ExplorerExtender.Icons.Delete.png");
            this.hIconTrash = Helper.GetIcon("ExplorerExtender.Icons.Trash.png");
            this.hIconRename = Helper.GetIcon("ExplorerExtender.Icons.Rename.png");
            this.hIconMove = Helper.GetIcon("ExplorerExtender.Icons.Move.png");
            this.hIconGroup = Helper.GetIcon("ExplorerExtender.Icons.Group.png");
            this.hIconUngroup = Helper.GetIcon("ExplorerExtender.Icons.Ungroup.png");
            this.hIconReplace = Helper.GetIcon("ExplorerExtender.Icons.Replace.png");
            this.hIconUrl = Helper.GetIcon("ExplorerExtender.Icons.Url.png");
            this.hIconHtml = Helper.GetIcon("ExplorerExtender.Icons.Html.png");
        }

        ~Main()
        {
            Helper.DeleteIconObject(ref this.hIconAdd);
            Helper.DeleteIconObject(ref this.hIconDelete);
            Helper.DeleteIconObject(ref this.hIconTrash);
            Helper.DeleteIconObject(ref this.hIconRename);
            Helper.DeleteIconObject(ref this.hIconMove);
            Helper.DeleteIconObject(ref this.hIconGroup);
            Helper.DeleteIconObject(ref this.hIconUngroup);
            Helper.DeleteIconObject(ref this.hIconReplace);
            Helper.DeleteIconObject(ref this.hIconUrl);
            Helper.DeleteIconObject(ref this.hIconHtml);
        }

        #region Methods

        #region Shell Registration

        [ComRegisterFunction]
        public static void Register(Type t)
        {
            ExplorerExtender.ShellExtReg.RegisterShellExtContextMenuHandler(t.GUID, "DIRECTORY", "ExplorerExtender.Main");
            ExplorerExtender.ShellExtReg.RegisterShellExtContextMenuHandler(t.GUID, @"DIRECTORY\Background", "ExplorerExtender.Main");
            ExplorerExtender.ShellExtReg.RegisterShellExtContextMenuHandler(t.GUID, "*", "ExplorerExtender.Main");
        }

        [ComUnregisterFunction]
        public static void Unregister(Type t)
        {
            ExplorerExtender.ShellExtReg.UnregisterShellExtContextMenuHandler(t.GUID, "DIRECTORY");
            ExplorerExtender.ShellExtReg.UnregisterShellExtContextMenuHandler(t.GUID, @"DIRECTORY\Background");
            ExplorerExtender.ShellExtReg.UnregisterShellExtContextMenuHandler(t.GUID, "*");
        }

        #endregion

        #region IShellExtInit
        public void Initialize(IntPtr pidlFolder, IntPtr pDataObj, IntPtr hKeyProgID)
        {
            Helper.GetSelectedItems(pidlFolder, pDataObj, hKeyProgID, ref this.SelectedItems, ref this.isClickedOnEmptyArea);
        }
        #endregion

        #region IContextMenu
        public void GetCommandString(UIntPtr idCmd, uint uFlags, IntPtr pReserved, StringBuilder pszName, uint cchMax)
        {
            uint idCommand = idCmd.ToUInt32();

            if (!this.Command.ContainsKey(idCommand))
                return;

            string command;
            string help;

            switch (this.Command[idCommand])
            {
                case "Break":
                    command = "Break";
                    help = "Move folder content to the current folder";
                    break;
                case "Build":
                    command = "Build";
                    help = "Group folder into 1 single folder";
                    break;
                case "Replace":
                    command = "Replace";
                    help = "Search and Replace Filename";
                    break;
                case "UrlDecode":
                    command = "UrlDecode";
                    help = "Url Decode Filename";
                    break;
                case "HtmlDecode":
                    command = "HtmlDecode";
                    help = "Html Decode Filename";
                    break;
                default:
                    return;
            }

            switch ((GCS)uFlags)
            {
                case GCS.GCS_VERBW:
                    if (command.Length > cchMax - 1)
                        Marshal.ThrowExceptionForHR(WinError.STRSAFE_E_INSUFFICIENT_BUFFER);
                    else
                    {
                        pszName.Clear();
                        pszName.Append(command);
                    }
                    break;
                case GCS.GCS_HELPTEXTW:
                    if (help.Length > cchMax - 1)
                        Marshal.ThrowExceptionForHR(WinError.STRSAFE_E_INSUFFICIENT_BUFFER);
                    else
                    {
                        pszName.Clear();
                        pszName.Append(help);
                    }
                    break;
            }
        }

        public void InvokeCommand(IntPtr pici)
        {
            uint iCommand = (uint)Helper.GetCommandOffsetId(pici);

            if (!this.Command.ContainsKey(iCommand))
                return;

            string str = this.Command[iCommand];
            
            if (str == "Break")
                FileOperations.BreakFolder(this.SelectedItems);
            else if (str == "Build")
                FileOperations.BuildFolder(this.SelectedItems);
            else if (str == "Replace")
            {
                this.SelectedItems.ForEach(fi =>
                {
                    FileInfo fileInfo = new FileInfo(fi);
                    fileInfo.ReplaceFileName();
                });
            }
            else if (str == "UrlDecode")
            {
                this.SelectedItems.ForEach(fi =>
                {
                    FileInfo fileInfo = new FileInfo(fi);
                    fileInfo.DecodeFileNameUrl();
                });
            }
            else if (str == "HtmlDecode")
            {
                this.SelectedItems.ForEach(fi =>
                {
                    FileInfo fileInfo = new FileInfo(fi);
                    fileInfo.DecodeFileNameHtml();
                });
            }
            else if (str == "FileMoverMove")
            {
                FileMover.ExecuteListAction(this.SelectedItems[0]);
            }
            else if (str == "FileMoverAdd")
            {
                this.SelectedItems.ForEach(i => FileMover.AddPathToList(i));
            }
            else if (str == "FileMoverClear")
            {
                FileMover.ClearList();
            }
            else
            {
                var list = FileMover.GetList();
                if (list.Contains(str))
                    FileMover.RemovePathFromList(str);
                else
                    Marshal.ThrowExceptionForHR(WinError.E_FAIL);
            }
        }

        public int QueryContextMenu(IntPtr hMenu, uint iMenu, uint idCmdFirst, uint idCmdLast, uint uFlags)
        {
            if (((uint)CMF.CMF_DEFAULTONLY & uFlags) != 0)
                return WinError.MAKE_HRESULT(WinError.SEVERITY_SUCCESS, 0, 0);

            #region Initialize Command List
            //this counter keep track of the # of commands
            uint idCommand = 0;
            this.Command = new Dictionary<uint, string>();
            #endregion

            #region Create Command Separator Menu Item
            //Menu Item can be added multiple time, so we only need one separator menuitem
            MENUITEMINFO miiSeparator = new MENUITEMINFO();
            Helper.InitializeMenuitemInfoAsSeparator(ref miiSeparator);
            #endregion

            bool onlyFolder = this.CheckOnlyFolder();

            #region Create Main Menu (Explorer Extender)
            //this counter keep track of the position in menu MainMenu
            uint iPosition_Menu_Main = 0;
            IntPtr h_Menu_Main = NativeMethods.CreatePopupMenu();

            MENUITEMINFO mii_Menu_Main = new MENUITEMINFO();
            Helper.InitializeMenuitemInfoAsSubMenu(ref mii_Menu_Main, h_Menu_Main, "Explorer Extender", IntPtr.Zero);

            if (!NativeMethods.InsertMenuItem(hMenu, iMenu, true, ref mii_Menu_Main))
                return Marshal.GetHRForLastWin32Error();
            #endregion

            #region Add MenuItem (Break Folder)

            if (!this.isClickedOnEmptyArea && onlyFolder)
            {
                MENUITEMINFO miiBreakFolder = new MENUITEMINFO();

                Helper.InitializeMenuitemInfoAsMenuItem(ref miiBreakFolder, idCmdFirst + idCommand, this.SelectedItems.Count == 1 ? "Break Folder" : "Break Folders", this.hIconUngroup);

                if (!NativeMethods.InsertMenuItem(h_Menu_Main, iPosition_Menu_Main++, true, ref miiBreakFolder))
                    return Marshal.GetHRForLastWin32Error();

                this.Command.Add(idCommand++, "Break");
            }

            #endregion

            #region Add MenuItem (Group into new folder)

            if (!this.isClickedOnEmptyArea)
            {
                MENUITEMINFO miiBuildFolder = new MENUITEMINFO();
                Helper.InitializeMenuitemInfoAsMenuItem(ref miiBuildFolder, idCmdFirst + idCommand, "Group", this.hIconGroup);

                if (!NativeMethods.InsertMenuItem(h_Menu_Main, iPosition_Menu_Main++, true, ref miiBuildFolder))
                    return Marshal.GetHRForLastWin32Error();

                this.Command.Add(idCommand++, "Build");
            }

            #endregion

            #region Add Separator
            if (!this.isClickedOnEmptyArea)
            {
                if (!NativeMethods.InsertMenuItem(h_Menu_Main, iPosition_Menu_Main++, true, ref miiSeparator))
                    return Marshal.GetHRForLastWin32Error();
            } 
            #endregion

            if (!this.isClickedOnEmptyArea)
            {
                #region Create Submenu (Rename)
                //this counter keep track of the position in menu MainMenu
                uint iPosition_Menu_Rename = 0;
                IntPtr h_Menu_Rename = NativeMethods.CreatePopupMenu();

                MENUITEMINFO mii_Menu_Rename = new MENUITEMINFO();
                Helper.InitializeMenuitemInfoAsSubMenu(ref mii_Menu_Rename, h_Menu_Rename, "Rename", this.hIconRename);

                if (!NativeMethods.InsertMenuItem(h_Menu_Main, iPosition_Menu_Main++, true, ref mii_Menu_Rename))
                    return Marshal.GetHRForLastWin32Error();
                #endregion

                #region Add MenuItem (Search & Replace)
                MENUITEMINFO miiReplace = new MENUITEMINFO();
                Helper.InitializeMenuitemInfoAsMenuItem(ref miiReplace, idCmdFirst + idCommand, "Search and Replace", this.hIconReplace);

                if (!NativeMethods.InsertMenuItem(h_Menu_Rename, iPosition_Menu_Rename++, true, ref miiReplace))
                    return Marshal.GetHRForLastWin32Error();

                this.Command.Add(idCommand++, "Replace");
                #endregion

                #region Add MenuItem (URL Decode)
                MENUITEMINFO miiUrlEncode = new MENUITEMINFO();
                Helper.InitializeMenuitemInfoAsMenuItem(ref miiUrlEncode, idCmdFirst + idCommand, "Url Decode", this.hIconUrl);

                if (!NativeMethods.InsertMenuItem(h_Menu_Rename, iPosition_Menu_Rename++, true, ref miiUrlEncode))
                    return Marshal.GetHRForLastWin32Error();

                this.Command.Add(idCommand++, "UrlDecode");
                #endregion

                #region Add MenuItem (Html Decode)
                MENUITEMINFO miiHtmlEncode = new MENUITEMINFO();
                Helper.InitializeMenuitemInfoAsMenuItem(ref miiHtmlEncode, idCmdFirst + idCommand, "Html Decode", this.hIconHtml);

                if (!NativeMethods.InsertMenuItem(h_Menu_Rename, iPosition_Menu_Rename++, true, ref miiHtmlEncode))
                    return Marshal.GetHRForLastWin32Error();

                this.Command.Add(idCommand++, "HtmlDecode");
                #endregion
            }

            #region Create Submenu (File Mover)

            uint iPosition_Menu_FileMover = 0;
            IntPtr h_Menu_FileMover = NativeMethods.CreatePopupMenu();

            MENUITEMINFO mii_Menu_FileMover = new MENUITEMINFO();
            Helper.InitializeMenuitemInfoAsSubMenu(ref mii_Menu_FileMover, h_Menu_FileMover, "File Mover", IntPtr.Zero);

            if (!NativeMethods.InsertMenuItem(h_Menu_Main, iPosition_Menu_Main++, true, ref mii_Menu_FileMover))
                return Marshal.GetHRForLastWin32Error();

            #endregion

            int fileCount = FileMover.GetListCount();

            #region Add MenuItem (Move File)
            if (CheckOnlyOneFolder() && fileCount > 0)
            {
                MENUITEMINFO miiFileMoverMove = new MENUITEMINFO();
                Helper.InitializeMenuitemInfoAsMenuItem(ref miiFileMoverMove, idCmdFirst + idCommand, fileCount == 1 ? "Move File Here" : "Move Files here", IntPtr.Zero);

                if (!NativeMethods.InsertMenuItem(h_Menu_FileMover, iPosition_Menu_FileMover++, true, ref miiFileMoverMove))
                    return Marshal.GetHRForLastWin32Error();

                this.Command.Add(idCommand++, "FileMoverMove");

                if (!NativeMethods.InsertMenuItem(h_Menu_FileMover, iPosition_Menu_FileMover++, true, ref miiSeparator))
                    return Marshal.GetHRForLastWin32Error();
            }
            #endregion

            #region Add MenuItem (Add to List)
            MENUITEMINFO miiFileMoverAdd = new MENUITEMINFO();
            Helper.InitializeMenuitemInfoAsMenuItem(ref miiFileMoverAdd, idCmdFirst + idCommand, "Add to list", this.hIconAdd);

            if (!NativeMethods.InsertMenuItem(h_Menu_FileMover, iPosition_Menu_FileMover++, true, ref miiFileMoverAdd))
                return Marshal.GetHRForLastWin32Error();

            this.Command.Add(idCommand++, "FileMoverAdd");
            #endregion

            #region Add MenuItem (Clear List)
            if (fileCount > 0)
            {
                MENUITEMINFO miiFileMoverClear = new MENUITEMINFO();
                Helper.InitializeMenuitemInfoAsMenuItem(ref miiFileMoverClear, idCmdFirst + idCommand, "Clear List", this.hIconDelete);

                if (!NativeMethods.InsertMenuItem(h_Menu_FileMover, iPosition_Menu_FileMover++, true, ref miiFileMoverClear))
                    return Marshal.GetHRForLastWin32Error();

                this.Command.Add(idCommand++, "FileMoverClear");
            }
            #endregion

            if (!NativeMethods.InsertMenuItem(h_Menu_FileMover, iPosition_Menu_FileMover++, true, ref miiSeparator))
                return Marshal.GetHRForLastWin32Error();

            var fileList = FileMover.GetList();

            #region Add MenuItem (Remove File From List)
            foreach (var i in fileList)
            {
                MENUITEMINFO mii_MenuItem_Remove = new MENUITEMINFO();
                Helper.InitializeMenuitemInfoAsMenuItem(ref mii_MenuItem_Remove, idCmdFirst + idCommand, i, this.hIconTrash);

                if (!NativeMethods.InsertMenuItem(h_Menu_FileMover, iPosition_Menu_FileMover++, true, ref mii_MenuItem_Remove))
                    return Marshal.GetHRForLastWin32Error();

                this.Command.Add(idCommand++, i);
            }
            #endregion

            return WinError.MAKE_HRESULT(WinError.SEVERITY_SUCCESS, 0, idCommand);
        }
        #endregion

        #region Private Helpers

        private bool CheckOnlyFolder()
        {
            foreach (string s in this.SelectedItems)
                if (File.Exists(s))
                    return false;
            return true;
        }

        private bool CheckOnlyOneFolder()
        {
            return this.SelectedItems.Count == 1 && this.CheckOnlyFolder();
        }

        #endregion
        #endregion Methods
    }
}
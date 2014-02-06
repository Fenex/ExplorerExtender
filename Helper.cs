using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Linq;

namespace ExplorerExtender
{
    internal class Helper
    {
        #region Methods

        internal static int GetCommandOffsetId(IntPtr pici)
        {
            return NativeMethods.LowWord(((CMINVOKECOMMANDINFO)Marshal.PtrToStructure(pici, typeof(CMINVOKECOMMANDINFO))).verb.ToInt32());
        }

        internal static void GetSelectedItems(IntPtr pidlFolder, IntPtr pDataObj, IntPtr hKeyProgID, ref List<String> SelectedItems, ref bool isClickOnEmptyArea)
        {
            if (pDataObj == IntPtr.Zero && pidlFolder == IntPtr.Zero)
                throw new ArgumentException();
            else if (pDataObj == IntPtr.Zero)
            {
                //User Click on Empty Area of a Folder, pDataObj is empty while pidlFolder is the current path
                StringBuilder sb = new StringBuilder(260);
                if (!NativeMethods.SHGetPathFromIDListW(pidlFolder, sb))
                    Marshal.ThrowExceptionForHR(WinError.E_FAIL);
                else
                {
                    isClickOnEmptyArea = true;
                    SelectedItems = new List<string>();
                    SelectedItems.Add(sb.ToString());
                }
            }
            else
            {
                //User actually select some item, pDataObj is the list while pidlFolder is empty
                isClickOnEmptyArea = false;

                FORMATETC fe = new FORMATETC();
                fe.cfFormat = (short)CLIPFORMAT.CF_HDROP;
                fe.ptd = IntPtr.Zero;
                fe.dwAspect = DVASPECT.DVASPECT_CONTENT;
                fe.lindex = -1;
                fe.tymed = TYMED.TYMED_HGLOBAL;
                STGMEDIUM stm = new STGMEDIUM();

                IDataObject dataObject = (IDataObject)Marshal.GetObjectForIUnknown(pDataObj);
                dataObject.GetData(ref fe, out stm);

                try
                {
                    IntPtr hDrop = stm.unionmember;
                    if (hDrop == IntPtr.Zero)
                        throw new ArgumentException();

                    uint nFiles = NativeMethods.DragQueryFile(hDrop, UInt32.MaxValue, null, 0);

                    if (nFiles == 0)
                        Marshal.ThrowExceptionForHR(WinError.E_FAIL);

                    SelectedItems = new List<string>();

                    for (int i = 0; i < nFiles; i++)
                    {
                        StringBuilder sb = new StringBuilder(260);
                        if (NativeMethods.DragQueryFile(hDrop, (uint)i, sb, sb.Capacity) == 0)
                            Marshal.ThrowExceptionForHR(WinError.E_FAIL);
                        else
                            SelectedItems.Add(sb.ToString());
                    }
                }
                finally
                {
                    NativeMethods.ReleaseStgMedium(ref stm);
                }
            }
        }

        public static void InitializeMenuitemInfoAsMenuItem(ref MENUITEMINFO mii, uint commandId, string menuText, IntPtr icon)
        {
            mii.cbSize = (uint)Marshal.SizeOf(mii);
            mii.fMask = MIIM.MIIM_STRING | MIIM.MIIM_FTYPE | MIIM.MIIM_ID | MIIM.MIIM_STATE;
            mii.wID = commandId;
            mii.fType = MFT.MFT_STRING;
            mii.dwTypeData = menuText;
            mii.fState = MFS.MFS_ENABLED;

            if (icon != IntPtr.Zero)
            {
                mii.fMask |= MIIM.MIIM_BITMAP;
                mii.hbmpItem = icon;
            }
        }

        public static void InitializeMenuitemInfoAsSeparator(ref MENUITEMINFO mii)
        {
            mii.cbSize = (uint)Marshal.SizeOf(mii);
            mii.fMask = MIIM.MIIM_TYPE;
            mii.fType = MFT.MFT_SEPARATOR;
        }

        public static void InitializeMenuitemInfoAsSubMenu(ref MENUITEMINFO mii, IntPtr submenu, string menuText, IntPtr icon)
        {
            mii.cbSize = (uint)Marshal.SizeOf(mii);
            mii.fMask = MIIM.MIIM_STRING | MIIM.MIIM_ID | MIIM.MIIM_SUBMENU;
            mii.hSubMenu = submenu;
            mii.fType = MFT.MFT_STRING;
            mii.dwTypeData = menuText;
            mii.fState = MFS.MFS_ENABLED;

            if (icon != IntPtr.Zero)
            {
                mii.fMask |= MIIM.MIIM_BITMAP;
                mii.hbmpItem = icon;
            }
        }

        public static IntPtr GetIcon(string resource)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();

            if (assembly.GetManifestResourceNames().Contains(resource))
                using (Bitmap bitmap = new Bitmap(assembly.GetManifestResourceStream(resource)))
                {
                    //bitmap.MakeTransparent(Color.White);
                    return bitmap.GetHbitmap();
                }
            else
                return IntPtr.Zero;
        }

        public static void DeleteIconObject(ref IntPtr icon)
        {
            if (icon != IntPtr.Zero)
            {
                NativeMethods.DeleteObject(icon);
                icon = IntPtr.Zero;
            }
        }

        #endregion Methods
    }
}
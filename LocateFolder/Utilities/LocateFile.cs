using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace LocateFolder.Utilities 
{
    internal static class LocateFile
    {
        private static Guid IID_IShellFolder = typeof(IShellFolder).GUID;
        private static int pointerSize = Marshal.SizeOf(typeof(IntPtr));

        public static void FileOrFolder(string path, bool edit = false)
        {
            if (path == null)
            {
                throw new ArgumentNullException("path");
            }
            IntPtr pidlFolder = PathToAbsolutePIDL(path);
            try
            {
                SHOpenFolderAndSelectItems(pidlFolder, null, edit);
            }
            finally
            {
                NativeMethods.ILFree(pidlFolder);
            }
        }

        public static void FilesOrFolders(IEnumerable<FileSystemInfo> paths)
        {
            if (paths == null)
            {
                throw new ArgumentNullException("paths");
            }
            if (paths.Count<FileSystemInfo>() != 0)
            {
                foreach (
                    IGrouping<string, FileSystemInfo> grouping in
                    from p in paths group p by Path.GetDirectoryName(p.FullName))
                {
                    FilesOrFolders(Path.GetDirectoryName(grouping.First<FileSystemInfo>().FullName),
                        (from fsi in grouping select fsi.Name).ToList<string>());
                }
            }
        }

        public static void FilesOrFolders(IEnumerable<string> paths)
        {
            FilesOrFolders(PathToFileSystemInfo(paths));
        }

        public static void FilesOrFolders(params string[] paths)
        {
            FilesOrFolders((IEnumerable<string>)paths);
        }

        public static void FilesOrFolders(string parentDirectory, ICollection<string> filenames)
        {
            if (filenames == null)
            {
                throw new ArgumentNullException("filenames");
            }
            if (filenames.Count != 0)
            {
                IntPtr pidl = PathToAbsolutePIDL(parentDirectory);
                try
                {
                    IShellFolder parentFolder = PIDLToShellFolder(pidl);
                    List<IntPtr> list = new List<IntPtr>(filenames.Count);
                    foreach (string str in filenames)
                    {
                        list.Add(GetShellFolderChildrenRelativePIDL(parentFolder, str));
                    }
                    try
                    {
                        SHOpenFolderAndSelectItems(pidl, list.ToArray(), false);
                    }
                    finally
                    {
                        using (List<IntPtr>.Enumerator enumerator2 = list.GetEnumerator())
                        {
                            while (enumerator2.MoveNext())
                            {
                                NativeMethods.ILFree(enumerator2.Current);
                            }
                        }
                    }
                }
                finally
                {
                    NativeMethods.ILFree(pidl);
                }
            }
        }

        private static IntPtr GetShellFolderChildrenRelativePIDL(IShellFolder parentFolder, string displayName)
        {
            uint num;
            IntPtr ptr;
            NativeMethods.CreateBindCtx();
            parentFolder.ParseDisplayName(IntPtr.Zero, null, displayName, out num, out ptr, 0);
            return ptr;
        }

        private static IntPtr PathToAbsolutePIDL(string path) =>
            GetShellFolderChildrenRelativePIDL(NativeMethods.SHGetDesktopFolder(), path);

        private static IEnumerable<FileSystemInfo> PathToFileSystemInfo(IEnumerable<string> paths)
        {
            foreach (string iteratorVariable0 in paths)
            {
                string path = iteratorVariable0;
                if (path.EndsWith(Path.DirectorySeparatorChar.ToString()) ||
                    path.EndsWith(Path.AltDirectorySeparatorChar.ToString()))
                {
                    path = path.Remove(path.Length - 1);
                }
                if (Directory.Exists(path))
                {
                    yield return new DirectoryInfo(path);
                }
                else
                {
                    if (!File.Exists(path))
                    {
                        throw new FileNotFoundException("The specified file or folder doesn't exists : " + path, path);
                    }
                    yield return new FileInfo(path);
                }
            }
        }

        private static IShellFolder PIDLToShellFolder(IntPtr pidl) =>
            PIDLToShellFolder(NativeMethods.SHGetDesktopFolder(), pidl);

        private static IShellFolder PIDLToShellFolder(IShellFolder parent, IntPtr pidl)
        {
            IShellFolder folder;
            Marshal.ThrowExceptionForHR(parent.BindToObject(pidl, null, ref IID_IShellFolder, out folder));
            return folder;
        }

        private static void SHOpenFolderAndSelectItems(IntPtr pidlFolder, IntPtr[] apidl, bool edit)
        {
            NativeMethods.SHOpenFolderAndSelectItems(pidlFolder, apidl, edit ? 1 : 0);
        }


        [ComImport, Guid("000214F2-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IEnumIDList
        {
            [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            int Next(uint celt, IntPtr rgelt, out uint pceltFetched);

            [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            int Skip([In] uint celt);

            [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            int Reset();

            [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            int Clone([MarshalAs(UnmanagedType.Interface)] out IEnumIDList ppenum);
        }

        [ComImport, Guid("000214E6-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
         ComConversionLoss]
        internal interface IShellFolder
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void ParseDisplayName(IntPtr hwnd, [In, MarshalAs(UnmanagedType.Interface)] IBindCtx pbc,
                [In, MarshalAs(UnmanagedType.LPWStr)] string pszDisplayName, out uint pchEaten, out IntPtr ppidl,
                [In, Out] ref uint pdwAttributes);

            [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            int EnumObjects([In] IntPtr hwnd, [In] SHCONT grfFlags,
                [MarshalAs(UnmanagedType.Interface)] out IEnumIDList ppenumIDList);

            [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            int BindToObject([In] IntPtr pidl, [In, MarshalAs(UnmanagedType.Interface)] IBindCtx pbc, [In] ref Guid riid,
                [MarshalAs(UnmanagedType.Interface)] out IShellFolder ppv);

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void BindToStorage([In] ref IntPtr pidl, [In, MarshalAs(UnmanagedType.Interface)] IBindCtx pbc,
                [In] ref Guid riid,
                out IntPtr ppv);

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void CompareIDs([In] IntPtr lParam, [In] ref IntPtr pidl1, [In] ref IntPtr pidl2);

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void CreateViewObject([In] IntPtr hwndOwner, [In] ref Guid riid, out IntPtr ppv);

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void GetAttributesOf([In] uint cidl, [In] IntPtr apidl, [In, Out] ref uint rgfInOut);

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void GetUIObjectOf([In] IntPtr hwndOwner, [In] uint cidl, [In] IntPtr apidl, [In] ref Guid riid,
                [In, Out] ref uint rgfReserved, out IntPtr ppv);

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void GetDisplayNameOf([In] ref IntPtr pidl, [In] uint uFlags, out IntPtr pName);

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void SetNameOf([In] IntPtr hwnd, [In] ref IntPtr pidl, [In, MarshalAs(UnmanagedType.LPWStr)] string pszName,
                [In] uint uFlags, [Out] IntPtr ppidlOut);
        }

        private class NativeMethods
        {
            private static readonly int pointerSize = Marshal.SizeOf(typeof(IntPtr));

            public static IBindCtx CreateBindCtx()
            {
                IBindCtx ctx;
                Marshal.ThrowExceptionForHR(CreateBindCtx_(0, out ctx));
                return ctx;
            }

            [DllImport("ole32.dll", EntryPoint = "CreateBindCtx")]
            public static extern int CreateBindCtx_(int reserved, out IBindCtx ppbc);

            [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
            public static extern IntPtr ILCreateFromPath([In, MarshalAs(UnmanagedType.LPWStr)] string pszPath);

            [DllImport("shell32.dll")]
            public static extern void ILFree([In] IntPtr pidl);

            public static IShellFolder SHGetDesktopFolder()
            {
                IShellFolder folder;
                Marshal.ThrowExceptionForHR(SHGetDesktopFolder_(out folder));
                return folder;
            }

            [DllImport("shell32.dll", EntryPoint = "SHGetDesktopFolder", CharSet = CharSet.Unicode, SetLastError = true)
            ]
            private static extern int SHGetDesktopFolder_(
                [MarshalAs(UnmanagedType.Interface)] out IShellFolder ppshf);

            public static void SHOpenFolderAndSelectItems(IntPtr pidlFolder, IntPtr[] apidl, int dwFlags)
            {
                uint cidl = (apidl != null) ? ((uint)apidl.Length) : 0;
                Marshal.ThrowExceptionForHR(SHOpenFolderAndSelectItems_(pidlFolder, cidl, apidl, dwFlags));
            }

            [DllImport("shell32.dll", EntryPoint = "SHOpenFolderAndSelectItems")]
            private static extern int SHOpenFolderAndSelectItems_([In] IntPtr pidlFolder, uint cidl,
                [In, Optional, MarshalAs(UnmanagedType.LPArray)] IntPtr[] apidl, int dwFlags);
        }

        [Flags]
        internal enum SHCONT : ushort
        {
            SHCONTF_CHECKING_FOR_CHILDREN = 0x10,
            SHCONTF_ENABLE_ASYNC = 0x8000,
            SHCONTF_FASTITEMS = 0x2000,
            SHCONTF_FLATLIST = 0x4000,
            SHCONTF_FOLDERS = 0x20,
            SHCONTF_INCLUDEHIDDEN = 0x80,
            SHCONTF_INIT_ON_FIRST_NEXT = 0x100,
            SHCONTF_NAVIGATION_ENUM = 0x1000,
            SHCONTF_NETPRINTERSRCH = 0x200,
            SHCONTF_NONFOLDERS = 0x40,
            SHCONTF_SHAREABLE = 0x400,
            SHCONTF_STORAGE = 0x800
        }
    }
}


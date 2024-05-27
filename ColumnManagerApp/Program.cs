using System.Diagnostics;
using System.Runtime.InteropServices;
using Windows.Win32;
using Windows.Win32.Foundation;
using Windows.Win32.UI.Shell;
using Windows.Win32.UI.Shell.Common;
using Windows.Win32.UI.Shell.PropertiesSystem;

/// <summary>
/// IN_PROGRESS. Currently displays all (and, separately, visible only) windows explorer columns and their values.
/// </summary>
public class Program
{
   public const int S_OK = 0;
   public const int S_FALSE = 1;

   public static void Main(string[] args)
   {
      Log("CLI start");

      int hr = PInvoke.SHGetDesktopFolder(out IShellFolder ppshf);
      if (hr != 0)
      {
         Log($"SHGetDesktopFolder failed. hr: {hr:x}");
         return;
      }

      uint pdwAttributes = 0;
      string folder = args[0];
      IntPtr ptrFolder = Marshal.StringToHGlobalAuto(folder);
      try
      {
         object ppv2 = ParseBindCreate(ppshf, ref pdwAttributes, ptrFolder);
         //IShellView shellView = ppv2 as IShellView;
         //IFolderView folderView = ppv2 as IFolderView;
         IColumnManager cm = (IColumnManager)ppv2;

         LogAllColumns(cm, folder);
         LogVisibleColumns(cm, folder);

         Log("CLI end");
      }
      finally
      {
         Marshal.FreeHGlobal(ptrFolder);
      }
   }
   private static string GetPropertyValue(string filePath, PROPERTYKEY propertyKey)
   {
      string propertyValue = string.Empty;
      Guid shellItem2Guid = typeof(IShellItem2).GUID;
      PInvoke.SHCreateItemFromParsingName(filePath, null, shellItem2Guid, out object shellItemObj);

      if (shellItemObj is IShellItem2 shellItem2)
      {
         shellItem2.GetPropertyStore(GETPROPERTYSTOREFLAGS.GPS_DEFAULT, typeof(IPropertyStore).GUID, out object ppv);
         var propertyStore = (IPropertyStore)ppv;
         unsafe
         {
            try
            {
               propertyStore.GetValue(in propertyKey, out var propVariant);

               int hr = PInvoke.PSFormatForDisplayAlloc(in propertyKey, propVariant, PROPDESC_FORMAT_FLAGS.PDFF_DEFAULT, out PWSTR ppszDisplay); // free ppszDisplay?
               if (hr != S_OK && hr != S_FALSE)
               {
                  Marshal.ThrowExceptionForHR(hr);
               }
               PInvoke.PropVariantClear(ref propVariant);

               propertyValue = ppszDisplay.ToString();              
            }
            finally
            {
               Marshal.ReleaseComObject(propertyStore);
            }
         }
      }

      return propertyValue;
   }

   private static void LogAllColumns(IColumnManager cm, string folder)
   {
      LogColumns(cm, folder, CM_ENUM_FLAGS.CM_ENUM_ALL);
   }
   private static void LogVisibleColumns(IColumnManager cm, string folder)
   {
      LogColumns(cm, folder, CM_ENUM_FLAGS.CM_ENUM_VISIBLE);
   }

   private static void LogColumns(IColumnManager cm, string folder, CM_ENUM_FLAGS colEnumFlag)
   {
      CM_COLUMNINFO colInfo = new();
      colInfo.cbSize = (uint)Marshal.SizeOf(colInfo);
      colInfo.dwMask = (uint)(CM_MASK.CM_MASK_NAME | CM_MASK.CM_MASK_STATE | CM_MASK.CM_MASK_WIDTH | CM_MASK.CM_MASK_IDEALWIDTH | CM_MASK.CM_MASK_DEFAULTWIDTH);
      CM_COLUMNINFO colInfo2 = new()
      {
         cbSize = (uint)Marshal.SizeOf(colInfo),
         dwMask = (uint)(CM_MASK.CM_MASK_NAME | CM_MASK.CM_MASK_STATE | CM_MASK.CM_MASK_WIDTH | CM_MASK.CM_MASK_IDEALWIDTH | CM_MASK.CM_MASK_DEFAULTWIDTH)
      };

      cm.GetColumnCount(colEnumFlag, out uint colCount);
      Span<PROPERTYKEY> propertyKeys = stackalloc PROPERTYKEY[(int)colCount];
      cm.GetColumns(colEnumFlag, propertyKeys);
      propertyKeys.Sort((PROPERTYKEY x, PROPERTYKEY y) =>
      {
         cm.GetColumnInfo(x, ref colInfo);
         cm.GetColumnInfo(y, ref colInfo2);
         return string.Compare(colInfo.wszName.ToString(), colInfo2.wszName.ToString());
      });

      string colType = colEnumFlag == CM_ENUM_FLAGS.CM_ENUM_ALL ? "all" : "visible";
      foreach (var propertyKey in propertyKeys)
      {
         cm.GetColumnInfo(propertyKey, ref colInfo);
         string colName = colInfo.wszName.ToString();
         string colValue = GetPropertyValue(folder, propertyKey);
         Log($"{colType} column: {colName}={colValue}");
      }
   }

   private static object ParseBindCreate(IShellFolder ppshf, ref uint pdwAttributes, IntPtr ptrFolder)
   {
      unsafe
      {
         PWSTR pwstr = new((char*)ptrFolder);
         ITEMIDLIST** ppIds = stackalloc ITEMIDLIST*[1];
         ppshf.ParseDisplayName(HWND.Null, null, pwstr, null, ppIds, ref pdwAttributes);
         Guid* riid = stackalloc Guid[1];
         riid[0] = typeof(IShellFolder).GUID;
         ppshf.BindToObject(ppIds[0], null, riid, out object ppv);
         IShellFolder sf = (IShellFolder)ppv;

         Guid* viid = stackalloc Guid[1];
         viid[0] = typeof(IShellView).GUID;
         sf.CreateViewObject(HWND.Null, viid, out object ppv2);
         return ppv2;
      }
   }

   private static void Log(string message)
   {
      Debug.WriteLine(message);
      Console.WriteLine(message);
   }
}
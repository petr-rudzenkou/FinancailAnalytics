using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace FinancialAnalytics.Wrappers.Office
{

    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    public struct InterfaceInfo
    {
        [MarshalAs(UnmanagedType.IUnknown)]
        public object punk;
        public Guid iid;
        public ushort wMethod;
    }

    /// <seealso href="http://msdn.microsoft.com/en-us/library/windows/desktop/ms693324.aspx"/>
    ///<seealso href="http://msdn.microsoft.com/en-us/library/8sesy69e.aspx"/>
    ///<seealso href="http://msdn.microsoft.com/en-us/library/ms228772.aspx"/>
    [ComImport, ComConversionLoss, InterfaceType((short)1),
     Guid("00000016-0000-0000-C000-000000000046")]
    public interface IMessageFilter
    {
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall,
            MethodCodeType = MethodCodeType.Runtime)]
        int HandleInComingCall([In] uint dwCallType, [In] IntPtr htaskCaller,
                               [In] uint dwTickCount,
                               [In, MarshalAs(UnmanagedType.LPArray)] InterfaceInfo[]
                                   lpInterfaceInfo);

        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall,
            MethodCodeType = MethodCodeType.Runtime)]
        int RetryRejectedCall([In] IntPtr htaskCallee, [In] uint dwTickCount,
                              [In] uint dwRejectType);

        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall,
            MethodCodeType = MethodCodeType.Runtime)]
        int MessagePending([In] IntPtr htaskCallee, [In] uint dwTickCount,
                           [In] uint dwPendingType);
    }
}
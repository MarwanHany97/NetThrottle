using System;
using System.Runtime.InteropServices;

namespace NetThrottle
{
    // WinDivert adress struct - 80 bytes for v2.2
    // i kept getting crashes until i figured out the size needs to match exactly lol
    [StructLayout(LayoutKind.Explicit, Size = 80)]
    public struct WinDivertAddress
    {
        [FieldOffset(0)]  public long Timestamp;
        [FieldOffset(8)]  public uint Flags;        // bunch of bitflags packed in here
        [FieldOffset(12)] public uint Reserved2;
        [FieldOffset(16)] public uint IfIdx;
        [FieldOffset(20)] public uint SubIfIdx;

        // helper props to pull out the flag bits we care about
        public readonly bool Outbound => (Flags & 0x00020000u) != 0;
        public readonly bool IPv6     => (Flags & 0x00100000u) != 0;
    }

    /// <summary>
    /// P/Invoke wrappers for WinDivert.dll 
    /// Only importing the functions we actualy need
    /// </summary>
    public static class WinDivertNative
    {
        private const string DLL = "WinDivert.dll";

        public const int WINDIVERT_LAYER_NETWORK = 0;

        [DllImport(DLL, CallingConvention = CallingConvention.Cdecl, SetLastError = true)]
        public static extern IntPtr WinDivertOpen(
            [MarshalAs(UnmanagedType.LPStr)] string filter,
            int layer, short priority, ulong flags);

        [DllImport(DLL, CallingConvention = CallingConvention.Cdecl, SetLastError = true)]
        public static extern bool WinDivertRecv(
            IntPtr handle, IntPtr pPacket, uint packetLen,
            out uint recvLen, ref WinDivertAddress pAddr);

        [DllImport(DLL, CallingConvention = CallingConvention.Cdecl, SetLastError = true)]
        public static extern bool WinDivertSend(
            IntPtr handle, IntPtr pPacket, uint packetLen,
            out uint sendLen, ref WinDivertAddress pAddr);

        [DllImport(DLL, CallingConvention = CallingConvention.Cdecl, SetLastError = true)]
        public static extern bool WinDivertClose(IntPtr handle);

        [DllImport(DLL, CallingConvention = CallingConvention.Cdecl, SetLastError = true)]
        public static extern bool WinDivertHelperCalcChecksums(
            IntPtr pPacket, uint packetLen,
            ref WinDivertAddress pAddr, ulong flags);

        public static readonly IntPtr INVALID_HANDLE = new IntPtr(-1);
    }

    /// <summary>
    /// IP Helper API imports - we use these to figure out which port belongs to which process
    /// its kinda janky but its the only reliable way on windows without a kernel driver
    /// </summary>
    public static class IpHelper
    {
        // AF_INET = 2 (ipv4)
        // TCP_TABLE_OWNER_PID_ALL = 5
        // UDP_TABLE_OWNER_PID = 1

        [DllImport("iphlpapi.dll", SetLastError = true)]
        public static extern uint GetExtendedTcpTable(
            IntPtr pTcpTable, ref int pdwSize, bool bOrder,
            int ulAf, int tableClass, uint reserved);

        [DllImport("iphlpapi.dll", SetLastError = true)]
        public static extern uint GetExtendedUdpTable(
            IntPtr pUdpTable, ref int pdwSize, bool bOrder,
            int ulAf, int tableClass, uint reserved);
    }
}

using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net;
using System.Runtime.InteropServices;
using System.Threading;

namespace NetThrottle
{
    public class ProcessRule
    {
        public bool BlockAll { get; set; }
        public bool LimitDownload { get; set; }
        public int DownloadKBps { get; set; }
        public bool LimitUpload { get; set; }
        public int UploadKBps { get; set; }

        // adaptive mode - when on, the UI adjusts these to hit the target avg
        public bool Adaptive { get; set; }
        public double AdjustedDLRate { get; set; }  // bytes/sec, 0 = use DownloadKBps
        public double AdjustedULRate { get; set; }  // bytes/sec, 0 = use UploadKBps

        public bool HasAnyRule =>
            BlockAll || (LimitDownload && DownloadKBps > 0) || (LimitUpload && UploadKBps > 0);

        public ProcessRule Clone() => new()
        {
            BlockAll = BlockAll,
            LimitDownload = LimitDownload,
            DownloadKBps = DownloadKBps,
            LimitUpload = LimitUpload,
            UploadKBps = UploadKBps,
            Adaptive = Adaptive,
            AdjustedDLRate = AdjustedDLRate,
            AdjustedULRate = AdjustedULRate
        };
    }

    // global limiter - caps everything regardless of per-process rules
    public class GlobalRule
    {
        public bool BlockAll { get; set; }
        public bool LimitDownload { get; set; }
        public int DownloadKBps { get; set; }
        public bool LimitUpload { get; set; }
        public int UploadKBps { get; set; }
        public bool Adaptive { get; set; }
        public double AdjustedDLRate { get; set; }
        public double AdjustedULRate { get; set; }
    }

    /// <summary>
    /// Token bucket algorythm for rate limiting.
    /// pretty standard - fill tokens over time, consume on packets, drop if empty.
    /// TCP retransmits and backs off naturaly which is what makes this whole thing work
    /// </summary>
    public class TokenBucket
    {
        private double _tokens;
        private double _maxTokens;
        private double _rate;
        private long _lastTicks;
        private readonly object _lock = new();

        public TokenBucket(double bytesPerSecond)
        {
            _rate = bytesPerSecond;
            _maxTokens = bytesPerSecond * 2;  // 2 sec burst
            _tokens = _maxTokens;
            _lastTicks = Stopwatch.GetTimestamp();
        }

        public double CurrentRate
        {
            get { lock (_lock) return _rate; }
        }

        public void SetRate(double bytesPerSecond)
        {
            lock (_lock)
            {
                _rate = bytesPerSecond;
                _maxTokens = bytesPerSecond * 2;
                if (_tokens > _maxTokens) _tokens = _maxTokens;
            }
        }

        public bool TryConsume(int bytes)
        {
            lock (_lock)
            {
                long now = Stopwatch.GetTimestamp();
                double elapsed = (double)(now - _lastTicks) / Stopwatch.Frequency;
                _lastTicks = now;
                _tokens = Math.Min(_maxTokens, _tokens + _rate * elapsed);
                if (_tokens >= bytes)
                {
                    _tokens -= bytes;
                    return true;
                }
                return false;
            }
        }
    }

    public class ByteCounters
    {
        public long DownloadBytes;
        public long UploadBytes;
    }

    /// <summary>
    /// The main engine. Intercepts packets via WinDivert, maps ports to PIDs,
    /// and applies global + per-process rate limiting / blocking.
    /// 
    /// Flow:
    /// 1. WinDivert grabs a packet
    /// 2. Parse IP header for src/dst ports
    /// 3. Look up which process owns that port (IP Helper API)
    /// 4. Track bytes for speed display
    /// 5. Apply global limit first, then per-process
    /// 6. Reinject or drop
    /// </summary>
    public class NetworkInterceptor
    {
        private IntPtr _handle = IntPtr.Zero;
        private Thread? _thread;
        private volatile bool _stopping;

        // per-process stuff
        private readonly ConcurrentDictionary<int, ProcessRule> _rules = new();
        private volatile Dictionary<ushort, int> _tcpPortPid = new();
        private volatile Dictionary<ushort, int> _udpPortPid = new();
        private readonly ConcurrentDictionary<(int pid, bool upload), TokenBucket> _buckets = new();
        private DateTime _lastMapRefresh = DateTime.MinValue;
        private readonly ConcurrentDictionary<int, ByteCounters> _byteCounters = new();

        // global limiter
        private volatile GlobalRule _globalRule = new();
        private TokenBucket? _globalDLBucket;
        private TokenBucket? _globalULBucket;
        private long _globalDLBytes;
        private long _globalULBytes;

        private long _packetsProcessed;
        private long _packetsDropped;

        public bool IsRunning => _thread != null && _thread.IsAlive;
        public long PacketsProcessed => Interlocked.Read(ref _packetsProcessed);
        public long PacketsDropped => Interlocked.Read(ref _packetsDropped);

        // --- global rule ---

        public void SetGlobalRule(GlobalRule rule) => _globalRule = rule;
        public GlobalRule GetGlobalRule() => _globalRule;

        public (long dl, long ul) SnapshotGlobalCounters()
        {
            long dl = Interlocked.Exchange(ref _globalDLBytes, 0);
            long ul = Interlocked.Exchange(ref _globalULBytes, 0);
            return (dl, ul);
        }

        // --- start/stop ---

        public string? Start()
        {
            if (IsRunning) return null;

            string exeDir = AppDomain.CurrentDomain.BaseDirectory;
            if (!System.IO.File.Exists(System.IO.Path.Combine(exeDir, "WinDivert.dll")))
                return "WinDivert.dll not found next to the executable.";
            if (!System.IO.File.Exists(System.IO.Path.Combine(exeDir, "WinDivert64.sys")))
                return "WinDivert64.sys not found next to the executable.";

            _stopping = false;
            _packetsProcessed = 0;
            _packetsDropped = 0;
            _buckets.Clear();
            _byteCounters.Clear();
            _globalDLBucket = null;
            _globalULBucket = null;
            Interlocked.Exchange(ref _globalDLBytes, 0);
            Interlocked.Exchange(ref _globalULBytes, 0);

            _handle = WinDivertNative.WinDivertOpen(
                "ip and (tcp or udp)",
                WinDivertNative.WINDIVERT_LAYER_NETWORK, 0, 0);

            if (_handle == WinDivertNative.INVALID_HANDLE)
            {
                int err = Marshal.GetLastWin32Error();
                return err switch
                {
                    5 => "Access denied. Run as Administrator.",
                    2 => "WinDivert driver not found.",
                    _ => $"WinDivert failed to open (error {err})."
                };
            }

            _thread = new Thread(RecvLoop) { Name = "NetThrottle-Recv", IsBackground = true };
            _thread.Start();
            return null;
        }

        public void Stop()
        {
            if (!IsRunning) return;
            _stopping = true;
            if (_handle != IntPtr.Zero)
            {
                WinDivertNative.WinDivertClose(_handle);
                _handle = IntPtr.Zero;
            }
            _thread?.Join(3000);
            _thread = null;
        }

        // --- per-process rules ---

        public void UpdateRule(int pid, ProcessRule rule)
        {
            if (rule.HasAnyRule)
                _rules[pid] = rule;
            else
            {
                _rules.TryRemove(pid, out _);
                // also clean up any leftover buckets for this pid
                _buckets.TryRemove((pid, true), out _);
                _buckets.TryRemove((pid, false), out _);
            }
        }

        public bool TryGetRule(int pid, out ProcessRule? rule) => _rules.TryGetValue(pid, out rule);

        public void UpdateRuleForPids(IEnumerable<int> pids, ProcessRule rule)
        {
            foreach (int pid in pids)
                UpdateRule(pid, rule.Clone());
        }

        /// <summary>
        /// Set the adaptive adjusted rate for a specific bucket directly.
        /// Called by the UI timer feedback loop when adaptive mode is on.
        /// </summary>
        public void SetBucketRate(int pid, bool upload, double bytesPerSec)
        {
            var key = (pid, upload);
            if (_buckets.TryGetValue(key, out var bucket))
                bucket.SetRate(bytesPerSec);
        }

        /// <summary>same but for the global buckets</summary>
        public void SetGlobalBucketRate(bool upload, double bytesPerSec)
        {
            if (upload)
                _globalULBucket?.SetRate(bytesPerSec);
            else
                _globalDLBucket?.SetRate(bytesPerSec);
        }

        /// <summary>grab byte counts since last call, per pid</summary>
        public Dictionary<int, (long dl, long ul)> SnapshotAndResetCounters()
        {
            var result = new Dictionary<int, (long dl, long ul)>();
            foreach (var kvp in _byteCounters)
            {
                long dl = Interlocked.Exchange(ref kvp.Value.DownloadBytes, 0);
                long ul = Interlocked.Exchange(ref kvp.Value.UploadBytes, 0);
                result[kvp.Key] = (dl, ul);
            }
            return result;
        }

        public List<(int pid, string name)> GetNetworkProcesses()
        {
            var tcpMap = BuildTcpPortPidMap();
            var udpMap = BuildUdpPortPidMap();
            var pids = new HashSet<int>(tcpMap.Values);
            pids.UnionWith(udpMap.Values);
            pids.Remove(0);

            var result = new List<(int pid, string name)>();
            foreach (int pid in pids)
            {
                try
                {
                    using var proc = Process.GetProcessById(pid);
                    result.Add((pid, proc.ProcessName));
                }
                catch { }
            }

            // also include procs with active rules even if no current connections
            foreach (var kvp in _rules)
            {
                if (pids.Contains(kvp.Key)) continue;
                try
                {
                    using var proc = Process.GetProcessById(kvp.Key);
                    result.Add((kvp.Key, proc.ProcessName));
                }
                catch { }
            }

            result.Sort((a, b) => string.Compare(a.name, b.name, StringComparison.OrdinalIgnoreCase));
            return result;
        }

        // ===========================================================
        //  MAIN PACKET LOOP
        // ===========================================================

        private void RecvLoop()
        {
            const int BUF_SIZE = 65535;
            IntPtr packetBuf = Marshal.AllocHGlobal(BUF_SIZE);

            try
            {
                while (!_stopping)
                {
                    var addr = new WinDivertAddress();
                    if (!WinDivertNative.WinDivertRecv(_handle, packetBuf, BUF_SIZE, out uint readLen, ref addr))
                        break;

                    Interlocked.Increment(ref _packetsProcessed);

                    // skip ipv6
                    if (addr.IPv6) { Reinject(packetBuf, readLen, ref addr); continue; }

                    // refresh port->pid map every 1.5s
                    if ((DateTime.UtcNow - _lastMapRefresh).TotalSeconds >= 1.5)
                    {
                        _tcpPortPid = BuildTcpPortPidMap();
                        _udpPortPid = BuildUdpPortPidMap();
                        _lastMapRefresh = DateTime.UtcNow;
                    }

                    byte ihl = (byte)((Marshal.ReadByte(packetBuf, 0) & 0x0F) * 4);
                    byte protocol = Marshal.ReadByte(packetBuf, 9);
                    int totalLen = (int)readLen;

                    if (ihl + 4 > totalLen || (protocol != 6 && protocol != 17))
                    {
                        Reinject(packetBuf, readLen, ref addr);
                        continue;
                    }

                    ushort srcPort = (ushort)IPAddress.NetworkToHostOrder(Marshal.ReadInt16(packetBuf, ihl));
                    ushort dstPort = (ushort)IPAddress.NetworkToHostOrder(Marshal.ReadInt16(packetBuf, ihl + 2));

                    bool isOutbound = addr.Outbound;
                    ushort localPort = isOutbound ? srcPort : dstPort;

                    int pid = -1;
                    var portMap = protocol == 6 ? _tcpPortPid : _udpPortPid;
                    portMap.TryGetValue(localPort, out pid);

                    // always track bytes for speed display
                    if (pid > 0)
                    {
                        var counters = _byteCounters.GetOrAdd(pid, _ => new ByteCounters());
                        if (isOutbound)
                            Interlocked.Add(ref counters.UploadBytes, totalLen);
                        else
                            Interlocked.Add(ref counters.DownloadBytes, totalLen);
                    }

                    // global totals
                    if (isOutbound)
                        Interlocked.Add(ref _globalULBytes, totalLen);
                    else
                        Interlocked.Add(ref _globalDLBytes, totalLen);

                    // ---- GLOBAL RULES ----
                    var gRule = _globalRule;
                    if (gRule.BlockAll)
                    {
                        Interlocked.Increment(ref _packetsDropped);
                        continue;
                    }

                    if (isOutbound && gRule.LimitUpload && gRule.UploadKBps > 0)
                    {
                        // figure out what rate to use
                        double rate = (gRule.Adaptive && gRule.AdjustedULRate > 0)
                            ? gRule.AdjustedULRate
                            : gRule.UploadKBps * 1024.0;

                        if (_globalULBucket == null)
                            _globalULBucket = new TokenBucket(rate);
                        else
                            _globalULBucket.SetRate(rate);

                        if (!_globalULBucket.TryConsume(totalLen))
                        {
                            Interlocked.Increment(ref _packetsDropped);
                            continue;
                        }
                    }
                    else if (!isOutbound && gRule.LimitDownload && gRule.DownloadKBps > 0)
                    {
                        double rate = (gRule.Adaptive && gRule.AdjustedDLRate > 0)
                            ? gRule.AdjustedDLRate
                            : gRule.DownloadKBps * 1024.0;

                        if (_globalDLBucket == null)
                            _globalDLBucket = new TokenBucket(rate);
                        else
                            _globalDLBucket.SetRate(rate);

                        if (!_globalDLBucket.TryConsume(totalLen))
                        {
                            Interlocked.Increment(ref _packetsDropped);
                            continue;
                        }
                    }

                    // ---- PER-PROCESS RULES ----
                    if (pid > 0 && _rules.TryGetValue(pid, out var rule))
                    {
                        if (rule.BlockAll)
                        {
                            Interlocked.Increment(ref _packetsDropped);
                            continue;
                        }

                        bool isUpload = isOutbound;
                        bool shouldLimit = isUpload
                            ? (rule.LimitUpload && rule.UploadKBps > 0)
                            : (rule.LimitDownload && rule.DownloadKBps > 0);

                        if (shouldLimit)
                        {
                            int targetKBps = isUpload ? rule.UploadKBps : rule.DownloadKBps;

                            // if adaptive, use the adjusted rate if we have one
                            double rate;
                            if (rule.Adaptive)
                            {
                                double adj = isUpload ? rule.AdjustedULRate : rule.AdjustedDLRate;
                                rate = adj > 0 ? adj : targetKBps * 1024.0;
                            }
                            else
                            {
                                rate = targetKBps * 1024.0;
                            }

                            var key = (pid, isUpload);
                            if (!_buckets.TryGetValue(key, out var bucket))
                            {
                                bucket = new TokenBucket(rate);
                                _buckets[key] = bucket;
                            }
                            else
                                bucket.SetRate(rate);

                            if (!bucket.TryConsume(totalLen))
                            {
                                Interlocked.Increment(ref _packetsDropped);
                                continue;
                            }
                        }
                    }

                    Reinject(packetBuf, readLen, ref addr);
                }
            }
            catch (Exception ex)
            {
                if (!_stopping) Debug.WriteLine($"RecvLoop error: {ex}");
            }
            finally { Marshal.FreeHGlobal(packetBuf); }
        }

        private void Reinject(IntPtr packetBuf, uint readLen, ref WinDivertAddress addr)
        {
            if (_handle == IntPtr.Zero || _stopping) return;
            try { WinDivertNative.WinDivertHelperCalcChecksums(packetBuf, readLen, ref addr, 0); } catch { }
            WinDivertNative.WinDivertSend(_handle, packetBuf, readLen, out _, ref addr);
        }

        // ===========================================================
        //  PORT -> PID MAPPING (IP Helper API)
        // ===========================================================

        private static Dictionary<ushort, int> BuildTcpPortPidMap()
        {
            var map = new Dictionary<ushort, int>();
            int size = 0;
            IpHelper.GetExtendedTcpTable(IntPtr.Zero, ref size, false, 2, 5, 0);
            if (size == 0) return map;
            IntPtr table = Marshal.AllocHGlobal(size);
            try
            {
                if (IpHelper.GetExtendedTcpTable(table, ref size, false, 2, 5, 0) != 0) return map;
                int count = Marshal.ReadInt32(table, 0);
                int offset = 4;
                for (int i = 0; i < count; i++)
                {
                    int dwLocalPort = Marshal.ReadInt32(table, offset + 8);
                    int dwOwningPid = Marshal.ReadInt32(table, offset + 20);
                    ushort port = (ushort)IPAddress.NetworkToHostOrder((short)(dwLocalPort & 0xFFFF));
                    if (dwOwningPid > 0 && port > 0) map[port] = dwOwningPid;
                    offset += 24;
                }
            }
            finally { Marshal.FreeHGlobal(table); }
            return map;
        }

        private static Dictionary<ushort, int> BuildUdpPortPidMap()
        {
            var map = new Dictionary<ushort, int>();
            int size = 0;
            IpHelper.GetExtendedUdpTable(IntPtr.Zero, ref size, false, 2, 1, 0);
            if (size == 0) return map;
            IntPtr table = Marshal.AllocHGlobal(size);
            try
            {
                if (IpHelper.GetExtendedUdpTable(table, ref size, false, 2, 1, 0) != 0) return map;
                int count = Marshal.ReadInt32(table, 0);
                int offset = 4;
                for (int i = 0; i < count; i++)
                {
                    int dwLocalPort = Marshal.ReadInt32(table, offset + 4);
                    int dwOwningPid = Marshal.ReadInt32(table, offset + 8);
                    ushort port = (ushort)IPAddress.NetworkToHostOrder((short)(dwLocalPort & 0xFFFF));
                    if (dwOwningPid > 0 && port > 0) map[port] = dwOwningPid;
                    offset += 12;
                }
            }
            finally { Marshal.FreeHGlobal(table); }
            return map;
        }
    }
}

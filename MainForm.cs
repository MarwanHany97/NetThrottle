using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Windows.Forms;

namespace NetThrottle
{
    public class RowTag
    {
        public bool IsGlobal { get; set; }
        public bool IsGroup { get; set; }
        public string ProcessName { get; set; } = "";
        public List<int> Pids { get; set; } = new();
        public bool Expanded { get; set; }
        public double DLSpeed { get; set; }  // smoothed (rolling avg)
        public double ULSpeed { get; set; }
    }

    public static class Theme
    {
        public static readonly Color WindowBg      = Color.FromArgb(250, 251, 253);
        public static readonly Color GridBg         = Color.White;
        public static readonly Color GridLine       = Color.FromArgb(232, 234, 238);
        public static readonly Color GlobalRowBg    = Color.FromArgb(30, 58, 95);
        public static readonly Color GlobalRowFg    = Color.White;
        public static readonly Color GroupRowBg     = Color.FromArgb(237, 241, 248);
        public static readonly Color GroupRowFg     = Color.FromArgb(30, 40, 60);
        public static readonly Color ChildRowBg     = Color.White;
        public static readonly Color ChildRowFg     = Color.FromArgb(50, 55, 65);
        public static readonly Color SelectBg       = Color.FromArgb(55, 120, 210);
        public static readonly Color SelectFg       = Color.White;
        public static readonly Color HeaderBg       = Color.FromArgb(245, 247, 250);
        public static readonly Color HeaderFg       = Color.FromArgb(80, 88, 102);
        public static readonly Color SortedHeaderBg = Color.FromArgb(210, 225, 248);
        public static readonly Color SortedHeaderFg = Color.FromArgb(30, 60, 130);
        public static readonly Color ToolBarBg      = Color.FromArgb(245, 247, 250);
        public static readonly Color StatusBg       = Color.FromArgb(238, 240, 244);
        public static readonly Color StatusFg       = Color.FromArgb(90, 96, 110);
        public static readonly Color StatusActiveFg = Color.FromArgb(30, 130, 76);
    }

    public class MainForm : Form
    {
        private readonly NetworkInterceptor _interceptor = new();
        private DataGridView _grid = null!;
        private ToolStripButton _btnStart = null!;
        private ToolStripButton _btnStop = null!;
        private ToolStripStatusLabel _statusLabel = null!;
        private System.Windows.Forms.Timer _refreshTimer = null!;
        private bool _suppressEvents;

        // sorting
        private string _sortColumn = "Process";
        private bool _sortAscending = true;

        private const int TIMER_MS = 1000;
        private const int AVG_WINDOW = 5;  // 5 second rolling average

        // rolling average buffers - keyed by PID
        // each queue holds the last N 1-second speed samples
        private readonly Dictionary<int, Queue<double>> _dlHistory = new();
        private readonly Dictionary<int, Queue<double>> _ulHistory = new();
        // same for global
        private readonly Queue<double> _globalDLHistory = new();
        private readonly Queue<double> _globalULHistory = new();

        // the smoothed speeds (what we display)
        private Dictionary<int, (double dl, double ul)> _pidSpeeds = new();
        private double _globalDLSpeed, _globalULSpeed;

        // adaptive feedback state - tracks the adjusted rate per pid/direction
        // we store the current adjusted bytes/sec that the feedback loop has settled on
        private readonly Dictionary<(int pid, bool upload), double> _adaptiveRates = new();
        private double _adaptiveGlobalDL, _adaptiveGlobalUL;

        // tray
        private NotifyIcon _trayIcon = null!;
        private ContextMenuStrip _trayMenu = null!;
        private bool _forceClose;

        // columns that count as "rule columns" - only react to changes on these
        private static readonly HashSet<string> RuleColumns = new()
            { "LimitDL", "DLLimit", "LimitUL", "ULLimit", "Block", "Adaptive" };

        public MainForm()
        {
            SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint, true);
            AutoScaleMode = AutoScaleMode.Dpi;
            InitializeComponent();
            SetupTrayIcon();
            RefreshProcessList();
        }

        // ==============================
        //  TRAY ICON
        // ==============================

        private void SetupTrayIcon()
        {
            _trayMenu = new ContextMenuStrip();

            var openItem = new ToolStripMenuItem("Open NetThrottle");
            openItem.Font = new Font(openItem.Font, FontStyle.Bold);
            openItem.Click += (_, _) => ShowFromTray();

            var statusItem = new ToolStripMenuItem("Status: Stopped") { Enabled = false, Name = "TrayStatus" };

            var exitItem = new ToolStripMenuItem("Exit");
            exitItem.Click += (_, _) => DoFullExit();

            _trayMenu.Items.AddRange(new ToolStripItem[] {
                openItem, new ToolStripSeparator(),
                statusItem, new ToolStripSeparator(),
                exitItem
            });

            _trayIcon = new NotifyIcon
            {
                Text = "NetThrottle",
                Icon = MakeTrayIcon(),
                ContextMenuStrip = _trayMenu,
                Visible = false
            };
            _trayIcon.DoubleClick += (_, _) => ShowFromTray();
        }

        private static Icon MakeTrayIcon()
        {
            using var bmp = new Bitmap(32, 32);
            using var g = Graphics.FromImage(bmp);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using var bg = new SolidBrush(Color.FromArgb(40, 90, 180));
            g.FillEllipse(bg, 1, 1, 30, 30);
            using var font = new Font("Segoe UI", 16f, FontStyle.Bold);
            var sz = g.MeasureString("N", font);
            g.DrawString("N", font, Brushes.White, (32 - sz.Width) / 2f, (32 - sz.Height) / 2f);
            return Icon.FromHandle(bmp.GetHicon());
        }

        private void ShowFromTray()
        {
            Show();
            WindowState = FormWindowState.Normal;
            Activate();
            _trayIcon.Visible = false;
        }

        private void HideToTray()
        {
            Hide();
            _trayIcon.Visible = true;
            _trayIcon.ShowBalloonTip(2000, "NetThrottle",
                "Running in background. Right-click tray icon to exit.", ToolTipIcon.Info);
        }

        private void DoFullExit()
        {
            _interceptor.SetGlobalRule(new GlobalRule());
            ClearAllProcessRules();
            _interceptor.Stop();
            _forceClose = true;
            _trayIcon.Visible = false;
            _trayIcon.Dispose();
            _refreshTimer.Stop();
            _refreshTimer.Dispose();
            Application.Exit();
        }

        private void ClearAllProcessRules()
        {
            var empty = new ProcessRule();
            for (int i = 0; i < _grid.Rows.Count; i++)
            {
                var tag = _grid.Rows[i].Tag as RowTag;
                if (tag == null || tag.IsGlobal) continue;
                if (tag.IsGroup)
                    _interceptor.UpdateRuleForPids(tag.Pids, empty);
                else
                    foreach (int pid in tag.Pids)
                        _interceptor.UpdateRule(pid, empty);
            }
        }

        private void SyncTrayStatus(bool running)
        {
            _trayIcon.Text = running ? "NetThrottle — Running" : "NetThrottle — Stopped";
            var item = _trayMenu.Items.Find("TrayStatus", false).FirstOrDefault();
            if (item != null) item.Text = running ? "Status: Running ●" : "Status: Stopped";
        }

        // ==============================
        //  FORM SETUP
        // ==============================

        private void InitializeComponent()
        {
            Text = "NetThrottle";
            Size = new Size(1120, 680);
            MinimumSize = new Size(960, 440);
            StartPosition = FormStartPosition.CenterScreen;
            BackColor = Theme.WindowBg;
            Icon = SystemIcons.Shield;
            Font = new Font("Segoe UI", 9f);

            var toolStrip = new ToolStrip
            {
                GripStyle = ToolStripGripStyle.Hidden,
                BackColor = Theme.ToolBarBg,
                Renderer = new ToolStripProfessionalRenderer(new ProfessionalColorTable()),
                Padding = new Padding(6, 3, 6, 3),
                RenderMode = ToolStripRenderMode.ManagerRenderMode
            };

            _btnStart = MakeBtn("▶  Start", "Start network interception");
            _btnStart.Click += BtnStart_Click;
            _btnStop = MakeBtn("■  Stop", "Stop network interception");
            _btnStop.Enabled = false;
            _btnStop.Click += BtnStop_Click;
            var btnRefresh = MakeBtn("↻  Refresh", "Refresh process list");
            btnRefresh.Click += (_, _) => RefreshProcessList();
            var btnCollapse = MakeBtn("⊟  Collapse All", "Collapse all groups");
            btnCollapse.Click += (_, _) => SetAllExpanded(false);
            var btnExpand = MakeBtn("⊞  Expand All", "Expand all groups");
            btnExpand.Click += (_, _) => SetAllExpanded(true);

            toolStrip.Items.AddRange(new ToolStripItem[] {
                _btnStart, _btnStop, new ToolStripSeparator(),
                btnRefresh, new ToolStripSeparator(),
                btnCollapse, btnExpand
            });

            var statusStrip = new StatusStrip { BackColor = Theme.StatusBg, SizingGrip = true };
            _statusLabel = new ToolStripStatusLabel("  Interceptor stopped")
            {
                Spring = true, TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = Theme.StatusFg, Font = new Font("Segoe UI", 8.5f)
            };
            statusStrip.Items.Add(_statusLabel);

            SetupGrid();

            Controls.Add(_grid);
            Controls.Add(toolStrip);
            Controls.Add(statusStrip);

            _refreshTimer = new System.Windows.Forms.Timer { Interval = TIMER_MS };
            _refreshTimer.Tick += RefreshTimer_Tick;
            _refreshTimer.Start();

            FormClosing += OnFormClosing;
        }

        private void OnFormClosing(object? sender, FormClosingEventArgs e)
        {
            if (_forceClose) return;
            if (_interceptor.IsRunning)
            {
                e.Cancel = true;
                HideToTray();
            }
            else
            {
                e.Cancel = true;
                DoFullExit();
            }
        }

        private static ToolStripButton MakeBtn(string text, string tip) =>
            new(text)
            {
                DisplayStyle = ToolStripItemDisplayStyle.Text,
                ToolTipText = tip,
                Padding = new Padding(4, 2, 4, 2),
                Font = new Font("Segoe UI", 9f)
            };

        // ==============================
        //  GRID SETUP
        // ==============================

        private void SetupGrid()
        {
            _grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                RowHeadersVisible = false,
                BackgroundColor = Theme.GridBg,
                GridColor = Theme.GridLine,
                BorderStyle = BorderStyle.None,
                CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                ColumnHeadersHeight = 34,
                RowTemplate = { Height = 30 },
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    Font = new Font("Segoe UI", 9f),
                    ForeColor = Theme.ChildRowFg,
                    BackColor = Theme.ChildRowBg,
                    SelectionBackColor = Theme.SelectBg,
                    SelectionForeColor = Theme.SelectFg,
                    Padding = new Padding(4, 2, 4, 2)
                },
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Theme.HeaderBg,
                    ForeColor = Theme.HeaderFg,
                    Font = new Font("Segoe UI Semibold", 8.5f),
                    Alignment = DataGridViewContentAlignment.MiddleLeft,
                    Padding = new Padding(4, 0, 4, 0),
                    SelectionBackColor = Theme.HeaderBg,
                    SelectionForeColor = Theme.HeaderFg
                }
            };

            _grid.Columns.AddRange(new DataGridViewColumn[]
            {
                TextCol("Process",    "Process",     180, 3.0f, true),
                TextCol("PID",        "PID",          55, 0.7f, true),
                TextCol("DLSpeedCol", "↓ DL Speed",   85, 1.1f, true, DataGridViewContentAlignment.MiddleRight),
                TextCol("ULSpeedCol", "↑ UL Speed",   85, 1.1f, true, DataGridViewContentAlignment.MiddleRight),
                CheckCol("LimitDL",   "Limit DL",     55),
                TextCol("DLLimit",    "DL KB/s",      60, 0.7f, false),
                CheckCol("LimitUL",   "Limit UL",     55),
                TextCol("ULLimit",    "UL KB/s",      60, 0.7f, false),
                CheckCol("Block",     "Block",         45),
                CheckCol("Adaptive",  "Adaptive",      60)
            });

            // commit checkbox immediately - dont wait for focus to move
            _grid.CurrentCellDirtyStateChanged += (_, _) =>
            {
                if (_grid.IsCurrentCellDirty) _grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
            };
            _grid.CellValueChanged += OnCellValueChanged;
            _grid.CellValidating += OnCellValidating;
            _grid.CellClick += OnCellClick;
            _grid.ColumnHeaderMouseClick += OnHeaderClick;
            _grid.CellFormatting += OnCellFormatting;
            _grid.CellPainting += OnCellPainting;
        }

        private static DataGridViewTextBoxColumn TextCol(string name, string header,
            int minW, float weight, bool ro,
            DataGridViewContentAlignment align = DataGridViewContentAlignment.MiddleLeft) =>
            new()
            {
                Name = name, HeaderText = header, MinimumWidth = minW,
                FillWeight = weight, ReadOnly = ro,
                SortMode = ro ? DataGridViewColumnSortMode.Programmatic : DataGridViewColumnSortMode.NotSortable,
                DefaultCellStyle = new DataGridViewCellStyle { Alignment = align }
            };

        private static DataGridViewCheckBoxColumn CheckCol(string name, string header, int minW) =>
            new()
            {
                Name = name, HeaderText = header, MinimumWidth = minW,
                FillWeight = 0.55f, SortMode = DataGridViewColumnSortMode.NotSortable
            };

        // ==============================
        //  START / STOP
        // ==============================

        private void BtnStart_Click(object? sender, EventArgs e)
        {
            string? err = _interceptor.Start();
            if (err != null)
            {
                MessageBox.Show(err, "NetThrottle", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            _btnStart.Enabled = false;
            _btnStop.Enabled = true;
            _statusLabel.ForeColor = Theme.StatusActiveFg;
            _statusLabel.Text = "  Interceptor running";
            SyncTrayStatus(true);

            // force-push all rules right now so they take effect immediately
            // this fixes the issue where rules set before starting dont get applied
            PushAllRules();
        }

        private void BtnStop_Click(object? sender, EventArgs e)
        {
            _interceptor.Stop();
            _btnStart.Enabled = true;
            _btnStop.Enabled = false;
            _statusLabel.ForeColor = Theme.StatusFg;
            _statusLabel.Text = "  Interceptor stopped";
            SyncTrayStatus(false);
            _adaptiveRates.Clear();
            _adaptiveGlobalDL = 0;
            _adaptiveGlobalUL = 0;
        }

        // ==============================
        //  TIMER - speed updates + adaptive feedback
        // ==============================

        private void RefreshTimer_Tick(object? sender, EventArgs e)
        {
            double sec = TIMER_MS / 1000.0;

            // --- collect raw 1-second samples ---

            var snap = _interceptor.SnapshotAndResetCounters();

            // push raw samples into rolling buffers and compute averages
            _pidSpeeds.Clear();
            foreach (var kv in snap)
            {
                double rawDL = kv.Value.dl / sec;
                double rawUL = kv.Value.ul / sec;

                PushSample(_dlHistory, kv.Key, rawDL);
                PushSample(_ulHistory, kv.Key, rawUL);

                double avgDL = GetAverage(_dlHistory, kv.Key);
                double avgUL = GetAverage(_ulHistory, kv.Key);
                _pidSpeeds[kv.Key] = (avgDL, avgUL);
            }

            // also update history for pids that had zero traffic this tick
            // (so their averages decay properly)
            foreach (var pid in _dlHistory.Keys.ToList())
            {
                if (!snap.ContainsKey(pid))
                {
                    PushSample(_dlHistory, pid, 0);
                    PushSample(_ulHistory, pid, 0);
                    _pidSpeeds[pid] = (GetAverage(_dlHistory, pid), GetAverage(_ulHistory, pid));
                }
            }

            // global
            var gs = _interceptor.SnapshotGlobalCounters();
            PushSampleQ(_globalDLHistory, gs.dl / sec);
            PushSampleQ(_globalULHistory, gs.ul / sec);
            _globalDLSpeed = GetAverageQ(_globalDLHistory);
            _globalULSpeed = GetAverageQ(_globalULHistory);

            // --- update grid display ---
            // only update speed columns which are read-only - no need to suppress events
            // since OnCellValueChanged checks for RuleColumns explicitly
            foreach (DataGridViewRow row in _grid.Rows)
            {
                var tag = row.Tag as RowTag;
                if (tag == null) continue;

                if (tag.IsGlobal)
                {
                    tag.DLSpeed = _globalDLSpeed;
                    tag.ULSpeed = _globalULSpeed;
                }
                else
                {
                    double dl = 0, ul = 0;
                    foreach (int pid in tag.Pids)
                        if (_pidSpeeds.TryGetValue(pid, out var s))
                        { dl += s.dl; ul += s.ul; }
                    tag.DLSpeed = dl;
                    tag.ULSpeed = ul;
                }

                row.Cells["DLSpeedCol"].Value = FmtSpeed(tag.DLSpeed);
                row.Cells["ULSpeedCol"].Value = FmtSpeed(tag.ULSpeed);
            }

            // --- adaptive feedback loop ---
            // for every process/direction with adaptive enabled, adjust the rate
            // to try and make the rolling average match the target
            RunAdaptiveFeedback();

            // --- status bar ---
            if (_interceptor.IsRunning)
            {
                _statusLabel.Text = $"  Interceptor running   ↓ {FmtSpeed(_globalDLSpeed)}   " +
                    $"↑ {FmtSpeed(_globalULSpeed)}   |   " +
                    $"Packets: {_interceptor.PacketsProcessed:N0}   Dropped: {_interceptor.PacketsDropped:N0}";
            }
        }

        // --- rolling average helpers ---

        private void PushSample(Dictionary<int, Queue<double>> dict, int pid, double value)
        {
            if (!dict.TryGetValue(pid, out var q))
            {
                q = new Queue<double>();
                dict[pid] = q;
            }
            q.Enqueue(value);
            while (q.Count > AVG_WINDOW) q.Dequeue();
        }

        private double GetAverage(Dictionary<int, Queue<double>> dict, int pid)
        {
            if (!dict.TryGetValue(pid, out var q) || q.Count == 0) return 0;
            return q.Average();
        }

        private void PushSampleQ(Queue<double> q, double value)
        {
            q.Enqueue(value);
            while (q.Count > AVG_WINDOW) q.Dequeue();
        }

        private double GetAverageQ(Queue<double> q) => q.Count == 0 ? 0 : q.Average();

        // --- adaptive feedback ---
        // basic proportional controller:
        // if avg is above target, lower the bucket rate to compensate
        // if avg is below target, ease the rate back up toward the target
        // this converges pretty quickly (a few seconds)

        private void RunAdaptiveFeedback()
        {
            if (!_interceptor.IsRunning) return;

            // per-process
            for (int i = 0; i < _grid.Rows.Count; i++)
            {
                var tag = _grid.Rows[i].Tag as RowTag;
                if (tag == null || tag.IsGlobal) continue;
                if (!tag.IsGroup && !tag.Pids.Any()) continue;

                bool adaptive = _grid.Rows[i].Cells["Adaptive"].Value is true;
                if (!adaptive) continue;

                bool limitDL = _grid.Rows[i].Cells["LimitDL"].Value is true;
                int dlKBps = ParseKB(_grid.Rows[i].Cells["DLLimit"].Value);
                bool limitUL = _grid.Rows[i].Cells["LimitUL"].Value is true;
                int ulKBps = ParseKB(_grid.Rows[i].Cells["ULLimit"].Value);

                foreach (int pid in tag.Pids)
                {
                    _pidSpeeds.TryGetValue(pid, out var spd);

                    if (limitDL && dlKBps > 0)
                        AdjustRate(pid, false, spd.dl, dlKBps * 1024.0);

                    if (limitUL && ulKBps > 0)
                        AdjustRate(pid, true, spd.ul, ulKBps * 1024.0);
                }
            }

            // global
            var gRule = _interceptor.GetGlobalRule();
            if (gRule.Adaptive)
            {
                if (gRule.LimitDownload && gRule.DownloadKBps > 0)
                {
                    double target = gRule.DownloadKBps * 1024.0;
                    _adaptiveGlobalDL = ComputeAdaptiveRate(_adaptiveGlobalDL, _globalDLSpeed, target);
                    var gr = _interceptor.GetGlobalRule();
                    gr.AdjustedDLRate = _adaptiveGlobalDL;
                    _interceptor.SetGlobalRule(gr);
                }
                if (gRule.LimitUpload && gRule.UploadKBps > 0)
                {
                    double target = gRule.UploadKBps * 1024.0;
                    _adaptiveGlobalUL = ComputeAdaptiveRate(_adaptiveGlobalUL, _globalULSpeed, target);
                    var gr = _interceptor.GetGlobalRule();
                    gr.AdjustedULRate = _adaptiveGlobalUL;
                    _interceptor.SetGlobalRule(gr);
                }
            }
        }

        private void AdjustRate(int pid, bool upload, double avgSpeed, double targetBytesPerSec)
        {
            var key = (pid, upload);
            double current = _adaptiveRates.TryGetValue(key, out var c) ? c : targetBytesPerSec;
            double adjusted = ComputeAdaptiveRate(current, avgSpeed, targetBytesPerSec);
            _adaptiveRates[key] = adjusted;

            // write the adjusted rate into the rule so the recv loop uses it
            if (_interceptor.TryGetRule(pid, out var rule) && rule != null)
            {
                if (upload)
                    rule.AdjustedULRate = adjusted;
                else
                    rule.AdjustedDLRate = adjusted;
                rule.Adaptive = true;
            }
        }

        /// <summary>
        /// Proportional feedback controller. Nudges the bucket rate to make
        /// the measured average converge toward the target.
        /// 
        /// If avg > target -> we're overshooting, tighten the rate
        /// If avg < target -> we're undershooting, loosen the rate (but never above target)
        /// </summary>
        private static double ComputeAdaptiveRate(double currentRate, double measuredAvg, double target)
        {
            // init to target if we havent started yet
            if (currentRate <= 0) currentRate = target;

            // if theres basically no traffic, just keep current rate
            if (measuredAvg < 100) return currentRate;

            double ratio = measuredAvg / target;

            double newRate;
            if (ratio > 1.02)
            {
                // over target - scale down proportionally with some dampening
                // use a stronger correction the further we are from target
                double correction = target / measuredAvg;
                newRate = currentRate * (0.3 + 0.7 * correction);
            }
            else if (ratio < 0.90)
            {
                // way under target - ease up faster
                newRate = currentRate * 1.15;
            }
            else if (ratio < 0.98)
            {
                // slightly under - ease up gently
                newRate = currentRate * 1.05;
            }
            else
            {
                // within 2% tolerance - leave it alone
                return currentRate;
            }

            // clamp: never go below 5% of target (would starve the connection)
            // and never go above 100% of target (would overshoot)
            double floor = target * 0.05;
            double ceiling = target;
            return Math.Clamp(newRate, floor, ceiling);
        }

        private static string FmtSpeed(double bps)
        {
            if (bps < 1) return "0 B/s";
            if (bps < 1024) return $"{bps:0} B/s";
            double kb = bps / 1024.0;
            if (kb < 1024) return $"{kb:0.0} KB/s";
            return $"{(kb / 1024.0):0.00} MB/s";
        }

        // ==============================
        //  ROW STYLING
        // ==============================

        private void OnCellFormatting(object? sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            var tag = _grid.Rows[e.RowIndex].Tag as RowTag;
            if (tag == null) return;

            if (tag.IsGlobal)
            {
                e.CellStyle.BackColor = Theme.GlobalRowBg;
                e.CellStyle.ForeColor = Theme.GlobalRowFg;
                e.CellStyle.SelectionBackColor = Color.FromArgb(45, 75, 120);
                e.CellStyle.SelectionForeColor = Color.White;
                e.CellStyle.Font = new Font("Segoe UI Semibold", 9f);
            }
            else if (tag.IsGroup)
            {
                e.CellStyle.BackColor = Theme.GroupRowBg;
                e.CellStyle.ForeColor = Theme.GroupRowFg;
                e.CellStyle.Font = new Font("Segoe UI Semibold", 9f);
            }
        }

        // sorted column header gets a blue highlight
        private void OnCellPainting(object? sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex != -1 || e.ColumnIndex < 0) return;

            string col = _grid.Columns[e.ColumnIndex].Name;
            bool sortable = col == "Process" || col == "PID" || col == "DLSpeedCol" || col == "ULSpeedCol";
            if (!sortable || col != _sortColumn) return;

            e.PaintBackground(e.ClipBounds, false);
            using (var brush = new SolidBrush(Theme.SortedHeaderBg))
                e.Graphics!.FillRectangle(brush, e.CellBounds);
            using (var pen = new Pen(Color.FromArgb(100, 140, 210), 2f))
                e.Graphics!.DrawLine(pen,
                    e.CellBounds.Left, e.CellBounds.Bottom - 1,
                    e.CellBounds.Right, e.CellBounds.Bottom - 1);

            string text = _grid.Columns[e.ColumnIndex].HeaderText + (_sortAscending ? " ▲" : " ▼");
            TextRenderer.DrawText(e.Graphics, text,
                new Font("Segoe UI Semibold", 8.5f), e.CellBounds, Theme.SortedHeaderFg,
                TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.EndEllipsis);
            e.Handled = true;
        }

        // ==============================
        //  EXPAND / COLLAPSE
        // ==============================

        private void OnCellClick(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || _grid.Columns[e.ColumnIndex].Name != "Process") return;
            var tag = _grid.Rows[e.RowIndex].Tag as RowTag;
            if (tag == null || !tag.IsGroup || tag.Pids.Count <= 1) return;
            tag.Expanded = !tag.Expanded;
            UpdateGroupLabel(e.RowIndex, tag);
            ApplyChildVis(e.RowIndex, tag);
        }

        private void UpdateGroupLabel(int idx, RowTag tag)
        {
            string arrow = tag.Expanded ? "▼" : "▶";
            _grid.Rows[idx].Cells["Process"].Value = $" {arrow}  {tag.ProcessName} ({tag.Pids.Count})";
        }

        private void ApplyChildVis(int gi, RowTag gt)
        {
            for (int i = gi + 1; i < _grid.Rows.Count; i++)
            {
                var ct = _grid.Rows[i].Tag as RowTag;
                if (ct == null || ct.IsGroup || ct.IsGlobal) break;
                if (ct.ProcessName != gt.ProcessName) break;
                _grid.Rows[i].Visible = gt.Expanded;
            }
        }

        private void SetAllExpanded(bool expand)
        {
            foreach (DataGridViewRow row in _grid.Rows)
            {
                var tag = row.Tag as RowTag;
                if (tag is { IsGroup: true } && tag.Pids.Count > 1)
                {
                    tag.Expanded = expand;
                    UpdateGroupLabel(row.Index, tag);
                    ApplyChildVis(row.Index, tag);
                }
            }
        }

        // ==============================
        //  SORTING
        // ==============================

        private void OnHeaderClick(object? sender, DataGridViewCellMouseEventArgs e)
        {
            string col = _grid.Columns[e.ColumnIndex].Name;
            if (col != "Process" && col != "PID" && col != "DLSpeedCol" && col != "ULSpeedCol") return;
            if (_sortColumn == col)
                _sortAscending = !_sortAscending;
            else { _sortColumn = col; _sortAscending = true; }
            foreach (DataGridViewColumn c in _grid.Columns)
                c.HeaderCell.SortGlyphDirection = SortOrder.None;
            RebuildGrid();
        }

        // ==============================
        //  CELL EDITS
        // ==============================

        private void OnCellValueChanged(object? sender, DataGridViewCellEventArgs e)
        {
            if (_suppressEvents || e.RowIndex < 0) return;

            // THIS IS THE KEY FIX: only react to rule columns
            // speed column updates (from the timer) wont trigger rule pushes anymore
            // which was causing all sorts of weird timing issues
            string colName = _grid.Columns[e.ColumnIndex].Name;
            if (!RuleColumns.Contains(colName)) return;

            var tag = _grid.Rows[e.RowIndex].Tag as RowTag;
            if (tag == null) return;

            if (tag.IsGlobal)
            {
                var gr = ReadGlobalRule(e.RowIndex);
                _interceptor.SetGlobalRule(gr);
                // reset adaptive state when rule changes
                if (!gr.Adaptive) { _adaptiveGlobalDL = 0; _adaptiveGlobalUL = 0; }
                return;
            }

            var rule = ReadRule(e.RowIndex);

            if (tag.IsGroup)
            {
                _interceptor.UpdateRuleForPids(tag.Pids, rule);

                // mirror to child rows
                _suppressEvents = true;
                for (int i = e.RowIndex + 1; i < _grid.Rows.Count; i++)
                {
                    var ct = _grid.Rows[i].Tag as RowTag;
                    if (ct == null || ct.IsGroup || ct.IsGlobal) break;
                    if (ct.ProcessName != tag.ProcessName) break;
                    WriteRule(i, rule);
                }
                _suppressEvents = false;

                // reset adaptive state for all pids in this group when rule changes
                if (!rule.Adaptive)
                    foreach (int pid in tag.Pids)
                    {
                        _adaptiveRates.Remove((pid, true));
                        _adaptiveRates.Remove((pid, false));
                    }
            }
            else
            {
                foreach (int pid in tag.Pids)
                {
                    _interceptor.UpdateRule(pid, rule.Clone());
                    if (!rule.Adaptive)
                    {
                        _adaptiveRates.Remove((pid, true));
                        _adaptiveRates.Remove((pid, false));
                    }
                }
            }
        }

        private void OnCellValidating(object? sender, DataGridViewCellValidatingEventArgs e)
        {
            string col = _grid.Columns[e.ColumnIndex].Name;
            if (col != "DLLimit" && col != "ULLimit") return;
            string val = e.FormattedValue?.ToString() ?? "";
            if (string.IsNullOrWhiteSpace(val)) return;
            if (!int.TryParse(val, out int n) || n < 0)
            {
                e.Cancel = true;
                MessageBox.Show("Please enter a positive number (KB/s).",
                    "NetThrottle", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // ==============================
        //  BUILD / REBUILD GRID
        // ==============================

        private void RefreshProcessList()
        {
            PushAllRules();

            var wasExpanded = new HashSet<string>();
            foreach (DataGridViewRow row in _grid.Rows)
            {
                var tag = row.Tag as RowTag;
                if (tag is { IsGroup: true, Expanded: true })
                    wasExpanded.Add(tag.ProcessName);
            }

            _grid.Rows.Clear();
            AddGlobalRow();

            var procs = _interceptor.GetNetworkProcesses();
            var groups = procs
                .GroupBy(p => p.name, StringComparer.OrdinalIgnoreCase)
                .Select(g => new ProcGroup { Name = g.Key, Pids = g.Select(p => p.pid).ToList() })
                .ToList();
            groups = SortGroups(groups);

            _suppressEvents = true;
            try
            {
                foreach (var grp in groups)
                {
                    bool multi = grp.Pids.Count > 1;
                    bool expanded = wasExpanded.Contains(grp.Name);

                    int gi = _grid.Rows.Add();
                    var gr = _grid.Rows[gi];
                    gr.Tag = new RowTag
                    {
                        IsGroup = true, ProcessName = grp.Name,
                        Pids = grp.Pids, Expanded = expanded
                    };
                    gr.Cells["Process"].Value = multi
                        ? $" {(expanded ? "▼" : "▶")}  {grp.Name} ({grp.Pids.Count})"
                        : $"      {grp.Name}";
                    gr.Cells["PID"].Value = multi ? "" : grp.Pids[0].ToString();
                    WriteRule(gi, GetExistingRule(grp.Pids));

                    if (multi)
                    {
                        foreach (int pid in grp.Pids)
                        {
                            int ci = _grid.Rows.Add();
                            _grid.Rows[ci].Tag = new RowTag
                            {
                                ProcessName = grp.Name,
                                Pids = new List<int> { pid }
                            };
                            _grid.Rows[ci].Cells["Process"].Value = $"            PID {pid}";
                            _grid.Rows[ci].Cells["PID"].Value = pid;
                            _grid.Rows[ci].Visible = expanded;

                            if (_interceptor.TryGetRule(pid, out var cr) && cr != null)
                                WriteRule(ci, cr);
                            else
                                WriteRule(ci, new ProcessRule());
                        }
                    }
                }
            }
            finally { _suppressEvents = false; }
        }

        private void AddGlobalRow()
        {
            _suppressEvents = true;
            int i = _grid.Rows.Add();
            var row = _grid.Rows[i];
            row.Tag = new RowTag { IsGlobal = true, ProcessName = "GLOBAL" };
            row.Cells["Process"].Value = "🌐  GLOBAL (all traffic)";
            row.Cells["PID"].Value = "";
            row.Height = 34;

            var gr = _interceptor.GetGlobalRule();
            row.Cells["LimitDL"].Value = gr.LimitDownload;
            row.Cells["DLLimit"].Value = gr.DownloadKBps > 0 ? gr.DownloadKBps.ToString() : "";
            row.Cells["LimitUL"].Value = gr.LimitUpload;
            row.Cells["ULLimit"].Value = gr.UploadKBps > 0 ? gr.UploadKBps.ToString() : "";
            row.Cells["Block"].Value = gr.BlockAll;
            row.Cells["Adaptive"].Value = gr.Adaptive;
            _suppressEvents = false;
        }

        private void RebuildGrid()
        {
            PushAllRules();
            RefreshProcessList();
        }

        // --- sorting ---

        private class ProcGroup
        {
            public string Name = "";
            public List<int> Pids = new();
        }

        private List<ProcGroup> SortGroups(List<ProcGroup> groups)
        {
            IEnumerable<ProcGroup> sorted = _sortColumn switch
            {
                "PID" => _sortAscending
                    ? groups.OrderBy(g => g.Pids.Min())
                    : groups.OrderByDescending(g => g.Pids.Min()),
                "DLSpeedCol" => _sortAscending
                    ? groups.OrderBy(g => SumDL(g.Pids))
                    : groups.OrderByDescending(g => SumDL(g.Pids)),
                "ULSpeedCol" => _sortAscending
                    ? groups.OrderBy(g => SumUL(g.Pids))
                    : groups.OrderByDescending(g => SumUL(g.Pids)),
                _ => _sortAscending
                    ? groups.OrderBy(g => g.Name, StringComparer.OrdinalIgnoreCase)
                    : groups.OrderByDescending(g => g.Name, StringComparer.OrdinalIgnoreCase)
            };
            return sorted.ToList();
        }

        private double SumDL(List<int> pids) =>
            pids.Sum(p => _pidSpeeds.TryGetValue(p, out var s) ? s.dl : 0);
        private double SumUL(List<int> pids) =>
            pids.Sum(p => _pidSpeeds.TryGetValue(p, out var s) ? s.ul : 0);

        // --- rule read/write ---

        private ProcessRule GetExistingRule(List<int> pids)
        {
            foreach (int pid in pids)
                if (_interceptor.TryGetRule(pid, out var r) && r != null) return r;
            return new ProcessRule();
        }

        private ProcessRule ReadRule(int idx)
        {
            var row = _grid.Rows[idx];
            return new ProcessRule
            {
                LimitDownload = row.Cells["LimitDL"].Value is true,
                DownloadKBps = ParseKB(row.Cells["DLLimit"].Value),
                LimitUpload = row.Cells["LimitUL"].Value is true,
                UploadKBps = ParseKB(row.Cells["ULLimit"].Value),
                BlockAll = row.Cells["Block"].Value is true,
                Adaptive = row.Cells["Adaptive"].Value is true
            };
        }

        private GlobalRule ReadGlobalRule(int idx)
        {
            var row = _grid.Rows[idx];
            return new GlobalRule
            {
                LimitDownload = row.Cells["LimitDL"].Value is true,
                DownloadKBps = ParseKB(row.Cells["DLLimit"].Value),
                LimitUpload = row.Cells["LimitUL"].Value is true,
                UploadKBps = ParseKB(row.Cells["ULLimit"].Value),
                BlockAll = row.Cells["Block"].Value is true,
                Adaptive = row.Cells["Adaptive"].Value is true
            };
        }

        private void WriteRule(int idx, ProcessRule rule)
        {
            var row = _grid.Rows[idx];
            row.Cells["LimitDL"].Value = rule.LimitDownload;
            row.Cells["DLLimit"].Value = rule.DownloadKBps > 0 ? rule.DownloadKBps.ToString() : "";
            row.Cells["LimitUL"].Value = rule.LimitUpload;
            row.Cells["ULLimit"].Value = rule.UploadKBps > 0 ? rule.UploadKBps.ToString() : "";
            row.Cells["Block"].Value = rule.BlockAll;
            row.Cells["Adaptive"].Value = rule.Adaptive;
        }

        private void PushAllRules()
        {
            for (int i = 0; i < _grid.Rows.Count; i++)
            {
                var tag = _grid.Rows[i].Tag as RowTag;
                if (tag == null) continue;

                if (tag.IsGlobal)
                {
                    _interceptor.SetGlobalRule(ReadGlobalRule(i));
                    continue;
                }

                var rule = ReadRule(i);
                if (tag.IsGroup)
                    _interceptor.UpdateRuleForPids(tag.Pids, rule);
                else
                    foreach (int pid in tag.Pids)
                        _interceptor.UpdateRule(pid, rule.Clone());
            }
        }

        private static int ParseKB(object? v)
        {
            if (v == null) return 0;
            return int.TryParse(v.ToString(), out int n) && n > 0 ? n : 0;
        }
    }
}

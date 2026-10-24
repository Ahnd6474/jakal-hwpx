using System;
using System.Collections;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

sealed class ChildJob : IDisposable {
    [StructLayout(LayoutKind.Sequential)] struct Limits {
        public long ProcessTime, JobTime; public uint Flags;
        public UIntPtr MinWorkingSet, MaxWorkingSet; public uint ActiveProcesses;
        public UIntPtr Affinity; public uint Priority, Scheduling;
    }
    [StructLayout(LayoutKind.Sequential)] struct Io {
        public ulong ReadOperations, WriteOperations, OtherOperations, ReadBytes, WriteBytes, OtherBytes;
    }
    [StructLayout(LayoutKind.Sequential)] struct Extended {
        public Limits Basic; public Io Counters;
        public UIntPtr ProcessMemory, JobMemory, PeakProcessMemory, PeakJobMemory;
    }
    [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)] static extern IntPtr CreateJobObject(IntPtr attributes, string name);
    [DllImport("kernel32.dll", SetLastError=true)] static extern bool SetInformationJobObject(IntPtr job, int type, IntPtr info, uint length);
    [DllImport("kernel32.dll", SetLastError=true)] static extern bool AssignProcessToJobObject(IntPtr job, IntPtr process);
    [DllImport("kernel32.dll")] static extern bool CloseHandle(IntPtr handle);
    IntPtr handle;
    public ChildJob(Process process) {
        handle = CreateJobObject(IntPtr.Zero, null);
        Extended limits = new Extended(); limits.Basic.Flags = 0x2000;
        int size = Marshal.SizeOf(typeof(Extended)); IntPtr data = Marshal.AllocHGlobal(size);
        try {
            Marshal.StructureToPtr(limits, data, false);
            if (handle == IntPtr.Zero || !SetInformationJobObject(handle, 9, data, (uint)size) || !AssignProcessToJobObject(handle, process.Handle)) {
                Dispose(); throw new InvalidOperationException("편집기 프로세스를 안전하게 시작하지 못했습니다.");
            }
        } finally { Marshal.FreeHGlobal(data); }
    }
    public void Dispose() { if (handle != IntPtr.Zero) { CloseHandle(handle); handle = IntPtr.Zero; } }
}

sealed class EditorProcess : IDisposable {
    public readonly string DataRoot;
    public Process Process;
    ChildJob job;
    StreamWriter log;
    readonly object logLock = new object();
    readonly string package = AppDomain.CurrentDomain.BaseDirectory;
    public EditorProcess(string dataRoot) { DataRoot = Path.GetFullPath(dataRoot); }
    public static string Quote(string value) { return "\"" + value.Replace("\"", "\\\"").TrimEnd('\\') + "\""; }
    void SecureRoot() {
        if (DataRoot.TrimEnd('\\') == Path.GetPathRoot(DataRoot).TrimEnd('\\')) throw new IOException("저장 위치로 드라이브 전체를 지정할 수 없습니다.");
        Directory.CreateDirectory(DataRoot);
        if ((File.GetAttributes(DataRoot) & FileAttributes.ReparsePoint) != 0) throw new IOException("연결된 폴더는 저장 위치로 사용할 수 없습니다.");
        SecurityIdentifier user = WindowsIdentity.GetCurrent().User;
        DirectorySecurity acl = new DirectorySecurity(); acl.SetAccessRuleProtection(true, false);
        foreach (SecurityIdentifier sid in new SecurityIdentifier[] { user, new SecurityIdentifier("S-1-5-18"), new SecurityIdentifier("S-1-5-32-544") })
            acl.AddAccessRule(new FileSystemAccessRule(sid, FileSystemRights.FullControl, InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow));
        Directory.SetAccessControl(DataRoot, acl);
    }
    void Record(object sender, DataReceivedEventArgs e) {
        if (e.Data == null) return;
        lock (logLock) {
            if (log == null) return;
            if (log.BaseStream.Length > 1024 * 1024) { log.BaseStream.SetLength(0); log.BaseStream.Position = 0; }
            log.WriteLine(e.Data); log.Flush();
        }
    }
    public void Start(int port) {
        if (!File.Exists(Path.Combine(package, "runtime", "python", "python.exe")) || !File.Exists(Path.Combine(package, "app", "desktop_bootstrap.py")))
            throw new IOException("ZIP 파일의 압축을 모두 푼 뒤 Jakal.exe를 실행해 주세요.");
        SecureRoot();
        string logs = Path.Combine(DataRoot, "logs"); Directory.CreateDirectory(logs);
        log = new StreamWriter(Path.Combine(logs, "application.log"), false, new UTF8Encoding(false));
        ProcessStartInfo start = new ProcessStartInfo(Path.Combine(package, "runtime", "python", "python.exe"));
        start.Arguments = "-I -B -X utf8 " + Quote(Path.Combine(package, "app", "desktop_bootstrap.py")) + " --data-root " + Quote(DataRoot) + " --port " + port;
        start.WorkingDirectory = Path.Combine(package, "app");
        start.UseShellExecute = false; start.CreateNoWindow = true;
        start.RedirectStandardInput = true; start.RedirectStandardOutput = true; start.RedirectStandardError = true;
        start.StandardOutputEncoding = Encoding.UTF8; start.StandardErrorEncoding = Encoding.UTF8;
        ArrayList remove = new ArrayList();
        foreach (string key in start.EnvironmentVariables.Keys)
            if (key.StartsWith("JAKAL_", StringComparison.OrdinalIgnoreCase) || key.StartsWith("PYTHON", StringComparison.OrdinalIgnoreCase) || key == "NODE_OPTIONS") remove.Add(key);
        foreach (string key in remove) start.EnvironmentVariables.Remove(key);
        start.EnvironmentVariables["PYTHONIOENCODING"] = "utf-8";
        Process = new Process(); Process.StartInfo = start;
        Process.OutputDataReceived += Record; Process.ErrorDataReceived += Record;
        Process.Start();
        try { job = new ChildJob(Process); }
        catch { Process.Kill(); throw; }
        Process.BeginOutputReadLine(); Process.BeginErrorReadLine();
    }
    public int ReadyPort() {
        if (Process == null || Process.HasExited) return -1;
        try {
            string text = File.ReadAllText(Path.Combine(DataRoot, "state", "ready.json"), Encoding.UTF8);
            int pid = int.Parse(Regex.Match(text, "\"pid\"\\s*:\\s*(\\d+)").Groups[1].Value);
            int port = int.Parse(Regex.Match(text, "\"port\"\\s*:\\s*(\\d+)").Groups[1].Value);
            return pid == Process.Id && port > 0 && port < 65536 ? port : 0;
        } catch (IOException) { return 0; } catch (FormatException) { return 0; }
    }
    public void Dispose() {
        if (Process != null) {
            try {
                if (!Process.HasExited) { Process.StandardInput.WriteLine("stop"); Process.StandardInput.Flush(); if (!Process.WaitForExit(15000)) Process.Kill(); }
                Process.WaitForExit();
            } catch (InvalidOperationException) { } catch (IOException) { }
            if (job != null) { job.Dispose(); job = null; }
            Process.Dispose(); Process = null;
        }
        lock (logLock) { if (log != null) { log.Dispose(); log = null; } }
    }
}

sealed class EditorWindow : Form {
    readonly EditorProcess editor;
    readonly Label status;
    readonly Button open;
    readonly System.Windows.Forms.Timer poll = new System.Windows.Forms.Timer();
    int port; bool shuttingDown; bool launched;
    readonly bool autoOpen;
    readonly bool selfTest;
    public bool WasReady { get { return launched; } }
    protected override bool ShowWithoutActivation { get { return selfTest; } }
    public const string WindowTitle = "Jakal 문서 편집기";
    public EditorWindow(string dataRoot, bool autoOpenBrowser, bool testMode = false) {
        editor = new EditorProcess(dataRoot); autoOpen = autoOpenBrowser; selfTest = testMode;
        if (selfTest) { ShowInTaskbar = false; Opacity = 0; }
        Text = WindowTitle; Font = new Font("맑은 고딕", 9F); ClientSize = new Size(520, 244);
        FormBorderStyle = FormBorderStyle.FixedSingle; MaximizeBox = false; StartPosition = FormStartPosition.CenterScreen;
        Label title = new Label { Text = "Jakal 문서 편집기", Font = new Font(Font.FontFamily, 16F, FontStyle.Bold), AutoSize = true, Location = new Point(24, 23) };
        Label hint = new Label { Text = "브라우저에서 HWP · HWPX 문서를 열고 편집하세요.", AutoSize = true, Location = new Point(26, 69) };
        status = new Label { Text = "편집기를 준비하고 있습니다…", AutoSize = false, Size = new Size(468, 35), Location = new Point(26, 102) };
        open = new Button { Text = "편집기 열기", Enabled = false, Location = new Point(26, 148), Size = new Size(145, 38) };
        open.Click += delegate { OpenBrowser(); };
        Button folder = new Button { Text = "저장 폴더 열기", Location = new Point(182, 148), Size = new Size(145, 38) };
        folder.Click += delegate { if (Directory.Exists(dataRoot)) Process.Start(new ProcessStartInfo(dataRoot) { UseShellExecute = true }); };
        Button close = new Button { Text = "종료", Location = new Point(339, 148), Size = new Size(145, 38) };
        close.Click += delegate { Close(); };
        Label footer = new Label { Text = "종료 전에는 편집 중인 문서를 저장해 주세요.", AutoSize = true, ForeColor = SystemColors.GrayText, Location = new Point(26, 205) };
        LinkLabel guide = new LinkLabel { Text = "사용 안내", AutoSize = true, Location = new Point(428, 205) };
        guide.LinkClicked += delegate { string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "guide.html"); if (File.Exists(path)) Process.Start(new ProcessStartInfo(path) { UseShellExecute = true }); };
        Controls.AddRange(new Control[] { title, hint, status, open, folder, close, footer, guide });
        poll.Interval = 250; poll.Tick += delegate {
            int ready = editor.ReadyPort();
            if (ready > 0 && !launched) { launched = true; port = ready; status.Text = "실행 중  ·  http://127.0.0.1:" + port; open.Enabled = true; if (autoOpen) OpenBrowser(); if (selfTest) Close(); }
            if (ready < 0) { poll.Stop(); open.Enabled = false; status.Text = "시작하지 못했거나 편집기가 종료되었습니다. 저장 폴더의 로그를 확인해 주세요."; if (selfTest) Close(); }
        };
        Shown += delegate { try { editor.Start(selfTest ? 0 : 18987); poll.Start(); } catch (Exception error) { status.Text = "실행하지 못했습니다: " + error.Message; if (selfTest) Close(); } };
        FormClosing += delegate(object sender, FormClosingEventArgs e) {
            if (shuttingDown) return;
            if (launched && !selfTest && MessageBox.Show(this, "문서를 저장한 뒤 종료해 주세요.\n편집기를 종료할까요?", WindowTitle, MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) { e.Cancel = true; return; }
            e.Cancel = true; shuttingDown = true; poll.Stop(); Enabled = false; status.Text = "편집기를 종료하고 있습니다…";
            ThreadPool.QueueUserWorkItem(delegate { editor.Dispose(); BeginInvoke((Action)delegate { Close(); }); });
        };
    }
    void OpenBrowser() { if (port > 0) Process.Start(new ProcessStartInfo("http://127.0.0.1:" + port + "/") { UseShellExecute = true }); }
}

static class Program {
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern IntPtr FindWindow(string cls, string title);
    [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr window);
    [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr window, int command);
    [STAThread] static int Main(string[] args) {
        string dataRoot = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "JakalWorkspace");
        bool headless = Array.IndexOf(args, "--headless") >= 0;
        bool uiCheck = Array.IndexOf(args, "--self-test-ui") >= 0;
        for (int i = 0; i + 1 < args.Length; i++) if (args[i] == "--data-root" && (headless || uiCheck)) dataRoot = args[i + 1];
        if (headless) {
            try {
                using (EditorProcess editor = new EditorProcess(dataRoot)) {
                    editor.Start(0); DateTime deadline = DateTime.UtcNow.AddSeconds(90);
                    while (editor.ReadyPort() == 0 && DateTime.UtcNow < deadline) Thread.Sleep(100);
                    if (editor.ReadyPort() <= 0) return 1;
                    Console.ReadLine();
                }
                return 0;
            } catch (Exception error) {
                try { Directory.CreateDirectory(Path.Combine(dataRoot, "logs")); File.WriteAllText(Path.Combine(dataRoot, "logs", "launcher-error.log"), error.ToString(), Encoding.UTF8); } catch { }
                return 1;
            }
        }
        if (uiCheck) {
            Application.EnableVisualStyles(); Application.SetCompatibleTextRenderingDefault(false);
            EditorWindow window = new EditorWindow(dataRoot, false, true);
            Application.Run(window);
            return window.WasReady ? 0 : 1;
        }
        string identity = WindowsIdentity.GetCurrent().User.Value;
        bool created;
        using (Mutex mutex = new Mutex(true, "Local\\JakalWorkspace-" + identity, out created)) {
            if (!created) { IntPtr window = FindWindow(null, EditorWindow.WindowTitle); if (window != IntPtr.Zero) { ShowWindow(window, 9); SetForegroundWindow(window); } return 0; }
            Application.EnableVisualStyles(); Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new EditorWindow(dataRoot, Array.IndexOf(args, "--no-browser") < 0));
            mutex.ReleaseMutex();
        }
        return 0;
    }
}

// PC 사용량 보고서 생성기
// - Windows 이벤트 로그(Winlogon 7001/7002, LocalSessionManager 21/23/24/25, 전원 이벤트)로 계정별 로그인·실사용 시간 계산
// - 레지스트리 UserAssist(계정별) + BAM 으로 실행한 앱 수집
// - template.html(리소스)에 데이터를 넣어 바탕화면에 HTML 보고서 저장 후 열기
// 빌드: build.ps1 (.NET Framework 4.x csc, C# 5)

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.Eventing.Reader;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using System.Xml;
using Microsoft.Win32;

class Ev
{
    public DateTime T;
    public string Prov = "";
    public int Id;
    public Dictionary<string, string> D = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    public List<string> P = new List<string>();
    public string Get(string k) { string v; return D.TryGetValue(k, out v) ? v : null; }
}

class Iv
{
    public string U; public DateTime A, B; public string R;
    public Iv(string u, DateTime a, DateTime b, string r) { U = u; A = a; B = b; R = r; }
}

class AppRec
{
    public string Name, Cat, Path, Key;
    public int? Runs; public double? Focus; public DateTime Last;
}

static class Program
{
    static readonly List<string> Warnings = new List<string>();
    static readonly CultureInfo Ko = new CultureInfo("ko-KR");
    static DateTime From, To, Now;
    static bool IsAdmin;
    static string PeriodName = "";
    static bool Anon;
    static readonly Dictionary<string, string> AnonMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

    static string UserLabel(int i)
    {
        if (i < 26) return "사용자 " + (char)('A' + i);
        return "사용자 " + (i + 1);
    }

    static string MapUser(string u)
    {
        if (!Anon || string.IsNullOrEmpty(u)) return u;
        string v;
        return AnonMap.TryGetValue(u, out v) ? v : "사용자";
    }

    static readonly List<string> SecretWords = new List<string>();
    static string Clean(string s)
    {
        if (!Anon || string.IsNullOrEmpty(s)) return s;
        string r = s;
        foreach (var w in SecretWords)
        {
            if (w.Length < 2) continue;
            r = Regex.Replace(r, "(?<![\\w])" + Regex.Escape(w) + "(?![\\w])", "사용자", RegexOptions.IgnoreCase);
        }
        r = Regex.Replace(r, @"[A-Za-z]:\\Users\\[^\\\s]+", @"<사용자 폴더>", RegexOptions.IgnoreCase);
        return r;
    }

    static void BuildAnonMap(List<string> users)
    {
        AnonMap.Clear();
        SecretWords.Clear();
        for (int i = 0; i < users.Count; i++) AnonMap[users[i]] = UserLabel(i);
        foreach (var u in users) SecretWords.Add(u);
        try { if (!string.IsNullOrEmpty(Environment.UserName)) SecretWords.Add(Environment.UserName); } catch { }
        try { if (!string.IsNullOrEmpty(Environment.MachineName)) SecretWords.Add(Environment.MachineName); } catch { }
        // 사용자 이름이 계정 표시 이름과 다를 수 있어 프로필 표시 이름도 추가
        try
        {
            using (var pr = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"))
            {
                if (pr != null)
                    foreach (string sid in pr.GetSubKeyNames())
                    {
                        if (!IsUserSid(sid)) continue;
                        string n = SidName(sid);
                        if (!string.IsNullOrEmpty(n)) SecretWords.Add(n);
                    }
            }
        }
        catch { }
    }

    static readonly string[] PeriodLabels = { "1주일", "2주일", "3주일", "4주일", "1달", "2달", "6개월", "1년" };
    static readonly string[] PeriodKeys = { "1w", "2w", "3w", "4w", "1m", "2m", "6m", "1y" };

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern uint QueryDosDevice(string lpDeviceName, StringBuilder lpTargetPath, int ucchMax);

    [STAThread]
    static int Main(string[] args)
    {
        try { Console.OutputEncoding = Encoding.UTF8; } catch { }
        Console.WriteLine("PC 사용량 보고서 생성기");
        Console.WriteLine("----------------------");
        bool noOpen = false; string outPath = null; bool noUi = false; bool periodSet = false;
        Now = DateTime.Now;
        DateTime today = Now.Date;
        int dow = ((int)today.DayOfWeek + 6) % 7; // 월=0
        DateTime monday = today.AddDays(-dow);
        From = monday; To = Now;

        try
        {
            for (int i = 0; i < args.Length; i++)
            {
                string a = args[i].ToLowerInvariant();
                if (a == "--lastweek" || a == "/lastweek") { From = monday.AddDays(-7); To = monday; PeriodName = "지난주"; periodSet = true; }
                else if (a == "--days" && i + 1 < args.Length) { int n = int.Parse(args[++i]); From = today.AddDays(-(n - 1)); To = Now; PeriodName = "최근 " + n + "일"; periodSet = true; }
                else if (a == "--period" && i + 1 < args.Length) { ApplyPeriod(args[++i]); periodSet = true; }
                else if (a == "--from" && i + 1 < args.Length) { From = DateTime.ParseExact(args[++i], "yyyy-MM-dd", CultureInfo.InvariantCulture); periodSet = true; }
                else if (a == "--to" && i + 1 < args.Length) { DateTime t = DateTime.ParseExact(args[++i], "yyyy-MM-dd", CultureInfo.InvariantCulture).AddDays(1); To = t > Now ? Now : t; periodSet = true; }
                else if (a == "--out" && i + 1 < args.Length) { outPath = args[++i]; }
                else if (a == "--no-open") { noOpen = true; }
                else if (a == "--no-ui") { noUi = true; }
                else if (a == "--anon" || a == "--share") { Anon = true; }
                else if (a == "--check-audit")
                {
                    string o; int rc = RunCmd(Path.Combine(Environment.SystemDirectory, "auditpol.exe"), "/get /subcategory:" + AuditGuid + " /r", out o);
                    Console.WriteLine("rc=" + rc); Console.WriteLine(o.Replace("\0", "")); Console.WriteLine("enabled=" + AuditEnabled());
                    return 0;
                }
                else if (a == "--setup-idle")
                {
                    int m = 10; int tmp;
                    if (i + 1 < args.Length && int.TryParse(args[i + 1], out tmp)) { m = tmp; i++; }
                    IsAdmin = new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator);
                    SetupIdle(m);
                    return 0;
                }
                else if (a == "--help" || a == "/?" || a == "-h")
                {
                    Console.WriteLine("사용법: PC사용량보고서.exe [--period 1w|2w|3w|4w|1m|2m|6m|1y | --lastweek | --days N | --from yyyy-MM-dd --to yyyy-MM-dd] [--out 파일.html] [--no-open] [--no-ui]");
                    Console.WriteLine("  기본: 기간 선택 창이 뜹니다 (1주일·2주일·3주일·4주일·1달·2달·6개월·1년)");
                    Console.WriteLine("  --period    기간을 바로 지정: 1w·2w·3w·4w(주 단위), 1m·2m·6m·1y(개월 단위)");
                    Console.WriteLine("  --no-ui     기간 선택 창 없이 이번 주 월요일 00:00 ~ 지금으로 생성");
                    Console.WriteLine("  --anon      공유용: 계정명을 '사용자 A/B/C'로, PC 이름과 앱 경로를 지워서 만듭니다");
                    Console.WriteLine("  --setup-idle [분]  입력 없는 시간(기본 10분 이상)을 실사용에서 빼도록 이 PC를 설정");
                    return 0;
                }
            }
            if (!periodSet && !noUi && Environment.UserInteractive)
            {
                string picked;
                if (!ShowPeriodDialog(out picked)) return 0;
                ApplyPeriod(picked);
            }
            if (To <= From) throw new Exception("기간이 올바르지 않습니다 (시작이 끝보다 늦음).");

            IsAdmin = new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator);
            if (!IsAdmin) Warnings.Add("관리자 권한 없이 실행되어 다른 계정의 앱 기록과 자리 비움 기록을 읽지 못했습니다. exe를 오른쪽 클릭 → ‘관리자 권한으로 실행’ 하면 모두 포함됩니다.");

            bool idleOn = IsAdmin && AuditEnabled();
            if (IsAdmin && !idleOn && !Console.IsInputRedirected)
            {
                Console.WriteLine();
                Console.WriteLine("키보드·마우스 입력이 없는 시간을 실사용에서 빼려면 이 PC에 설정이 필요합니다.");
                Console.WriteLine("(모든 계정에 10분 무입력 시 빈 화면 보호기 + 잠금/화면 보호기 기록. 비밀번호는 걸지 않습니다.)");
                Console.Write("지금 설정할까요? (Y/N): ");
                string ans = (Console.ReadLine() ?? "").Trim().ToLowerInvariant();
                if (ans == "y" || ans == "yes" || ans == "ㅛ") { SetupIdle(10); idleOn = AuditEnabled(); }
                Console.WriteLine();
            }
            if (!idleOn && IsAdmin) Warnings.Add("자리 비움(입력 없는 시간) 제외가 꺼져 있어, 실사용에 PC를 켜 두고 자리를 비운 시간도 포함됩니다. exe를 다시 실행해 설정을 켜면 그 이후 기록부터 적용됩니다.");

            Console.WriteLine("기간: {0} ~ {1}{2}", From.ToString("yyyy-MM-dd HH:mm"), To.ToString("yyyy-MM-dd HH:mm"),
                string.IsNullOrEmpty(PeriodName) ? "" : "  (" + PeriodName + ")");
            Console.WriteLine("[1/4] 이벤트 로그 읽는 중...");
            DateTime since = From.AddDays(-14);
            var sys = Read("System", new[] { 7001, 7002, 6005, 6006, 6008, 1074, 42, 107, 506, 507, 1, 12, 13 }, since)
                .Where(e =>
                    (e.Prov == "Microsoft-Windows-Winlogon" && (e.Id == 7001 || e.Id == 7002)) ||
                    (e.Prov == "EventLog" && (e.Id == 6005 || e.Id == 6006 || e.Id == 6008)) ||
                    (e.Prov == "User32" && e.Id == 1074) ||
                    (e.Prov == "Microsoft-Windows-Kernel-Power" && (e.Id == 42 || e.Id == 107 || e.Id == 506 || e.Id == 507)) ||
                    (e.Prov == "Microsoft-Windows-Power-Troubleshooter" && e.Id == 1) ||
                    (e.Prov == "Microsoft-Windows-Kernel-General" && (e.Id == 12 || e.Id == 13)))
                .OrderBy(e => e.T).ToList();
            var lsm = Read("Microsoft-Windows-TerminalServices-LocalSessionManager/Operational", new[] { 21, 23, 24, 25 }, since)
                .OrderBy(e => e.T).ToList();

            // ---- 전원: 절전 구간, 부팅, 종료 ----
            var sleeps = new List<KeyValuePair<DateTime, DateTime>>();
            foreach (var e in sys.Where(x => x.Prov.EndsWith("Power-Troubleshooter")))
            {
                DateTime s, w;
                if (TryUtc(e.Get("SleepTime"), out s) && TryUtc(e.Get("WakeTime"), out w) && w > s)
                    sleeps.Add(new KeyValuePair<DateTime, DateTime>(s, w));
            }
            if (sleeps.Count == 0)
            {   // 대체: Kernel-Power 42 → 다음 107
                DateTime? open = null;
                foreach (var e in sys.Where(x => x.Prov.EndsWith("Kernel-Power")))
                {
                    if (e.Id == 42) open = e.T;
                    else if (e.Id == 107 && open.HasValue) { sleeps.Add(new KeyValuePair<DateTime, DateTime>(open.Value, e.T)); open = null; }
                }
            }
            {   // 모던 스탠바이 506 → 507
                DateTime? open = null;
                foreach (var e in sys.Where(x => x.Prov.EndsWith("Kernel-Power") && (x.Id == 506 || x.Id == 507)))
                {
                    if (e.Id == 506) open = e.T;
                    else if (open.HasValue) { if ((e.T - open.Value).TotalMinutes >= 1) sleeps.Add(new KeyValuePair<DateTime, DateTime>(open.Value, e.T)); open = null; }
                }
            }
            var boots = sys.Where(x => x.Prov.EndsWith("Kernel-General") && x.Id == 12).Select(x => x.T).ToList();
            if (boots.Count == 0) boots = sys.Where(x => x.Prov == "EventLog" && x.Id == 6005).Select(x => x.T).ToList();
            var shutdowns = sys.Where(x => (x.Prov == "EventLog" && x.Id == 6006) || (x.Prov.EndsWith("Kernel-General") && x.Id == 13)).Select(x => x.T).OrderBy(t => t).ToList();
            var allTimes = sys.Select(x => x.T).Concat(lsm.Select(x => x.T)).OrderBy(t => t).ToList();
            var restarts = sys.Where(x => x.Prov == "User32" && x.Id == 1074).ToList();

            Func<DateTime, DateTime, DateTime> endBeforeBoot = (start, boot) =>
            {
                var sd = shutdowns.Where(t => t > start && t < boot).ToList();
                if (sd.Count > 0) return sd.Min();
                var last = allTimes.Where(t => t > start && t < boot).ToList();
                return last.Count > 0 ? last.Max() : start;
            };
            Func<DateTime, string> classify = t =>
            {
                if (sleeps.Any(s => s.Key >= t.AddSeconds(-15) && s.Key <= t.AddSeconds(40))) return "절전";
                var r = restarts.FirstOrDefault(x => x.T >= t.AddSeconds(-240) && x.T <= t.AddSeconds(10));
                if (r != null)
                {
                    string ty = (r.P.Count > 4 ? r.P[4] : "") ?? "";
                    return (ty.Contains("다시") || ty.ToLowerInvariant().Contains("restart")) ? "재시작" : "종료";
                }
                if (shutdowns.Any(s => s >= t && s <= t.AddSeconds(120))) return "종료";
                return "로그아웃";
            };

            // ---- 계정 이름 ----
            var display = new Dictionary<string, string>(); // key → 표시 이름
            Func<string, string> reg = name =>
            {
                string k = name.ToLowerInvariant();
                if (!display.ContainsKey(k)) display[k] = name;
                return display[k];
            };

            // ---- 로그인 세션 (Winlogon) ----
            Console.WriteLine("[2/4] 로그인·실사용 시간 계산 중...");
            var sessions = new List<Iv>();
            var wl = sys.Where(x => x.Prov == "Microsoft-Windows-Winlogon").ToList();
            var timeline = wl.Select(e => new { T = e.T, E = e, Boot = false })
                .Concat(boots.Select(b => new { T = b, E = (Ev)null, Boot = true }))
                .OrderBy(x => x.T).ThenBy(x => x.Boot ? 0 : 1).ToList();
            var openS = new Dictionary<string, DateTime>();
            foreach (var x in timeline)
            {
                if (x.Boot)
                {
                    foreach (var kv in openS.ToList()) sessions.Add(new Iv(kv.Key, kv.Value, endBeforeBoot(kv.Value, x.T), "종료"));
                    openS.Clear(); continue;
                }
                string sid = x.E.Get("UserSid");
                if (string.IsNullOrEmpty(sid) || !IsUserSid(sid)) continue;
                string u = reg(SidName(sid));
                if (x.E.Id == 7001)
                {
                    if (openS.ContainsKey(u)) sessions.Add(new Iv(u, openS[u], x.T, "로그아웃"));
                    openS[u] = x.T;
                }
                else if (openS.ContainsKey(u))
                {
                    sessions.Add(new Iv(u, openS[u], x.T, classify(x.T)));
                    openS.Remove(u);
                }
            }
            foreach (var kv in openS) sessions.Add(new Iv(kv.Key, kv.Value, Now, "사용 중"));

            // ---- 실사용 (LocalSessionManager) ----
            var active = new List<Iv>();
            var lsmUser = lsm.Where(e => IsRealUser(e.Get("User"))).ToList();
            if (lsmUser.Count == 0)
            {
                Warnings.Add("계정 전환 기록(LocalSessionManager 로그)이 없어 실사용 시간을 로그인 시간으로 대신했습니다.");
                active = sessions.Select(s => new Iv(s.U, s.A, s.B, s.R)).ToList();
            }
            else
            {
                var tl2 = lsmUser.Select(e => new { T = e.T, E = e, Boot = false })
                    .Concat(boots.Select(b => new { T = b, E = (Ev)null, Boot = true }))
                    .OrderBy(x => x.T).ThenBy(x => x.Boot ? 0 : 1).ToList();
                var openA = new Dictionary<string, KeyValuePair<string, DateTime>>(); // 세션ID → (계정, 시작)
                foreach (var x in tl2)
                {
                    if (x.Boot)
                    {
                        foreach (var kv in openA.Values.ToList()) active.Add(new Iv(kv.Key, kv.Value, endBeforeBoot(kv.Value, x.T), ""));
                        openA.Clear(); continue;
                    }
                    string sess = x.E.Get("SessionID") ?? "";
                    string u = reg(ShortName(x.E.Get("User")));
                    if (x.E.Id == 21 || x.E.Id == 25)
                    {
                        if (!openA.ContainsKey(sess)) openA[sess] = new KeyValuePair<string, DateTime>(u, x.T);
                    }
                    else if (openA.ContainsKey(sess))
                    {
                        var o = openA[sess];
                        active.Add(new Iv(o.Key, o.Value, x.T, ""));
                        openA.Remove(sess);
                    }
                }
                foreach (var kv in openA.Values) active.Add(new Iv(kv.Key, kv.Value, Now, ""));
                if (sessions.Count == 0)
                {   // Winlogon 기록이 없으면 LSM 로그온~로그오프로 세션 구성
                    Warnings.Add("Winlogon 로그인 기록이 없어 계정 전환 기록으로 로그인 세션을 구성했습니다.");
                    sessions = active.Select(s => new Iv(s.U, s.A, s.B, "")).ToList();
                }
            }

            sessions = Clip(Subtract(sessions, sleeps));
            active = Clip(Subtract(active, sleeps));

            // ---- 자리 비움: 화면 잠금(4800→4801), 화면 보호기(4802→4803, 시작 = 보호기 - 대기시간) ----
            var idleMin = new Dictionary<string, double>();
            if (IsAdmin)
            {
                var sec = Read("Security", new[] { 4800, 4801, 4802, 4803 }, since).OrderBy(e => e.T).ToList();
                int waitMin = IdleMinutesSetting();
                var cuts = new Dictionary<string, List<KeyValuePair<DateTime, DateTime>>>();
                var openLock = new Dictionary<string, DateTime>();
                var openSaver = new Dictionary<string, DateTime>();
                Func<string, DateTime, DateTime> closeAt = (u, start) =>
                {
                    DateTime end = Now;
                    var nb = boots.Where(b => b > start).ToList(); if (nb.Count > 0 && nb.Min() < end) end = nb.Min();
                    var nl = lsmUser.Where(e => (e.Id == 21) && e.T > start && string.Equals(ShortName(e.Get("User")), u, StringComparison.OrdinalIgnoreCase)).ToList();
                    if (nl.Count > 0 && nl.Min(e => e.T) < end) end = nl.Min(e => e.T);
                    return end;
                };
                foreach (var e in sec)
                {
                    string un = e.Get("TargetUserName") ?? e.Get("SubjectUserName");
                    if (string.IsNullOrEmpty(un)) continue;
                    string u = reg(un);
                    if (!cuts.ContainsKey(u)) cuts[u] = new List<KeyValuePair<DateTime, DateTime>>();
                    if (e.Id == 4800) { if (!openLock.ContainsKey(u)) openLock[u] = e.T; }
                    else if (e.Id == 4801) { if (openLock.ContainsKey(u)) { cuts[u].Add(new KeyValuePair<DateTime, DateTime>(openLock[u], e.T)); openLock.Remove(u); } }
                    else if (e.Id == 4802) { if (!openSaver.ContainsKey(u)) openSaver[u] = e.T.AddMinutes(-waitMin); }
                    else if (e.Id == 4803) { if (openSaver.ContainsKey(u)) { cuts[u].Add(new KeyValuePair<DateTime, DateTime>(openSaver[u], e.T)); openSaver.Remove(u); } }
                }
                foreach (var kv in openLock) cuts[kv.Key].Add(new KeyValuePair<DateTime, DateTime>(kv.Value, closeAt(kv.Key, kv.Value)));
                foreach (var kv in openSaver) cuts[kv.Key].Add(new KeyValuePair<DateTime, DateTime>(kv.Value, closeAt(kv.Key, kv.Value.AddMinutes(waitMin))));

                var after = new List<Iv>();
                foreach (var g in active.GroupBy(s => s.U))
                {
                    double before = g.Sum(s => (s.B - s.A).TotalMinutes);
                    List<KeyValuePair<DateTime, DateTime>> c;
                    var res = cuts.TryGetValue(g.Key, out c) ? Subtract(g.ToList(), c) : g.ToList();
                    after.AddRange(res);
                    idleMin[g.Key] = before - res.Sum(s => (s.B - s.A).TotalMinutes);
                }
                active = after.Where(s => (s.B - s.A).TotalSeconds >= 1).ToList();
            }

            var users = display.Values.Where(u => sessions.Any(s => s.U == u) || active.Any(s => s.U == u))
                .OrderByDescending(u => active.Where(s => s.U == u).Sum(s => (s.B - s.A).TotalSeconds)).ToList();

            if (sys.Count == 0) Warnings.Add("시스템 이벤트 로그에서 기록을 찾지 못했습니다.");
            DateTime oldest = allTimes.Count > 0 ? allTimes.Min() : Now;
            if (oldest > From.AddHours(12) && allTimes.Count > 0)
                Warnings.Add("이벤트 로그가 " + oldest.ToString("M/d HH:mm") + "부터만 남아 있어 그 이전 기록은 빠졌을 수 있습니다.");

            // ---- 앱 ----
            Console.WriteLine("[3/4] 실행한 앱 수집 중...");
            var apps = new Dictionary<string, object>();
            var other = new Dictionary<string, object>();
            var top = new Dictionary<string, string>();
            if (Anon) BuildAnonMap(users);
            CollectApps(users, apps, other, top);

            // ---- 출력 ----
            Console.WriteLine("[4/4] 보고서 만드는 중...");
            Func<DateTime, string> iso = d => d.ToString("yyyy-MM-dd'T'HH:mm:ss", CultureInfo.InvariantCulture);
            var data = new Dictionary<string, object>();
            data["computer"] = Anon ? "" : Environment.MachineName;
            data["share"] = Anon;
            data["from"] = iso(From);
            data["to"] = iso(To);
            data["fromText"] = From.ToString("yyyy-MM-dd(ddd)", Ko);
            data["toText"] = (To >= Now.AddMinutes(-1)) ? To.ToString("yyyy-MM-dd(ddd) HH:mm", Ko) : To.AddSeconds(-1).ToString("yyyy-MM-dd(ddd)", Ko);
            data["generated"] = Now.ToString("yyyy-MM-dd HH:mm", Ko);
            data["period"] = PeriodName ?? "";
            data["users"] = users.Select(MapUser).ToList();
            data["sessions"] = sessions.OrderBy(s => s.A).Select(s => new { u = MapUser(s.U), a = iso(s.A), b = iso(s.B), r = s.R }).ToList();
            data["active"] = active.OrderBy(s => s.A).Select(s => new { u = MapUser(s.U), a = iso(s.A), b = iso(s.B) }).ToList();
            data["apps"] = apps;
            data["other"] = other;
            data["top"] = top;            data["warnings"] = Warnings;
            data["idleOn"] = idleOn;
            data["idleWait"] = IdleMinutesSetting();
            data["idle"] = idleMin.ToDictionary(kv => kv.Key, kv => (object)Math.Round(kv.Value, 1));

            var js = new JavaScriptSerializer(); js.MaxJsonLength = int.MaxValue;
            string json = js.Serialize(data).Replace("</", "<\\/");
            string tpl;
            using (var st = Assembly.GetExecutingAssembly().GetManifestResourceStream("template.html"))
            using (var sr = new StreamReader(st, Encoding.UTF8)) tpl = sr.ReadToEnd();
            string html = tpl.Replace("/*__DATA__*/null", json);

            if (outPath == null)
            {
                string desk = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                outPath = Path.Combine(desk, string.Format("PC사용량_보고서_{0}_{1}_{2}.html",
                    Environment.MachineName, From.ToString("yyyyMMdd"), To.AddSeconds(-1).ToString("yyyyMMdd")));
            }
            File.WriteAllText(outPath, html, new UTF8Encoding(false));

            Console.WriteLine();
            foreach (var u in users)
            {
                double real = active.Where(s => s.U == u).Sum(s => (s.B - s.A).TotalMinutes);
                double login = sessions.Where(s => s.U == u).Sum(s => (s.B - s.A).TotalMinutes);
                Console.WriteLine("  {0,-14} 실사용 {1,4}시간 {2,2}분  (로그인 유지 {3}시간 {4}분)", u,
                    (int)(real / 60), (int)Math.Round(real % 60), (int)(login / 60), (int)Math.Round(login % 60));
            }
            foreach (var w in Warnings) Console.WriteLine("  ! " + w);
            Console.WriteLine();
            Console.WriteLine("저장: " + outPath);
            if (!noOpen)
            {
                try { Process.Start(new ProcessStartInfo("explorer.exe", "\"" + outPath + "\"") { UseShellExecute = true }); }
                catch { try { Process.Start(outPath); } catch { } }
                System.Threading.Thread.Sleep(1500);
            }
            return 0;
        }
        catch (Exception ex)
        {
            Console.WriteLine();
            Console.WriteLine("오류: " + ex.Message);
            Console.WriteLine(ex.StackTrace);
            Console.WriteLine("Enter 키를 누르면 닫힙니다.");
            try { Console.ReadLine(); } catch { }
            return 1;
        }
    }

    // ================= 기간 선택 =================
    static void CalcPeriod(string key, out DateTime from, out DateTime to, out string label)
    {
        string k = (key ?? "").Trim().ToLowerInvariant();
        DateTime today = Now.Date;
        from = today; to = Now; label = "";
        switch (k)
        {
            case "1w": from = today.AddDays(-6); label = "최근 1주일"; break;
            case "2w": from = today.AddDays(-13); label = "최근 2주일"; break;
            case "3w": from = today.AddDays(-20); label = "최근 3주일"; break;
            case "4w": from = today.AddDays(-27); label = "최근 4주일"; break;
            case "1m": from = today.AddMonths(-1); label = "최근 1개월"; break;
            case "2m": from = today.AddMonths(-2); label = "최근 2개월"; break;
            case "6m": from = today.AddMonths(-6); label = "최근 6개월"; break;
            case "1y": from = today.AddMonths(-12); label = "최근 1년"; break;
            default:
                throw new Exception("알 수 없는 기간입니다: " + key + " (1w, 2w, 3w, 4w, 1m, 2m, 6m, 1y)");
        }
    }

    static void ApplyPeriod(string key)
    {
        DateTime f, t; string label;
        CalcPeriod(key, out f, out t, out label);
        From = f; To = t; PeriodName = label;
    }

    static string RangeText(string key)
    {
        DateTime f, t; string label;
        CalcPeriod(key, out f, out t, out label);
        int days = (int)(t.Date - f.Date).TotalDays + 1;
        return string.Format("{0} ~ {1}  ({2}일)", f.ToString("yyyy-MM-dd"), t.ToString("yyyy-MM-dd"), Math.Max(days, 1));
    }

    static bool ShowPeriodDialog(out string picked)
    {
        picked = null;
        try
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
        }
        catch { }

        using (Form f = new Form())
        {
            f.Text = "PC 사용량 보고서 - 기간 선택";
            f.FormBorderStyle = FormBorderStyle.FixedDialog;
            f.StartPosition = FormStartPosition.CenterScreen;
            f.MaximizeBox = false;
            f.MinimizeBox = false;
            f.ClientSize = new Size(364, 300);
            f.Font = new Font("Malgun Gothic", 9.75f);

            Label title = new Label();
            title.Text = "보고서 기간을 선택하세요";
            title.Font = new Font("Malgun Gothic", 11f, FontStyle.Bold);
            title.SetBounds(22, 18, 320, 24);
            f.Controls.Add(title);

            List<RadioButton> radios = new List<RadioButton>();
            Label preview = new Label();
            preview.SetBounds(22, 186, 320, 40);
            preview.ForeColor = SystemColors.GrayText;
            preview.Font = new Font("Malgun Gothic", 9f);

            for (int i = 0; i < PeriodKeys.Length; i++)
            {
                RadioButton rb = new RadioButton();
                rb.Text = PeriodLabels[i];
                rb.Tag = PeriodKeys[i];
                rb.SetBounds(22 + (i % 2) * 168, 56 + (i / 2) * 30, 160, 24);
                if (i == 0) rb.Checked = true;
                rb.CheckedChanged += delegate(object s, EventArgs e)
                {
                    RadioButton cur = (RadioButton)s;
                    if (cur.Checked) preview.Text = RangeText((string)cur.Tag);
                };
                radios.Add(rb);
                f.Controls.Add(rb);
            }
            f.Controls.Add(preview);
            preview.Text = RangeText(PeriodKeys[0]);

            Button ok = new Button();
            ok.Text = "보고서 만들기";
            ok.SetBounds(140, 240, 110, 32);
            ok.DialogResult = DialogResult.OK;
            f.Controls.Add(ok);

            Button cancel = new Button();
            cancel.Text = "취소";
            cancel.SetBounds(258, 240, 84, 32);
            cancel.DialogResult = DialogResult.Cancel;
            f.Controls.Add(cancel);

            f.AcceptButton = ok;
            f.CancelButton = cancel;
            f.ActiveControl = radios[0];

            if (f.ShowDialog() != DialogResult.OK) return false;
            for (int i = 0; i < radios.Count; i++)
            {
                if (radios[i].Checked) { picked = (string)radios[i].Tag; break; }
            }
            return picked != null;
        }
    }

    // ================= 자리 비움 설정 =================
    const string AuditGuid = "{0CCE921C-69AE-11D9-BED3-505054503030}"; // 기타 로그온/로그오프 이벤트 (4800~4803)
    const string ConfKey = @"SOFTWARE\PCUsageReport";

    [DllImport("user32.dll", SetLastError = true)]
    static extern bool SystemParametersInfo(uint uiAction, uint uiParam, IntPtr pvParam, uint fWinIni);

    static int RunCmd(string exe, string args, out string output)
    {
        output = "";
        try
        {
            var psi = new ProcessStartInfo(exe, args) { UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true };
            using (var p = Process.Start(psi))
            {
                var ms = new MemoryStream(); p.StandardOutput.BaseStream.CopyTo(ms);
                string err = p.StandardError.ReadToEnd(); p.WaitForExit(15000);
                byte[] b = ms.ToArray();
                bool utf16 = b.Length > 3 && ((b[0] == 0xFF && b[1] == 0xFE) || (b[1] == 0 && b[3] == 0));
                output = (utf16 ? Encoding.Unicode.GetString(b) : Encoding.Default.GetString(b)).TrimStart('﻿') + err;
                return p.ExitCode;
            }
        }
        catch (Exception ex) { output = ex.Message; return -1; }
    }

    static bool AuditEnabled()
    {
        string o;
        if (RunCmd(Path.Combine(Environment.SystemDirectory, "auditpol.exe"), "/get /subcategory:" + AuditGuid + " /r", out o) != 0) return false;
        // CSV: Machine Name,Policy Target,Subcategory,Subcategory GUID,Inclusion Setting,...
        o = o.Replace("\0", "");
        foreach (var line in o.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
        {
            if (line.IndexOf(AuditGuid, StringComparison.OrdinalIgnoreCase) < 0) continue;
            var cols = line.Split(',');
            if (cols.Length < 5) continue;
            string inc = cols[4].Trim().ToLowerInvariant();
            return inc.Contains("success") || inc.Contains("성공");
        }
        return false;
    }

    static int IdleMinutesSetting()
    {
        try { using (var k = Registry.LocalMachine.OpenSubKey(ConfKey)) if (k != null) { object v = k.GetValue("IdleMinutes"); if (v is int && (int)v > 0) return (int)v; } }
        catch { }
        return 10;
    }

    static void SetupIdle(int minutes)
    {
        if (minutes < 1) minutes = 10;
        Console.WriteLine("자리 비움 제외 설정 ({0}분 무입력 기준)", minutes);
        if (!IsAdmin) { Console.WriteLine("  관리자 권한이 필요합니다."); return; }
        string o;
        int rc = RunCmd(Path.Combine(Environment.SystemDirectory, "auditpol.exe"), "/set /subcategory:" + AuditGuid + " /success:enable", out o);
        Console.WriteLine(rc == 0 ? "  [완료] 잠금·화면 보호기 기록 켜기" : "  [실패] 감사 정책: " + o.Trim());
        try { using (var k = Registry.LocalMachine.CreateSubKey(ConfKey)) k.SetValue("IdleMinutes", minutes, RegistryValueKind.DWord); } catch { }

        string scr = Path.Combine(Environment.SystemDirectory, "scrnsave.scr");
        string curSid = WindowsIdentity.GetCurrent().User.Value;
        using (var pl = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"))
        {
            foreach (string sid in pl.GetSubKeyNames())
            {
                if (!IsUserSid(sid)) continue;
                string profile = "";
                try { using (var k = pl.OpenSubKey(sid)) profile = Environment.ExpandEnvironmentVariables((k.GetValue("ProfileImagePath") ?? "").ToString()); } catch { }
                string hive = sid; bool loaded = false;
                using (var t = Registry.Users.OpenSubKey(sid)) if (t == null) hive = null;
                if (hive == null)
                {
                    string dat = Path.Combine(profile, "NTUSER.DAT");
                    if (!File.Exists(dat)) continue;
                    hive = "PCUR_" + Math.Abs(sid.GetHashCode());
                    if (RunReg("load \"HKU\\" + hive + "\" \"" + dat + "\"") != 0) { Console.WriteLine("  [실패] " + SidName(sid) + ": 레지스트리를 열지 못함"); continue; }
                    loaded = true;
                }
                try
                {
                    using (var d = Registry.Users.CreateSubKey(hive + @"\Control Panel\Desktop"))
                    {
                        d.SetValue("ScreenSaveActive", "1");
                        d.SetValue("ScreenSaveTimeOut", (minutes * 60).ToString());
                        if (string.IsNullOrEmpty((d.GetValue("SCRNSAVE.EXE") ?? "").ToString())) d.SetValue("SCRNSAVE.EXE", scr);
                        if (d.GetValue("ScreenSaverIsSecure") == null) d.SetValue("ScreenSaverIsSecure", "0");
                    }
                    Console.WriteLine("  [완료] " + SidName(sid) + ": {0}분 무입력 시 화면 보호기", minutes);
                }
                catch (Exception ex) { Console.WriteLine("  [실패] " + SidName(sid) + ": " + ex.Message); }
                finally
                {
                    if (loaded) { GC.Collect(); GC.WaitForPendingFinalizers(); for (int i = 0; i < 5 && RunReg("unload \"HKU\\" + hive + "\"") != 0; i++) System.Threading.Thread.Sleep(500); }
                }
            }
        }
        // 지금 로그인한 계정에는 바로 적용 (다른 계정은 다음 로그인부터)
        SystemParametersInfo(0x000F, (uint)(minutes * 60), IntPtr.Zero, 3);
        SystemParametersInfo(0x0011, 1, IntPtr.Zero, 3);
        Console.WriteLine("  이 시점 이후의 기록부터 자리 비움 시간이 빠집니다. 다른 계정은 다음 로그인부터 적용됩니다.");
    }

    // ================= 이벤트 로그 =================
    static List<Ev> Read(string log, int[] ids, DateTime since)
    {
        var res = new List<Ev>();
        string idq = string.Join(" or ", ids.Select(i => "EventID=" + i));
        string xp = "*[System[(" + idq + ") and TimeCreated[@SystemTime>='" +
                    since.ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'", CultureInfo.InvariantCulture) + "']]]";
        try
        {
            using (var r = new EventLogReader(new EventLogQuery(log, PathType.LogName, xp)))
            {
                EventRecord e;
                while ((e = r.ReadEvent()) != null)
                {
                    using (e)
                    {
                        if (!e.TimeCreated.HasValue) continue;
                        var ev = new Ev { T = e.TimeCreated.Value, Id = e.Id, Prov = e.ProviderName ?? "" };
                        try
                        {
                            var doc = new XmlDocument(); doc.LoadXml(e.ToXml());
                            foreach (XmlNode n in doc.SelectNodes("//*[local-name()='EventData']/*[local-name()='Data']"))
                            {
                                var a = n.Attributes != null ? n.Attributes["Name"] : null;
                                if (a != null) ev.D[a.Value] = n.InnerText;
                                ev.P.Add(n.InnerText);
                            }
                            var ud = doc.SelectSingleNode("//*[local-name()='UserData']/*");
                            if (ud != null) foreach (XmlNode c in ud.ChildNodes) ev.D[c.LocalName] = c.InnerText;
                        }
                        catch { }
                        res.Add(ev);
                    }
                }
            }
        }
        catch (EventLogNotFoundException) { }
        catch (Exception ex) { Warnings.Add(log + " 로그를 읽지 못했습니다: " + ex.Message); }
        return res;
    }

    static bool TryUtc(string s, out DateTime t)
    {
        t = DateTime.MinValue;
        if (string.IsNullOrEmpty(s) || s.Length < 19) return false;
        DateTime u;
        if (!DateTime.TryParseExact(s.Substring(0, 19), "yyyy-MM-dd'T'HH:mm:ss", CultureInfo.InvariantCulture,
            DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out u)) return false;
        t = DateTime.SpecifyKind(u, DateTimeKind.Utc).ToLocalTime();
        return true;
    }

    static bool IsUserSid(string sid) { return sid.StartsWith("S-1-5-21-") || sid.StartsWith("S-1-12-1-"); }

    static bool IsRealUser(string u)
    {
        if (string.IsNullOrEmpty(u)) return false;
        string l = u.ToLowerInvariant();
        return !(l.StartsWith("window manager\\") || l.StartsWith("font driver host\\") || l.StartsWith("nt authority\\") || l.StartsWith("nt service\\"));
    }

    static string ShortName(string u)
    {
        if (string.IsNullOrEmpty(u)) return "?";
        int i = u.LastIndexOf('\\');
        return i >= 0 ? u.Substring(i + 1) : u;
    }

    static readonly Dictionary<string, string> sidCache = new Dictionary<string, string>();
    static string SidName(string sid)
    {
        string n;
        if (sidCache.TryGetValue(sid, out n)) return n;
        try { n = ShortName(new SecurityIdentifier(sid).Translate(typeof(NTAccount)).Value); }
        catch
        {
            n = null;
            try
            {
                using (var k = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\" + sid))
                    if (k != null) n = Path.GetFileName(Environment.ExpandEnvironmentVariables((k.GetValue("ProfileImagePath") ?? "").ToString()));
            }
            catch { }
            if (string.IsNullOrEmpty(n)) n = sid;
        }
        sidCache[sid] = n;
        return n;
    }

    // ================= 구간 계산 =================
    static List<Iv> Subtract(List<Iv> src, List<KeyValuePair<DateTime, DateTime>> cuts)
    {
        var res = new List<Iv>();
        foreach (var iv in src)
        {
            var segs = new List<KeyValuePair<DateTime, DateTime>> { new KeyValuePair<DateTime, DateTime>(iv.A, iv.B) };
            foreach (var c in cuts)
            {
                var next = new List<KeyValuePair<DateTime, DateTime>>();
                foreach (var s in segs)
                {
                    if (c.Value <= s.Key || c.Key >= s.Value) { next.Add(s); continue; }
                    if (c.Key > s.Key) next.Add(new KeyValuePair<DateTime, DateTime>(s.Key, c.Key));
                    if (c.Value < s.Value) next.Add(new KeyValuePair<DateTime, DateTime>(c.Value, s.Value));
                }
                segs = next;
            }
            foreach (var s in segs.OrderBy(x => x.Key))
                res.Add(new Iv(iv.U, s.Key, s.Value, s.Value < iv.B ? "절전" : iv.R));
        }
        return res;
    }

    static List<Iv> Clip(List<Iv> src)
    {
        var res = new List<Iv>();
        foreach (var iv in src)
        {
            DateTime a = iv.A < From ? From : iv.A, b = iv.B > To ? To : iv.B;
            if (b > a) res.Add(new Iv(iv.U, a, b, iv.B > To ? (To >= Now.AddMinutes(-1) ? "사용 중" : "기간 끝") : iv.R));
        }
        return res;
    }

    // ================= 앱 =================
    static void CollectApps(List<string> users, Dictionary<string, object> appsOut, Dictionary<string, object> otherOut, Dictionary<string, string> topOut)
    {
        var dev = DeviceMap();
        string self = "";
        try { self = Process.GetCurrentProcess().MainModule.FileName.ToLowerInvariant(); } catch { }
        string curSid = WindowsIdentity.GetCurrent().User.Value;

        using (var pl = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"))
        {
            if (pl == null) return;
            foreach (string sid in pl.GetSubKeyNames())
            {
                if (!IsUserSid(sid)) continue;
                string uname = SidName(sid);
                string user = users.FirstOrDefault(u => string.Equals(u, uname, StringComparison.OrdinalIgnoreCase));
                if (user == null) continue;
                string profile = "";
                try { using (var k = pl.OpenSubKey(sid)) profile = Environment.ExpandEnvironmentVariables((k.GetValue("ProfileImagePath") ?? "").ToString()); } catch { }

                var main = new Dictionary<string, AppRec>();
                var bamOnly = new Dictionary<string, AppRec>();

                // UserAssist
                if (IsAdmin || sid == curSid)
                {
                    foreach (var ua in ReadUserAssist(sid, profile))
                    {
                        var rec = Describe(ua.Key, profile, dev, self);
                        if (rec == null) continue;
                        rec.Runs = ua.Value.Runs; rec.Focus = ua.Value.Focus; rec.Last = ua.Value.Last;
                        Merge(main, rec);
                    }
                }
                // BAM
                foreach (var b in ReadBam(sid))
                {
                    var rec = Describe(b.Key, profile, dev, self);
                    if (rec == null) continue;
                    rec.Last = b.Value;
                    AppRec m;
                    if (main.TryGetValue(rec.Key, out m)) { if (rec.Last > m.Last) m.Last = rec.Last; }
                    else Merge(bamOnly, rec);
                }

                var list = main.Values.OrderByDescending(a => a.Focus ?? -1).ThenByDescending(a => a.Last).ToList();
                appsOut[MapUser(user)] = list.Select(a => (object)new
                {
                    n = Clean(a.Name), c = a.Cat, r = a.Runs, f = a.Focus.HasValue ? (double?)Math.Round(a.Focus.Value, 1) : null,
                    l = a.Last.ToString("yyyy-MM-dd'T'HH:mm:ss", CultureInfo.InvariantCulture), p = Anon ? null : a.Path
                }).ToList();
                otherOut[MapUser(user)] = bamOnly.Values.OrderByDescending(a => a.Last).Take(60).Select(a => (object)new
                {
                    n = Clean(a.Name), l = a.Last.ToString("yyyy-MM-dd'T'HH:mm:ss", CultureInfo.InvariantCulture), p = Anon ? null : a.Path
                }).ToList();
                var topNames = list.Where(a => (a.Focus ?? 0) > 0 && a.Cat != "시스템").Take(3).Select(a => Clean(a.Name)).ToList();
                if (topNames.Count == 0) topNames = list.Where(a => a.Cat != "시스템").OrderByDescending(a => a.Runs ?? 0).Take(3).Select(a => Clean(a.Name)).ToList();
                if (topNames.Count > 0) topOut[MapUser(user)] = string.Join(" · ", topNames);
            }
        }
    }

    static void Merge(Dictionary<string, AppRec> d, AppRec r)
    {
        AppRec m;
        if (!d.TryGetValue(r.Key, out m)) { d[r.Key] = r; return; }
        if (r.Last > m.Last) m.Last = r.Last;
        if (r.Runs.HasValue && (!m.Runs.HasValue || r.Runs > m.Runs)) m.Runs = r.Runs;
        if (r.Focus.HasValue && (!m.Focus.HasValue || r.Focus > m.Focus)) m.Focus = r.Focus;
        if (string.IsNullOrEmpty(m.Path) || (m.Path.EndsWith(".lnk", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrEmpty(r.Path))) m.Path = r.Path;
        if (m.Cat == "기타" && r.Cat != "기타") m.Cat = r.Cat;
    }

    class UaVal { public int? Runs; public double? Focus; public DateTime Last; }

    static List<KeyValuePair<string, UaVal>> ReadUserAssist(string sid, string profile)
    {
        var res = new List<KeyValuePair<string, UaVal>>();
        string hiveName = sid; bool loaded = false;
        using (var t = Registry.Users.OpenSubKey(sid)) { if (t == null) hiveName = null; }
        if (hiveName == null)
        {
            if (!IsAdmin) return res;
            string dat = Path.Combine(profile, "NTUSER.DAT");
            if (!File.Exists(dat)) return res;
            hiveName = "PCUR_" + Math.Abs(sid.GetHashCode());
            if (RunReg("load \"HKU\\" + hiveName + "\" \"" + dat + "\"") != 0) { Warnings.Add(SidName(sid) + " 계정의 레지스트리를 열지 못했습니다."); return res; }
            loaded = true;
        }
        try
        {
            using (var ua = Registry.Users.OpenSubKey(hiveName + @"\Software\Microsoft\Windows\CurrentVersion\Explorer\UserAssist"))
            {
                if (ua != null)
                    foreach (string g in ua.GetSubKeyNames())
                        using (var cnt = ua.OpenSubKey(g + @"\Count"))
                        {
                            if (cnt == null) continue;
                            foreach (string vn in cnt.GetValueNames())
                            {
                                var d = cnt.GetValue(vn) as byte[];
                                if (d == null || d.Length < 68) continue;
                                long ft = BitConverter.ToInt64(d, 60);
                                if (ft <= 0) continue;
                                DateTime last;
                                try { last = DateTime.FromFileTime(ft); } catch { continue; }
                                if (last < From || last > To) continue;
                                int runs = BitConverter.ToInt32(d, 4);
                                int focusMs = BitConverter.ToInt32(d, 12);
                                res.Add(new KeyValuePair<string, UaVal>(Rot13(vn), new UaVal
                                {
                                    Runs = runs > 0 ? (int?)runs : null,
                                    Focus = focusMs > 0 ? (double?)(focusMs / 60000.0) : null,
                                    Last = last
                                }));
                            }
                        }
            }
        }
        catch (Exception ex) { Warnings.Add(SidName(sid) + " 계정 앱 기록 읽기 실패: " + ex.Message); }
        finally
        {
            if (loaded)
            {
                GC.Collect(); GC.WaitForPendingFinalizers();
                for (int i = 0; i < 5 && RunReg("unload \"HKU\\" + hiveName + "\"") != 0; i++) System.Threading.Thread.Sleep(500);
            }
        }
        return res;
    }

    static List<KeyValuePair<string, DateTime>> ReadBam(string sid)
    {
        var res = new List<KeyValuePair<string, DateTime>>();
        foreach (string p in new[] { @"SYSTEM\CurrentControlSet\Services\bam\State\UserSettings\", @"SYSTEM\CurrentControlSet\Services\bam\UserSettings\" })
        {
            try
            {
                using (var k = Registry.LocalMachine.OpenSubKey(p + sid))
                {
                    if (k == null) continue;
                    foreach (string vn in k.GetValueNames())
                    {
                        var d = k.GetValue(vn) as byte[];
                        if (d == null || d.Length < 8) continue;
                        DateTime t;
                        try { t = DateTime.FromFileTime(BitConverter.ToInt64(d, 0)); } catch { continue; }
                        if (t >= From && t <= To) res.Add(new KeyValuePair<string, DateTime>(vn, t));
                    }
                    if (res.Count > 0) break;
                }
            }
            catch { }
        }
        return res;
    }

    static int RunReg(string args)
    {
        try
        {
            var psi = new ProcessStartInfo(Path.Combine(Environment.SystemDirectory, "reg.exe"), args)
            { UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true };
            using (var p = Process.Start(psi)) { p.StandardOutput.ReadToEnd(); p.StandardError.ReadToEnd(); p.WaitForExit(15000); return p.ExitCode; }
        }
        catch { return -1; }
    }

    static string Rot13(string s)
    {
        var c = s.ToCharArray();
        for (int i = 0; i < c.Length; i++)
        {
            char x = c[i];
            if (x >= 'A' && x <= 'Z') c[i] = (char)('A' + (x - 'A' + 13) % 26);
            else if (x >= 'a' && x <= 'z') c[i] = (char)('a' + (x - 'a' + 13) % 26);
        }
        return new string(c);
    }

    static Dictionary<string, string> DeviceMap()
    {
        var m = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var drive in Environment.GetLogicalDrives())
        {
            string dl = drive.TrimEnd('\\');
            var sb = new StringBuilder(512);
            if (QueryDosDevice(dl, sb, sb.Capacity) > 0) m[sb.ToString()] = dl;
        }
        return m;
    }

    // ---------- 앱 이름·분류 ----------
    static readonly string[,] AumidMap = {
        {"Microsoft.Windows.Explorer","파일 탐색기"}, {"Microsoft.Windows.ControlPanel","제어판"}, {"windows.immersivecontrolpanel","설정"},
        {"Microsoft.Windows.RemoteDesktop","원격 데스크톱 연결"}, {"Microsoft.Windows.Photos","사진"}, {"Microsoft.Windows.SecHealthUI","Windows 보안"},
        {"Microsoft.SecHealthUI","Windows 보안"}, {"Microsoft.VisualStudioCode","Visual Studio Code"}, {"Microsoft.WindowsNotepad","메모장"},
        {"Microsoft.WindowsStore","Microsoft Store"}, {"Microsoft.XboxGamingOverlay","Xbox Game Bar"}, {"Microsoft.GamingApp","Xbox 앱"},
        {"Microsoft.WindowsTerminal","Windows 터미널"}, {"NearbyShare","Quick Share"}, {"Microsoft.MicrosoftOfficeHub","Microsoft 365"},
        {"Microsoft.ScreenSketch","캡처 도구"}, {"Microsoft.WindowsCalculator","계산기"}, {"Microsoft.Paint","그림판"},
        {"Microsoft.ZuneMusic","미디어 플레이어"}, {"Microsoft.ZuneVideo","영화 및 TV"}, {"MSTeams","Microsoft Teams"}, {"Microsoft.Teams","Microsoft Teams"},
        {"Microsoft.Copilot","Copilot"}, {"Microsoft.Windows.Copilot","Copilot"}, {"5319275A.WhatsAppDesktop","WhatsApp"}, {"SpotifyAB.SpotifyMusic","Spotify"},
        {"Claude_","Claude"}, {"Claude","Claude"}, {"Microsoft.OutlookForWindows","Outlook (new)"}, {"Microsoft.Todos","Microsoft To Do"},
        {"Microsoft.WindowsAlarms","알람 및 시계"}, {"Microsoft.BingWeather","날씨"}, {"Microsoft.MicrosoftStickyNotes","스티커 메모"},
        {"Microsoft.WindowsCamera","카메라"}, {"Microsoft.YourPhone","휴대폰과 연결"}, {"MSEdge","Microsoft Edge"}, {"Chrome","Google Chrome"},
        {"308046B0AF4A39CB","Firefox"}, {"Microsoft.Office.WINWORD","Word"}, {"Microsoft.Office.EXCEL","Excel"}, {"Microsoft.Office.POWERPNT","PowerPoint"},
        {"Microsoft.Office.OUTLOOK","Outlook"}, {"Microsoft.Office.ONENOTE","OneNote"}, {"Microsoft.Office.MSACCESS","Access"},
    };
    static readonly string[] SystemPkgPrefix = {
        "Microsoft.Windows.", "MicrosoftWindows.", "Microsoft.LockApp", "Microsoft.AAD.", "Microsoft.Win32WebViewHost", "Microsoft.XboxGameCallableUI",
        "Microsoft.Edge.GameAssist", "MicrosoftCorporationII.", "Microsoft.UI.", "Microsoft.VCLibs", "Microsoft.NET", "Microsoft.Services.",
        "NcsiUwpApp", "Microsoft.ECApp", "Microsoft.CredDialogHost", "Microsoft.AccountsControl", "Microsoft.BioEnrollment", "Microsoft.AutoGenerated.",
        "Microsoft.DesktopAppInstaller", "Microsoft.StorePurchaseApp", "Microsoft.XboxIdentityProvider", "Microsoft.Xbox.TCUI", "Microsoft.WidgetsPlatformRuntime",
        "Microsoft.StartExperiencesApp", "Microsoft.ApplicationCompatibilityEnhancements", "Microsoft.Winget", "windows.", "Windows.", "UEME_",
    };
    static readonly Dictionary<string, string> ExeMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) {
        {"msedge.exe","Microsoft Edge"}, {"chrome.exe","Google Chrome"}, {"firefox.exe","Firefox"}, {"whale.exe","네이버 웨일"}, {"opera.exe","Opera"}, {"brave.exe","Brave"},
        {"code.exe","Visual Studio Code"}, {"devenv.exe","Visual Studio"}, {"kakaotalk.exe","카카오톡"}, {"robloxplayerbeta.exe","Roblox"}, {"robloxplayerlauncher.exe","Roblox"},
        {"valorant-win64-shipping.exe","VALORANT"}, {"valorant.exe","VALORANT"}, {"riot client.exe","Riot Client"}, {"riotclientux.exe","Riot Client"},
        {"leagueclient.exe","리그 오브 레전드"}, {"league of legends.exe","리그 오브 레전드"}, {"steam.exe","Steam"}, {"epicgameslauncher.exe","Epic Games 런처"},
        {"discord.exe","Discord"}, {"notepad.exe","메모장"}, {"mspaint.exe","그림판"}, {"mstsc.exe","원격 데스크톱 연결"}, {"cmd.exe","명령 프롬프트"},
        {"powershell.exe","PowerShell"}, {"pwsh.exe","PowerShell 7"}, {"taskmgr.exe","작업 관리자"}, {"regedit.exe","레지스트리 편집기"},
        {"winword.exe","Word"}, {"excel.exe","Excel"}, {"powerpnt.exe","PowerPoint"}, {"outlook.exe","Outlook"}, {"onenote.exe","OneNote"},
        {"hwp.exe","한글"}, {"minecraft.exe","Minecraft"}, {"minecraftlauncher.exe","Minecraft"}, {"zoom.exe","Zoom"}, {"telegram.exe","Telegram"},
        {"ms-teams.exe","Microsoft Teams"}, {"acrobat.exe","Adobe Acrobat"}, {"acrord32.exe","Adobe Acrobat Reader"}, {"potplayermini64.exe","팟플레이어"},
        {"potplayermini.exe","팟플레이어"}, {"vlc.exe","VLC"}, {"snippingtool.exe","캡처 도구"}, {"calc.exe","계산기"}, {"explorer.exe","파일 탐색기"},
        {"battle.net.exe","Battle.net"}, {"spotify.exe","Spotify"}, {"slack.exe","Slack"}, {"obs64.exe","OBS Studio"},
    };
    static readonly Dictionary<string, string> Alias = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) {
        {"roblox player","Roblox"}, {"file explorer","파일 탐색기"}, {"remote desktop connection","원격 데스크톱 연결"}, {"command prompt","명령 프롬프트"},
        {"control panel","제어판"}, {"windows powershell","PowerShell"}, {"notepad","메모장"}, {"task manager","작업 관리자"}, {"kakaotalk","카카오톡"},
        {"microsoft word","Word"}, {"microsoft excel","Excel"}, {"microsoft powerpoint","PowerPoint"}, {"microsoft outlook","Outlook"},
        {"riot client services","Riot Client"}, {"visual studio code","Visual Studio Code"}, {"google chrome","Google Chrome"}, {"microsoft edge","Microsoft Edge"},
        {"word","Word"}, {"excel","Excel"}, {"powerpoint","PowerPoint"}, {"outlook","Outlook"}, {"settings","설정"}, {"calculator","계산기"},
        {"snipping tool","캡처 도구"}, {"paint","그림판"}, {"microsoft store","Microsoft Store"}, {"windows 탐색기","파일 탐색기"},
    };
    static readonly HashSet<string> WinWhitelist = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        "notepad.exe","mspaint.exe","mstsc.exe","cmd.exe","powershell.exe","pwsh.exe","taskmgr.exe","regedit.exe","snippingtool.exe","calc.exe","write.exe","magnify.exe","osk.exe","charmap.exe","mmc.exe","control.exe"
    };
    static readonly string[] ExeBlack = {
        "update","unins","setup.tmp","helper","crashpad","crashhandler","crashreport","webview","broker","notification","conhost","dllhost","rundll32",
        "svchost","taskhost","sihost","searchhost","shellhost","openwith","pickerhost","dataexchangehost","consent.exe","werfault","ctfmon","smartscreen",
        "textinputhost","msiexec","reg.exe","git-credential","agent","service","daemon","tray","bootstrap","elevation","watchdog","monitor","overlay"
    };
    static readonly string[] PathBlack = { @"\driverstore\", @"\nvidia corporation\", @"\common files\", @"\windowsapps\", @"\installer\", @"\$windows" };

    static string Folder(string guid, string profile)
    {
        string pf = Environment.GetEnvironmentVariable("ProgramW6432") ?? Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
        string pf86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
        string win = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
        string pd = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
        switch (guid.ToUpperInvariant())
        {
            case "{6D809377-6AF0-444B-8957-A3773F02200E}": return pf;
            case "{7C5A40EF-A0FB-4BFC-874A-C0F2E0B9FA8E}": return pf86;
            case "{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}": return Path.Combine(win, "System32");
            case "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}": return Path.Combine(win, "SysWOW64");
            case "{F38BF404-1D43-42F2-9305-67DE0B28FC23}": return win;
            case "{0139D44E-6AFE-49F2-8690-3DAFCAE6FFB8}": return Path.Combine(pd, @"Microsoft\Windows\Start Menu\Programs");
            case "{62AB5D82-FDC1-4DC3-A9DD-070D1D495D97}": return pd;
            case "{A77F5D77-2E2B-44C3-A6A2-ABA601054A51}": return Path.Combine(profile, @"AppData\Roaming\Microsoft\Windows\Start Menu\Programs");
            case "{9E3995AB-1F9C-4F13-B827-48B24B6C7174}": return Path.Combine(profile, @"AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned");
            case "{5E6C858F-0E22-4760-9AFE-EA3317B67173}": return profile;
            case "{F1B32785-6FBA-4FCF-9D55-7B8E7F157091}": return Path.Combine(profile, @"AppData\Local");
            case "{3EB685DB-65F9-4CF6-A03A-E3EF65729F3D}": return Path.Combine(profile, @"AppData\Roaming");
            case "{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}": return Path.Combine(profile, "Desktop");
            case "{374DE290-123F-4565-9164-39C4925E467B}": return Path.Combine(profile, "Downloads");
        }
        return null;
    }

    static AppRec Describe(string raw, string profile, Dictionary<string, string> dev, string self)
    {
        if (string.IsNullOrEmpty(raw)) return null;
        string s = raw.Trim();
        if (s == "Version" || s == "SequenceNumber") return null;

        // 알려진 폴더 GUID → 실제 경로
        var gm = Regex.Match(s, @"^(\{[0-9A-Fa-f\-]{36}\})(.*)$");
        if (gm.Success) { string f = Folder(gm.Groups[1].Value, profile); if (f != null) s = f + gm.Groups[2].Value; }
        // \Device\HarddiskVolumeN → 드라이브 문자
        if (s.StartsWith(@"\Device\", StringComparison.OrdinalIgnoreCase))
            foreach (var kv in dev)
                if (s.StartsWith(kv.Key + @"\", StringComparison.OrdinalIgnoreCase)) { s = kv.Value + s.Substring(kv.Key.Length); break; }

        string name = null, path = null;
        bool isPath = s.Length > 2 && (s[1] == ':' || s.StartsWith(@"\\") || s.StartsWith(@"\Device\", StringComparison.OrdinalIgnoreCase));
        if (isPath)
        {
            path = s;
            string lower = s.ToLowerInvariant();
            if (lower == self) return null;
            string fn = Path.GetFileName(s);
            string fl = fn.ToLowerInvariant();
            if (fl.EndsWith(".lnk"))
            {
                name = Path.GetFileNameWithoutExtension(fn);
                string nl = name.ToLowerInvariant();
                if (nl.Contains("uninstall") || nl.Contains("제거") || nl.Contains("readme") || nl.Contains("website") || nl.Contains("help")) return null;
            }
            else
            {
                string win = Environment.GetFolderPath(Environment.SpecialFolder.Windows).ToLowerInvariant() + "\\";
                if (lower.StartsWith(win) && !WinWhitelist.Contains(fl)) return null;
                if (PathBlack.Any(p => lower.Contains(p))) return null;
                if (!ExeMap.ContainsKey(fl) && ExeBlack.Any(b => fl.Contains(b))) return null;
                if (!ExeMap.TryGetValue(fl, out name))
                {
                    try
                    {
                        if (File.Exists(s))
                        {
                            var vi = FileVersionInfo.GetVersionInfo(s);
                            string d = (vi.FileDescription ?? "").Trim();
                            if (d.Length == 0 || d.Length > 40) d = (vi.ProductName ?? "").Trim();
                            if (d.Length > 0 && d.Length <= 40) name = d;
                        }
                    }
                    catch { }
                    if (string.IsNullOrEmpty(name)) name = Path.GetFileNameWithoutExtension(fn);
                }
            }
        }
        else
        {
            // AppUserModelID / 패키지 이름
            for (int i = 0; i < AumidMap.GetLength(0); i++)
                if (s.StartsWith(AumidMap[i, 0], StringComparison.OrdinalIgnoreCase)) { name = AumidMap[i, 1]; break; }
            if (name == null)
            {
                if (SystemPkgPrefix.Any(p => s.StartsWith(p, StringComparison.OrdinalIgnoreCase))) return null;
                var om = Regex.Match(s, @"^Microsoft\.Office\.(\w+)\.EXE", RegexOptions.IgnoreCase);
                if (om.Success) name = om.Groups[1].Value;
                else
                {
                    string pkg = s.Split('!')[0];
                    int us = pkg.IndexOf('_'); if (us > 0) pkg = pkg.Substring(0, us);
                    int dot = pkg.LastIndexOf('.');
                    if (dot > 0 && dot < pkg.Length - 1) pkg = pkg.Substring(dot + 1);   // 게시자 접두어 제거
                    if (pkg.Length == 0 || pkg.Length > 40 || pkg.Contains("{")) return null;
                    if (Regex.IsMatch(pkg, "^[0-9A-Fa-f]{6,}$")) return null;            // 의미 없는 ID
                    name = Regex.Replace(pkg, "([a-z])([A-Z])", "$1 $2");
                }
            }
        }
        if (string.IsNullOrEmpty(name)) return null;
        string al; if (Alias.TryGetValue(name, out al)) name = al;
        return new AppRec { Name = name, Path = path, Key = name.ToLowerInvariant(), Cat = Category(name + " " + (path ?? s)) };
    }

    static string Category(string t)
    {
        t = t.ToLowerInvariant();
        Func<string[], bool> any = ks => ks.Any(k => t.Contains(k));
        if (any(new[] { "claude", "chatgpt", "copilot", "gemini", "perplexity" })) return "AI";
        if (any(new[] { "\\games\\", "steam", "riot", "valorant", "roblox", "epic games", "xbox", "battle.net", "minecraft", "nexon", "넥슨", "-shipping.exe", "blizzard", "ubisoft", "ea app", "리그 오브 레전드", "게임", "game" })) return "게임";
        if (any(new[] { "edge", "chrome", "firefox", "whale", "웨일", "opera", "brave", "ego" })) return "브라우저";
        if (any(new[] { "word", "excel", "powerpoint", "powerpnt", "outlook", "onenote", "office", "microsoft 365", "hwp", "한글", "acrobat", "pdf", "to do", "스티커" })) return "오피스";
        if (any(new[] { "kakao", "카카오", "discord", "telegram", "whatsapp", "slack", "teams", "zoom", "line" })) return "메신저";
        if (any(new[] { "visual studio", "terminal", "터미널", "powershell", "명령 프롬프트", "git", "python", "intellij", "pycharm", "android studio", "\\jetbrains\\" })) return "개발";
        if (any(new[] { "spotify", "vlc", "팟플레이어", "potplayer", "melon", "멜론", "netflix", "미디어", "사진", "영화", "obs" })) return "미디어";
        if (any(new[] { "탐색기", "설정", "제어판", "store", "작업 관리자", "레지스트리", "windows 보안", "휴대폰과 연결" })) return "시스템";
        return "도구";
    }
}

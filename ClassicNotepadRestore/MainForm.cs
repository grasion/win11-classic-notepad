using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;

namespace ClassicNotepadRestore;

public class MainForm : Form
{
    // ── 색상 (Catppuccin Mocha) ──
    static readonly Color ColBg       = Color.FromArgb(30, 30, 46);
    static readonly Color ColBgDark   = Color.FromArgb(17, 17, 27);
    static readonly Color ColSurface  = Color.FromArgb(49, 50, 68);
    static readonly Color ColSurfaceH = Color.FromArgb(69, 71, 90);
    static readonly Color ColText     = Color.FromArgb(205, 214, 244);
    static readonly Color ColSubtext  = Color.FromArgb(166, 173, 200);
    static readonly Color ColOverlay  = Color.FromArgb(108, 112, 134);
    static readonly Color ColBlue     = Color.FromArgb(137, 180, 250);
    static readonly Color ColGreen    = Color.FromArgb(166, 227, 161);
    static readonly Color ColYellow   = Color.FromArgb(249, 226, 175);
    static readonly Color ColRed      = Color.FromArgb(243, 139, 168);
    static readonly Color ColPeach    = Color.FromArgb(250, 179, 135);

    private RichTextBox _log = null!;
    private Label _statusLabel = null!;
    private Label _lblTitle = null!;
    private Label _lblDesc = null!;
    private Label _lblCurrentState = null!;
    private Button _btnToggle = null!;

    // 현재 상태: true = 신버전 메모장 사용 중, false = 구버전 사용 중
    private bool _isNewNotepad = true;

    public MainForm()
    {
        InitializeUI();
        DetectCurrentState();
    }

    private void InitializeUI()
    {
        Text = "Win11 메모장 버전 전환 도구";
        Size = new Size(640, 650);
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedSingle;
        MaximizeBox = false;
        BackColor = ColBg;
        Font = new Font("Segoe UI", 10f);

        // ── 폼 아이콘 설정 (임베디드 리소스에서 로드) ──
        try
        {
            var asm = System.Reflection.Assembly.GetExecutingAssembly();
            var iconResName = Array.Find(asm.GetManifestResourceNames(),
                n => n.EndsWith("app.ico", StringComparison.OrdinalIgnoreCase));
            if (iconResName != null)
            {
                using var stream = asm.GetManifestResourceStream(iconResName);
                if (stream != null)
                    Icon = new Icon(stream);
            }
        }
        catch { }

        // ── 제목 ──
        _lblTitle = new Label
        {
            Text = "\U0001f4dd  메모장 버전 전환 도구",
            Font = new Font("Segoe UI", 18f, FontStyle.Bold),
            ForeColor = ColText,
            AutoSize = true,
            Location = new Point(28, 20),
            BackColor = Color.Transparent
        };
        Controls.Add(_lblTitle);

        _lblDesc = new Label
        {
            Text = "윈도우 11 신버전 / 구버전(클래식) 메모장을 자유롭게 전환합니다.",
            Font = new Font("Segoe UI", 9f),
            ForeColor = ColSubtext,
            AutoSize = true,
            Location = new Point(30, 56),
            BackColor = Color.Transparent
        };
        Controls.Add(_lblDesc);

        // ── 현재 상태 패널 ──
        var pnlState = new Panel
        {
            Size = new Size(576, 60),
            Location = new Point(28, 86),
            BackColor = ColBgDark
        };
        _lblCurrentState = new Label
        {
            Text = "   검사 중...",
            Font = new Font("Segoe UI", 13f, FontStyle.Bold),
            ForeColor = ColSubtext,
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft
        };
        pnlState.Controls.Add(_lblCurrentState);
        Controls.Add(pnlState);

        // ── 메인 토글 버튼 ──
        _btnToggle = new Button
        {
            Text = "...",
            Size = new Size(576, 54),
            Location = new Point(28, 158),
            FlatStyle = FlatStyle.Flat,
            BackColor = ColBlue,
            ForeColor = ColBgDark,
            Font = new Font("Segoe UI", 14f, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        _btnToggle.FlatAppearance.BorderSize = 0;
        _btnToggle.FlatAppearance.MouseOverBackColor = Color.FromArgb(180, 210, 251);
        _btnToggle.Click += async (s, e) => await ToggleAsync();
        Controls.Add(_btnToggle);

        // ── 상태 표시줄 ──
        var pnlStatus = new Panel
        {
            Size = new Size(576, 34),
            Location = new Point(28, 224),
            BackColor = ColBgDark
        };
        _statusLabel = new Label
        {
            Text = "   \u23F3  준비 완료",
            Font = new Font("Segoe UI", 10f),
            ForeColor = ColSubtext,
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft
        };
        pnlStatus.Controls.Add(_statusLabel);
        Controls.Add(pnlStatus);

        // ── 로그 ──
        _log = new RichTextBox
        {
            Size = new Size(576, 220),
            Location = new Point(28, 268),
            BackColor = ColBgDark,
            ForeColor = ColSubtext,
            Font = new Font("Consolas", 9.5f),
            BorderStyle = BorderStyle.None,
            ReadOnly = true,
            ScrollBars = RichTextBoxScrollBars.Vertical
        };
        Controls.Add(_log);

        // ── 하단 안내 ──
        var lblFooter = new Label
        {
            Text = "* 윈도우 대규모 업데이트 후 설정이 초기화될 수 있습니다. 그때 다시 실행하세요.",
            Font = new Font("Segoe UI", 8f),
            ForeColor = ColOverlay,
            AutoSize = true,
            Location = new Point(28, 494),
            BackColor = Color.Transparent
        };
        Controls.Add(lblFooter);

        // ── 하단 버튼: 커피 사주기 + 블로그 ──
        var btnCoffee = new Button
        {
            Text = "\u2615  개발자에게 커피 사주기",
            Size = new Size(278, 40),
            Location = new Point(28, 518),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(249, 226, 175),
            ForeColor = ColBgDark,
            Font = new Font("Segoe UI", 10f, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnCoffee.FlatAppearance.BorderSize = 0;
        btnCoffee.FlatAppearance.MouseOverBackColor = Color.FromArgb(255, 240, 200);
        btnCoffee.Click += (s, e) => ShowQrDialog();
        Controls.Add(btnCoffee);

        var btnBlog = new Button
        {
            Text = "\U0001f310  개발자 블로그 보기",
            Size = new Size(278, 40),
            Location = new Point(326, 518),
            FlatStyle = FlatStyle.Flat,
            BackColor = ColSurface,
            ForeColor = ColText,
            Font = new Font("Segoe UI", 10f, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnBlog.FlatAppearance.BorderSize = 0;
        btnBlog.FlatAppearance.MouseOverBackColor = ColSurfaceH;
        btnBlog.Click += (s, e) =>
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "https://1st-life-2nd.tistory.com/",
                    UseShellExecute = true
                });
            }
            catch { }
        };
        Controls.Add(btnBlog);

        // ── 최하단 저작권 ──
        var lblCopy = new Label
        {
            Text = "* 윈도우 대규모 업데이트 후 설정이 초기화될 수 있습니다. 그때 다시 실행하세요.",
            Font = new Font("Segoe UI", 8f),
            ForeColor = ColOverlay,
            AutoSize = true,
            Location = new Point(28, 566),
            BackColor = Color.Transparent
        };
        Controls.Add(lblCopy);
    }

    // ══════════════════════════════════════
    // QR 코드 팝업
    // ══════════════════════════════════════
    private void ShowQrDialog()
    {
        var dlg = new Form
        {
            Text = "개발자에게 커피 사주기 \u2615",
            Size = new Size(380, 480),
            StartPosition = FormStartPosition.CenterParent,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            MaximizeBox = false,
            MinimizeBox = false,
            BackColor = ColBg,
            ShowInTaskbar = false
        };

        var lblTitle = new Label
        {
            Text = "\u2615  커피 한 잔의 응원!",
            Font = new Font("Segoe UI", 16f, FontStyle.Bold),
            ForeColor = ColYellow,
            AutoSize = true,
            Location = new Point(70, 16),
            BackColor = Color.Transparent
        };
        dlg.Controls.Add(lblTitle);

        var lblMsg = new Label
        {
            Text = "아래 QR코드를 스캔하여\n개발자에게 커피를 선물해주세요!",
            Font = new Font("Segoe UI", 10f),
            ForeColor = ColSubtext,
            Size = new Size(340, 44),
            Location = new Point(20, 52),
            TextAlign = ContentAlignment.MiddleCenter,
            BackColor = Color.Transparent
        };
        dlg.Controls.Add(lblMsg);

        // QR 이미지 로드 시도: exe 옆 qr.png → 임베디드 리소스 → 흰색 플레이스홀더
        Image? qrImage = null;

        // 1) exe와 같은 폴더의 qr.png
        var exeDir = AppContext.BaseDirectory;
        var qrPath = Path.Combine(exeDir, "qr.png");
        if (File.Exists(qrPath))
        {
            try { qrImage = Image.FromFile(qrPath); } catch { }
        }

        // 2) 임베디드 리소스
        if (qrImage == null)
        {
            var asm = System.Reflection.Assembly.GetExecutingAssembly();
            var resName = Array.Find(asm.GetManifestResourceNames(),
                n => n.EndsWith("qr.png", StringComparison.OrdinalIgnoreCase));
            if (resName != null)
            {
                using var stream = asm.GetManifestResourceStream(resName);
                if (stream != null)
                    qrImage = Image.FromStream(stream);
            }
        }

        if (qrImage != null)
        {
            var picBox = new PictureBox
            {
                Image = qrImage,
                SizeMode = PictureBoxSizeMode.Zoom,
                Size = new Size(280, 280),
                Location = new Point(35, 104),
                BackColor = Color.White
            };
            dlg.Controls.Add(picBox);
        }
        else
        {
            // QR 이미지가 없을 때 안내
            var lblNoQr = new Label
            {
                Text = "QR 이미지를 찾을 수 없습니다.\n\nexe와 같은 폴더에\n'qr.png' 파일을 넣어주세요.",
                Font = new Font("Segoe UI", 11f),
                ForeColor = ColOverlay,
                Size = new Size(280, 280),
                Location = new Point(35, 104),
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = ColBgDark
            };
            dlg.Controls.Add(lblNoQr);
        }

        var lblThank = new Label
        {
            Text = "감사합니다! \u2764",
            Font = new Font("Segoe UI", 12f, FontStyle.Bold),
            ForeColor = ColPeach,
            AutoSize = true,
            Location = new Point(120, 394),
            BackColor = Color.Transparent
        };
        dlg.Controls.Add(lblThank);

        dlg.ShowDialog(this);

        // 다이얼로그 닫힐 때 이미지 리소스 해제
        qrImage?.Dispose();
    }

    // ══════════════════════════════════════
    // 상태 감지
    // ══════════════════════════════════════
    private void DetectCurrentState()
    {
        Log("=== 현재 메모장 상태 검사 ===", ColBlue);

        var (newInstalled, newVersion) = CheckNewNotepad();
        var noOpenWith = CheckNoOpenWith();

        try
        {
            using var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion");
            var build = key?.GetValue("CurrentBuild")?.ToString() ?? "?";
            Log($"OS 빌드: {build}", ColOverlay);
        }
        catch { }

        if (newInstalled)
        {
            _isNewNotepad = true;
            Log($"신버전 메모장: 설치됨 v{newVersion}", ColYellow);
            if (noOpenWith) Log("NoOpenWith: 존재함", ColYellow);
            _lblCurrentState.Text = "   \U0001f7e1  현재: 신버전 메모장 사용 중";
            _lblCurrentState.ForeColor = ColYellow;
            _btnToggle.Text = "\u2B07  구버전(클래식) 메모장으로 전환";
            _btnToggle.BackColor = ColGreen;
            _btnToggle.FlatAppearance.MouseOverBackColor = Color.FromArgb(190, 240, 185);
        }
        else
        {
            _isNewNotepad = false;
            Log("신버전 메모장: 미설치 (구버전 사용 중)", ColGreen);
            if (!noOpenWith) Log("NoOpenWith: 없음 (정상)", ColGreen);
            _lblCurrentState.Text = "   \U0001f7e2  현재: 구버전(클래식) 메모장 사용 중";
            _lblCurrentState.ForeColor = ColGreen;
            _btnToggle.Text = "\u2B06  신버전 메모장으로 복원";
            _btnToggle.BackColor = ColPeach;
            _btnToggle.FlatAppearance.MouseOverBackColor = Color.FromArgb(255, 200, 165);
        }

        Log("", ColOverlay);
    }

    private void RefreshStateUI()
    {
        var (newInstalled, _) = CheckNewNotepad();
        _isNewNotepad = newInstalled;

        if (_isNewNotepad)
        {
            _lblCurrentState.Text = "   \U0001f7e1  현재: 신버전 메모장 사용 중";
            _lblCurrentState.ForeColor = ColYellow;
            _btnToggle.Text = "\u2B07  구버전(클래식) 메모장으로 전환";
            _btnToggle.BackColor = ColGreen;
            _btnToggle.FlatAppearance.MouseOverBackColor = Color.FromArgb(190, 240, 185);
        }
        else
        {
            _lblCurrentState.Text = "   \U0001f7e2  현재: 구버전(클래식) 메모장 사용 중";
            _lblCurrentState.ForeColor = ColGreen;
            _btnToggle.Text = "\u2B06  신버전 메모장으로 복원";
            _btnToggle.BackColor = ColPeach;
            _btnToggle.FlatAppearance.MouseOverBackColor = Color.FromArgb(255, 200, 165);
        }
    }

    // ══════════════════════════════════════
    // 토글 실행
    // ══════════════════════════════════════
    private async Task ToggleAsync()
    {
        _btnToggle.Enabled = false;

        if (_isNewNotepad)
        {
            await SwitchToClassicAsync();
        }
        else
        {
            await SwitchToNewAsync();
        }

        // UI 갱신
        if (InvokeRequired)
            Invoke(RefreshStateUI);
        else
            RefreshStateUI();

        _btnToggle.Enabled = true;
    }

    // ══════════════════════════════════════
    // 구버전으로 전환
    // ══════════════════════════════════════
    private async Task SwitchToClassicAsync()
    {
        SetStatus("\U0001f504  구버전 메모장으로 전환 중...", ColBlue);
        Log("=== 구버전(클래식) 메모장으로 전환 시작 ===", ColBlue);
        Log("  [주의] 작업표시줄 고정을 위해 Explorer가 재시작됩니다.", ColYellow);

        bool ok = true;
        await Task.Run(() =>
        {
            // 1. NoOpenWith 삭제
            if (!Step_RemoveNoOpenWith()) ok = false;
            // 2. 신버전 제거
            if (!Step_RemoveNewNotepad()) ok = false;
            // 3. 파일 연결
            if (!Step_SetFileAssocClassic()) ok = false;
            // 4. 작업표시줄 고정
            Step_PinToTaskbar();
        });

        if (ok)
        {
            SetStatus("\u2705  구버전 메모장으로 전환 완료!", ColGreen);
            Log("=== 전환 완료 ===", ColGreen);
        }
        else
        {
            SetStatus("\u26A0  일부 작업에서 문제 발생. 로그를 확인하세요.", ColYellow);
            Log("=== 일부 작업 실패 ===", ColYellow);
        }
    }

    // ══════════════════════════════════════
    // 신버전으로 복원
    // ══════════════════════════════════════
    private async Task SwitchToNewAsync()
    {
        SetStatus("\U0001f504  신버전 메모장으로 복원 중...", ColBlue);
        Log("=== 신버전 메모장으로 복원 시작 ===", ColBlue);

        bool ok = true;
        await Task.Run(() =>
        {
            // 1. NoOpenWith 복원
            if (!Step_RestoreNoOpenWith()) ok = false;
            // 2. 신버전 재설치
            if (!Step_InstallNewNotepad()) ok = false;
        });

        if (ok)
        {
            SetStatus("\u2705  신버전 메모장으로 복원 완료!", ColGreen);
            Log("=== 복원 완료 ===", ColGreen);
        }
        else
        {
            SetStatus("\u26A0  일부 작업에서 문제 발생. 로그를 확인하세요.", ColYellow);
            Log("=== 일부 작업 실패 ===", ColYellow);
        }
    }

    // ══════════════════════════════════════
    // 개별 단계들
    // ══════════════════════════════════════

    private bool Step_RemoveNoOpenWith()
    {
        Log("-- [1] 레지스트리 NoOpenWith 삭제 --", ColYellow);
        bool success = true;

        var regPaths = new (RegistryKey root, string path, string label)[]
        {
            (Registry.ClassesRoot, @"Applications\notepad.exe", "HKCR"),
            (Registry.LocalMachine, @"SOFTWARE\Classes\Applications\notepad.exe", "HKLM")
        };

        foreach (var (root, path, label) in regPaths)
        {
            try
            {
                using var key = root.OpenSubKey(path, writable: true);
                if (key != null && key.GetValue("NoOpenWith") != null)
                {
                    key.DeleteValue("NoOpenWith");
                    Log($"  [{label}] NoOpenWith 삭제 완료", ColGreen);
                }
                else
                {
                    Log($"  [{label}] NoOpenWith 이미 없음 (정상)", ColGreen);
                }
            }
            catch (Exception ex)
            {
                Log($"  [{label}] 삭제 실패: {ex.Message}", ColRed);
                success = false;
            }
        }
        return success;
    }

    private bool Step_RestoreNoOpenWith()
    {
        Log("-- [1] 레지스트리 NoOpenWith 복원 --", ColYellow);
        try
        {
            using var key = Registry.ClassesRoot.OpenSubKey(@"Applications\notepad.exe", writable: true)
                         ?? Registry.ClassesRoot.CreateSubKey(@"Applications\notepad.exe");
            key.SetValue("NoOpenWith", "", RegistryValueKind.String);
            Log("  [HKCR] NoOpenWith 복원 완료", ColGreen);
            return true;
        }
        catch (Exception ex)
        {
            Log($"  NoOpenWith 복원 실패: {ex.Message}", ColRed);
            return false;
        }
    }

    private bool Step_RemoveNewNotepad()
    {
        Log("-- [2] 신버전 메모장 앱 제거 --", ColYellow);

        var (installed, version) = CheckNewNotepad();
        if (!installed)
        {
            Log("  신버전 메모장이 이미 제거되어 있음", ColGreen);
            return true;
        }

        Log($"  발견: Microsoft.WindowsNotepad v{version}", ColOverlay);

        try
        {
            RunPowerShell("Get-AppxPackage -Name Microsoft.WindowsNotepad | Remove-AppxPackage -ErrorAction Stop");
            Log("  신버전 메모장 제거 완료", ColGreen);
        }
        catch (Exception ex)
        {
            Log($"  제거 실패: {ex.Message}", ColRed);
            return false;
        }

        try
        {
            RunPowerShell("Get-AppxProvisionedPackage -Online | Where-Object {$_.DisplayName -eq 'Microsoft.WindowsNotepad'} | Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue");
            Log("  프로비전 패키지 제거 시도 완료", ColGreen);
        }
        catch { Log("  프로비전 패키지 제거 실패 (일반적으로 문제없음)", ColOverlay); }

        return true;
    }

    private bool Step_InstallNewNotepad()
    {
        Log("-- [2] 신버전 메모장 재설치 --", ColYellow);

        var (installed, _) = CheckNewNotepad();
        if (installed)
        {
            Log("  신버전 메모장이 이미 설치되어 있음", ColGreen);
            return true;
        }

        Log("  Microsoft Store에서 신버전 메모장 설치 시도 중...", ColOverlay);

        try
        {
            // winget으로 설치 시도
            var result = RunCmd("winget install --id 9MSMLRH6LZF3 --source msstore --accept-package-agreements --accept-source-agreements");
            Log("  winget 설치 명령 실행 완료", ColOverlay);

            // 설치 확인
            System.Threading.Thread.Sleep(3000);
            var (check, ver) = CheckNewNotepad();
            if (check)
            {
                Log($"  신버전 메모장 설치 완료: v{ver}", ColGreen);
                return true;
            }
            else
            {
                // winget 실패시 Add-AppxPackage 방식 시도
                Log("  winget 설치 확인 안됨, 다른 방법 시도 중...", ColOverlay);
                RunPowerShell("Get-AppxPackage -AllUsers -Name Microsoft.WindowsNotepad | ForEach-Object { Add-AppxPackage -DisableDevelopmentMode -Register \\\"$($_.InstallLocation)\\AppXManifest.xml\\\" -ErrorAction SilentlyContinue }");
                System.Threading.Thread.Sleep(2000);
                var (check2, ver2) = CheckNewNotepad();
                if (check2)
                {
                    Log($"  신버전 메모장 복원 완료: v{ver2}", ColGreen);
                    return true;
                }

                Log("  자동 설치 실패. Microsoft Store에서 수동 설치하세요.", ColYellow);
                Log("  Store 앱 열기: ms-windows-store://pdp/?ProductId=9MSMLRH6LZF3", ColBlue);
                // Store 열기
                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = "ms-windows-store://pdp/?ProductId=9MSMLRH6LZF3",
                        UseShellExecute = true
                    });
                    Log("  Microsoft Store 페이지를 열었습니다.", ColGreen);
                }
                catch { }
                return false;
            }
        }
        catch (Exception ex)
        {
            Log($"  설치 실패: {ex.Message}", ColRed);
            Log("  Microsoft Store에서 '메모장'을 검색하여 수동 설치하세요.", ColYellow);
            return false;
        }
    }

    private bool Step_SetFileAssocClassic()
    {
        Log("-- [3] .txt 파일 연결 설정 --", ColYellow);

        var notepadPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "System32", "notepad.exe");
        if (!File.Exists(notepadPath))
        {
            Log($"  구버전 메모장을 찾을 수 없음: {notepadPath}", ColRed);
            return false;
        }

        Log($"  구버전 메모장 확인: {notepadPath}", ColOverlay);

        try
        {
            RunCmd($"ftype txtfile=\"{notepadPath}\" \"%1\"");
            RunCmd("assoc .txt=txtfile");
            Log("  시스템 수준 .txt 파일 연결 설정 완료", ColGreen);

            string[] extras = [".log", ".ini", ".cfg", ".inf"];
            foreach (var ext in extras)
            {
                try
                {
                    RunCmd($"assoc {ext}=txtfile");
                    Log($"  {ext} 파일도 연결 완료", ColGreen);
                }
                catch { Log($"  {ext} 연결 실패 (무시 가능)", ColOverlay); }
            }

            Log("", ColOverlay);
            Log("  [안내] 사용자 기본 앱은 수동 변경 필요:", ColBlue);
            Log("  .txt 우클릭 > 연결 프로그램 > 다른 앱 선택 > 메모장", ColText);
        }
        catch (Exception ex)
        {
            Log($"  파일 연결 설정 실패: {ex.Message}", ColRed);
            return false;
        }
        return true;
    }

    private void Step_PinToTaskbar()
    {
        Log("-- [4] 구버전 메모장 작업표시줄 고정 --", ColYellow);

        var notepadPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "System32", "notepad.exe");
        if (!File.Exists(notepadPath))
        {
            Log("  메모장 경로를 찾을 수 없음", ColRed);
            return;
        }

        try
        {
            // 1단계: 시작 메뉴 Programs 폴더에 바로가기 생성 (LayoutModification에 필요)
            var startMenuPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                @"Microsoft\Windows\Start Menu\Programs"
            );
            var lnkPath = Path.Combine(startMenuPath, "클래식 메모장.lnk");

            var createLnkCmd = $@"
$ws = New-Object -ComObject WScript.Shell
$sc = $ws.CreateShortcut('{lnkPath.Replace("'", "''")}')
$sc.TargetPath = '{notepadPath}'
$sc.WorkingDirectory = '{Path.GetDirectoryName(notepadPath)}'
$sc.IconLocation = '{notepadPath},0'
$sc.Description = 'Classic Notepad'
$sc.Save()
";
            RunPowerShell(createLnkCmd);

            if (File.Exists(lnkPath))
                Log("  시작 메뉴에 바로가기 생성 완료", ColGreen);
            else
                Log("  시작 메뉴 바로가기 생성 실패", ColRed);

            // 2단계: LayoutModification.xml 생성하여 작업표시줄에 고정
            var xmlPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                @"Microsoft\Windows\Shell\LayoutModification.xml"
            );

            var xmlContent = $@"<?xml version=""1.0"" encoding=""utf-8""?>
<LayoutModificationTemplate
    xmlns:defaultlayout=""http://schemas.microsoft.com/Start/2014/FullDefaultLayout""
    xmlns:start=""http://schemas.microsoft.com/Start/2014/StartLayout""
    xmlns:taskbar=""http://schemas.microsoft.com/Start/2014/TaskbarLayout""
    xmlns=""http://schemas.microsoft.com/Start/2014/LayoutModification""
    Version=""1"">
  <CustomTaskbarLayoutCollection PinListPlacement=""Replace"">
    <defaultlayout:TaskbarLayout>
      <taskbar:TaskbarPinList>
        <taskbar:DesktopApp DesktopApplicationLinkPath=""{lnkPath}""/>
      </taskbar:TaskbarPinList>
    </defaultlayout:TaskbarLayout>
  </CustomTaskbarLayoutCollection>
</LayoutModificationTemplate>";

            File.WriteAllText(xmlPath, xmlContent, System.Text.Encoding.UTF8);
            Log("  LayoutModification.xml 생성 완료", ColGreen);

            // 3단계: 레이아웃 캐시 삭제
            var defaultLayoutsPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                @"Microsoft\Windows\Shell\DefaultLayouts"
            );
            if (Directory.Exists(defaultLayoutsPath))
            {
                try { Directory.Delete(defaultLayoutsPath, true); } catch { }
            }

            // 4단계: Explorer 재시작 (작업표시줄 레이아웃 적용)
            Log("  Explorer 재시작 중 (작업표시줄 적용)...", ColOverlay);

            foreach (var proc in Process.GetProcessesByName("explorer"))
            {
                try { proc.Kill(); } catch { }
            }

            // Explorer가 자동 재시작되지 않을 경우를 대비
            System.Threading.Thread.Sleep(2000);
            var explorerProcs = Process.GetProcessesByName("explorer");
            if (explorerProcs.Length == 0)
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "explorer.exe"),
                    UseShellExecute = true
                });
            }

            // Explorer가 완전히 로드될 때까지 대기
            System.Threading.Thread.Sleep(5000);

            // 5단계: XML 파일 정리 (레이아웃 잠금 방지)
            try
            {
                if (File.Exists(xmlPath))
                    File.Delete(xmlPath);
                Log("  LayoutModification.xml 정리 완료", ColGreen);
            }
            catch { }

            Log("  작업표시줄 고정 완료!", ColGreen);
        }
        catch (Exception ex)
        {
            Log($"  작업표시줄 고정 실패: {ex.Message}", ColRed);
            Log("  [수동 고정 방법]", ColBlue);
            Log("  Win+R > notepad.exe 실행 > 작업표시줄 아이콘 우클릭 > 작업 표시줄에 고정", ColText);
        }
    }

    // ══════════════════════════════════════
    // 유틸리티
    // ══════════════════════════════════════

    private void Log(string message, Color? color = null)
    {
        if (_log.InvokeRequired)
        {
            _log.Invoke(() => Log(message, color));
            return;
        }
        var c = color ?? ColSubtext;
        var ts = DateTime.Now.ToString("HH:mm:ss");
        _log.SelectionStart = _log.TextLength;
        _log.SelectionColor = ColOverlay;
        _log.AppendText($"[{ts}] ");
        _log.SelectionStart = _log.TextLength;
        _log.SelectionColor = c;
        _log.AppendText(message + "\r\n");
        _log.ScrollToCaret();
    }

    private void SetStatus(string text, Color? color = null)
    {
        if (_statusLabel.InvokeRequired)
        {
            _statusLabel.Invoke(() => SetStatus(text, color));
            return;
        }
        _statusLabel.Text = "   " + text;
        _statusLabel.ForeColor = color ?? ColSubtext;
    }

    private string RunPowerShell(string command)
    {
        var psi = new ProcessStartInfo
        {
            FileName = "powershell.exe",
            Arguments = $"-NoProfile -ExecutionPolicy Bypass -Command \"{command}\"",
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        using var proc = Process.Start(psi)!;
        var output = proc.StandardOutput.ReadToEnd();
        var error = proc.StandardError.ReadToEnd();
        proc.WaitForExit();
        return string.IsNullOrEmpty(output) ? error : output;
    }

    private string RunCmd(string command)
    {
        var psi = new ProcessStartInfo
        {
            FileName = "cmd.exe",
            Arguments = $"/c {command}",
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        using var proc = Process.Start(psi)!;
        var output = proc.StandardOutput.ReadToEnd();
        proc.WaitForExit();
        return output;
    }

    private bool CheckNoOpenWith()
    {
        try
        {
            using var key = Registry.ClassesRoot.OpenSubKey(@"Applications\notepad.exe");
            return key?.GetValue("NoOpenWith") != null;
        }
        catch { return false; }
    }

    private (bool installed, string version) CheckNewNotepad()
    {
        try
        {
            var result = RunPowerShell("Get-AppxPackage -Name Microsoft.WindowsNotepad | Select-Object -ExpandProperty Version");
            var ver = result.Trim();
            if (!string.IsNullOrEmpty(ver) && !ver.Contains("error", StringComparison.OrdinalIgnoreCase)
                && !ver.Contains("cmdlet", StringComparison.OrdinalIgnoreCase))
                return (true, ver);
        }
        catch { }
        return (false, "");
    }
}

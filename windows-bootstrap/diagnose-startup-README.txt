用途：NSIS 原生外壳启动前的离线诊断，不启动 EXE。预置字节数与哈希对应早期故障样本，不是当前发布包的验收标准；需要主动启动并采集引导日志时，应使用配套调试包中的 Start-Local-Debug 入口。

最少步骤：
1. 将 CMD 和 PS1 两个文件一起放到 Windows 上；可与报错 EXE 放同一目录。
2. 普通双击 diagnose-startup.cmd，不用管理员运行。
3. 窗口会显示报告路径；保留该目录中的 startup-report.json 给本次私下排查使用。

如果企业策略禁止脚本，停止运行，把窗口中的完整错误回传。不要调整策略、改用其他执行方式或以管理员身份重试。

提权运行时会跳过所有写探测；路径存在 junction、符号链接、无法确认的祖先或网络路径时也会跳过写探测并报告。报告记录故障 EXE 的字节数和哈希，已知基准是 199039041 字节。路径中的已知用户目录前缀按不区分大小写方式脱敏。

CharFloat Windows startup diagnostic (not an installer)

1. Put diagnose-startup.cmd and diagnose-startup.ps1 together. Optionally place the failing CharFloat-2.2.3-win-x64.exe beside them for a read-only SHA-256 check.
2. Run diagnose-startup.cmd as the normal logged-in user. Do not run as administrator or change TEMP, TMP, ExecutionPolicy, antivirus, or enterprise restrictions.
3. The console prints a report path under the current user's LocalApplicationData/CharFloat Startup Diagnostics. Keep startup-report.json for this private debugging session. It includes local folder paths; do not publish it.

The tool calls GetTempPathW/GetTempFileNameW and records real Win32 error numbers. It writes/deletes only its own newly created probe files within confirmed user directories. It does not probe Windows/Temp, launch the installer or app, install anything, or upload data.

Add-Type compilation uses an explicit CompilerParameters.TempFiles directory. No process/user/machine environment value or security policy is changed. Original TMP/TEMP and the managed API probe are recorded before compilation, so compilation failure does not erase that evidence.

If PowerShell or Add-Type is blocked by existing enterprise policy, stop and provide the displayed message. There are no bypass flags and no downloads.

This is a diagnostic candidate, not a validated fix. The current Mac environment cannot run Windows APIs or verify actual PowerShell/enterprise behavior. The GetTempPathW result and configuredTempProbe/privateUserDirectoryProbe Win32 error fields are the minimum evidence still needed.

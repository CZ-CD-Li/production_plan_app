param(
    [ValidateSet("onedir", "onefile")]
    [string]$Mode = "onedir",
    [switch]$SkipZip
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$appName = "生产计划协同管理平台"
$spec = Join-Path $PSScriptRoot "production_plan_app.spec"

Push-Location $root
try {
    Write-Host "[1/4] 生成程序图标" -ForegroundColor Cyan
    python build_tools\make_icon.py
    if ($LASTEXITCODE -ne 0) { throw "图标生成失败" }

    Write-Host "[2/4] PyInstaller 打包（$Mode）" -ForegroundColor Cyan
    $staleDemo = Join-Path $root "dist\demo_data"
    if (Test-Path $staleDemo) { Remove-Item $staleDemo -Recurse -Force }
    if ($Mode -eq "onedir") {
        python -m PyInstaller --noconfirm --clean $spec
        if ($LASTEXITCODE -ne 0) { throw "打包失败" }
        $dist = Join-Path $root "dist\$appName"
    }
    else {
        $excludes = @(
            "tkinter", "unittest", "pydoc", "doctest", "pandas",
            "matplotlib", "scipy", "setuptools", "pip", "pytest",
            "PyQt5.QtWebEngine", "PyQt5.QtWebEngineCore", "PyQt5.QtWebEngineWidgets",
            "PyQt5.QtWebSockets", "PyQt5.QtWebChannel", "PyQt5.QtQuick",
            "PyQt5.QtQuickWidgets", "PyQt5.QtQml", "PyQt5.QtMultimedia",
            "PyQt5.QtMultimediaWidgets", "PyQt5.QtBluetooth", "PyQt5.QtNfc",
            "PyQt5.QtPositioning", "PyQt5.QtLocation", "PyQt5.QtSerialPort",
            "PyQt5.QtSql", "PyQt5.QtTest", "PyQt5.QtXml", "PyQt5.QtXmlPatterns",
            "PyQt5.QtDesigner", "PyQt5.QtHelp", "PyQt5.QtOpenGL", "PyQt5.QtPrintSupport"
        )
        $args = @(
            "--noconfirm", "--clean", "--onefile", "--windowed",
            "--name", $appName,
            "--icon", (Join-Path $PSScriptRoot "app.ico"),
            "--version-file", (Join-Path $PSScriptRoot "version_info.txt"),
            "--add-data", "demo_data;demo_data",
            "--collect-data", "rapidocr_onnxruntime",
            "--hidden-import", "rapidocr_onnxruntime",
            "--hidden-import", "onnxruntime",
            "--hidden-import", "cv2",
            "--hidden-import", "pymupdf",
            "--optimize", "1"
        )
        foreach ($item in $excludes) { $args += @("--exclude-module", $item) }
        $args += "main.py"
        python -m PyInstaller @args
        if ($LASTEXITCODE -ne 0) { throw "打包失败" }
        $dist = Join-Path $root "dist"
    }

    Write-Host "[3/4] 附带演示数据" -ForegroundColor Cyan
    $target = if ($Mode -eq "onedir") { $dist } else { Join-Path $root "dist" }
    $demoTarget = Join-Path $target "demo_data"
    if (-not (Test-Path $demoTarget) -and (Test-Path (Join-Path $root "demo_data"))) {
        Copy-Item (Join-Path $root "demo_data") $demoTarget -Recurse -Force
    }
    $spcTemplates = @("SPCsheet示例.xlsx", "SPC单值极差图模板.xlsx", "五菱SPC模板.xlsx")
    foreach ($name in $spcTemplates) {
        $spcTemplate = Join-Path $root $name
        if (Test-Path $spcTemplate) {
            Copy-Item $spcTemplate (Join-Path $target $name) -Force
        }
    }

    Write-Host "[4/4] 冒烟启动测试" -ForegroundColor Cyan
    $exe = if ($Mode -eq "onedir") { Join-Path $dist "$appName.exe" } else { Join-Path $dist "$appName.exe" }
    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    $process = Start-Process -FilePath $exe -ArgumentList "--smoke-test" -PassThru -WindowStyle Hidden
    $process.WaitForExit(20000) | Out-Null
    $watch.Stop()
    if (-not $process.HasExited) {
        $process.Kill()
        throw "程序启动超时（20 秒未退出）"
    }
    if ($process.ExitCode -ne 0) { throw "冒烟测试失败，退出码 $($process.ExitCode)" }
    Write-Host ("启动可用，总耗时 {0:N2} 秒（含 2.2 秒自动退出等待）" -f $watch.Elapsed.TotalSeconds) -ForegroundColor Green

    if ($Mode -eq "onedir") {
        # 冒烟测试会在 exe 旁边生成 data 目录，不随安装包发布（首次运行自动创建）
        $runtimeData = Join-Path $dist "data"
        if (Test-Path $runtimeData) {
            Start-Sleep -Milliseconds 600
            Remove-Item $runtimeData -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    if (-not $SkipZip) {
        if ($Mode -eq "onedir") {
            $zip = Join-Path $root "dist\$appName.zip"
            if (Test-Path $zip) { Remove-Item $zip -Force }
            for ($attempt = 1; $attempt -le 5; $attempt++) {
                try {
                    Compress-Archive -Path (Join-Path $dist "*") -DestinationPath $zip -CompressionLevel Optimal
                    break
                }
                catch {
                    if ($attempt -eq 5) { throw }
                    Start-Sleep -Seconds 1
                }
            }
        }
        else {
            $zip = Join-Path $root "dist\${appName}_单文件.zip"
            if (Test-Path $zip) { Remove-Item $zip -Force }
            Compress-Archive -Path (Join-Path $dist "$appName.exe") -DestinationPath $zip
        }
        Write-Host "发布包：$zip" -ForegroundColor Green
    }

    # 删除 build 目录里的中间产物 exe，避免误双击（缺少 _internal 会报 Python DLL 加载失败）
    $intermediateExe = Join-Path $root "build\production_plan_app\$appName.exe"
    if (Test-Path $intermediateExe) { Remove-Item $intermediateExe -Force }
    Write-Host "请运行：$dist\$appName.exe" -ForegroundColor Yellow
}
finally {
    Pop-Location
}

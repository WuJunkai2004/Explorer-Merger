import os
import subprocess
import sys
from pathlib import Path


def find_vcvarsall():
    program_files_x86 = os.environ.get("ProgramFiles(x86)", "C:\\Program Files (x86)")
    vswhere = (
        Path(program_files_x86)
        / "Microsoft Visual Studio"
        / "Installer"
        / "vswhere.exe"
    )

    if not vswhere.exists():
        print("[-] vswhere.exe not found")
        return None

    try:
        output = subprocess.check_output(
            [
                str(vswhere),
                "-latest",
                "-products",
                "*",
                "-requires",
                "Microsoft.VisualStudio.Component.VC.Tools.x86.x64",
                "-property",
                "installationPath",
            ],
            encoding="utf-8",
        ).strip()

        if output:
            vcvarsall = Path(output) / "VC" / "Auxiliary" / "Build" / "vcvarsall.bat"
            if vcvarsall.exists():
                return vcvarsall
    except Exception as e:
        print(f"[-] Error locating MSVC: {e}")

    return None


def main():
    vcvarsall = find_vcvarsall()
    if not vcvarsall:
        print("[-] Could not find MSVC installation (vcvarsall.bat).")
        sys.exit(1)

    print(f"[+] Found MSVC environment: {vcvarsall}")

    # 确保 dist 目录存在
    dist_dir = Path("dist")
    dist_dir.mkdir(exist_ok=True)

    # 构造编译命令
    # 通过 cmd /c 调用 vcvarsall.bat 设置环境，然后紧接着运行 cl.exe
    compile_cmd = f'"{vcvarsall}" x64 && cl /EHsc /W4 /Zi main.cpp /Fe:dist\\Explorer-Merger-Debug.exe'

    print(f"[+] Executing: {compile_cmd}")

    # 使用 shell=True 因为需要 && 逻辑和 bat 执行
    result = subprocess.run(compile_cmd, shell=True)
    sys.exit(result.returncode)


if __name__ == "__main__":
    main()

import subprocess
from version_manager import update_version, update_main_py
import os
import glob


def find_package_dlls(base_path, package_name):
    """
    指定されたパッケージに関連するDLLファイルを再帰的に検索する
    """
    site_packages = os.path.join(base_path, "Lib", "site-packages")
    all_dlls = []

    # パッケージのルートディレクトリを検索
    package_paths = []
    for root, dirs, files in os.walk(site_packages):
        if os.path.basename(root).startswith(package_name):
            package_paths.append(root)

    # 各パスでDLLを検索
    for path in package_paths:
        for root, dirs, files in os.walk(path):
            dlls = [os.path.join(root, f) for f in files if f.endswith('.dll')]
            if dlls:
                print(f"Found DLLs in {root}:")
                for dll in dlls:
                    print(f"  - {dll}")
                all_dlls.extend(dlls)

    return all_dlls


def build_executable():
    new_version = update_version()
    update_main_py(new_version)

    # 環境変数のパスを設定
    venv_path = r"C:\Users\yokam\PycharmProjects\CSV2XL\.venv"

    # DLLファイルを検索
    print("Searching for DLLs...")
    pandas_dlls = find_package_dlls(venv_path, "pandas")
    numpy_dlls = find_package_dlls(venv_path, "numpy")

    if not pandas_dlls and not numpy_dlls:
        print("警告: DLLファイルが見つかりません。別の方法を試みます。")
        # 代替パスを試す
        alt_paths = [
            os.path.join(venv_path, "Lib", "site-packages"),
            os.path.join(venv_path, "libs"),
            os.path.join(venv_path, "DLLs"),
            os.path.join(os.path.dirname(os.__file__), "site-packages")
        ]

        for path in alt_paths:
            print(f"Checking alternative path: {path}")
            if os.path.exists(path):
                print(f"Directory contents of {path}:")
                for item in os.listdir(path):
                    print(f"  - {item}")

    # PyInstallerのオプションを設定
    pyinstaller_options = [
        "pyinstaller",
        "--name=ICLData",
        "--windowed",
        "--icon=assets/CSV2XL.ico",
        "--add-data", "config.ini:.",
    ]

    # DLLファイルを個別に追加
    for dll in pandas_dlls:
        relative_path = os.path.relpath(os.path.dirname(dll),
                                        os.path.join(venv_path, "Lib", "site-packages"))
        pyinstaller_options.extend(["--add-binary", f"{dll};{relative_path}"])

    for dll in numpy_dlls:
        relative_path = os.path.relpath(os.path.dirname(dll),
                                        os.path.join(venv_path, "Lib", "site-packages"))
        pyinstaller_options.extend(["--add-binary", f"{dll};{relative_path}"])

    # 残りのオプションを追加
    pyinstaller_options.extend([
        "--hidden-import", "pandas._libs.aggregations",
        "--hidden-import", "pandas._libs.window.aggregations",
        "--hidden-import", "pandas._libs.window",
        "--hidden-import", "pandas._libs.groupby",
        "--hidden-import", "numpy",
        "--hidden-import", "scipy",
        "--collect-all", "pandas",
        "--collect-all", "numpy",
        "--clean",
        "--log-level", "DEBUG",
        "--noconfirm",
        "main.py"
    ])

    try:
        # PyInstallerを実行
        print("Running PyInstaller with options:", " ".join(pyinstaller_options))
        result = subprocess.run(
            pyinstaller_options,
            check=True,
            capture_output=True,
            text=True
        )

        print(f"Executable built successfully. Version: {new_version}")
        print("Build output:")
        print(result.stdout)

    except subprocess.CalledProcessError as e:
        print(f"Error building executable: {e}")
        print("Build error output:")
        print(e.stderr)
        raise


if __name__ == "__main__":
    build_executable()

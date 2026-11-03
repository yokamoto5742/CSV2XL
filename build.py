import subprocess


def build_executable():

    subprocess.run([
        "pyinstaller",
        "--name=CSV2XL",
        "--windowed",
        "--icon=assets/CSV2XL.ico",
        "--add-data", "utils/config.ini;.",
        "main.py"
    ])

    print(f"Executable built successfully.")


if __name__ == "__main__":
    build_executable()
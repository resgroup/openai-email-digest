import PyInstaller.__main__

PyInstaller.__main__.run(
    [
        "app.py",
        "--onefile",
        "--windowed",
        "--noconfirm",
        "--name=Email Summariser",
        "--icon=icon.ico",
        "--add-data=icon.ico;.",
        "--add-data=icon.png;.",
    ]
)

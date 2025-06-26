# https://pyinstaller.org/en/stable/usage.html
import PyInstaller.__main__

PyInstaller.__main__.run([
    'main.py',
    '--windowed',
    '--noconsole',
    '--icon=test_icon.icns',
    '--add-data=test_header.png:.'
])
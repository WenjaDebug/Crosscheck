import os


def test_windows_scripts_are_there():
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    for name in ("VoiceAgent.bat", "ExcelAddin.bat", os.path.join("scripts", "install-windows.ps1"),
                 os.path.join("scripts", "start-windows.ps1")):
        assert os.path.isfile(os.path.join(root, name)), name
    # Windows PowerShell liest Skripte ohne BOM nicht als UTF-8, dann wären die Umlaute kaputt.
    for name in ("install-windows.ps1", "start-windows.ps1"):
        with open(os.path.join(root, "scripts", name), "rb") as f:
            assert f.read(3) == b"\xef\xbb\xbf", name

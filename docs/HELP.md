# When something does not work

For every known problem, this page says how you recognise it and what to do. Gegenprobe almost always prints a sentence in its window (Terminal on a Mac, the black window on Windows) saying what is missing; those messages are in German, so the German wording is quoted below. If nothing here helps, take a screenshot of that window and ask with it. Once Gegenprobe runs, the start page has an **error report** under **Mehr → Zustand** (More → Status) to send along; it contains nothing from your workbooks.

*Deutsch: [HILFE.md](HILFE.md)*

**Contents:** [Python missing or too old](#python-missing-or-too-old) · [No internet on the first start](#no-internet-on-the-first-start) · [Mac will not open the file](#mac-will-not-open-the-file) · [Windows warns about the file](#windows-warns-about-the-file) · [Windows: "still inside the ZIP file"](#windows-still-inside-the-zip-file) · [Mac asks for your password](#mac-asks-for-your-password) · [Mac: Terminal may not write into Excel's folder](#mac-terminal-may-not-write-into-excels-folder) · [Mac asks about Downloads, Desktop or Documents](#mac-asks-about-downloads-desktop-or-documents) · [Windows security warning about the certificate](#windows-security-warning-about-the-certificate) · [The Checker button is missing in Excel](#the-checker-button-is-missing-in-excel) · [Side panel blank, add-in error or old version](#side-panel-blank-add-in-error-or-old-version) · [Excel 2016 or 2019](#excel-2016-or-2019) · [Port in use](#port-in-use) · [Antivirus or blocked scripts](#antivirus-or-blocked-scripts) · [Folder moved or renamed](#folder-moved-or-renamed) · [New version](#new-version) · [Removing the add-in](#removing-the-add-in) · [Uninstalling everything](#uninstalling-everything)

## Python missing or too old

**How you recognise it:** The window says "Gegenprobe braucht Python (ein kostenloses Programm), und das fehlt auf diesem Mac/PC noch." (Gegenprobe needs Python, a free program, and it is missing), and the browser opens python.org.

Gegenprobe needs Python 3.10 or newer. It uses one that is already there: on a Mac from python.org (`/Library/Frameworks/Python.framework`), from Homebrew or otherwise on the PATH; on Windows from the usual python.org folders, through the `py` launcher or on the PATH. The Python that comes with macOS (`/usr/bin/python3`) is usually 3.9 and too old. Gegenprobe only asks it when the Command Line Tools are installed, so macOS does not pop up its offer to install them. On Windows a `python.exe` in the `WindowsApps` folder does not count: it is only a shortcut to the Microsoft Store.

**What to do:**

- **Mac:** On the page that opened (python.org/downloads/macos), click **macOS 64-bit universal2 installer** next to "Python 3.13…". Double-click the downloaded file and click through with **Continue** and **Install**; the Mac asks for your password. That needs an account with admin rights; on a company or hospital Mac, IT does this. Then double-click **ExcelAddin.command** (or **VoiceAgent.command**) again. If Homebrew is installed, Gegenprobe installs Python 3.13 with it by itself.
- **Windows:** Gegenprobe first tries to install Python 3.13 with winget (just for you, no admin rights). If that does not work (older Windows 10 has no winget), it opens python.org/downloads/windows: click **Windows installer (64-bit)** next to "Python 3.13…" and double-click the file. Without admin rights, untick **Use admin privileges when installing py.exe** at the bottom. Then **Install Now**. Then double-click **ExcelAddin.bat** (or **VoiceAgent.bat**) again.

A specific Python can be set with the environment variable `VOICEAGENT_PYTHON`.

## No internet on the first start

**How you recognise it:** "Ich komme nicht ins Internet. Beim ersten Start lädt Gegenprobe einmal ihre drei Pakete (FastAPI, Uvicorn, pypdf)." (I cannot reach the internet; the first start downloads three packages once.)

**What to do:** Connect the computer to the internet and double-click the file again. Only the installation needs internet (Python and the three packages from pypi.org); after that Gegenprobe works fully offline. If a company or hospital network blocks pypi.org, IT has to allow it once, or do the installation on another network.

## Mac will not open the file

**How you recognise it:** After double-clicking **ExcelAddin.command** or **VoiceAgent.command**, macOS says the file was not opened because Apple could not verify it is free of malware, or that it comes from an unidentified developer. This happens once for every file from the internet that is not registered with Apple.

**What to do:**

- **macOS 15 (Sequoia) and newer:** Click **Done**. Open **System Settings → Privacy & Security**, scroll to the bottom and click **Open Anyway** next to the file, confirm with your password and click **Open** once more.
- **macOS 14 and older:** Right-click (or Control-click) the file, choose **Open**, and click **Open** again in the dialog.

After that, double-clicking works normally. Gegenprobe never switches this check off itself.

**"… could not be executed because you do not have appropriate access privileges":** Unpacking lost the execute permission. The GitHub ZIP carries it (macOS's Archive Utility, i.e. double-clicking the ZIP, keeps it); some other unzip tools drop it. Easiest: unpack the ZIP again by double-clicking it. Otherwise open Terminal (Applications → Utilities → Terminal), type `chmod +x ` (with a space), drag the file from Finder into the Terminal window and press Return.

## Windows warns about the file

**How you recognise it:** After double-clicking **ExcelAddin.bat** or **VoiceAgent.bat** you see "Windows protected your PC" or "Open File – Security Warning: Do you want to run this file?". Windows does this for every unsigned file from the internet.

**What to do:** For "Windows protected your PC", click **More info**, then **Run anyway**. For the security warning, click **Run**. To avoid the warning altogether: before unpacking, right-click the ZIP file → **Properties** → tick **Unblock** at the bottom → **OK**, then unpack.

## Windows: "still inside the ZIP file"

**How you recognise it:** "Gegenprobe steckt noch in der ZIP-Datei. Bitte zuerst entpacken" (Gegenprobe is still inside the ZIP file, please unpack it first). On Windows, double-clicking a ZIP only shows its contents without unpacking.

**What to do:** Right-click the ZIP file → **Extract All…** → choose a place (your user folder is best) → **Extract**. Then double-click inside the unpacked folder.

## Mac asks for your password

**How you recognise it:** ExcelAddin.command says "2/4 Jetzt fragt der Mac nach deinem Passwort, damit er dem Zertifikat vertraut" (now the Mac asks for your password so it trusts the certificate), and a macOS dialog asks for your password.

**What to do:** Enter your Mac password (the one you log in with). Excel only opens add-ins over a secure connection. For that, Gegenprobe creates its own certificate authority for this Mac only (in the `.voiceagent` folder in your home folder; its key never leaves the Mac) and adds it to your keychain. Gegenprobe only issues a certificate for `localhost` with it. This happens once. If you cancelled, double-click ExcelAddin.command again.

## Mac: Terminal may not write into Excel's folder

**How you recognise it:** "macOS lässt das Terminal nicht in den Ordner von Excel schreiben" (macOS does not let Terminal write into Excel's folder), and the **Full Disk Access** setting opens.

**What to do:** As the window says: switch Terminal on there (if it is missing: **+** at the bottom, then Applications → Utilities → Terminal), quit Terminal completely (Cmd-Q) and double-click ExcelAddin.command again. Afterwards you can switch Full Disk Access off again.

## Mac asks about Downloads, Desktop or Documents

**How you recognise it:** "Terminal would like to access files in your Downloads folder", or on the first setup the tip "Dieser Ordner liegt in Downloads, auf dem Schreibtisch, in Dokumente oder in iCloud" (this folder is in Downloads, on the Desktop, in Documents or in iCloud).

**What to do:** Best put the Gegenprobe folder into your home folder before the first start (Finder: **Go → Home**). macOS protects those folders, and the checker in the background can get stuck there. If everything is already set up, see [Folder moved or renamed](#folder-moved-or-renamed). Otherwise answer the question with **OK**.

## Windows security warning about the certificate

**How you recognise it:** ExcelAddin.bat says "2/4 Jetzt fragt Windows einmal, ob es dem Zertifikat „VoiceAgent lokal (nur dieser Rechner)“ vertrauen soll" (Windows now asks once whether to trust the certificate "VoiceAgent lokal (nur dieser Rechner)"), and Windows shows a **Security Warning** with that name.

**What to do:** Click **Yes**. The certificate authority was created on this PC, its key never leaves it, and Gegenprobe only issues a certificate for `localhost` with it. It goes only into your user's certificates; no admin rights needed. If you clicked **No**, double-click ExcelAddin.bat again.

## The Checker button is missing in Excel

**How you recognise it:** ExcelAddin.command or ExcelAddin.bat said "Fertig" (done), but there is no **Prüfer** (in English Excel: **Checker**) button under **Home**.

**What to do:**

1. Quit Excel completely, not just the window: Cmd-Q on a Mac, close all Excel windows on Windows. Reopen Excel and open a workbook.
2. **Mac:** Under **Home**, click **Add-ins** at the far right and choose **Gegenprobe Checker** (German Excel: **Gegenprobe Prüfer**). After that the **Checker** button is under **Home**.
3. **Windows:** Under **Home**, click **Add-ins** (or **Insert → My Add-ins**) and choose **Gegenprobe Checker** (German Excel: **Gegenprobe Prüfer**) under developer add-ins.
4. If that does not help, double-click ExcelAddin again and read what it says.

## Side panel blank, add-in error or old version

**How you recognise it:** The panel on the right stays white, Excel reports "Add-in error", or something that should be there after an update is missing.

**What to do:**

1. **Double-click ExcelAddin again.** It checks whether the checker runs in the background, says why not (not started, crashed, folder renamed, switched off in macOS) and restarts it. On a Mac it opens **System Settings → General → Login Items** if needed: switch on python3.13 (or Python) under **Allow in the Background**. What the checker itself reports is in `data/pruefer-dienst.log` in the Gegenprobe folder.
2. **Clear Excel's cache.** Gegenprobe gives every new version of the panel a new address, which is usually enough. Otherwise:
   - **Mac:** Open Excel's menu at the top right of the panel header (not Gegenprobe's own **i**) and choose **Clear Web Cache**, then close and reopen the panel.
   - **Windows:** Quit Excel completely and reopen it. If that does not help: quit Excel, type `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` into Explorer's address bar, press Enter, delete everything in it and open Excel again.

## Excel 2016 or 2019

**How you recognise it:** On Windows with Excel 2016 or 2019, the side panel stays blank or shows an error.

**What to do:** These versions still show add-ins with the old Internet Explorer, and the panel does not run there. Use the checker in the browser: double-click **VoiceAgent.bat** and choose **Excel-Prüfer im Browser** on the start page. It reads the open workbook directly from Excel. Excel 2021, 2024 and Microsoft 365 can show the panel.

## Port in use

**How you recognise it:**

- "Port 8766 ist belegt. Läuft noch VoiceAgent.command?" (Mac) or "Läuft noch VoiceAgent.bat?" (Windows) while setting up with ExcelAddin. Port 8766 belongs to the side panel.
- "Port 8765 hat schon ein anderes Programm. Gegenprobe nimmt diesmal Port …" (another program has port 8765; Gegenprobe uses port … this time) when VoiceAgent starts. Gegenprobe then takes the next free port by itself, and the browser opens the right address.
- "Gegenprobe läuft schon, in einem anderen Fenster." (already running in another window). A second double-click then only opens the browser.

**What to do:** For port 8766, close the VoiceAgent window (or press Ctrl + C there) and double-click ExcelAddin again. Otherwise nothing: Gegenprobe moves aside by itself.

## Antivirus or blocked scripts

**How you recognise it:** Antivirus reports `ExcelAddin.bat`, `powershell` or `pythonw.exe`, or the window says "… cannot be loaded because running scripts is disabled on this system".

**What to do:** On a company or hospital computer, ask IT. You can tell them what Gegenprobe does: all launch files and the code are readable text files. On Windows, `VoiceAgent.bat` runs `scripts\start-windows.ps1` with `powershell -ExecutionPolicy Bypass` (for that one call only). It installs Python with winget and three packages from pypi.org into the `.venv` folder. ExcelAddin.bat creates a certificate for `localhost` (`%USERPROFILE%\.voiceagent\addin`), adds it with `certutil -user` and registers the add-in under `HKCU\Software\Microsoft\Office\16.0\WEF\Developer`. Under `HKCU\Software\Microsoft\Windows\CurrentVersion\Run` ("Gegenprobe Pruefer") it starts the checker at every login with `pythonw.exe -m voiceagent.dienst`, without a window. It only listens on `127.0.0.1`. On a Mac, ExcelAddin.command sets up the same with a LaunchAgent (`~/Library/LaunchAgents/io.github.wenjadebug.voiceagent.pruefer.plist`) and the keychain. No workbook data goes anywhere.

## Folder moved or renamed

**How you recognise it:** After moving or renaming the folder, VoiceAgent no longer starts (for example "bad interpreter" or "No such file"), or the side panel reports an add-in error.

**What to do:** The private Python environment is tied to the folder's location. Easiest is to move the folder back. Otherwise delete the `.venv` folder inside the Gegenprobe folder (hidden on a Mac: Cmd-Shift-period in Finder shows it) and double-click **ExcelAddin** (or **VoiceAgent**) again. That sets everything up again at the new place.

## New version

**How you recognise it:** When VoiceAgent starts it says "Es gibt eine neuere Version von Gegenprobe (vom …)" (there is a newer version, from …), also on the start page under **Mehr → Zustand**.

If you downloaded Gegenprobe as a ZIP, it never downloads code by itself. At most once a day it asks GitHub when the newest version was made, without an account and without sending anything from the folder or your workbooks. Switch it off with the environment variable `VOICEAGENT_NO_UPDATE=1`. If you got Gegenprobe with `git clone`, new versions arrive by themselves at start.

**What to do:**

1. Rename the old folder, for example to "Gegenprobe old".
2. Download and unpack the new ZIP as the first time, put the new folder exactly where the old one was and give it exactly the old name.
3. Bring your own column rules: copy `data/spaltenregeln.json` from the old folder into the `data` folder of the new one (if it exists).
4. Double-click **ExcelAddin** in the new folder. Then delete the old folder.

## Removing the add-in

This takes Gegenprobe out of Excel and out of the background. It needs one extra word, `--entfernen` (remove):

- **Mac:** Open Terminal (Applications → Utilities → Terminal). Drag **ExcelAddin.command** from Finder into the Terminal window, type a space and `--entfernen`, press Return. Then restart Excel. The certificate stays in the keychain. To take it out: open **Keychain Access**, **login** on the left, **Certificates** at the top, delete "VoiceAgent lokal (nur dieser Rechner)".
- **Windows:** Open the Gegenprobe folder in Explorer, type `cmd` into the address bar, press Enter. In the black window type `ExcelAddin.bat --entfernen`, press Enter. Then quit Excel completely and reopen it. The certificate stays in Windows so that setting up again does not ask once more. To take it out: Windows key, "Manage user certificates", **Trusted Root Certification Authorities → Certificates**, delete "VoiceAgent lokal (nur dieser Rechner)".

## Uninstalling everything

1. First [remove the add-in](#removing-the-add-in) while the folder is still there.
2. Delete the Gegenprobe folder (on a Mac: move it to the Trash). It holds everything, including the private Python environment `.venv` and the `data` folder with your column rules, reports and the log.
3. Delete the `.voiceagent` folder in your home folder (certificates and keys). It is hidden on a Mac: Cmd-Shift-period in Finder shows it. On Windows: type `%USERPROFILE%` into Explorer's address bar.
4. Take the certificate out of the keychain or Windows as described above.
5. Python itself can stay; it does no harm. To remove it: on Windows **Settings → Apps**, uninstall "Python 3.13"; on a Mac it lives under **Applications → Python 3.13** and in `/Library/Frameworks/Python.framework`.

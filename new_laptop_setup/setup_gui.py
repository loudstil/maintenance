import subprocess
import sys
import tkinter as tk
from tkinter import messagebox, filedialog
from cr import ConsoleRedirector

def set_execution_policy():
    subprocess.run(["powershell", "Set-ExecutionPolicy", "RemoteSigned", "-Scope", "Process"])

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def get_current_computer_name():
    try:
        result = subprocess.run(["hostname"], stdout=subprocess.PIPE, text=True)
        return result.stdout.strip()
    except Exception as e:
        print(f"Error getting computer name: {e}")
        return ""

def run_script():
    if not is_admin():
        messagebox.showerror("Admin Privileges Required", "Please run the script as an administrator.")
        return

    set_execution_policy()

    current_computer_name = get_current_computer_name()
    if not current_computer_name:
        messagebox.showerror("Error", "Unable to retrieve current computer name.")
        return

    new_computer_name = computer_name_entry.get()
    if new_computer_name != current_computer_name :
        subprocess.run(["powershell", "Rename-Computer", "-NewName", new_computer_name, "-Force", "-Restart"])
    else :
        essagebox.showinfo("Info", "Computer name was'nt updated.")


    if makom_user_checkbox.get():
        subprocess.run(["powershell", "New-LocalUser", "-Name", "Makom", "-Description", "Standard User", "-Password", "$null"])

    if makom_ligdol_user_checkbox.get():
        subprocess.run(["powershell", "New-LocalUser", "-Name", "Makom Ligdol", "-Password", "4729", "-Description", "Administrator", "-AccountNeverExpires", "-UserMayNotChangePassword"])
        subprocess.run(["powershell", "Add-LocalGroupMember", "-Group", "Administrators", "-Member", "Makom Ligdol"])

    if wifi_connection_checkbox.get():
        subprocess.run(["netsh", "wlan", "connect", "name=MakomLigdol", "key=MakomLigdol2019"])

    if office_installation_checkbox.get():
        office_product_key = office_product_key_entry.get()
        office_config_path = office_config_path_entry.get()
        setup_exe_path = setup_exe_path_entry.get()

        subprocess.run(["powershell", "Start-Process", "-FilePath", setup_exe_path, "-ArgumentList", "/configure", office_config_path, "-ProductKey", office_product_key, "-Wait"])

    if chrome_installation_checkbox.get():
        chrome_installer_url = "https://dl.google.com/chrome/install/latest/chrome_installer.exe"
        subprocess.run(["powershell", "Invoke-WebRequest", "-Uri", chrome_installer_url, "-OutFile", "$env:TEMP\\chrome_installer.exe"])
        subprocess.run(["powershell", "Start-Process", "-FilePath", "$env:TEMP\\chrome_installer.exe", "-Wait"])

    if vlc_installation_checkbox.get():
        vlc_installer_url = "https://get.videolan.org/vlc/3.0.11/win64/vlc-3.0.11-win64.exe"
        subprocess.run(["powershell", "Invoke-WebRequest", "-Uri", vlc_installer_url, "-OutFile", "$env:TEMP\\vlc_installer.exe"])
        subprocess.run(["powershell", "Start-Process", "-FilePath", "$env:TEMP\\vlc_installer.exe", "-Wait"])

    if zoom_installation_checkbox.get():
        zoom_installer_url = "https://zoom.us/client/latest/ZoomInstallerFull.msi"
        subprocess.run(["powershell", "Invoke-WebRequest", "-Uri", zoom_installer_url, "-OutFile", "$env:TEMP\\zoom_installer.msi"])
        subprocess.run(["msiexec.exe", "/i", "$env:TEMP\\zoom_installer.msi", "/qn", "-Wait"])

    if windows_update_checkbox.get():
        subprocess.run(["powershell", "Install-WindowsUpdate", "-AcceptAll", "-AutoReboot"])

    if shortcuts_creation_checkbox.get():
        desktop_path = subprocess.check_output(["[System.IO.Path]::Combine([System.Environment]::GetFolderPath('Desktop'), 'Shortcuts')"], shell=True, text=True).strip()
        subprocess.run(["powershell", "New-Item", "-ItemType", "Directory", "-Path", desktop_path, "-Force"])

        shortcuts = [
            ("$env:ProgramFiles\\Microsoft Office\\root\\Office16\\WINWORD.EXE", "Word.lnk"),
            ("$env:ProgramFiles\\Microsoft Office\\root\\Office16\\EXCEL.EXE", "Excel.lnk"),
            ("$env:ProgramFiles\\Microsoft Office\\root\\Office16\\POWERPNT.EXE", "PowerPoint.lnk"),
            ("$env:ProgramFiles\\Google\\Chrome\\Application\\chrome.exe", "Google Chrome.lnk"),
            ("$env:ProgramFiles\\VideoLAN\\VLC\\vlc.exe", "VLC Media Player.lnk"),
            ("$env:ProgramFiles\\Zoom\\bin\\Zoom.exe", "Zoom.lnk"),
        ]

        for shortcut in shortcuts:
            shell = subprocess.Popen(["powershell", "[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null; $WshShell = New-Object -ComObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut([System.IO.Path]::Combine('" + desktop_path + "', '" + shortcut[1] + "')); $Shortcut.TargetPath = '" + shortcut[0] + "'; $Shortcut.Save()"])

    messagebox.showinfo("Script Execution", "Script execution completed successfully.")

def browse_file(entry):
    filepath = filedialog.askopenfilename()
    entry.delete(0, tk.END)
    entry.insert(tk.END, filepath)

# Create main window
root = tk.Tk()
root.title("Installation Script GUI")

# Create and grid widgets for the first column (input elements)
computer_name_label = tk.Label(root, text="Enter Computer Name:")
computer_name_label.grid(row=0, column=0, sticky="w", padx=10, pady=5)

computer_name_entry = tk.Entry(root)
computer_name_entry.grid(row=0, column=1, padx=10, pady=5)

makom_user_checkbox = tk.Checkbutton(root, text="Create Makom User (No Admin Privileges)")
makom_user_checkbox.grid(row=1, column=0, columnspan=2, sticky="w", padx=10, pady=5)

makom_ligdol_user_checkbox = tk.Checkbutton(root, text="Create Makom Ligdol User (Administrator)")
makom_ligdol_user_checkbox.grid(row=2, column=0, columnspan=2, sticky="w", padx=10, pady=5)

wifi_connection_checkbox = tk.Checkbutton(root, text="Connect to WiFi (MakomLigdol)")
wifi_connection_checkbox.grid(row=3, column=0, columnspan=2, sticky="w", padx=10, pady=5)

office_installation_checkbox = tk.Checkbutton(root, text="Install Microsoft Office")
office_installation_checkbox.grid(row=4, column=0, columnspan=2, sticky="w", padx=10, pady=5)

# Additional input boxes for Office installation
office_product_key_label = tk.Label(root, text="Enter Office Product Key:")
office_product_key_label.grid(row=5, column=0, sticky="w", padx=10, pady=5)

office_product_key_entry = tk.Entry(root)
office_product_key_entry.grid(row=5, column=1, padx=10, pady=5)

office_config_path_label = tk.Label(root, text="Enter Office Configuration XML File Path:")
office_config_path_label.grid(row=6, column=0, sticky="w", padx=10, pady=5)

office_config_path_entry = tk.Entry(root)
office_config_path_entry.grid(row=6, column=1, padx=10, pady=5)

office_config_path_browse_button = tk.Button(root, text="Browse", command=lambda: browse_file(office_config_path_entry))
office_config_path_browse_button.grid(row=6, column=2, padx=10, pady=5)

setup_exe_path_label = tk.Label(root, text="Enter Setup.exe File Path:")
setup_exe_path_label.grid(row=7, column=0, sticky="w", padx=10, pady=5)

setup_exe_path_entry = tk.Entry(root)
setup_exe_path_entry.grid(row=7, column=1, padx=10, pady=5)

setup_exe_path_browse_button = tk.Button(root, text="Browse", command=lambda: browse_file(setup_exe_path_entry))
setup_exe_path_browse_button.grid(row=7, column=2, padx=10, pady=5)

chrome_installation_checkbox = tk.Checkbutton(root, text="Install Google Chrome")
chrome_installation_checkbox.grid(row=8, column=0, columnspan=2, sticky="w", padx=10, pady=5)

vlc_installation_checkbox = tk.Checkbutton(root, text="Install VLC Player")
vlc_installation_checkbox.grid(row=9, column=0, columnspan=2, sticky="w", padx=10, pady=5)

zoom_installation_checkbox = tk.Checkbutton(root, text="Install Zoom")
zoom_installation_checkbox.grid(row=10, column=0, columnspan=2, sticky="w", padx=10, pady=5)

windows_update_checkbox = tk.Checkbutton(root, text="Run Windows Update")
windows_update_checkbox.grid(row=11, column=0, columnspan=2, sticky="w", padx=10, pady=5)

shortcuts_creation_checkbox = tk.Checkbutton(root, text="Create Shortcuts for Installed Apps")
shortcuts_creation_checkbox.grid(row=12, column=0, columnspan=2, sticky="w", padx=10, pady=5)

run_button = tk.Button(root, text="Run Script", command=run_script)
run_button.grid(row=13, column=0, columnspan=2, pady=10)

# Create a text widget for the console (output)
console_text = tk.Text(root, height=15, width=50, wrap="word", state=tk.DISABLED)
console_text.grid(row=0, column=2, rowspan=14, columnspan=2, padx=10, pady=5, sticky="nsew")

# Create a scrollbar for the console_text
scrollbar = tk.Scrollbar(root, command=console_text.yview)
scrollbar.grid(row=0, column=4, rowspan=14, sticky="nsew")

# Configure the console_text to use the scrollbar
console_text.config(yscrollcommand=scrollbar.set)

# Redirect sys.stdout to the console_text widget
sys.stdout = ConsoleRedirector(console_text)

# Start the Tkinter event loop
root.mainloop()

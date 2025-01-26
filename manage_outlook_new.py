import os
import sys
import subprocess
import platform
import winreg
import shutil
import ctypes

def is_admin():
    """ Verifica si el script se está ejecutando con permisos de administrador. """
    try:
        return os.getuid() == 0
    except AttributeError:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0

def check_windows_version():
    """ Verifica si el sistema operativo es Windows XP o Windows 7. """
    version = platform.version()
    if "XP" in version or "Windows 7" in version:
        print("[-] Este script no puede ejecutarse en Windows XP o Windows 7.")
        sys.exit(1)

def disable_outlook_new():
    """ Deshabilita Outlook New mediante claves de registro. """
    try:
        arch = platform.architecture()[0]
        reg_path = r"SOFTWARE\Microsoft\Office\Outlook"
        if arch == "64bit":
            reg_path = r"SOFTWARE\WOW6432Node\Microsoft\Office\Outlook"

        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path, 0, winreg.KEY_WRITE) as key:
            winreg.SetValueEx(key, "HideNewOutlookToggle", 0, winreg.REG_DWORD, 1)
        print("[+] Outlook New ha sido deshabilitado correctamente.")
    except Exception as e:
        print(f"[-] Error al deshabilitar Outlook New: {e}")

def uninstall_outlook_new():
    """ Desinstala Outlook New si está instalado. """
    try:
        if platform.system() == "Windows":
            subprocess.run(["powershell", "-Command", "Get-AppxPackage *OutlookForWindows* | Remove-AppxPackage"], check=True)
        print("[+] Outlook New ha sido desinstalado correctamente.")
    except subprocess.CalledProcessError as e:
        print(f"[-] Error al desinstalar Outlook New: {e}")

def block_future_installations():
    """ Bloquea la reinstalación automática de Outlook New mediante claves de registro. """
    try:
        arch = platform.architecture()[0]
        reg_path = r"SOFTWARE\Policies\Microsoft\Office\16.0\Outlook\Options\General"
        if arch == "64bit":
            reg_path = r"SOFTWARE\WOW6432Node\Policies\Microsoft\Office\16.0\Outlook\Options\General"

        with winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, reg_path) as key:
            winreg.SetValueEx(key, "DisableNewOutlookToggle", 0, winreg.REG_DWORD, 1)
        print("[+] Se ha bloqueado la reinstalación automática de Outlook New.")
    except Exception as e:
        print(f"[-] Error al bloquear la reinstalación: {e}")

def disable_mail_app():
    """ Deshabilita la aplicación 'mail' de Windows utilizando políticas de grupo locales. """
    try:
        subprocess.run(["powershell", "-Command", "New-ItemProperty -Path 'HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\Appx' -Name 'DisableMailApp' -Value 1 -PropertyType DWORD -Force"], check=True)
        print("[+] Aplicación 'mail' deshabilitada utilizando políticas de grupo locales.")
    except subprocess.CalledProcessError as e:
        print(f"[-] Error al deshabilitar la aplicación 'mail': {e}")

def disable_store_reinstallation():
    """ Bloquea la reinstalación desde Microsoft Store. """
    try:
        subprocess.run(["powershell", "-Command", "reg add HKLM\\SOFTWARE\\Policies\\Microsoft\\WindowsStore /v DisableStoreApps /t REG_DWORD /d 1 /f"], check=True)
        subprocess.run(["powershell", "-Command", "reg add HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\Appx /v AllowAllTrustedApps /t REG_DWORD /d 0 /f"], check=True)
        print("[+] Bloqueada la reinstalación desde Microsoft Store.")
    except subprocess.CalledProcessError as e:
        print(f"[-] Error al bloquear la reinstalación desde Microsoft Store: {e}")

def disable_microsoft_store():
    """ Deshabilita completamente la Microsoft Store. """
    try:
        # Deshabilitar la ejecución de la Microsoft Store
        subprocess.run(["powershell", "-Command", "reg add HKLM\\SOFTWARE\\Policies\\Microsoft\\WindowsStore /v RemoveWindowsStore /t REG_DWORD /d 1 /f"], check=True)
        print("[+] Microsoft Store ha sido deshabilitada correctamente.")
    except subprocess.CalledProcessError as e:
        print(f"[-] Error al deshabilitar Microsoft Store: {e}")

def uninstall_microsoft_store():
    """ Desinstala la aplicación de la Microsoft Store si está instalada. """
    try:
        subprocess.run(["powershell", "-Command", "Get-AppxPackage *windowsstore* | Remove-AppxPackage"], check=True)
        print("[+] Aplicación de la Microsoft Store desinstalada correctamente.")
    except subprocess.CalledProcessError as e:
        print(f"[-] Error al desinstalar la aplicación de la Microsoft Store: {e}")

def uninstall_mail_app():
    """ Desinstala la aplicación de correo si está instalada. """
    try:
        subprocess.run(["powershell", "-Command", "Get-AppxPackage *windowscommunicationsapps* | Remove-AppxPackage"], check=True)
        print("[+] Aplicación de correo desinstalada correctamente.")
    except subprocess.CalledProcessError as e:
        print(f"[-] Error al desinstalar la aplicación de correo: {e}")

def schedule_periodic_removal():
    """ Programa una tarea para eliminar Outlook New periódicamente. """
    try:
        task_cmd = f'powershell.exe -Command "Get-AppxPackage *OutlookForWindows* | Remove-AppxPackage"'
        subprocess.run(["schtasks", "/Create", "/SC", "ONLOGON", "/TN", "RemoveOutlookNew", "/TR", task_cmd, "/F"], check=True)
        print("[+] Se ha programado la eliminación periódica de Outlook New.")
    except subprocess.CalledProcessError as e:
        print(f"[-] Error al programar la eliminación periódica: {e}")

def update_group_policies():
    """ Actualiza las políticas de grupo locales para aplicar los cambios inmediatamente. """
    try:
        subprocess.run(["gpupdate", "/force"], check=True)
        print("[+] Políticas de grupo actualizadas correctamente.")
    except subprocess.CalledProcessError as e:
        print(f"[-] Error al actualizar las políticas de grupo: {e}")

def revert_changes():
    """ Revierte los cambios realizados restaurando configuraciones originales. """
    try:
        # Eliminar claves de registro
        arch = platform.architecture()[0]
        reg_paths = [
            r"SOFTWARE\Microsoft\Office\Outlook",
            r"SOFTWARE\Policies\Microsoft\Office\16.0\Outlook\Options\General",
            r"SOFTWARE\Policies\Microsoft\Windows\Appx",
            r"SOFTWARE\Policies\Microsoft\WindowsStore"
        ]
        if arch == "64bit":
            reg_paths = [
                r"SOFTWARE\WOW6432Node\Microsoft\Office\Outlook",
                r"SOFTWARE\WOW6432Node\Policies\Microsoft\Office\16.0\Outlook\Options\General",
                r"SOFTWARE\Policies\Microsoft\Windows\Appx",
                r"SOFTWARE\Policies\Microsoft\WindowsStore"
            ]

        for reg_path in reg_paths:
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path, 0, winreg.KEY_READ) as key:
                    winreg.DeleteKey(key, "")
                print(f"[+] Clave de registro {reg_path} eliminada correctamente.")
            except FileNotFoundError:
                print(f"[-] Clave de registro {reg_path} no encontrada.")
            except PermissionError:
                print(f"[-] Permiso denegado al eliminar la clave de registro {reg_path}.")
            except Exception as e:
                print(f"[-] Error al eliminar la clave de registro {reg_path}: {e}")

        # Eliminar tarea programada
        subprocess.run(["schtasks", "/Delete", "/TN", "RemoveOutlookNew", "/F"], check=True)

        # Reinstalar la Microsoft Store
        subprocess.run([
            "powershell", "-Command",
            "Get-AppxPackage -allusers Microsoft.WindowsStore | Foreach {Add-AppxPackage -DisableDevelopmentMode -Register \"$($_.InstallLocation)\\AppXManifest.xml\"}"
        ], check=True)
        print("[+] Microsoft Store ha sido reinstalada correctamente.")

        # Reinstalar la aplicación de correo
        subprocess.run([
            "powershell", "-Command",
            "Get-AppxPackage -allusers Microsoft.WindowsCommunicationsApps | Foreach {Add-AppxPackage -DisableDevelopmentMode -Register \"$($_.InstallLocation)\\AppXManifest.xml\"}"
        ], check=True)
        print("[+] Aplicación de correo ha sido reinstalada correctamente.")

        print("[+] Todos los cambios han sido revertidos correctamente.")
    except Exception as e:
        print(f"[-] Error al revertir los cambios: {e}")

def show_help():
    """ Muestra las instrucciones de uso. """
    help_message = """
    Uso: manage_outlook_new.py [opción]

    Opciones:
      1   Deshabilita Outlook New.
      2   Desinstala Outlook New.
      3   Bloquea futuras instalaciones de Outlook New.
      4   Revierte todos los cambios realizados.
      6   Bloquea la reinstalación desde Microsoft Store.
      7   Programa una tarea para eliminar Outlook New periódicamente.
      8   Deshabilita la aplicación de correo integrada en Windows.
      9   Actualiza las políticas de grupo locales.
      10  Desinstala la aplicación de correo integrada en Windows.
      11  Deshabilita completamente la Microsoft Store.
      12  Desinstala la aplicación de la Microsoft Store.
      /? Muestra este mensaje de ayuda.
    """
    print(help_message)

def main():
    if not is_admin():
        print("[-] Este script debe ejecutarse con permisos de administrador.")
        sys.exit(1)

    check_windows_version()

    if len(sys.argv) == 1:
        print("[*] Ejecutando acciones predeterminadas (1, 2, 3, 8, 6, 10, 11 y 12)...")
        disable_outlook_new()
        uninstall_outlook_new()
        block_future_installations()
        disable_mail_app()
        disable_store_reinstallation()
        disable_microsoft_store()
        uninstall_mail_app()
        uninstall_microsoft_store()
        schedule_periodic_removal()
        update_group_policies()
    else:
        option = sys.argv[1]
        if option == "/?":
            show_help()
        elif option == "1":
            disable_outlook_new()
        elif option == "2":
            uninstall_outlook_new()
        elif option == "3":
            block_future_installations()
        elif option == "4":
            revert_changes()
        elif option == "8":
            disable_mail_app()
        elif option == "6":
            disable_store_reinstallation()
        elif option == "7":
            schedule_periodic_removal()
        elif option == "9":
            update_group_policies()
        elif option == "10":
            uninstall_mail_app()
        elif option == "11":
            disable_microsoft_store()
        elif option == "12":
            uninstall_microsoft_store()
        else:
            print("[-] Opción no válida. Usa /? para ver las instrucciones de uso.")

if __name__ == "__main__":
    main()

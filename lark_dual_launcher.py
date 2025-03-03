import os
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox
import json
import sys

# 配置文件路径
CONFIG_FILE = "lark_config.json"

def user_exists(username):
    """检查用户是否已存在"""
    try:
        subprocess.run(['net', 'user', username], shell=True, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        return True
    except subprocess.CalledProcessError:
        return False

def create_user(username, password):
    """创建新用户"""
    try:
        subprocess.run(['net', 'user', username, password, '/add'], shell=True, check=True)
        return True
    except subprocess.CalledProcessError as e:
        messagebox.showerror("错误", f"无法创建用户: {e}")
        return False

def run_as_user(username, password, command):
    """以新用户身份运行命令"""
    try:
        # 使用 PowerShell 的 Start-Process 命令来运行程序
        ps_command = f'Start-Process -FilePath "{command}" -Credential (New-Object System.Management.Automation.PSCredential("{username}", (ConvertTo-SecureString "{password}" -AsPlainText -Force)))'
        subprocess.run(['powershell', '-Command', ps_command], shell=True, check=True)
        return True
    except subprocess.CalledProcessError as e:
        messagebox.showerror("错误", f"无法以新用户身份运行飞书: {e}")
        return False

def save_config(username, password, lark_path):
    """保存配置到文件"""
    config = {
        "username": username,
        "password": password,
        "lark_path": lark_path
    }
    with open(CONFIG_FILE, 'w') as f:
        json.dump(config, f)

def load_config():
    """从文件加载配置"""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r') as f:
                return json.load(f)
        except json.JSONDecodeError:
            # 如果 JSON 文件格式错误，返回 None
            return None
    return None

class LarkDualLauncherApp:
    def __init__(self, root):
        self.root = root
        self.root.title("飞书双开助手")

        # 加载配置
        self.config = load_config()

        # 如果配置存在且完整，直接启动飞书
        if self.config and self.config.get("username") and self.config.get("password") and self.config.get("lark_path"):
            self.launch_lark()
            return

        # 否则显示配置界面
        self.show_config_ui()

    def show_config_ui(self):
        """显示配置界面"""
        # 用户名输入
        self.label_username = tk.Label(self.root, text="用户名:")
        self.label_username.grid(row=0, column=0, padx=10, pady=10)
        self.entry_username = tk.Entry(self.root)
        self.entry_username.grid(row=0, column=1, padx=10, pady=10)
        if self.config:
            self.entry_username.insert(0, self.config.get("username", ""))

        # 密码输入
        self.label_password = tk.Label(self.root, text="密码:")
        self.label_password.grid(row=1, column=0, padx=10, pady=10)
        self.entry_password = tk.Entry(self.root, show="*")
        self.entry_password.grid(row=1, column=1, padx=10, pady=10)
        if self.config:
            self.entry_password.insert(0, self.config.get("password", ""))

        # 飞书路径选择
        self.label_lark_path = tk.Label(self.root, text="飞书路径:")
        self.label_lark_path.grid(row=2, column=0, padx=10, pady=10)
        self.entry_lark_path = tk.Entry(self.root)
        self.entry_lark_path.grid(row=2, column=1, padx=10, pady=10)
        if self.config:
            self.entry_lark_path.insert(0, self.config.get("lark_path", ""))
        self.button_browse = tk.Button(self.root, text="浏览", command=self.browse_lark_path)
        self.button_browse.grid(row=2, column=2, padx=10, pady=10)

        # 保存配置按钮
        self.button_save = tk.Button(self.root, text="保存并启动", command=self.save_and_launch)
        self.button_save.grid(row=3, column=0, columnspan=2, padx=10, pady=10)

    def browse_lark_path(self):
        """选择飞书路径"""
        lark_path = filedialog.askopenfilename(title="选择飞书应用程序", filetypes=[("Executable files", "*.exe")])
        if lark_path:
            self.entry_lark_path.delete(0, tk.END)
            self.entry_lark_path.insert(0, lark_path)

    def save_and_launch(self):
        """保存配置并启动飞书"""
        username = self.entry_username.get()
        password = self.entry_password.get()
        lark_path = self.entry_lark_path.get()

        if not username or not password or not lark_path:
            messagebox.showerror("错误", "请填写所有字段")
            return

        if not os.path.exists(lark_path):
            messagebox.showerror("错误", "飞书路径无效")
            return

        # 保存配置
        save_config(username, password, lark_path)

        # 启动飞书
        self.launch_lark()

    def launch_lark(self):
        """启动飞书"""
        username = self.config.get("username")
        password = self.config.get("password")
        lark_path = self.config.get("lark_path")

        # 检查用户是否已存在
        if not user_exists(username):
            if not create_user(username, password):
                return  # 如果创建用户失败，直接返回

        # 用户创建成功后，直接启动飞书
        if run_as_user(username, password, lark_path):
            messagebox.showinfo("成功", "飞书已启动")
            self.root.quit()  # 关闭 Tkinter 主窗口
            sys.exit()  # 终止程序
        else:
            messagebox.showerror("错误", "无法以新用户身份运行飞书")

if __name__ == "__main__":
    root = tk.Tk()
    app = LarkDualLauncherApp(root)
    root.mainloop()
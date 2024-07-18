# Korabli Localization Installer GUI
# Copyright © 2024 MikhailTapio
#
# This program is free software: you can redistribute it and/or modify it under the terms of the GNU Affero General
# Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option)
# any later version.
#
# This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied
# warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License along with this program. If not,
# see <https://www.gnu.org/licenses/>.
import os
import shutil
import subprocess
import sys
import threading
import tkinter as tk
import urllib.request
# pip install urllib3==1.25.11
# The newer urllib has break changes.
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from tkinter import filedialog
from tkinter import ttk
from typing import Any

import requests

locale_config = '''<locale_config>
    <locale_id>ru</locale_id>
    <text_path>../res/texts</text_path>
    <text_domain>global</text_domain>
    <lang_mapping>
        <lang acceptLang="ru" egs="ru" fonts="CN" full="schinese" languageBar="true" localeRfcName="ru" short="ru" />
    </lang_mapping>
</locale_config>
'''

base_path: str = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
resource_path: str = os.path.join(base_path, "resources")
_run_dir: str = ''
is_installing: bool = False


class LocalizationInstaller:
    game_version: tk.StringVar
    localization_status: tk.StringVar
    is_release: tk.BooleanVar
    download_source: tk.StringVar
    builtin_mods_selection: tk.BooleanVar
    mod_selection: tk.BooleanVar
    mo_path: tk.StringVar
    install_progress: tk.StringVar
    download_info: tk.StringVar

    def __init__(self, parent: tk.Tk):
        mkdir('l10n_installer/cache')
        mkdir('l10n_installer/downloads')
        mkdir('l10n_installer/mods')
        mkdir('l10n_installer/processed')
        self.root = parent
        self.root.title('LocalizedKorabli汉化安装器')
        half_screen_width = int(self.root.winfo_screenwidth() / 2) - 150
        half_screen_height = int(self.root.winfo_screenheight() / 2) - 150
        self.root.geometry(f'+{half_screen_width}+{half_screen_height}')

        self.game_version = tk.StringVar()
        self.localization_status = tk.StringVar()
        self.is_release = tk.BooleanVar()
        self.download_source = tk.StringVar()
        self.builtin_mods_selection = tk.BooleanVar()
        self.mod_selection = tk.BooleanVar()
        self.mo_path = tk.StringVar()
        self.install_progress = tk.StringVar()
        self.game_launcher_status = tk.StringVar()
        self.download_info = tk.StringVar()

        # 第一行：游戏版本
        tk.Label(parent, textvariable=self.game_version) \
            .grid(row=0, column=0, columnspan=3, sticky=tk.W)
        self.game_version.set('游戏版本：' + self.get_run_dir())

        # 第二行：汉化状态
        tk.Label(parent, textvariable=self.localization_status) \
            .grid(row=1, column=0, columnspan=3, sticky=tk.W)
        self.localization_status.set('汉化状态：' + '未安装')

        # 第三行：游戏类型
        tk.Label(parent, text='游戏类型：').grid(row=2, column=0, sticky=tk.W)

        # 游戏类型选项
        ttk.Radiobutton(parent, text='正式服', variable=self.is_release, value=True) \
            .grid(row=2, column=1, sticky=tk.W)
        ttk.Radiobutton(parent, text='测试（PT）服', variable=self.is_release, value=False) \
            .grid(row=2, column=2, sticky=tk.W)

        # 第四行：下载源
        tk.Label(parent, text='汉化来源：').grid(row=3, column=0, sticky=tk.W)
        # 下载源选项
        ttk.Radiobutton(parent, text='Gitee', variable=self.download_source, value='gitee') \
            .grid(row=3, column=1, sticky=tk.W)
        ttk.Radiobutton(parent, text='GitHub', variable=self.download_source, value='github') \
            .grid(row=3, column=2, sticky=tk.W)
        ttk.Radiobutton(parent, text='本地文件', variable=self.download_source, value='local') \
            .grid(row=3, column=3, sticky=tk.W)

        # 第五行：体验增强包/汉化修改包
        ttk.Checkbutton(parent, text='安装体验增强包', variable=self.builtin_mods_selection) \
            .grid(row=5, column=0, sticky=tk.W)
        ttk.Checkbutton(parent, text='安装汉化修改包', variable=self.mod_selection) \
            .grid(row=5, column=1, columnspan=2, sticky=tk.W)

        # 第六行：安装路径选择/下载进度
        self.install_path_entry = tk.Entry(parent, textvariable=self.mo_path, width=30)
        self.install_path_button = tk.Button(parent, text='选择文件', command=self.choose_mo)
        self.download_progress_label = tk.Label(parent, text='下载进度：')
        self.download_progress_info = tk.Label(parent, textvariable=self.download_info)

        # 第七行：安装/更新按钮
        self.install_button = tk.Button(parent, text='安装或更新', command=self.install_update)
        self.install_button.grid(row=6, column=0)

        # 安装进度
        tk.Label(parent, textvariable=self.install_progress).grid(row=6, column=1, sticky=tk.W)

        # 第八行：启动游戏
        self.launch_button = tk.Button(parent, text='启动客户端', command=launch_game)
        self.launch_button.grid(row=7, column=0)

        # 启动器状态
        tk.Label(parent, textvariable=self.game_launcher_status).grid(row=7, column=1, sticky=tk.W)

        # 根据下载源选项显示或隐藏安装路径选择
        self.download_source.trace('w', self.toggle_install_path)

        self.builtin_mods_selection.set(True)
        self.mod_selection.set(True)
        self.is_release.set(True)
        self.download_source.set('gitee')
        self.download_info.set('准备')
        self.install_progress.set('安装进度：' + '等待中')
        self.game_launcher_status.set(find_launcher())

    def choose_mo(self):
        mo_path = filedialog.askopenfilename(initialdir='.', filetypes=[('MO文件', '*.mo')])
        if mo_path:
            self.mo_path.set(mo_path)

    def install_update(self):
        global is_installing
        if is_installing:
            return
        is_installing = True
        thread = threading.Thread(target=self.do_install_update())
        thread.start()

    def do_install_update(self):
        run_dir = self.get_run_dir()
        try:
            int(run_dir)
        except ValueError:
            return
        is_release = self.is_release.get()
        target_path = Path('bin').joinpath(run_dir).joinpath('res_mods' if is_release else 'res')
        mkdir(target_path)
        self.install_progress.set('安装locale_config')
        if not is_release:
            old_cfg = target_path.joinpath('locale_config.xml')
            old_cfg_renamed = target_path.joinpath('locale_config.xml.old')
            if not os.path.isfile(old_cfg_renamed):
                if os.path.isfile(old_cfg):
                    shutil.copy(old_cfg, old_cfg_renamed)
            with open(old_cfg, "w", encoding="utf-8") as file:
                file.write(locale_config)
        else:
            with open(target_path.joinpath('locale_config.xml'), "w", encoding="utf-8") as file:
                file.write(locale_config)
        self.install_progress.set('安装locale_config——完成')
        proxies = {scheme: proxy for scheme, proxy in urllib.request.getproxies().items()}
        if is_release:
            # EE
            if self.builtin_mods_selection.get():
                self.install_progress.set('安装体验增强包')
                output_file = 'l10n_installer/downloads/LK_EE.zip'
                self.download_info.set('下载体验增强包——连接中')
                ee_ready = False
                try:
                    response = requests.get('https://gitee.com/localized-korabli/Korabli-LESTA-L10N/raw/main'
                                            '/BuiltInMods/LKExperienceEnhancement.zip', stream=True, proxies=proxies)
                    status = response.status_code
                    if status == 200:
                        self.download_info.set('下载体验增强包——下载中')
                        with open(output_file, 'wb') as f:
                            for chunk in response.iter_content(chunk_size=1024):
                                if chunk:
                                    f.write(chunk)
                        ee_ready = True
                        self.download_info.set("下载体验增强包——完成")
                    else:
                        self.download_info.set(f'下载体验增强包——失败（{status}）')
                except requests.exceptions.RequestException:
                    self.download_info.set('下载体验增强包——请求异常')
                if ee_ready:
                    with zipfile.ZipFile(output_file, 'r') as ee_zip:
                        ee_zip.extractall(target_path)
                    self.install_progress.set('安装体验增强包——完成')
                else:
                    self.install_progress.set('安装体验增强包——失败')
        self.install_progress.set('安装汉化')

        global is_installing
        is_installing = False

    def get_run_dir(self) -> str:
        global _run_dir
        if _run_dir and _run_dir != '':
            return _run_dir
        if os.path.isfile('game_info.xml'):
            game_info = ET.parse('game_info.xml')
            for version in game_info.findall('.//version'):
                if version.get('name') == 'locale':
                    _run_dir = str(version.get('installed').split('.')[-1])
                    return _run_dir
            _run_dir = '未知'
            return _run_dir
        _run_dir = '未在运行目录下找到战舰世界客户端！'
        return _run_dir

    def toggle_install_path(self, *args):
        if self.download_source.get() == 'local':
            self.install_path_entry.grid(row=4, column=0, columnspan=3, pady=12)
            self.install_path_button.grid(row=4, column=3)
            self.download_progress_label.grid_forget()
            self.download_progress_info.grid_forget()
        else:
            self.download_progress_label.grid(row=4, column=0, pady=11, sticky=tk.W)
            self.download_progress_info.grid(row=4, column=1, pady=11, sticky=tk.W)
            self.install_path_entry.grid_forget()
            self.install_path_button.grid_forget()


def mkdir(t_dir: Any):
    os.makedirs(t_dir, exist_ok=True)


def launch_game():
    if os.path.isfile('lgc_api.exe'):
        subprocess.run('lgc_api.exe')


def find_launcher() -> str:
    if os.path.isfile('lgc_api.exe'):
        return '莱服客户端'
    return '未找到客户端'


if __name__ == '__main__':
    root = tk.Tk()
    app = LocalizationInstaller(root)
    root.mainloop()

# pyinstaller -w -i icon.ico --onefile --add-data "resources\*;resources" installer_gui.py --clean

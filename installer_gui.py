# Korabli Localization Installer GUI
# Copyright © 2024 澪刻LocalizedKorabli
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
import ctypes
import hashlib
import json
import os
import shutil
import string
import subprocess
import sys
import threading
import time
import tkinter as tk
import urllib.request
import webbrowser
import winreg
# pip install urllib3==1.25.11
# The newer urllib has break changes.
import xml.etree.ElementTree as ET
import zipfile
from datetime import datetime
from optparse import OptionParser
from pathlib import Path
from tkinter import filedialog, font
from typing import Any, Dict, List, Tuple, Optional, Union

import polib
import pythoncom
import requests
import ttkbootstrap as ttk
import winshell
from tktooltip import ToolTip
from ttkbootstrap.dialogs.dialogs import Messagebox

mods_link = 'https://tapio.lanzn.com/b0nxzso2b'
project_repo_link = 'https://github.com/LocalizedKorabli/Korabli-LESTA-L10N/'
installer_repo_link = 'https://github.com/LocalizedKorabli/L10nInstallerGUI/'

version = '0.2.1'

locale_config = '''<locale_config>
    <locale_id>ru</locale_id>
    <text_path>../res/texts</text_path>
    <text_domain>global</text_domain>
    <lang_mapping>
        <lang acceptLang="ru" egs="ru" fonts="CN" full="schinese" languageBar="true" localeRfcName="ru" short="ru" />
    </lang_mapping>
</locale_config>
'''

download_routes = {
    'r': {
        'gitee': {
            'url': 'https://gitee.com/localized-korabli/Korabli-LESTA-L10N/raw/main/Localizations/latest/',
            'direct': False
        },
        'gitlab': {
            'url': 'https://gitlab.com/localizedkorabli/korabli-lesta-l10n/-/raw/main/Localizations/latest/',
            'direct': False
        },
        'github': {
            'url': 'https://github.com/LocalizedKorabli/Korabli-LESTA-L10N/raw/main/Localizations/latest/',
            'direct': False
        }
    },
    'pt': {
        'gitee': {
            'url': 'https://gitee.com/localized-korabli/Korabli-LESTA-L10N-PublicTest/raw/Localizations/Localizations'
                   '/latest/',
            'direct': False
        },
        'gitlab': {
            'url': 'https://gitlab.com/localizedkorabli/korabli-lesta-l10n-publictest/-/raw/Localizations'
                   '/Localizations/latest/',
            'direct': False
        },
        'github': {
            'url': 'https://github.com/LocalizedKorabli/Korabli-LESTA-L10N-PublicTest/raw/Localizations/Localizations'
                   '/latest/',
            'direct': False
        }
    }
}

server_regions_dict: Dict[str, Tuple[str, bool]] = {
    'WOWS.RU.PRODUCTION': ('ru', True),
    'WOWS.RPT.PRODUCTION': ('ru', False),
    'WOWS.WW.PRODUCTION': ('zh_sg', True),
    'WOWS.PT.PRODUCTION': ('zh_sg', False),
    'WOWS.CN.PRODUCTION': ('zh_cn', True)
}

launcher_dict: Dict[str, str] = {
    'lgc_api.exe': '莱服客户端',
    'wgc_api.exe': '直营服客户端',
    'wgc360_api.exe': '国服客户端'
}

base_path: str = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
resource_path: str = os.path.join(base_path, 'resources')
game_path_current = '<程序运行目录>'
game_path_unknown = '<请选择游戏目录>'
au_shortcut_path_desktop = '<用户桌面/默认名称>'

msg_game_path_may_be_invalid = '''您选择的游戏目录缺失必要的游戏文件，
可能并非战舰世界安装目录。'''
msg_auto_update_notification = '''自动更新原理：
程序在用户指定的位置生成快捷方式，
用户通过快捷方式启动安装器自动更新模式，
当远程汉化包版本新于本地版本时，
安装器执行更新操作，随后自动启动游戏。

注意：
1. 仅当您通过生成的快捷方式启动游戏时，才会执行自动更新。

2. 更新脚本的目标游戏路径、游戏类型、汉化来源等参数以您
在本次安装汉化时的设置为准；

3. 汉化来源被指定为本地文件时，将不会生成更新脚本；

4. 快捷方式默认生成为"用户桌面/启动战舰世界.lnk"，
您也可以自行选择生成位置。'''

tooltip_auto_search_clients = '''自动检测电脑上安装的
战舰世界（莱服）客户端'''

tooltip_select_game_path = '手动选择战舰世界客户端路径'

tooltip_detect_game_type = '自动检测游戏区服和类型'

tooltip_src_gitee = '从适合大陆用户的Gitee线路（镜像）下载汉化包'

tooltip_src_gitlab = '从适合全球用户的GitLab线路（镜像）下载汉化包（大陆用户访问可能较慢）'

tooltip_src_github = '从适合港澳台/国外用户的GitHub线路（源仓库）下载汉化包'

tooltip_src_local = '选择本地汉化包'

tooltip_mo_path_selection = '手动选择要安装的汉化包文件'

tooltip_ee_selection = '''体验增强包以战舰世界模组形式
提供了一些视觉修改，如中文开屏logo、
解除战斗载入界面战舰名称字数显示限制等'''

tooltip_mods_selection = '''模组（汉化修改包）允许用户
对汉化文件做局部修改，
以获得更好的游戏体验。
与澪刻安装器联动的模组
需要您打开此选项。'''

tooltip_isolation = '''在版本隔离被启用的情况下，
程序将从游戏目录下的l10n_installer/mods
文件夹中读取要应用的模组（汉化修改包）'''

tooltip_mods_dir = '打开模组（汉化修改包）文件夹'

tooltip_download_mods = '打开基于蓝奏云的模组中心'

tooltip_au = '''若自动更新选项被勾选，在汉化安装成功后，
将根据本次安装设置生成快捷方式，
用于启动自动更新模式下的安装器'''

tooltip_launch_game = '启动战舰世界客户端'

tooltip_about = '查看澪刻汉化托管仓库'

tooltip_source_code = '查看本安装器代码仓库'

tooltip_license = '本开源项目使用GNU AGPL 3.0许可证'


class LocalizationInstallerAuto:
    no_gui: bool
    no_run: bool
    game_path: Path
    is_release: bool
    use_ee: bool
    use_mods: bool
    isolation: bool
    download_src: str
    server_region: str
    installing: bool = True

    install_progress_bar: ttk.Progressbar
    install_progress: tk.DoubleVar

    def __init__(self, parent: tk.Tk, options: Any):
        self.no_gui = bool(options.no_gui)
        if self.no_gui:
            parent.overrideredirect(True)
            parent.withdraw()
        self.root = parent
        self.root.title(f'汉化安装器[自动更新模式]v{version}')
        self.no_run = bool(options.no_run)
        self.game_path = Path(options.game_path)
        self.is_release = bool(options.is_release)
        self.use_ee = bool(options.use_ee)
        self.use_mods = bool(options.use_mods)
        self.isolation = bool(options.isolation)
        self.download_src = options.download_src
        self.server_region = options.server_region
        self.install_progress = tk.DoubleVar()

        ttk.Label(self.root, text='自动更新：') \
            .grid(row=0, column=0, columnspan=1, padx=10, pady=10, sticky=tk.W)
        self.install_progress_bar = ttk.Progressbar(parent, variable=self.install_progress, maximum=100.0,
                                                    style='success-striped', length=780)
        self.install_progress_bar.grid(row=1, column=0, columnspan=4, padx=10, pady=10)

        self.safely_set_install_progress(progress=0.0)

        self.update_timer()
        # print(f'{self.game_path}+{self.is_release}+{self.use_ee}+{self.use_mods}+{self.download_src}+{self.server_region}')

        tr = threading.Thread(target=self.do_install_update)
        tr.start()

    def do_install_update(self):
        _install_update(self,
                        self.game_path, self.is_release,
                        self.use_ee, self.use_mods,
                        self.isolation, self.download_src,
                        self.server_region)
        self.installing = False

    def safely_set_install_progress(self, progress: Optional[float] = None):
        self.root.after(0, self.install_progress.set(progress))

    def update_timer(self):
        if not self.installing:
            self.root.destroy()
        self.root.after(1000, self.update_timer)

    def on_closed(self):
        if not self.no_run:
            subprocess.run(find_launcher(self.game_path)[0])


class LocalizationInstaller:
    # Components
    install_progress_bar: ttk.Progressbar

    # GUI Related Variables
    localization_status_1st: tk.StringVar
    localization_status_2nd: tk.StringVar
    server_region: tk.StringVar
    is_release: tk.BooleanVar
    download_source: tk.StringVar
    ee_selection: tk.BooleanVar
    mods_selection: tk.BooleanVar
    isolation: tk.BooleanVar
    mo_path: tk.StringVar
    install_progress_text: tk.StringVar
    download_progress_text: tk.StringVar
    install_progress: tk.DoubleVar
    game_path: tk.StringVar

    # Variables
    global_settings: Dict[str, Any] = None
    choice: Dict[str, Any] = None
    last_installed_l10n_version = ''
    run_dirs: Dict[str, str] = {}
    is_installing: bool = False
    game_launcher_file: Optional[Path] = None
    available_game_paths: List[str] = []
    user_desktop_path: Optional[Path] = None

    def __init__(self, parent: tk.Tk):
        self.root = parent
        self.root.title(f'汉化安装器v{version}')

        self.game_path = tk.StringVar()
        self.localization_status_1st = tk.StringVar()
        self.localization_status_2nd = tk.StringVar()
        self.server_region = tk.StringVar()
        self.is_release = tk.BooleanVar()
        self.download_source = tk.StringVar()
        self.ee_selection = tk.BooleanVar()
        self.mods_selection = tk.BooleanVar()
        self.isolation = tk.BooleanVar()
        self.mo_path = tk.StringVar()
        self.install_progress_text = tk.StringVar()
        self.game_launcher_status = tk.StringVar()
        self.download_progress_text = tk.StringVar()
        self.install_progress = tk.DoubleVar()
        self.gen_auto_update = tk.BooleanVar()
        self.gen_auto_update_path = tk.StringVar()

        # 游戏目录
        ttk.Label(parent, text='游戏目录：') \
            .grid(row=0, column=0, columnspan=1, padx=5, pady=5, sticky=tk.W)

        self.game_path_combo = ttk.Combobox(parent, width=26, textvariable=self.game_path, state='readonly')
        self.game_path_combo.grid(row=0, column=1, columnspan=3, padx=5, pady=5, sticky=tk.W)
        ToolTip(self.game_path_combo, msg=lambda: get_str_from_optional_path(self.get_game_path(), '游戏目录：'),
                delay=1.0)

        self.refresh_path_combo()

        self.auto_search = ttk.Button(parent, text='自动检测', command=lambda: self.find_game(overwrite=False),
                                      style='success')
        self.auto_search.grid(row=1, column=2, columnspan=1)
        ToolTip(self.auto_search, msg=tooltip_auto_search_clients, delay=1.0)

        self.game_path_button = ttk.Button(parent, text='选择目录', command=self.choose_path, style='warning')
        self.game_path_button.grid(row=1, column=3, columnspan=1)
        ToolTip(self.game_path_button, msg=tooltip_select_game_path, delay=1.0)

        # 游戏版本
        ttk.Label(parent, text='游戏版本/汉化版本') \
            .grid(row=1, column=0, columnspan=4, padx=5, pady=5, sticky=tk.W)

        # 汉化状态
        self.localization_status_label_1st = ttk.Label(parent, textvariable=self.localization_status_1st)
        self.localization_status_label_1st.grid(row=3, column=0, columnspan=4, padx=5, pady=5, sticky=tk.W)
        self.localization_status_label_2nd = ttk.Label(parent, textvariable=self.localization_status_2nd)
        self.localization_status_label_2nd.grid(row=4, column=0, columnspan=4, padx=5, pady=5, sticky=tk.W)

        # 游戏区服
        ttk.Label(parent, text='游戏区服：').grid(row=5, column=0, padx=5, pady=5, sticky=tk.W)

        # 游戏区服选项
        ttk.Radiobutton(parent, text='莱服', variable=self.server_region, value='ru') \
            .grid(row=5, column=1, padx=5, pady=5, sticky=tk.W)
        ttk.Radiobutton(parent, text='直营服', variable=self.server_region, value='zh_sg', style='warning') \
            .grid(row=5, column=2, padx=5, pady=5, sticky=tk.W)
        ttk.Radiobutton(parent, text='国服', variable=self.server_region, value='zh_cn', style='danger') \
            .grid(row=5, column=3, padx=5, pady=5, sticky=tk.W)

        # 游戏类型
        ttk.Label(parent, text='游戏类型：').grid(row=6, column=0, padx=5, pady=5, sticky=tk.W)

        # 游戏类型选项
        ttk.Radiobutton(parent, text='正式服', variable=self.is_release, value=True, style='success') \
            .grid(row=6, column=1, padx=5, pady=5, sticky=tk.W)
        ttk.Radiobutton(parent, text='PT服', variable=self.is_release, value=False, style='danger') \
            .grid(row=6, column=2, padx=5, pady=5, sticky=tk.W)

        self.detect_game_type_button = ttk.Button(parent, text='自动检测',
                                                  command=lambda: self.detect_game_status(manually=True))
        self.detect_game_type_button.grid(row=6, column=3)
        ToolTip(self.detect_game_type_button, tooltip_detect_game_type, delay=1.0)

        # 下载源
        ttk.Label(parent, text='汉化来源：').grid(row=7, column=0, padx=5, pady=5, sticky=tk.W)
        # 下载源选项
        self.gitee_button = ttk.Radiobutton(parent, text='Gitee', variable=self.download_source,
                                            value='gitee', style='danger')
        self.gitee_button.grid(row=8, column=0, padx=5, pady=5, sticky=tk.W)
        ToolTip(self.gitee_button, tooltip_src_gitee, delay=1.0)
        self.gitlab_button = ttk.Radiobutton(parent, text='GitLab', variable=self.download_source,
                                             value='gitlab', style='warning')
        self.gitlab_button.grid(row=8, column=1, padx=5, pady=5, sticky=tk.W)
        ToolTip(self.gitlab_button, tooltip_src_gitlab, delay=1.0)
        self.github_button = ttk.Radiobutton(parent, text='GitHub', variable=self.download_source,
                                             value='github', style='dark')
        self.github_button.grid(row=8, column=2, padx=5, pady=5, sticky=tk.W)
        ToolTip(self.github_button, tooltip_src_github, delay=1.0)
        self.local_file_button = ttk.Radiobutton(parent, text='本地文件', variable=self.download_source,
                                                 value='local')
        self.local_file_button.grid(row=8, column=3, padx=5, pady=5, sticky=tk.W)
        ToolTip(self.local_file_button, tooltip_src_local, delay=1.0)

        # 体验增强包/汉化修改包
        self.ee_check_button = ttk.Checkbutton(parent, text='安装体验增强包', variable=self.ee_selection)
        self.ee_check_button.grid(row=10, column=0, columnspan=2, padx=10, pady=5, sticky=tk.W)
        ToolTip(self.ee_check_button, tooltip_ee_selection, delay=1.0)
        self.mods_check_button = ttk.Checkbutton(parent, text='安装模组（汉化修改包）', variable=self.mods_selection)
        self.mods_check_button.grid(row=11, column=0, columnspan=2, padx=10, pady=5, sticky=tk.W)
        ToolTip(self.mods_check_button, tooltip_mods_selection, delay=1.0)
        self.isolation_check_button = ttk.Checkbutton(parent, text='版本隔离', variable=self.isolation)
        self.isolation_check_button.grid(row=12, column=0, columnspan=2, padx=10, pady=5, sticky=tk.W)
        ToolTip(self.isolation_check_button, tooltip_isolation, delay=1.0)
        self.mods_dir_button = ttk.Button(parent, text='模组目录', command=self.open_mods_folder)
        self.mods_dir_button.grid(row=12, column=2, columnspan=1)
        ToolTip(self.mods_dir_button, tooltip_mods_dir, delay=1.0)
        self.download_mods_button = ttk.Button(parent, text='下载模组', style='info',
                                               command=lambda: webbrowser.open_new_tab(mods_link))
        self.download_mods_button.grid(row=12, column=3, columnspan=1)
        ToolTip(self.download_mods_button, tooltip_download_mods, delay=1.0)

        # 安装路径选择/下载进度
        self.mo_path_entry = ttk.Entry(parent, textvariable=self.mo_path, width=28, state='readonly')
        self.mo_path_selection_button = ttk.Button(parent, text='选择文件', command=self.choose_mo)
        ToolTip(self.mo_path_selection_button, tooltip_mo_path_selection, delay=1.0)
        self.download_progress_label = ttk.Label(parent, text='下载进度：')
        self.download_progress_info = ttk.Label(parent, textvariable=self.download_progress_text)

        # 安装/更新按钮
        self.install_button = ttk.Button(parent, text='安装汉化', command=self.install_update,
                                         style=ttk.SUCCESS)
        self.install_button.grid(row=13, column=0, pady=5)

        # 安装进度
        ttk.Label(parent, textvariable=self.install_progress_text).grid(row=13, column=1, columnspan=3,
                                                                        padx=5, sticky=tk.W)

        self.install_progress_bar = ttk.Progressbar(parent, variable=self.install_progress, maximum=100.0,
                                                    style='success-striped', length=400)
        self.install_progress_bar.grid(row=14, column=0, columnspan=4, padx=10, pady=5)

        # 自动更新
        self.gen_auto_update_button = ttk.Checkbutton(parent, text='自动更新', variable=self.gen_auto_update)
        self.gen_auto_update_button.grid(row=15, column=0, columnspan=1, padx=10, pady=10)
        ToolTip(self.gen_auto_update_button, tooltip_au, delay=1.0)

        self.gen_auto_update_entry = ttk.Entry(parent, textvariable=self.gen_auto_update_path, width=18,
                                               state='readonly')
        ToolTip(self.gen_auto_update_entry, msg=lambda: get_str_from_optional_path(
            self.get_au_shortcut_path(), '快捷方式生成目录：'
        ), delay=1.0)

        self.choose_au_shortcut_path_button = ttk.Button(parent, text='选择位置', command=self.choose_au_shortcut_path,
                                                         style='dark')

        # 启动游戏
        self.launch_button = ttk.Button(parent, text='启动游戏', command=self.launch_game, style=ttk.WARNING)
        self.launch_button.grid(row=16, column=0, pady=5)
        ToolTip(self.launch_button, tooltip_launch_game, delay=1.0)

        # 启动器状态
        ttk.Label(parent, textvariable=self.game_launcher_status).grid(row=16, column=1, columnspan=3,
                                                                       padx=5, sticky=tk.W)

        # 相关链接
        self.about_button = ttk.Button(parent, text='关于项目',
                                       command=lambda: webbrowser.open_new_tab(project_repo_link),
                                       style=ttk.INFO)
        self.about_button.grid(row=17, column=0, pady=5)
        ToolTip(self.about_button, tooltip_about, delay=1.0)

        self.src_button = ttk.Button(parent, text='代码仓库',
                                     command=lambda: webbrowser.open_new_tab(installer_repo_link),
                                     style=ttk.DANGER)
        self.src_button.grid(row=17, column=1, pady=5, padx=5)
        ToolTip(self.src_button, tooltip_source_code, delay=1.0)

        # 版权声明
        self.license_text = ttk.Label(parent, text='© 2024 澪刻本地化')
        self.license_text.grid(row=17, column=2, columnspan=3, pady=5)
        ToolTip(self.license_text, tooltip_license, delay=1.0)

        # 根据汉化来源选项显示或隐藏安装路径选择
        self.download_source.trace('w', self.on_download_source_changed)
        # 更换游戏路径时，刷新数据
        self.game_path.trace('w', self.on_game_path_changed)
        # 非莱斯塔正式服客户端无需安装体验增强包
        self.server_region.trace('w', self.on_server_region_or_game_type_changed)
        self.is_release.trace('w', self.on_server_region_or_game_type_changed)
        # 自动更新被勾选时，显示快捷方式生成目录选项
        self.gen_auto_update.trace('w', self.on_au_selected)

        mkdir('l10n_installer/downloads')
        mkdir('l10n_installer/mods')
        mkdir('l10n_installer/processed')
        mkdir('l10n_installer/settings')

        global_settings = self.parse_global_settings()
        last_saved_paths = global_settings.get('available_game_paths')
        if isinstance(last_saved_paths, list):
            for saved_path in last_saved_paths:
                self.available_game_paths.append(saved_path)
            self.refresh_path_combo()
        self.game_path.set(global_settings.get('last_game_path'))
        self.game_path_combo.current()
        self.find_game()
        self.on_game_path_changed()

        self.gen_auto_update_path.set(au_shortcut_path_desktop)

        self.reset_progress()
        self.game_launcher_status.set(find_launcher(self.get_game_path())[1])

    def reset_progress(self):
        self.safely_set_download_progress_text('等待中')
        self.safely_set_install_progress_text('等待中')
        self.safely_set_install_progress(progress=0.0)

    def safely_set_download_progress_text(self, msg: str):
        self.root.after(0, self.download_progress_text.set(msg))

    def safely_set_install_progress_text(self, msg: str):
        self.root.after(0, self.install_progress_text.set('进度：' + msg))

    def safely_set_install_progress(self, progress: Optional[float] = None):
        self.root.after(0, self.install_progress.set(progress))

    def refresh_path_combo(self):
        self.game_path_combo['values'] = list(dict.fromkeys(self.available_game_paths))

    def get_game_path(self, find: bool = True) -> Optional[Path]:
        game_path_str = self.game_path.get()
        if game_path_str == game_path_unknown:
            return self.find_game() if find else None
        if game_path_str == game_path_current:
            return Path('.')
        return Path(game_path_str)

    def on_download_source_changed(self, *args):
        if self.download_source.get() == 'local':
            self.mo_path_entry.grid(row=9, column=0, columnspan=3)
            self.mo_path_selection_button.grid(row=9, column=3)
            self.download_progress_label.grid_forget()
            self.download_progress_info.grid_forget()
            self.gen_auto_update.set(False)
            self.gen_auto_update_button.configure(state='disabled')
        else:
            self.download_progress_label.grid(row=9, column=0, padx=5, pady=5, sticky=tk.W)
            self.download_progress_info.grid(row=9, column=1, pady=5, columnspan=3, sticky=tk.W)
            self.mo_path_entry.grid_forget()
            self.mo_path_selection_button.grid_forget()
            self.gen_auto_update_button.configure(state='')

    def on_game_path_changed(self, *args) -> None:
        self.reset_progress()
        self.gen_auto_update.set(False)
        if self.game_path.get() == game_path_unknown:
            return

        game_path = self.get_game_path()
        if not game_path:
            return

        self.detect_game_status()
        parse_game_version(self, self.get_game_path())
        self.game_launcher_file = find_launcher(self.get_game_path())[0]

        mkdir(game_path.joinpath('l10n_installer/settings'))
        mkdir(game_path.joinpath('l10n_installer/mods'))

        choice = self.parse_choice(use_cache=False)
        if not choice:
            return

        self.server_region.set(choice.get('server_region', 'ru'))
        self.is_release.set(choice.get('is_release', True))
        self.download_source.set(choice.get('download_source', 'gitee'))
        self.ee_selection.set(choice.get('use_ee', True))
        self.mods_selection.set(choice.get('apply_mods', True))
        self.isolation.set(choice.get('isolation', False))

    def on_server_region_or_game_type_changed(self, *args):
        self.ee_check_button.configure(state=('' if self.supports_ee() else 'disabled'))

    # def on_mods_selection_changed(self, *args):
    #     self.isolation_check_button.configure(state='' if self.mods_selection.get() else 'disabled')

    def on_au_selected(self, *args):
        if self.gen_auto_update.get():
            Messagebox.ok(msg_auto_update_notification)
            self.gen_auto_update_entry.grid(row=15, column=1, columnspan=2)
            self.choose_au_shortcut_path_button.grid(row=15, column=3, columnspan=1)
        else:
            self.gen_auto_update_entry.grid_forget()
            self.choose_au_shortcut_path_button.grid_forget()

    def get_au_shortcut_path(self) -> Optional[Path]:
        shortcut_str = self.gen_auto_update_path.get()

        if shortcut_str == au_shortcut_path_desktop:
            desktop_path = self.find_user_desktop()
            if not desktop_path:
                return None
            return desktop_path.joinpath('启动战舰世界.lnk')

        return Path(shortcut_str)

    def supports_ee(self):
        return self.server_region.get() == 'ru' and self.is_release.get()

    def find_game(self, overwrite: bool = True) -> Optional[Path]:
        found_in_reg = self.find_from_reg()
        found_manually = self.find_manually()
        game_path_str = self.game_path.get()
        if not overwrite:
            return None
        if game_path_str != game_path_unknown:
            game_path = Path(game_path_str)
            if is_valid_game_path(game_path):
                return game_path
        if is_valid_game_path(Path('.')):
            self.game_path.set(game_path_current)
            self.available_game_paths.append(game_path_current)
            self.refresh_path_combo()
            return Path('.')
        if found_in_reg:
            reg_first = found_in_reg[0]
            self.game_path.set(str(reg_first.absolute()))
            return reg_first
        # Manually
        if found_manually:
            manually_first = found_manually[0]
            self.game_path.set(str(manually_first.absolute()))
            return manually_first
        return None

    def find_from_reg(self) -> List[Path]:
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Classes\lgc\DefaultIcon') as key:
                lgc_dir_str, _ = winreg.QueryValueEx(key, '')
                if lgc_dir_str is None:
                    # Try the default value
                    lgc_dir_str = r'C:\ProgramData\Lesta\GameCenter\lgc.exe'
                if ',' in lgc_dir_str:
                    lgc_dir_str = lgc_dir_str.split(',')[0]
                preferences_path = Path(lgc_dir_str).parent.joinpath('preferences.xml')
                if not preferences_path.is_file():
                    return []
                pref_root = ET.parse(preferences_path).getroot()
                games_block = pref_root.find('.//application/games_manager/games')
                games = games_block.findall('.//game')
                if not games:
                    return []
                path_strs = [game.find('working_dir').text for game in games if game.find('working_dir') is not None]
                path_strs_filtered = [dir_str for dir_str in path_strs if is_valid_game_path(Path(dir_str))]
                for path_str in path_strs_filtered:
                    self.available_game_paths.append(path_str)
                self.refresh_path_combo()
                return [Path(dir_str) for dir_str in path_strs_filtered]
        except Exception:
            pass

    def find_manually(self) -> List[Path]:
        found_manually: List[Path] = []
        drives = find_all_drives()
        for drive in drives:
            for possible_path in [
                Path(drive).joinpath('Games').joinpath('Korabli'),
                Path(drive).joinpath('Games').joinpath('Korabli_PT'),
                Path(drive).joinpath('Korabli'),
                Path(drive).joinpath('Korabli_PT'),
            ]:
                if is_valid_game_path(possible_path):
                    found_manually.append(possible_path)
        for found_path in found_manually:
            self.available_game_paths.append(str(found_path.absolute()))
        self.refresh_path_combo()
        return found_manually

    def popup_result(self, nothing_wrong: bool):
        if nothing_wrong:
            msg_response = Messagebox.show_question('汉化安装完成。是否启动游戏？', '安装完成', alert=True, buttons=[
                '启动游戏:primary',
                '返回主页:secondary'
            ])
            if msg_response == '启动游戏':
                self.launch_game()
        else:
            Messagebox.show_error('汉化安装失败。请检查您的网络，选择合适的汉化来源重试。', '安装失败')

    def open_mods_folder(self) -> None:
        instance_path = self.get_game_path() if self.isolation.get() else Path('.')
        if not instance_path:
            return
        mods_folder = instance_path.joinpath('l10n_installer').joinpath('mods')
        mkdir(mods_folder)
        subprocess.run(['explorer', mods_folder.absolute()])

    def choose_path(self):
        game_path_chosen = filedialog.askdirectory(
            title='选择游戏目录',
            initialdir='.'
        )
        if game_path_chosen:
            if not is_valid_game_path(Path(game_path_chosen)):
                Messagebox.show_warning(msg_game_path_may_be_invalid)
            else:
                self.available_game_paths.append(game_path_chosen)
                self.refresh_path_combo()
            self.game_path.set(game_path_chosen)

    def choose_mo(self):
        mo_path = filedialog.askopenfilename(
            title='选择汉化包文件',
            initialdir='.',
            filetypes=[('汉化包', ['*.mo', '*.zip']),
                       ('MO汉化文件', '*.mo'),
                       ('打包的汉化文件', '*.zip')]
        )
        if mo_path:
            self.mo_path.set(mo_path)

    def choose_au_shortcut_path(self):
        desktop_path = self.find_user_desktop()
        default_shortcut_path = desktop_path.joinpath('启动战舰世界.lnk')
        shortcut_path = filedialog.asksaveasfilename(
            title='选择自动更新快捷方式生成位置',
            initialfile=default_shortcut_path,
            initialdir=desktop_path,
            filetypes=[('快捷方式', '*.lnk')]
        )
        if shortcut_path:
            if str(Path(shortcut_path).absolute()) == str(self.get_au_shortcut_path().absolute()):
                return
            self.gen_auto_update_path.set(shortcut_path)

    def find_user_desktop(self) -> Optional[Path]:
        if not self.user_desktop_path:
            self.user_desktop_path = Path(os.path.expanduser("~/Desktop"))
        return self.user_desktop_path

    def install_update(self):
        if self.is_installing:
            Messagebox.show_warning('安装已在进行！', '安装汉化')
            return
        self.is_installing = True
        tr = threading.Thread(target=self.do_install_update)
        tr.start()

    def do_install_update(self) -> None:
        if _install_update(self) and self.gen_auto_update.get():
            self.gen_au_script_and_shortcut()

    def gen_au_script_and_shortcut(self) -> None:
        src = self.download_source.get()
        if src == 'local':
            return
        game_path = str(self.get_game_path().absolute())
        release = '--release' if self.is_release.get() else ''
        ee = '--ee' if self.ee_selection.get() else ''
        mods = '--mods' if self.mods_selection.get() else ''
        isolation = '--isolation' if self.isolation.get() else ''
        region = self.server_region.get()

        pythoncom.CoInitialize()
        with winshell.shortcut(str(self.get_au_shortcut_path().absolute())) as link:
            link.path = str(sys.executable)
            link.description = '自动更新汉化并启动战舰世界'
            link.icon_location = str(Path(game_path).joinpath('WorldOfWarships.exe').absolute()), 0
            link.arguments = f'--auto --gamepath "{game_path}" {release} {ee} {mods} {isolation} ' \
                             f'--region {region} --src "{src}"'
        pythoncom.CoUninitialize()

    def get_choice_template(self):
        self.detect_game_status()
        return {
            'server_region': self.server_region.get(),
            'is_release': self.is_release.get(),
            'download_source': 'gitee',
            'use_ee': True,
            'apply_mods': True,
            'isolation': False
        }

    def get_global_settings_template(self):
        return {
            'last_game_path': game_path_unknown,
            'available_game_paths': [
                game_path_current
            ]
        }

    def detect_game_status(self, manually: bool = False):
        if not manually:
            self.server_region.set('ru')
            self.is_release.set(True)
        game_path = self.get_game_path()
        if not game_path:
            return
        game_info_file = game_path.joinpath('game_info.xml')
        if not game_info_file.is_file():
            return
        game_info = ET.parse(game_info_file)
        game_id = game_info.find('.//game/id')
        if game_id is None:
            return
        game_type: (str, bool) = server_regions_dict.get(game_id.text, ('ru', 'PT.PRODUCTION' not in game_id.text))
        self.server_region.set(game_type[0])
        self.is_release.set(game_type[1])

    def launch_game(self) -> None:
        if not self.game_launcher_file or not self.game_launcher_file.is_file():
            self.game_launcher_file = find_launcher(self.get_game_path())[0]
        if not self.game_launcher_file or not self.game_launcher_file.is_file():
            Messagebox.show_warning('未找到客户端！', '启动游戏')
            return
        subprocess.run(self.game_launcher_file)

    def parse_global_settings(self):
        if self.global_settings:
            return self.global_settings
        self.global_settings = {}
        global_settings_file = Path('l10n_installer/settings/global.json')
        if global_settings_file.is_file():
            try:
                with open(global_settings_file, 'r', encoding='utf-8') as f:
                    self.global_settings = json.load(f)
            except Exception:
                pass
        self.check_global_settings()
        return self.global_settings

    def check_global_settings(self):
        template = self.get_global_settings_template()
        for entry in ['last_game_path', 'available_game_paths']:
            if entry not in self.global_settings.keys():
                self.global_settings[entry] = template[entry]

    def save_global_settings(self):
        if self.global_settings:
            self.global_settings['last_game_path'] = self.game_path.get()
            self.global_settings['available_game_paths'] = list(dict.fromkeys(self.available_game_paths))
            self.global_settings['last_installed_l10n_version'] = self.last_installed_l10n_version
            with open('l10n_installer/settings/global.json', 'w', encoding='utf-8') as f:
                json.dump(self.global_settings, f, ensure_ascii=False, indent=4)

    def parse_choice(self, use_cache: bool = True) -> Optional[Dict[str, str]]:
        if use_cache and self.choice:
            return self.choice
        self.choice = {}
        game_path = self.get_game_path(find=False)
        if not game_path:
            return None
        choice_file = game_path.joinpath('l10n_installer/settings/choice.json')
        if os.path.isfile(choice_file):
            try:
                with open(choice_file, 'r', encoding='utf-8') as f:
                    self.choice = json.load(f)
            except Exception:
                pass
        self.check_choice()
        return self.choice

    def check_choice(self):
        template = self.get_choice_template()
        for choice in ['server_region', 'is_release', 'download_source', 'use_ee', 'apply_mods', 'isolation']:
            if choice not in self.choice.keys():
                self.choice[choice] = template[choice]

    def save_choice(self) -> None:
        if self.choice and self.game_path.get() != game_path_unknown:
            game_path = self.get_game_path(find=False)
            if not game_path:
                return
            self.choice['is_release'] = self.is_release.get()
            self.choice['download_source'] = self.download_source.get()
            self.choice['use_ee'] = self.ee_selection.get()
            self.choice['apply_mods'] = self.mods_selection.get()
            self.choice['isolation'] = self.isolation.get()
            with open(game_path.joinpath('l10n_installer/settings/choice.json'), 'w', encoding='utf-8') as f:
                json.dump(self.choice, f, ensure_ascii=False, indent=4)

    def on_closed(self):
        self.save_global_settings()


def mkdir(t_dir: Any):
    os.makedirs(t_dir, exist_ok=True)


def process_modification_file(source_mo, mod_path: str,
                              json_mods_d: Dict[str, Union[str, List[str]]], json_mods_m: Dict[str, str]) -> bool:
    translated = None
    is_json_mod = False
    if mod_path.endswith('po'):
        translated = polib.pofile(mod_path)
    elif mod_path.endswith('mo'):
        translated = polib.mofile(mod_path)
    elif mod_path.endswith('l10nmod') or mod_path.endswith('json'):
        is_json_mod = True

    if translated:
        translation_dict_singular: Dict[str, polib.MOEntry] = {entry.msgid: entry for entry in translated if
                                                               entry.msgid and not entry.msgid_plural}
        translation_dict_plural: Dict[str, polib.MOEntry] = {entry.msgid: entry for entry in translated if
                                                             entry.msgid and entry.msgid_plural}
        singular_count = len(translation_dict_singular)
        plural_count = len(translation_dict_plural)
        for entry in source_mo:
            if not entry.msgid:
                continue
            if singular_count == 0 and plural_count == 0:
                break
            if entry.msgid_plural and entry.msgid_plural in translation_dict_plural:
                target_strs = translation_dict_plural[entry.msgid].msgstr_plural
                del translation_dict_plural[entry.msgid]
                plural_count -= 1
                entry.msgstr_plural = target_strs
            elif entry.msgid and entry.msgid in translation_dict_singular:
                target_str = translation_dict_singular[entry.msgid].msgstr
                del translation_dict_singular[entry.msgid]
                singular_count -= 1
                if entry.msgid == 'IDS_RIGHTS_RESERVED':
                    continue
                entry.msgstr = target_str
        if singular_count > 0 or plural_count > 0:
            for s_e in translation_dict_singular:
                source_mo.append(translation_dict_singular[s_e])
            for p_e in translation_dict_plural:
                source_mo.append(translation_dict_plural[p_e])
    elif is_json_mod:
        try:
            with open(mod_path, 'r', encoding='utf-8') as f:
                json_mod = json.load(f)
            append_json_mod(json_mod, json_mods_d, json_mods_m)
        except Exception:
            pass
    return translated or is_json_mod


def append_json_mod(json_mod: Dict[str, Any],
                    json_mods_d: Dict[str, Union[str, List[str]]], json_mods_m: Dict[str, str]):
    if 'replace' in json_mod.keys():
        replaces = json_mod.get('replace')
        if isinstance(replaces, Dict):
            for r_k in replaces.keys():
                r_v = replaces[r_k]
                if isinstance(r_v, str) or isinstance(r_v, List):
                    json_mods_d[r_k] = r_v
    if 'words' in json_mod.keys():
        words = json_mod.get('words')
        if isinstance(words, Dict):
            for w_k in words.keys():
                w_v = words[w_k]
                if isinstance(w_v, str):
                    json_mods_m[w_k] = w_v


def process_json_mods(source_mo, json_mods_d_replace: Dict[str, Union[str, List[str]]],
                      json_mods_m_replace: Dict[str, str]):
    for entry in source_mo:
        if not entry.msgid:
            continue
        if entry.msgid_plural:
            for m_k in json_mods_m_replace:
                m_v = json_mods_m_replace[m_k]
                msgstrs: Dict[int, str] = entry.msgstr_plural
                should_modify = False
                for i in msgstrs.keys():
                    if m_k in msgstrs.get(i):
                        msgstrs[i] = msgstrs.get(i).replace(m_k, m_v)
                        should_modify = True
                if should_modify:
                    entry.msgstr_plural = msgstrs
            for d_k in json_mods_d_replace:
                if entry.msgid == d_k:
                    target_text = json_mods_d_replace[d_k]
                    if isinstance(target_text, str):
                        list_l = len(entry.msgstr_plural)
                        entry.msgstr_plural = {i: target_text for i in range(list_l)}
                    elif isinstance(target_text, List):
                        entry.msgstr_plural = {i: target_text[i] for i in range(len(target_text))}
        else:
            for m_k in json_mods_m_replace:
                m_v = json_mods_m_replace[m_k]
                msgstr = entry.msgstr
                if m_k in msgstr:
                    entry.msgstr = msgstr.replace(m_k, m_v)
            for d_k in json_mods_d_replace:
                if entry.msgid == d_k:
                    target_text = json_mods_d_replace[d_k]
                    if isinstance(target_text, str):
                        entry.msgstr = target_text


def is_valid_game_path(game_path: Path) -> bool:
    game_info_file = game_path.joinpath('game_info.xml')
    if not game_info_file.is_file() or not game_path.joinpath('bin').is_dir():
        return False
    try:
        game_info = ET.parse(game_info_file)
        game_id = game_info.find('.//game/id')
        if game_id is None:
            return False
        return 'WOWS' in game_id.text
    except Exception:
        return False


def is_valid_build_dir(build_dir: Path) -> bool:
    res_dir = Path(build_dir).joinpath('res')
    if not res_dir.is_dir():
        return False
    return os.path.isfile(res_dir.joinpath('locale_config.xml'))


def check_sha256(mo_path: Path, sha256: str):
    return get_sha256_for_mo(mo_path) == sha256


def find_all_drives() -> List[str]:
    return ['%s:/' % d for d in string.ascii_uppercase if os.path.exists('%s:' % d)]


def get_sha256_for_mo(mo_path: Path):
    with open(mo_path, 'rb') as file:
        return hashlib.sha256(file.read()).hexdigest()


# Returns whether the update was successful
def _install_update(
        gui: Union[LocalizationInstaller, LocalizationInstallerAuto],
        game_path: Path = None,
        is_release: bool = False,
        use_ee: bool = False,
        use_mods: bool = False,
        isolation: bool = False,
        download_src: str = '',
        server_region: str = 'ru'
) -> bool:
    full_gui = isinstance(gui, LocalizationInstaller)
    if full_gui:
        game_path = gui.get_game_path()
        is_release = gui.is_release.get()
        use_ee = gui.ee_selection.get()
        use_mods = gui.mods_selection.get()
        isolation = gui.isolation.get()
        download_src = gui.download_source.get()
        server_region = gui.server_region.get()
    gui.safely_set_install_progress(progress=0.0)
    if not game_path or not is_valid_game_path(game_path):
        if full_gui:
            gui.root.after(0, Messagebox.show_error, '游戏目录不可用，无法安装。', '安装汉化')
            gui.is_installing = False
        else:
            Messagebox.show_error('游戏目录不可用，无法更新汉化。', '自动更新')
        return False
    if full_gui:
        run_dirs = gui.run_dirs.keys()
    else:
        run_dirs = parse_game_version(None, game_path)[0]
    if not run_dirs or len(run_dirs) == 0:
        if full_gui:
            gui.root.after(0, Messagebox.show_error, '未发现游戏版本，无法安装。', '安装汉化')
            gui.is_installing = False
        else:
            Messagebox.show_error('未发现游戏版本，无法更新汉化。', '自动更新')
        return False
    for run_dir in run_dirs:
        target_path = game_path.joinpath('bin').joinpath(run_dir).joinpath('res_mods' if is_release else 'res')
        mkdir(target_path)
        if full_gui:
            gui.safely_set_install_progress_text('安装locale_config')
        if not is_release:
            old_cfg = target_path.joinpath('locale_config.xml')
            old_cfg_backup = target_path.joinpath('locale_config.xml.old')
            if not os.path.isfile(old_cfg_backup) and os.path.isfile(old_cfg):
                shutil.copy(old_cfg, old_cfg_backup)
            with open(old_cfg, 'w', encoding='utf-8') as file:
                file.write(locale_config)
        else:
            with open(target_path.joinpath('locale_config.xml'), 'w', encoding='utf-8') as file:
                file.write(locale_config)
    gui.safely_set_install_progress(progress=20.0)
    if full_gui:
        gui.safely_set_install_progress_text('安装locale_config——完成')

    proxies = {scheme: proxy for scheme, proxy in urllib.request.getproxies().items()}

    if is_release:
        # EE
        if full_gui:
            use_ee = gui.supports_ee() and gui.ee_selection.get()
        if use_ee:
            if full_gui:
                gui.safely_set_install_progress_text('安装体验增强包')
                gui.safely_set_download_progress_text('下载体验增强包——连接中')
            output_file = Path('l10n_installer').joinpath('downloads').joinpath('LK_EE.zip')
            ee_ready = False
            try:
                response = requests.get('https://gitee.com/localized-korabli/Korabli-LESTA-L10N/raw/main'
                                        '/BuiltInMods/LKExperienceEnhancement.zip', stream=True,
                                        proxies=proxies, timeout=5000)
                status = response.status_code
                if status == 200:
                    if full_gui:
                        gui.safely_set_download_progress_text('下载体验增强包——下载中')
                    with open(output_file, 'wb') as f:
                        for chunk in response.iter_content(chunk_size=1024):
                            if chunk:
                                f.write(chunk)
                    ee_ready = True
                    if full_gui:
                        gui.safely_set_download_progress_text('下载体验增强包——完成')
                elif full_gui:
                    gui.safely_set_download_progress_text(f'下载体验增强包——失败（{status}）')
            except requests.exceptions.RequestException:
                if full_gui:
                    gui.safely_set_download_progress_text('下载体验增强包——请求异常')
            if ee_ready:
                for run_dir in run_dirs:
                    target_path = game_path.joinpath('bin').joinpath(run_dir).joinpath(
                        'res_mods' if is_release else 'res')
                    with zipfile.ZipFile(output_file, 'r') as mo_zip:
                        process_possible_gbk_zip(mo_zip)
                        mo_zip.extractall(target_path)
                if full_gui:
                    gui.safely_set_install_progress_text('安装体验增强包——完成')
            elif full_gui:
                gui.safely_set_install_progress_text('安装体验增强包——失败')
    # 汉化包
    gui.safely_set_install_progress(progress=30.0)
    if full_gui:
        gui.safely_set_install_progress_text('安装汉化包')
    nothing_wrong = True
    # remote_version = ''
    # downloaded_mo = ''
    local_src = download_src == 'local'
    fetched_file = None
    if not local_src:
        if full_gui:
            gui.safely_set_download_progress_text('下载汉化包')
        download_link_base = download_routes['r' if is_release else 'pt'][download_src]['url']
        # Check and Fetch
        downloaded: Tuple = check_version_and_fetch_mo(None,
                                                       None if full_gui else parse_game_version(None, game_path)[1],
                                                       download_link_base, proxies)
        if downloaded[2] is True:
            return True
        fetched_file = downloaded[0]
        remote_version = downloaded[1]
    else:
        if full_gui:
            fetched_file = gui.mo_path.get()
        remote_version = 'local'
    if not fetched_file:
        if full_gui:
            gui.safely_set_install_progress(0.0)
            gui.safely_set_install_progress_text('安装汉化包——文件异常')
        Messagebox.show_error('选择本地文件作为汉化来源时，\n请手动启动安装器进行安装。')
        return False
    dir_total = len(run_dirs)
    dir_progress = 1
    for run_dir in run_dirs:
        execution_time = str(time.time_ns())
        target_path = game_path.joinpath('bin').joinpath(run_dir).joinpath('res_mods' if is_release else 'res')
        version_info_file: Optional[Path] = None
        if fetched_file.endswith('.zip'):
            extracted_path = Path('l10n_installer').joinpath('downloads').joinpath('extracted_mo')
            mkdir(extracted_path)
            info_fetched = False
            with zipfile.ZipFile(fetched_file, 'r') as mo_zip:
                process_possible_gbk_zip(mo_zip)
                info_files = [info for info in mo_zip.filelist if info.filename.endswith('version.info')]
                if info_files:
                    info_file_name = info_files[0].filename
                    mo_zip.extract(info_file_name, extracted_path)
                    version_info_file = extracted_path.joinpath(info_file_name)
                    info_fetched = True
                mo_files = [mo for mo in mo_zip.filelist if mo.filename.endswith('.mo')]
                if mo_files:
                    mo_file_name = mo_files[0].filename
                    mo_zip.extract(mo_file_name, extracted_path)
                    fetched_file = os.path.join(extracted_path, mo_file_name)
            if info_fetched and version_info_file and version_info_file.is_file():
                with open(version_info_file, 'r', encoding='utf-8') as f:
                    remote_version = f.readline()
        if not fetched_file.endswith('.mo') or not os.path.isfile(fetched_file):
            if full_gui:
                gui.safely_set_install_progress_text('安装汉化包——文件异常')
            nothing_wrong = False
            break
        if nothing_wrong:
            if full_gui:
                gui.safely_set_download_progress_text('下载汉化包——完成')
            mods = get_mods(use_mods, game_path, run_dir, game_path if isolation else Path('.'))
            fetched_file = parse_and_apply_mods(gui, fetched_file, mods, execution_time, dir_progress, dir_total)
            cache_path = 'l10n_installer/cache'
            if os.path.isdir(cache_path):
                shutil.rmtree(cache_path)
            if fetched_file == '':
                if full_gui:
                    gui.safely_set_install_progress_text('安装汉化包——文件损坏')
                nothing_wrong = False
                break
        if nothing_wrong:
            if full_gui:
                gui.safely_set_install_progress_text(f'安装汉化包——移动文件({dir_progress}/{dir_total})')
            mo_dir = target_path.joinpath('texts').joinpath(server_region).joinpath('LC_MESSAGES')
            mkdir(mo_dir)
            old_mo = mo_dir.joinpath('global.mo')
            old_mo_backup = mo_dir.joinpath('global.mo.old')
            if not is_release:
                if not os.path.isfile(old_mo_backup) and os.path.isfile(old_mo):
                    shutil.copy(old_mo, old_mo_backup)
            shutil.copy(fetched_file, old_mo)

            info_path = game_path.joinpath('bin').joinpath(run_dir).joinpath('l10n')
            mkdir(info_path)
            installation_info_file = info_path.joinpath('installation.info')
            with open(installation_info_file, 'w', encoding='utf-8') as f:
                f.writelines([
                    remote_version,
                    '\n',
                    str(old_mo.absolute()),
                    '\n',
                    get_sha256_for_mo(old_mo)
                ])
                try:
                    float(remote_version)
                except ValueError:
                    f.write(f'\n{time.time()}')
        dir_progress += 1
    if full_gui:
        gui.is_installing = False
    if nothing_wrong:
        gui.safely_set_install_progress(progress=100.0)
        if full_gui:
            gui.available_game_paths.append(gui.game_path.get())
            gui.save_global_settings()
            gui.save_choice()
    if full_gui:
        gui.safely_set_install_progress_text('完成！' if nothing_wrong else '失败！')
        parse_game_version(gui, gui.get_game_path())
        gui.root.after(0, gui.popup_result, nothing_wrong)
    return nothing_wrong


# Returns (output_file: str, remote_version: str, should_skip: bool)
def check_version_and_fetch_mo(
        gui: Optional[LocalizationInstaller],
        l10n_versions: Optional[List[str]],
        download_link_base: str,
        proxies: Dict
) -> (str, str, bool):
    full_gui = gui is not None
    remote_version: str = 'latest'
    if full_gui:
        gui.safely_set_download_progress_text('下载汉化包——获取版本')
    try:
        response = requests.get(download_link_base + 'version.info', stream=True, proxies=proxies, timeout=5000)
        status = response.status_code
        if status == 200:
            info_file = 'l10n_installer/downloads/version.info'
            with open(info_file, 'wb') as f:
                for chunk in response.iter_content(chunk_size=1024):
                    if chunk:
                        f.write(chunk)
            with open(info_file, 'r', encoding='utf-8') as f:
                remote_version = f.readline().strip()
                if full_gui:
                    gui.safely_set_download_progress_text(f'下载汉化包——最新={remote_version}')
                elif compare_with_local(remote_version, l10n_versions):
                    return '', '', True

    except requests.exceptions.RequestException:
        pass
    valid_version = remote_version != 'latest'
    if not valid_version:
        if full_gui:
            gui.safely_set_download_progress_text(f'下载汉化包——版本获取失败')
    mo_file_name = f'{remote_version}.mo'
    output_file = f'l10n_installer/downloads/{mo_file_name}'
    # Check existed
    if full_gui:
        if valid_version and gui.last_installed_l10n_version == remote_version:
            if os.path.isfile(output_file):
                try:
                    if polib.mofile(output_file):
                        gui.safely_set_download_progress_text(f'下载汉化包——使用已下载文件')
                        return output_file, remote_version, False
                except Exception:
                    pass
    # Download from remote
    output_file = download_mo_from_remote(download_link_base + mo_file_name, output_file, proxies)
    if valid_version and output_file == '':
        # valid_version = False
        remote_version = 'latest'
        mo_file_name = f'{remote_version}.mo'
        output_file = f'l10n_installer/downloads/{mo_file_name}'
        output_file = download_mo_from_remote(download_link_base + mo_file_name, output_file, proxies)
    return output_file, remote_version, False


# Returns False if update needed
def compare_with_local(remote_version: str, local_versions: Optional[List[str]]):
    if not local_versions:
        return False
    try:
        rv = remote_version.split('.')
        rgv = int(rv[0])
        rlv = int(rv[1])
    except ValueError:
        return False
    for local_version in local_versions:
        try:
            lv = local_version.split('.')
            lgv = int(lv[0])
            llv = int(lv[1])
        except ValueError:
            return False
        if lgv < rgv:
            return False
        if llv < rlv:
            return False
    return True


def get_mods(mods_selection: bool, game_path: Path, run_dir: str, instance_dir: Path) -> List[str]:
    if not mods_selection:
        return []
    installer_mods_dir = instance_dir.joinpath('l10n_installer').joinpath('mods')
    compat_mods_dir_0 = game_path.joinpath('bin').joinpath(run_dir).joinpath('res_mods') \
        .joinpath('texts').joinpath('l10n_mods')
    compat_mods_dir_1 = game_path.joinpath('bin').joinpath(run_dir).joinpath('res_mods') \
        .joinpath('texts').joinpath('mods')
    if not installer_mods_dir.is_dir() and not compat_mods_dir_0.is_dir() and not compat_mods_dir_1.is_dir():
        return []
    file_list = []
    mkdir('l10n_installer/cache')
    scan_mods(file_list, installer_mods_dir)
    scan_mods(file_list, compat_mods_dir_0)
    scan_mods(file_list, compat_mods_dir_1)
    return file_list


def scan_mods(file_list: List[str], mods_dir: Path) -> None:
    extracted_list = []
    for root0, _, files0 in os.walk(mods_dir):
        for name0 in files0:
            if name0.endswith(('.mo', '.po', '.json', '.l10nmod')):
                file_list.append(os.path.abspath(os.path.join(root0, name0)))
            elif name0.endswith('.zip'):
                try:
                    mod_name = os.path.basename(name0).replace('.zip', '')
                    extracted_cache = f'l10n_installer/cache/{mod_name}'
                    with zipfile.ZipFile(os.path.join(root0, name0), 'r') as mod_zip:
                        process_possible_gbk_zip(mod_zip)
                        mod_files = [mod_file for mod_file in mod_zip.filelist if
                                     mod_file.filename.split('/')[-1].endswith(('.mo', '.po', '.json', '.l10nmod'))]
                        for mod_file in mod_files:
                            mod_zip.extract(mod_file, extracted_cache)

                    extracted_list.append(extracted_cache)
                except Exception as ex:
                    print(ex)
        for extracted_cache in extracted_list:
            for root1, _, files1 in os.walk(extracted_cache):
                for name1 in files1:
                    if name1.endswith(('.mo', '.po', '.json', '.l10nmod')):
                        file_list.append(os.path.abspath(os.path.join(root1, name1)))


def parse_and_apply_mods(gui: Union[LocalizationInstaller, LocalizationInstallerAuto], downloaded_mo: str,
                         mods: List[str],
                         execution_time: str,
                         dir_progress: int,
                         dir_total: int) -> str:
    full_gui = isinstance(gui, LocalizationInstaller)
    try:
        downloaded_mo_instance = polib.mofile(downloaded_mo)
    except Exception:
        return ''
    mods_count = len(mods)
    if mods_count == 0:
        gui.safely_set_install_progress(90.0)
        return downloaded_mo
    if full_gui:
        gui.safely_set_install_progress_text(f'安装汉化包——应用模组({str(dir_progress)}/{str(dir_total)})')
    modded_file_name = f'l10n_installer/processed/modified_{execution_time}.mo'
    if os.path.isfile(modded_file_name):
        return modded_file_name
    applied_mods = 0
    json_mods_d_replace: Dict[str, Union[str, List[str]]] = {}
    json_mods_m_replace: Dict[str, str] = {}
    for mod in mods:
        try:
            applied = process_modification_file(downloaded_mo_instance, mod, json_mods_d_replace, json_mods_m_replace)
        except Exception:
            pass
        if applied:
            applied_mods += 1
            gui.safely_set_install_progress(
                30.0 + 60.0 * ((dir_progress - 1) / dir_total + applied_mods / (mods_count * dir_total))
            )
    process_json_mods(downloaded_mo_instance, json_mods_d_replace, json_mods_m_replace)

    for file in os.listdir('l10n_installer/processed/'):
        try:
            os.remove(file)
        except Exception:
            continue
    downloaded_mo_instance.save(modded_file_name)
    return modded_file_name


def download_mo_from_remote(download_link: str, output_file: str, proxies: Dict) -> str:
    try:
        response = requests.get(download_link, stream=True, proxies=proxies, timeout=5000)
        status = response.status_code
        if status == 200:
            with open(output_file, 'wb') as f:
                for chunk in response.iter_content(chunk_size=1024):
                    if chunk:
                        f.write(chunk)
            return output_file
    except requests.exceptions.RequestException:
        return ''


# Returns (run_dirs: List[str], installed_l10ns: List[str])
def parse_game_version(gui: Optional[LocalizationInstaller], game_path: Optional[Path]) -> Tuple[List[str], List[str]]:
    gui_on = gui is not None
    if gui_on:
        gui.run_dirs = {}
        gui.localization_status_1st.set('未发现游戏版本')
        gui.localization_status_2nd.set('')
    v_1st = ''
    v_2nd = ''
    if gui_on:
        game_path = gui.get_game_path()
    if not game_path:
        return [], []
    bin_path = game_path.joinpath('bin')
    if not bin_path.is_dir():
        return [], []
    for v_dir_b in os.listdir(bin_path):
        v_dir = str(v_dir_b)
        # v_dir_num = 0
        try:
            v_dir_num = int(v_dir)
            if not is_valid_build_dir(bin_path.joinpath(v_dir)):
                continue
        except ValueError:
            continue
        if not v_1st:
            v_1st = v_dir
        elif int(v_1st) > v_dir_num:
            if not v_2nd or int(v_2nd) < v_dir_num:
                v_2nd = v_dir
        else:
            v_2nd = v_1st
            v_1st = v_dir
    if not v_1st:
        return [], []
    _1st_l10n_status = get_local_l10n_version(gui, game_path, v_1st)
    if gui_on:
        gui.run_dirs[v_1st] = _1st_l10n_status[0]
        gui.localization_status_1st.set(_1st_l10n_status[1])
    if v_2nd:
        _2nd_l10n_status = get_local_l10n_version(gui, game_path, v_2nd)
        if gui_on:
            gui.run_dirs[v_2nd] = _2nd_l10n_status[0]
            gui.localization_status_2nd.set(_2nd_l10n_status[1])
        return [v_1st, v_2nd], [_1st_l10n_status[0], _2nd_l10n_status[0]]
    return [v_1st], [_1st_l10n_status[0]]


# 返回：(汉化版本号: str, 汉化状态: str)
def get_local_l10n_version(gui: Optional[LocalizationInstaller], game_path: Path, run_dir: str) -> (str, str):
    gui_on = gui is not None
    installation_info_file = game_path.joinpath('bin').joinpath(run_dir).joinpath('l10n') \
        .joinpath('installation.info')
    if not installation_info_file.is_file():
        return '', f'{run_dir}——未安装汉化'
    with open(installation_info_file, 'r', encoding='utf-8') as f:
        parsed_version = f.readline().strip()
        mo_path = Path(f.readline().strip())
        if not mo_path.is_file():
            return '', f'{run_dir}——未安装汉化'
        mo_sha256 = f.readline().strip()
    if gui_on:
        gui.last_installed_l10n_version = parsed_version
    not_parsable = False
    to_return = parsed_version, f'{run_dir}——{parsed_version}'
    try:
        float(parsed_version)
    except ValueError:
        not_parsable = True
    if not_parsable:
        try:
            inst_time = float(f.readline())
            time_formatted = datetime.fromtimestamp(inst_time).strftime('%Y-%m-%d %H:%M')
            to_return = '', f'{run_dir}——汉化版本未知，于{time_formatted}安装'
        except ValueError:
            to_return = '', f'{run_dir}——汉化版本未知，安装时间未知'

    return to_return if check_sha256(mo_path, mo_sha256) else ('', to_return[1] + '（被篡改）')


def get_str_from_optional_path(target_path: Optional[Path], prefix: str = '') -> str:
    if not target_path:
        return ''
    return prefix + str(target_path.absolute())


# Returns (launcher_file: Path, launcher_status: str)
def find_launcher(game_path: Optional[Path]) -> (Optional[Path], str):
    if game_path:
        for launcher in launcher_dict.keys():
            launcher_file = game_path.joinpath(launcher)
            if launcher_file.is_file():
                return launcher_file, launcher_dict.get(launcher)
    return None, '未找到客户端'


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False


def run():
    parser = OptionParser()
    parser.add_option('--auto', dest='auto', action='store_true', default=False)
    parser.add_option('--nogui', dest='no_gui', action='store_true', default=False)
    parser.add_option('--norun', dest='no_run', action='store_true', default=False)
    parser.add_option('--gamepath', dest='game_path')
    parser.add_option('--release', dest='is_release', action='store_true', default=False)
    parser.add_option('--ee', dest='use_ee', action='store_true', default=False)
    parser.add_option('--mods', dest='use_mods', action='store_true', default=False)
    parser.add_option('--isolation', dest='isolation', action='store_true', default=False)
    parser.add_option('--src', dest='download_src')
    parser.add_option('--region', dest='server_region')
    options, _ = parser.parse_args()

    if options.auto is False:
        root = ttk.Window()
        icon = os.path.join(resource_path, 'icon.ico')
        root.iconbitmap(default=icon)
        root.iconbitmap(bitmap=icon)
        configure_font()
        half_screen_width = int(root.winfo_screenwidth() / 2) - 234
        half_screen_height = int(root.winfo_screenheight() / 2) - 359
        root.geometry(f'+{half_screen_width}+{half_screen_height}')
        app = LocalizationInstaller(root)
        root.mainloop()
        app.on_closed()
    else:
        root = ttk.Window()
        icon = os.path.join(resource_path, 'icon.ico')
        root.iconbitmap(default=icon)
        root.iconbitmap(bitmap=icon)
        configure_font()
        scr_width = 800
        scr_height = 100
        half_screen_width = int((root.winfo_screenwidth() - scr_width) / 2)
        half_screen_height = int((root.winfo_screenheight() - scr_height) / 2)
        root.geometry(f'{scr_width}x{scr_height}+{half_screen_width}+{half_screen_height}')
        app = LocalizationInstallerAuto(root, options)
        root.mainloop()
        app.on_closed()


def configure_font():
    font_list = list(font.families())
    if 'SimHei' in font_list:
        do_configure_font('SimHei')
    elif '黑体' in font_list:
        do_configure_font('黑体')
    elif 'DengXian' in font_list:
        do_configure_font('DengXian')
    elif '等线' in font_list:
        do_configure_font('等线')


def do_configure_font(family: str):
    ttk.font.nametofont('TkDefaultFont').configure(family=family)
    ttk.font.nametofont('TkTextFont').configure(family=family)
    ttk.font.nametofont('TkFixedFont').configure(family=family)


def process_possible_gbk_zip(zip_file: zipfile.ZipFile):
    name_to_info = zip_file.NameToInfo
    for name, info in name_to_info.copy().items():
        real_name = name.encode('cp437').decode('gbk')
        if real_name != name:
            info.filename = real_name
            del name_to_info[name]
            name_to_info[real_name] = info
    return zip_file


if __name__ == '__main__':
    dev_env = sys.executable.endswith('python.exe')
    if dev_env:
        run()
    else:
        os.chdir(Path(sys.executable).parent)
        if is_admin():
            run()
        else:
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv[1:]), None, 1)

# pyinstaller -w -i resources/icon.ico --onefile --add-data "resources\*;resources" --version-file=version_file.txt installer_gui.py --clean

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
from pathlib import Path
from tkinter import filedialog
from typing import Any, Dict, List, Tuple, Optional

import polib
import requests
import ttkbootstrap as ttk
from ttkbootstrap.dialogs.dialogs import Messagebox

mods_link = 'https://tapio.lanzn.com/b0nxzso2b'
project_repo_link = 'https://github.com/LocalizedKorabli/Korabli-LESTA-L10N/'
installer_repo_link = 'https://github.com/LocalizedKorabli/L10nInstallerGUI/'

version = '0.1.0'

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
msg_game_path_may_be_invalid = '''您选择的游戏目录缺失必要的游戏文件，
可能并非战舰世界安装目录。'''


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
    mod_selection: tk.BooleanVar
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

    def __init__(self, parent: tk.Tk):
        self.root = parent
        self.root.title(f'汉化安装器v{version}')

        self.localization_status_1st = tk.StringVar()
        self.localization_status_2nd = tk.StringVar()
        self.server_region = tk.StringVar()
        self.is_release = tk.BooleanVar()
        self.download_source = tk.StringVar()
        self.ee_selection = tk.BooleanVar()
        self.mod_selection = tk.BooleanVar()
        self.mo_path = tk.StringVar()
        self.install_progress_text = tk.StringVar()
        self.game_launcher_status = tk.StringVar()
        self.download_progress_text = tk.StringVar()
        self.install_progress = tk.DoubleVar()
        self.game_path = tk.StringVar()

        # 游戏目录
        ttk.Label(parent, text='游戏目录：') \
            .grid(row=0, column=0, columnspan=1, pady=5, sticky=tk.W)

        self.game_path_combo = ttk.Combobox(root, width=26, textvariable=self.game_path, state='readonly')
        self.refresh_path_combo()

        self.game_path_combo.grid(row=0, column=1, columnspan=3, padx=5, pady=5, sticky=tk.W)

        self.auto_search = ttk.Button(parent, text='自动检测', command=lambda: self.find_game(overwrite=False),
                                      style='success')
        self.auto_search.grid(row=1, column=2, columnspan=1)

        self.game_path_button = ttk.Button(parent, text='选择目录', command=self.choose_path, style='warning')
        self.game_path_button.grid(row=1, column=3, columnspan=1)

        # 游戏版本
        ttk.Label(parent, text='游戏版本/汉化版本') \
            .grid(row=1, column=0, columnspan=4, pady=5, sticky=tk.W)

        # 汉化状态
        self.localization_status_label_1st = ttk.Label(parent, textvariable=self.localization_status_1st)
        self.localization_status_label_1st.grid(row=3, column=0, columnspan=4, pady=5, sticky=tk.W)
        self.localization_status_label_2nd = ttk.Label(parent, textvariable=self.localization_status_2nd)
        self.localization_status_label_2nd.grid(row=4, column=0, columnspan=4, pady=5, sticky=tk.W)
        self.parse_game_version()

        # 游戏区服
        ttk.Label(parent, text='游戏区服：').grid(row=5, column=0, pady=5, sticky=tk.W)

        # 游戏区服选项
        ttk.Radiobutton(parent, text='莱服', variable=self.server_region, value='ru') \
            .grid(row=5, column=1, sticky=tk.W)
        ttk.Radiobutton(parent, text='直营服', variable=self.server_region, value='zh_sg', style='warning') \
            .grid(row=5, column=2, sticky=tk.W)
        ttk.Radiobutton(parent, text='国服', variable=self.server_region, value='zh_cn', style='danger') \
            .grid(row=5, column=3, sticky=tk.W)

        # 游戏类型
        ttk.Label(parent, text='游戏类型：').grid(row=6, column=0, pady=5, sticky=tk.W)

        # 游戏类型选项
        ttk.Radiobutton(parent, text='正式服', variable=self.is_release, value=True, style='success') \
            .grid(row=6, column=1, sticky=tk.W)
        ttk.Radiobutton(parent, text='PT服', variable=self.is_release, value=False, style='danger') \
            .grid(row=6, column=2, sticky=tk.W)

        self.detect_game_type_button = ttk.Button(parent, text='自动检测',
                                                  command=lambda: self.detect_game_status(manually=True))
        self.detect_game_type_button.grid(row=6, column=3)

        # 下载源
        ttk.Label(parent, text='汉化来源：').grid(row=7, column=0, pady=5, sticky=tk.W)
        # 下载源选项
        ttk.Radiobutton(parent, text='Gitee', variable=self.download_source, value='gitee', style='danger') \
            .grid(row=7, column=1, sticky=tk.W)
        ttk.Radiobutton(parent, text='GitHub', variable=self.download_source, value='github', style='dark') \
            .grid(row=7, column=2, sticky=tk.W)
        ttk.Radiobutton(parent, text='本地文件', variable=self.download_source, value='local') \
            .grid(row=7, column=3, sticky=tk.W)

        # 体验增强包/汉化修改包
        self.ee_button = ttk.Checkbutton(parent, text='安装体验增强包', variable=self.ee_selection)
        self.ee_button.grid(row=9, column=0, columnspan=2, pady=5, sticky=tk.W)
        ttk.Checkbutton(parent, text='安装模组（汉化修改包）', variable=self.mod_selection) \
            .grid(row=30, column=0, columnspan=2, pady=5, sticky=tk.W)
        self.mods_button = ttk.Button(parent, text='模组目录', command=self.open_mods_folder)
        self.mods_button.grid(row=30, column=2, columnspan=1)
        self.download_mods_button = ttk.Button(parent, text='下载模组', style='info',
                                               command=lambda: webbrowser.open_new_tab(mods_link))
        self.download_mods_button.grid(row=30, column=3, columnspan=1)

        # 安装路径选择/下载进度
        self.install_path_entry = ttk.Entry(parent, textvariable=self.mo_path, width=20)
        self.install_path_button = ttk.Button(parent, text='选择文件', command=self.choose_mo)
        self.download_progress_label = ttk.Label(parent, text='下载进度：')
        self.download_progress_info = ttk.Label(parent, textvariable=self.download_progress_text)

        # 安装/更新按钮
        self.install_button = ttk.Button(parent, text='安装汉化', command=self.install_update,
                                         style=ttk.SUCCESS)
        self.install_button.grid(row=31, column=0, pady=5)

        # 安装进度
        ttk.Label(parent, textvariable=self.install_progress_text).grid(row=31, column=1, columnspan=3,
                                                                        padx=5, sticky=tk.W)

        self.install_progress_bar = ttk.Progressbar(parent, variable=self.install_progress, maximum=100.0,
                                                    style='success-striped', length=400)
        self.install_progress_bar.grid(row=32, column=0, columnspan=4, padx=10)

        # 启动游戏
        self.launch_button = ttk.Button(parent, text='启动游戏', command=self.launch_game, style=ttk.WARNING)
        self.launch_button.grid(row=33, column=0, pady=5)

        # 启动器状态
        ttk.Label(parent, textvariable=self.game_launcher_status).grid(row=33, column=1, columnspan=3,
                                                                       padx=5, sticky=tk.W)

        # 相关链接
        about_button = ttk.Button(parent, text='关于项目', command=lambda: webbrowser.open_new_tab(project_repo_link),
                                  style=ttk.INFO)
        about_button.grid(row=34, column=0, pady=5)

        src_button = ttk.Button(parent, text='代码仓库', command=lambda: webbrowser.open_new_tab(installer_repo_link),
                                style=ttk.DANGER)
        src_button.grid(row=34, column=1, pady=5, padx=5)

        # 版权声明
        ttk.Label(parent, text='© 2024 LocalizedKorabli').grid(row=34, column=2, columnspan=3, pady=5)

        # 根据下载源选项显示或隐藏安装路径选择
        self.download_source.trace('w', self.on_download_source_changed)
        # 更换游戏路径时，刷新数据
        self.game_path.trace('w', self.on_game_path_changed)
        # 非俄服客户端无需安装体验增强包
        self.server_region.trace('w', self.on_server_region_or_game_type_changed)
        self.is_release.trace('w', self.on_server_region_or_game_type_changed)

        mkdir('l10n_installer/cache')
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

        self.safely_set_download_progress_text('准备')
        self.safely_set_install_progress_text('准备')
        self.safely_set_install_progress(progress=0.0)
        self.game_launcher_status.set(self.find_launcher())

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

    def on_game_path_changed(self, *args) -> None:
        if self.game_path.get() == game_path_unknown:
            return

        game_path = self.get_game_path()
        if not game_path:
            return

        self.detect_game_status()
        self.parse_game_version()
        self.find_launcher()

        mkdir(game_path.joinpath('l10n_installer/settings'))
        mkdir(game_path.joinpath('l10n_installer/mods'))

        choice = self.parse_choice(use_cache=False)

        self.server_region.set(choice.get('server_region', 'ru'))
        self.is_release.set(choice.get('is_release', True))
        self.download_source.set(choice.get('download_source', 'gitee'))
        self.ee_selection.set(choice.get('use_ee', True))
        self.mod_selection.set(choice.get('apply_mods', True))

    def on_server_region_or_game_type_changed(self, *args):
        self.ee_button.configure(state=('' if self.supports_ee() else 'disabled'))

    def supports_ee(self):
        return self.server_region.get() == 'ru' and self.is_release.get()

    def find_game(self, overwrite: bool = True) -> Optional[Path]:
        found_in_reg = self.find_from_reg()
        found_manually = self.find_manually()
        game_path_str = self.game_path.get()
        if not overwrite:
            return None
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
                path_strs_filtered = [dir_str for dir_str in path_strs if
                                      'Tank' not in dir_str and 'GameCheck' not in dir_str]
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

    def on_download_source_changed(self, *args):
        if self.download_source.get() == 'local':
            self.install_path_entry.grid(row=8, column=0, columnspan=3)
            self.install_path_button.grid(row=8, column=3)
            self.download_progress_label.grid_forget()
            self.download_progress_info.grid_forget()
        else:
            self.download_progress_label.grid(row=8, column=0, pady=5, sticky=tk.W)
            self.download_progress_info.grid(row=8, column=1, pady=5, columnspan=3, sticky=tk.W)
            self.install_path_entry.grid_forget()
            self.install_path_button.grid_forget()

    def open_mods_folder(self):
        mods_folder = Path('l10n_installer').joinpath('mods')
        mkdir(mods_folder)
        subprocess.run(['explorer', mods_folder.absolute()])

    def choose_path(self):
        game_path_chosen = filedialog.askdirectory(initialdir='.')
        if game_path_chosen:
            if not is_valid_game_path(Path(game_path_chosen)):
                Messagebox.show_warning(msg_game_path_may_be_invalid)
            else:
                self.available_game_paths.append(game_path_chosen)
                self.refresh_path_combo()
            self.game_path.set(game_path_chosen)

    def choose_mo(self):
        mo_path = filedialog.askopenfilename(initialdir='.', filetypes=[('汉化包', ['*.mo', '*.zip']),
                                                                        ('MO汉化文件', '*.mo'),
                                                                        ('打包的汉化文件', '*.zip')])
        if mo_path:
            self.mo_path.set(mo_path)

    def install_update(self):
        if self.is_installing:
            Messagebox.show_warning('安装已在进行！', '安装汉化')
            return
        self.is_installing = True
        tr = threading.Thread(target=self.do_install_update)
        tr.start()

    def do_install_update(self) -> None:
        self.safely_set_install_progress(progress=0.0)
        game_path = self.get_game_path()
        if not is_valid_game_path(game_path):
            self.root.after(0, Messagebox.show_error, '游戏目录不可用，无法安装。', '安装汉化')
            self.is_installing = False
            return
        run_dirs = self.run_dirs.keys()
        if len(run_dirs) == 0:
            self.root.after(0, Messagebox.show_error, '未发现游戏版本，无法安装。', '安装汉化')
            self.is_installing = False
            return
        is_release = self.is_release.get()
        for run_dir in run_dirs:
            target_path = game_path.joinpath('bin').joinpath(run_dir).joinpath('res_mods' if is_release else 'res')
            mkdir(target_path)
            self.safely_set_install_progress_text('安装locale_config')
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
        self.safely_set_install_progress_text('安装locale_config——完成')
        self.safely_set_install_progress(progress=20.0)
        proxies = {scheme: proxy for scheme, proxy in urllib.request.getproxies().items()}
        if is_release:
            # EE
            if self.supports_ee() and self.ee_selection.get():
                self.safely_set_install_progress_text('安装体验增强包')
                output_file = 'l10n_installer/downloads/LK_EE.zip'
                self.safely_set_download_progress_text('下载体验增强包——连接中')
                ee_ready = False
                try:
                    response = requests.get('https://gitee.com/localized-korabli/Korabli-LESTA-L10N/raw/main'
                                            '/BuiltInMods/LKExperienceEnhancement.zip', stream=True, proxies=proxies)
                    status = response.status_code
                    if status == 200:
                        self.safely_set_download_progress_text('下载体验增强包——下载中')
                        with open(output_file, 'wb') as f:
                            for chunk in response.iter_content(chunk_size=1024):
                                if chunk:
                                    f.write(chunk)
                        ee_ready = True
                        self.safely_set_download_progress_text('下载体验增强包——完成')
                    else:
                        self.safely_set_download_progress_text(f'下载体验增强包——失败（{status}）')
                except requests.exceptions.RequestException:
                    self.safely_set_download_progress_text('下载体验增强包——请求异常')
                if ee_ready:
                    for run_dir in run_dirs:
                        target_path = game_path.joinpath('bin').joinpath(run_dir).joinpath(
                            'res_mods' if is_release else 'res')
                        with zipfile.ZipFile(output_file, 'r') as mo_zip:
                            mo_zip.extractall(target_path)
                    self.safely_set_install_progress_text('安装体验增强包——完成')
                else:
                    self.safely_set_install_progress_text('安装体验增强包——失败')
        # 汉化包
        self.safely_set_install_progress(progress=30.0)
        self.safely_set_install_progress_text('安装汉化包')
        download_src = self.download_source.get()
        nothing_wrong = True
        # remote_version = ''
        # downloaded_mo = ''
        local_src = download_src == 'local'
        if not local_src:
            self.safely_set_download_progress_text('下载汉化包')
            download_link_base = download_routes['r' if is_release else 'pt'][download_src]['url']
            # Check and Fetch
            downloaded: Tuple = self.check_version_and_fetch_mo(download_link_base, proxies)
            downloaded_file = downloaded[0]
            remote_version = downloaded[1]
        else:
            downloaded_file = self.mo_path.get()
            remote_version = 'local'
        execution_time = str(time.time_ns())
        for run_dir in run_dirs:
            target_path = game_path.joinpath('bin').joinpath(run_dir).joinpath('res_mods' if is_release else 'res')
            info_path = game_path.joinpath('bin').joinpath(run_dir).joinpath('l10n')
            mkdir(info_path)
            version_info_file = info_path.joinpath('version.info')
            if downloaded_file.endswith('.zip'):
                extracted_path = Path('l10n_installer').joinpath('downloads').joinpath('extracted_mo')
                mkdir(extracted_path)
                info_fetched = False
                with zipfile.ZipFile(downloaded_file, 'r') as mo_zip:
                    info_files = [info for info in mo_zip.filelist if info.filename.split('/')[-1] == 'version.info']
                    if info_files:
                        info_file_name = info_files[0].filename
                        mo_zip.extract(info_file_name, info_path)
                        info_fetched = True
                    mo_files = [mo for mo in mo_zip.filelist if mo.filename.endswith('.mo')]
                    if mo_files:
                        mo_file_name = mo_files[0].filename
                        mo_zip.extract(mo_file_name, extracted_path)
                        downloaded_file = os.path.join(extracted_path, mo_file_name)
                if info_fetched and version_info_file.is_file():
                    with open(version_info_file, 'r', encoding='utf-8') as f:
                        remote_version = f.readline()
            if not downloaded_file.endswith('.mo') or not os.path.isfile(downloaded_file):
                self.safely_set_install_progress_text('安装汉化包——文件异常')
                nothing_wrong = False
                break
            if nothing_wrong:
                self.safely_set_download_progress_text('下载汉化包——完成')
                mods = self.get_mods()
                downloaded_file = self.parse_and_apply_mods(downloaded_file, mods, execution_time)
                if downloaded_file == '':
                    self.safely_set_install_progress_text('安装汉化包——文件损坏')
                    nothing_wrong = False
                    break
            if nothing_wrong:
                self.safely_set_install_progress_text('安装汉化包——正在移动文件')
                mo_dir = target_path.joinpath('texts').joinpath(self.server_region.get()).joinpath('LC_MESSAGES')
                mkdir(mo_dir)
                old_mo = mo_dir.joinpath('global.mo')
                old_mo_backup = mo_dir.joinpath('global.mo.old')
                if not is_release:
                    if not os.path.isfile(old_mo_backup) and os.path.isfile(old_mo):
                        shutil.copy(old_mo, old_mo_backup)
                shutil.copy(downloaded_file, old_mo)
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
        self.is_installing = False
        if nothing_wrong:
            self.safely_set_install_progress(progress=100.0)
            self.available_game_paths.append(self.game_path.get())
            self.save_global_settings()
            self.save_choice()
        self.safely_set_install_progress_text('完成！' if nothing_wrong else '失败！')
        self.parse_game_version()
        self.root.after(0, self.popup_result, nothing_wrong)

    # Returns (output_file: str, remote_version: str)
    def check_version_and_fetch_mo(self, download_link_base: str, proxies: Dict) -> (str, str):
        remote_version: str = 'latest'
        self.safely_set_download_progress_text('下载汉化包——获取版本')
        try:
            response = requests.get(download_link_base + 'version.info', stream=True, proxies=proxies)
            status = response.status_code
            if status == 200:
                info_file = 'l10n_installer/downloads/version.info'
                with open(info_file, 'wb') as f:
                    for chunk in response.iter_content(chunk_size=1024):
                        if chunk:
                            f.write(chunk)
                with open(info_file, 'r', encoding='utf-8') as f:
                    remote_version = f.readline()
                    self.safely_set_download_progress_text(f'下载汉化包——最新={remote_version}')
        except requests.exceptions.RequestException:
            pass
        valid_version = remote_version != 'latest'
        if not valid_version:
            self.safely_set_download_progress_text(f'下载汉化包——版本获取失败')
        mo_file_name = f'{remote_version}.mo'
        output_file = f'l10n_installer/downloads/{mo_file_name}'
        # Check existed
        if valid_version and self.last_installed_l10n_version == remote_version:
            if os.path.isfile(output_file):
                try:
                    if polib.mofile(output_file):
                        self.safely_set_download_progress_text(f'下载汉化包——使用已下载文件')
                        return output_file, remote_version
                except Exception:
                    pass
        # Download from remote
        output_file = self.download_mo_from_remote(download_link_base + mo_file_name, output_file, proxies)
        if valid_version and output_file == '':
            # valid_version = False
            remote_version = 'latest'
            mo_file_name = f'{remote_version}.mo'
            output_file = f'l10n_installer/downloads/{mo_file_name}'
            output_file = self.download_mo_from_remote(download_link_base + mo_file_name, output_file, proxies)
        return output_file, remote_version

    def download_mo_from_remote(self, download_link: str, output_file: str, proxies: Dict) -> str:
        try:
            response = requests.get(download_link, stream=True, proxies=proxies)
            status = response.status_code
            if status == 200:
                with open(output_file, 'wb') as f:
                    for chunk in response.iter_content(chunk_size=1024):
                        if chunk:
                            f.write(chunk)
                return output_file
        except requests.exceptions.RequestException:
            return ''

    def parse_and_apply_mods(self, downloaded_mo: str, mods: List[str], execution_time: str) -> str:
        try:
            downloaded_mo_instance = polib.mofile(downloaded_mo)
        except Exception:
            return ''
        mods_count = len(mods)
        if mods_count == 0:
            self.safely_set_install_progress(90.0)
            return downloaded_mo
        self.safely_set_install_progress_text('安装汉化包——正在应用模组')
        modded_file_name = f'l10n_installer/processed/modified_{execution_time}.mo'
        if os.path.isfile(modded_file_name):
            return modded_file_name
        applied_mods = 0
        for mod in mods:
            try:
                process_modification_file(downloaded_mo_instance, mod)
            except Exception:
                pass
            applied_mods += 1
            self.safely_set_install_progress(30.0 + 60.0 * applied_mods / mods_count)
        for file in os.listdir('l10n_installer/processed/'):
            try:
                os.remove(file)
            except Exception:
                continue
        downloaded_mo_instance.save(modded_file_name)
        return modded_file_name

    def get_mods(self) -> List[str]:
        if not self.mod_selection.get() or not Path('l10n_installer/mods').is_dir():
            return []
        files = os.listdir('l10n_installer/mods')
        return [('l10n_installer/mods/' + file) for file in files if (file.endswith('.po') or file.endswith('.mo'))]

    def parse_game_version(self) -> None:
        self.run_dirs = {}
        self.localization_status_1st.set('未发现游戏版本')
        self.localization_status_2nd.set('')
        v_1st = ''
        v_2nd = ''
        game_path = self.get_game_path()
        if not game_path:
            return
        bin_path = game_path.joinpath('bin')
        if not bin_path.is_dir():
            return
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
            return
        _1st_l10n_status = self.get_local_l10n_version(v_1st)
        self.run_dirs[v_1st] = _1st_l10n_status[0]
        self.localization_status_1st.set(_1st_l10n_status[1])
        if v_2nd:
            _2nd_l10n_status = self.get_local_l10n_version(v_2nd)
            self.run_dirs[v_2nd] = _2nd_l10n_status[0]
            self.localization_status_2nd.set(_2nd_l10n_status[1])
        if not self.localization_status_2nd.get():
            self.localization_status_label_2nd.grid_forget()
        else:
            self.localization_status_label_2nd.grid(row=4, column=0, columnspan=4, pady=5, sticky=tk.W)

    def get_choice_template(self):
        self.detect_game_status()
        return {
            'server_region': self.server_region.get(),
            'is_release': self.is_release.get(),
            'download_source': 'gitee',
            'use_ee': True,
            'apply_mods': True
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

    # 返回：(汉化版本号: str, 汉化状态: str)
    def get_local_l10n_version(self, run_dir: str) -> (str, str):
        installation_info_file = self.get_game_path().joinpath('bin').joinpath(run_dir).joinpath('l10n') \
            .joinpath('installation.info')
        if not os.path.isfile(installation_info_file):
            return '', f'{run_dir}——未安装汉化'
        to_return: (str, str)
        with open(installation_info_file, 'r', encoding='utf-8') as f:
            parsed_version = f.readline().strip()
            mo_path = Path(f.readline().strip())
            if not mo_path.is_file():
                return '', f'{run_dir}——未安装汉化'
            mo_sha256 = f.readline().strip()
            self.last_installed_l10n_version = parsed_version
            not_parsable = False
            try:
                float(parsed_version)
                to_return = parsed_version, f'{run_dir}——{parsed_version}'
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

    def find_launcher(self) -> str:
        game_path = self.get_game_path()
        if game_path:
            for launcher in launcher_dict.keys():
                launcher_file = game_path.joinpath(launcher)
                if launcher_file.is_file():
                    self.game_launcher_file = launcher_file
                    return launcher_dict.get(launcher)

        return '未找到客户端'

    def launch_game(self) -> None:
        if not self.game_launcher_file or not self.game_launcher_file.is_file():
            self.find_launcher()
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

    def parse_choice(self, use_cache: bool = True) -> Dict[str, str]:
        if use_cache and self.choice:
            return self.choice
        self.choice = {}
        game_path = self.get_game_path(find=False)
        if not game_path:
            return
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
        for choice in ['server_region', 'is_release', 'download_source', 'use_ee', 'apply_mods']:
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
            self.choice['apply_mods'] = self.mod_selection.get()
            with open(game_path.joinpath('l10n_installer/settings/choice.json'), 'w', encoding='utf-8') as f:
                json.dump(self.choice, f, ensure_ascii=False, indent=4)

    def on_closed(self):
        self.save_global_settings()


def mkdir(t_dir: Any):
    os.makedirs(t_dir, exist_ok=True)


def process_modification_file(source_mo, translated_path: str):
    if translated_path.endswith('po'):
        translated = polib.pofile(translated_path)
    else:
        translated = polib.mofile(translated_path)
    source_dict_singular = {entry.msgid: entry.msgstr for entry in source_mo if entry.msgid and entry.msgid != ''}
    translation_dict_singular = {entry.msgid: entry.msgstr for entry in translated if entry.msgid and entry.msgid != ''}
    singular_count = len(translation_dict_singular)
    for entry in source_mo:
        if singular_count == 0:
            break
        if entry.msgid and entry.msgid in translation_dict_singular:
            target_str = translation_dict_singular[entry.msgid]
            del translation_dict_singular[entry.msgid]
            singular_count -= 1
            if entry.msgid == 'IDS_RIGHTS_RESERVED':
                continue
            entry.msgstr = target_str
    if singular_count > 0:
        for t_entry in translated:
            if t_entry.msgid and t_entry.msgid not in source_dict_singular:
                source_mo.append(t_entry)


def is_valid_game_path(game_path: Path):
    return game_path.joinpath('game_info.xml').is_file() and game_path.joinpath('bin').is_dir()


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


if __name__ == '__main__':
    root = ttk.Window()
    icon = os.path.join(resource_path, 'icon.ico')
    root.iconbitmap(default=icon)
    root.iconbitmap(bitmap=icon)
    half_screen_width = int(root.winfo_screenwidth() / 2) - 225
    half_screen_height = int(root.winfo_screenheight() / 2) - 300
    root.geometry(f'+{half_screen_width}+{half_screen_height}')
    app = LocalizationInstaller(root)
    root.mainloop()
    app.on_closed()

# pyinstaller -w -i resources/icon.ico --onefile --add-data "resources\*;resources" installer_gui.py --clean

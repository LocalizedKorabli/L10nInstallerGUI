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
import json
import os
import shutil
import subprocess
import sys
import threading
import time
import tkinter as tk
import urllib.request
import webbrowser
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

mods_link = 'https://tapio.lanzn.com/b0nxzso2b'
project_repo_link = 'https://github.com/LocalizedKorabli/Korabli-LESTA-L10N/'
installer_repo_link = 'https://github.com/LocalizedKorabli/L10nInstallerGUI/'

version = '0.0.3a'

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

base_path: str = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
resource_path: str = os.path.join(base_path, 'resources')


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

    # Variables
    choice: Dict[str, Any] = None
    installed_l10n_version = ''
    run_dirs: Dict[str, str] = {}
    is_installing: bool = False
    game_launcher_file: str = ''

    def __init__(self, parent: tk.Tk):
        mkdir('l10n_installer/cache')
        mkdir('l10n_installer/downloads')
        mkdir('l10n_installer/mods')
        mkdir('l10n_installer/processed')
        mkdir('l10n_installer/settings')
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

        # 游戏版本
        ttk.Label(parent, text='游戏版本——汉化版本：') \
            .grid(row=0, column=0, columnspan=4, pady=5, sticky=tk.W)

        # 汉化状态
        self.localization_status_label_1st = ttk.Label(parent, textvariable=self.localization_status_1st)
        self.localization_status_label_1st.grid(row=1, column=0, columnspan=4, pady=5, sticky=tk.W)
        self.localization_status_label_2nd = ttk.Label(parent, textvariable=self.localization_status_2nd)
        self.localization_status_label_2nd.grid(row=2, column=0, columnspan=4, pady=5, sticky=tk.W)
        self.parse_game_version()

        # 游戏区服
        ttk.Label(parent, text='游戏区服：').grid(row=3, column=0, pady=5, sticky=tk.W)

        # 游戏区服选项
        ttk.Radiobutton(parent, text='莱服', variable=self.server_region, value='ru') \
            .grid(row=3, column=1, sticky=tk.W)
        ttk.Radiobutton(parent, text='直营服', variable=self.server_region, value='zh_sg', style='warning') \
            .grid(row=3, column=2, sticky=tk.W)
        ttk.Radiobutton(parent, text='国服', variable=self.server_region, value='zh_cn', style='danger') \
            .grid(row=3, column=3, sticky=tk.W)

        # 游戏类型
        ttk.Label(parent, text='游戏类型：').grid(row=4, column=0, pady=5, sticky=tk.W)

        # 游戏类型选项
        ttk.Radiobutton(parent, text='正式服', variable=self.is_release, value=True, style='success') \
            .grid(row=4, column=1, sticky=tk.W)
        ttk.Radiobutton(parent, text='PT服', variable=self.is_release, value=False, style='danger') \
            .grid(row=4, column=2, sticky=tk.W)

        self.detect_game_type_button = ttk.Button(parent, text='自动检测',
                                                  command=lambda: self.detect_game_status(manually=True))
        self.detect_game_type_button.grid(row=4, column=3)

        # 下载源
        ttk.Label(parent, text='汉化来源：').grid(row=5, column=0, pady=5, sticky=tk.W)
        # 下载源选项
        ttk.Radiobutton(parent, text='Gitee', variable=self.download_source, value='gitee', style='danger') \
            .grid(row=5, column=1, sticky=tk.W)
        ttk.Radiobutton(parent, text='GitHub', variable=self.download_source, value='github', style='dark') \
            .grid(row=5, column=2, sticky=tk.W)
        ttk.Radiobutton(parent, text='本地文件', variable=self.download_source, value='local') \
            .grid(row=5, column=3, sticky=tk.W)

        # 体验增强包/汉化修改包
        ttk.Checkbutton(parent, text='安装体验增强包', variable=self.ee_selection) \
            .grid(row=7, column=0, columnspan=2, pady=5, sticky=tk.W)
        ttk.Checkbutton(parent, text='安装模组（汉化修改包）', variable=self.mod_selection) \
            .grid(row=8, column=0, columnspan=2, pady=5, sticky=tk.W)
        self.mods_button = ttk.Button(parent, text='模组目录', command=self.open_mods_folder)
        self.mods_button.grid(row=8, column=2, columnspan=1)
        self.download_mods_button = ttk.Button(parent, text='下载模组',
                                               command=lambda: webbrowser.open_new_tab(mods_link))
        self.download_mods_button.grid(row=8, column=3, columnspan=1)

        # 安装路径选择/下载进度
        self.install_path_entry = ttk.Entry(parent, textvariable=self.mo_path, width=20)
        self.install_path_button = ttk.Button(parent, text='选择文件', command=self.choose_mo)
        self.download_progress_label = ttk.Label(parent, text='下载进度：')
        self.download_progress_info = ttk.Label(parent, textvariable=self.download_progress_text)

        # 安装/更新按钮
        self.install_button = ttk.Button(parent, text='安装汉化', command=self.install_update,
                                         style=ttk.SUCCESS)
        self.install_button.grid(row=9, column=0, pady=5)

        # 安装进度
        ttk.Label(parent, textvariable=self.install_progress_text).grid(row=9, column=1, columnspan=3,
                                                                        padx=5, sticky=tk.W)

        self.install_progress_bar = ttk.Progressbar(parent, variable=self.install_progress, maximum=100.0,
                                                    style='success-striped', length=400)
        self.install_progress_bar.grid(row=10, column=0, columnspan=4, padx=10)

        # 启动游戏
        self.launch_button = ttk.Button(parent, text='启动游戏', command=self.launch_game, style=ttk.WARNING)
        self.launch_button.grid(row=11, column=0, pady=5)

        # 启动器状态
        ttk.Label(parent, textvariable=self.game_launcher_status).grid(row=11, column=1, columnspan=3,
                                                                       padx=5, sticky=tk.W)

        # 相关链接
        about_button = ttk.Button(parent, text='关于项目', command=lambda: webbrowser.open_new_tab(project_repo_link),
                                  style=ttk.INFO)
        about_button.grid(row=12, column=0, pady=5)

        src_button = ttk.Button(parent, text='代码仓库', command=lambda: webbrowser.open_new_tab(installer_repo_link),
                                style=ttk.DANGER)
        src_button.grid(row=12, column=1, pady=5, padx=5)

        # 版权声明
        ttk.Label(parent, text='© 2024 LocalizedKorabli').grid(row=12, column=2, columnspan=3, pady=5)

        # 根据下载源选项显示或隐藏安装路径选择
        self.download_source.trace('w', self.toggle_install_path)

        choice = self.parse_choice()

        self.server_region.set(choice.get('server_region', 'ru'))
        self.is_release.set(choice.get('is_release', True))
        self.download_source.set(choice.get('download_source', 'gitee'))
        self.ee_selection.set(choice.get('use_ee', True))
        self.mod_selection.set(choice.get('apply_mods', True))

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

    def toggle_install_path(self, *args):
        if self.download_source.get() == 'local':
            self.install_path_entry.grid(row=6, column=0, columnspan=3)
            self.install_path_button.grid(row=6, column=3)
            self.download_progress_label.grid_forget()
            self.download_progress_info.grid_forget()
        else:
            self.download_progress_label.grid(row=6, column=0, pady=5, sticky=tk.W)
            self.download_progress_info.grid(row=6, column=1, pady=5, columnspan=3, sticky=tk.W)
            self.install_path_entry.grid_forget()
            self.install_path_button.grid_forget()

    def open_mods_folder(self):
        mods_folder = Path('l10n_installer').joinpath('mods')
        mkdir(mods_folder)
        subprocess.run(['explorer', mods_folder.absolute()])

    def choose_mo(self):
        mo_path = filedialog.askopenfilename(initialdir='.', filetypes=[('汉化包', ['*.mo', '*.zip']),
                                                                        ('MO汉化文件', '*.mo'),
                                                                        ('打包的汉化文件', '*.zip')])
        if mo_path:
            self.mo_path.set(mo_path)

    def install_update(self):
        if self.is_installing:
            return
        self.is_installing = True
        self.save_choice()
        tr = threading.Thread(target=self.do_install_update)
        tr.start()

    def do_install_update(self) -> None:
        self.safely_set_install_progress(progress=0.0)
        run_dirs = self.run_dirs.keys()
        if len(run_dirs) == 0:
            return
        is_release = self.is_release.get()
        for run_dir in run_dirs:
            target_path = Path('bin').joinpath(run_dir).joinpath('res_mods' if is_release else 'res')
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
            if self.ee_selection.get():
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
                        target_path = Path('bin').joinpath(run_dir).joinpath('res_mods' if is_release else 'res')
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
            target_path = Path('bin').joinpath(run_dir).joinpath('res_mods' if is_release else 'res')
            info_path = Path('bin').joinpath(run_dir).joinpath('l10n')
            mkdir(info_path)
            info_file = info_path.joinpath('version.info')
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
                if info_fetched and info_file.is_file():
                    with open(info_file, 'r', encoding='utf-8') as f:
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
                with open(info_file, 'w', encoding='utf-8') as f:
                    f.write(remote_version)
                    try:
                        float(remote_version)
                    except ValueError:
                        f.write(f'\n{time.time()}')
        self.is_installing = False
        if nothing_wrong:
            self.safely_set_install_progress(progress=100.0)
        self.safely_set_install_progress_text('完成！' if nothing_wrong else '失败！')
        self.parse_game_version()

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
        if valid_version and self.installed_l10n_version == remote_version:
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
        if not self.mod_selection.get() or not os.path.isdir('l10n_installer/mods'):
            return []
        files = os.listdir('l10n_installer/mods')
        return [('l10n_installer/mods/' + file) for file in files if (file.endswith('.po') or file.endswith('.mo'))]

    def parse_game_version(self) -> None:
        self.run_dirs = {}
        self.localization_status_1st.set('未发现游戏版本')
        self.localization_status_2nd.set('')
        v_1st = ''
        v_2nd = ''
        for v_dir_b in os.listdir(Path('bin')):
            v_dir = str(v_dir_b)
            # v_dir_num = 0
            try:
                v_dir_num = int(v_dir)
                if not is_valid_build_dir(Path('bin').joinpath(v_dir)):
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
            self.localization_status_label_2nd.grid(row=2, column=0, columnspan=4, pady=5, sticky=tk.W)

    def get_choice_template(self):
        self.detect_game_status()
        return {
            'server_region': self.server_region.get(),
            'is_release': self.is_release.get(),
            'download_source': 'gitee',
            'use_ee': True,
            'apply_mods': True
        }

    def detect_game_status(self, manually: bool = False):
        if not manually:
            self.server_region.set('ru')
            self.is_release.set(True)
        if not os.path.isfile('game_info.xml'):
            return
        game_info = ET.parse('game_info.xml')
        game_id = game_info.find('.//game/id')
        if game_id is None:
            return
        game_type: (str, bool) = server_regions_dict.get(game_id.text, ('ru', 'PT.PRODUCTION' not in game_id.text))
        self.server_region.set(game_type[0])
        self.is_release.set(game_type[1])
        self.parse_game_version()

    # 返回：(汉化版本号: str, 汉化状态: str)
    def get_local_l10n_version(self, run_dir: str) -> (str, str):
        info_file = Path('bin').joinpath(run_dir).joinpath('l10n').joinpath('version.info')
        if not os.path.isfile(info_file):
            return '', f'{run_dir}——未安装汉化'
        with open(info_file, 'r', encoding='utf-8') as f:
            parsed_version = f.readline()
            self.installed_l10n_version = parsed_version
            try:
                float(parsed_version)
                return parsed_version, f'{run_dir}——{parsed_version}'
            except ValueError:
                pass
            try:
                inst_time = float(f.readline())
                time_formatted = datetime.fromtimestamp(inst_time).strftime('%Y-%m-%d %H:%M:%S')
                return '', f'{run_dir}——汉化版本未知，于{time_formatted}安装'
            except ValueError:
                return '', f'{run_dir}——汉化版本未知，安装时间未知'

    def find_launcher(self) -> str:
        if os.path.isfile('lgc_api.exe'):
            self.game_launcher_file = 'lgc_api.exe'
            return '莱服客户端'
        elif os.path.isfile('wgc_api.exe'):
            self.game_launcher_file = 'wgc_api.exe'
            return '直营服客户端'
        elif os.path.isfile('wgc360_api.exe'):
            self.game_launcher_file = 'wgc360_api.exe'
            return '国服客户端'
        return '未找到客户端'

    def launch_game(self) -> None:
        if not self.game_launcher_file or not os.path.isfile(self.game_launcher_file):
            self.find_launcher()
        if not self.game_launcher_file or not os.path.isfile(self.game_launcher_file):
            return
        subprocess.run(self.game_launcher_file)

    def parse_choice(self) -> Dict[str, str]:
        if self.choice:
            return self.choice
        self.choice = {}
        choice_file = 'l10n_installer/settings/choice.json'
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
        if self.choice:
            self.choice['is_release'] = self.is_release.get()
            self.choice['download_source'] = self.download_source.get()
            self.choice['use_ee'] = self.ee_selection.get()
            self.choice['apply_mods'] = self.mod_selection.get()
            with open('l10n_installer/settings/choice.json', 'w', encoding='utf-8') as f:
                json.dump(self.choice, f, ensure_ascii=False, indent=4)


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


def is_valid_build_dir(build_dir: Path) -> bool:
    res_dir = Path(build_dir).joinpath('res')
    if not os.path.isdir(res_dir):
        return False
    return os.path.isfile(res_dir.joinpath('locale_config.xml'))


if __name__ == '__main__':
    root = ttk.Window(iconphoto=None)
    root.iconbitmap(os.path.join(resource_path, 'icon.ico'))
    half_screen_width = int(root.winfo_screenwidth() / 2) - 200
    half_screen_height = int(root.winfo_screenheight() / 2) - 300
    root.geometry(f'+{half_screen_width}+{half_screen_height}')
    app = LocalizationInstaller(root)
    root.mainloop()

# pyinstaller -w -i resources/icon.ico --onefile --add-data "resources\*;resources" installer_gui.py --clean

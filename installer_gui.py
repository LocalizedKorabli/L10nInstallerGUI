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
import time
import tkinter as tk
import urllib.request
# pip install urllib3==1.25.11
# The newer urllib has break changes.
import xml.etree.ElementTree as ET
import zipfile
from datetime import datetime
from pathlib import Path
from tkinter import filedialog
from tkinter import ttk
from typing import Any, Dict, List, Tuple

import polib
import requests

# import ttkbootstrap

version = '0.0.1'

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

base_path: str = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
resource_path: str = os.path.join(base_path, 'resources')


class LocalizationInstaller:
    # GUI Related
    game_version: tk.StringVar
    localization_status: tk.StringVar
    is_release: tk.BooleanVar
    download_source: tk.StringVar
    builtin_mods_selection: tk.BooleanVar
    mod_selection: tk.BooleanVar
    mo_path: tk.StringVar
    install_progress: tk.StringVar
    download_info: tk.StringVar

    # Variables
    human_readable_version: str = ''
    installed_l10n_version = ''
    run_dir: str = ''
    is_installing: bool = False

    def __init__(self, parent: tk.Tk):
        mkdir('l10n_installer/cache')
        mkdir('l10n_installer/downloads')
        mkdir('l10n_installer/mods')
        mkdir('l10n_installer/processed')
        self.root = parent
        self.root.title(f'汉化安装器-v{version}')
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
        self.game_version.set('游戏版本：' + self.get_human_readable_version())

        # 第二行：汉化状态
        tk.Label(parent, textvariable=self.localization_status) \
            .grid(row=1, column=0, columnspan=3, sticky=tk.W)
        self.localization_status.set('汉化版本：' + self.get_local_l10n_version())

        # 第三行：游戏类型
        tk.Label(parent, text='游戏类型：').grid(row=2, column=0, sticky=tk.W)

        # 游戏类型选项
        ttk.Radiobutton(parent, text='正式服', variable=self.is_release, value=True) \
            .grid(row=2, column=1, sticky=tk.W)
        ttk.Radiobutton(parent, text='PT服', variable=self.is_release, value=False) \
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
            .grid(row=5, column=0, columnspan=2, sticky=tk.W)
        ttk.Checkbutton(parent, text='安装汉化修改包', variable=self.mod_selection) \
            .grid(row=6, column=0, columnspan=2, sticky=tk.W)
        self.mods_button = tk.Button(parent, text='打开汉化修改包文件夹', command=lambda: self.open_mods_folder())
        self.mods_button.grid(row=6, column=2, columnspan=2)

        # 第六行：安装路径选择/下载进度
        self.install_path_entry = tk.Entry(parent, textvariable=self.mo_path, width=20)
        self.install_path_button = tk.Button(parent, text='选择文件', command=self.choose_mo)
        self.download_progress_label = tk.Label(parent, text='下载进度：')
        self.download_progress_info = tk.Label(parent, textvariable=self.download_info)

        # 第七行：安装/更新按钮
        self.install_button = tk.Button(parent, text='安装或更新', command=self.install_update)
        self.install_button.grid(row=7, column=0)

        # 安装进度
        tk.Label(parent, textvariable=self.install_progress).grid(row=7, column=1, columnspan=2, sticky=tk.W)

        # 第八行：启动游戏
        self.launch_button = tk.Button(parent, text='启动客户端', command=launch_game)
        self.launch_button.grid(row=8, column=0)

        # 启动器状态
        tk.Label(parent, textvariable=self.game_launcher_status).grid(row=8, column=1, columnspan=2, sticky=tk.W)

        tk.Label(parent, text='Copyright © 2024 LocalizedKorabli').grid(row=9, column=0, columnspan=4, sticky='ew')

        # 根据下载源选项显示或隐藏安装路径选择
        self.download_source.trace('w', self.toggle_install_path)

        self.builtin_mods_selection.set(True)
        self.mod_selection.set(True)
        self.is_release.set(True)
        self.download_source.set('gitee')
        self.download_info.set('准备')
        self.install_progress.set('安装进度：' + '等待中')
        self.game_launcher_status.set(find_launcher())

    def open_mods_folder(self):
        mods_folder = Path('l10n_installer/mods')
        mkdir(mods_folder)
        subprocess.run(['explorer', mods_folder.absolute()])

    def choose_mo(self):
        mo_path = filedialog.askopenfilename(initialdir='.', filetypes=[('MO文件', '*.mo')])
        if mo_path:
            self.mo_path.set(mo_path)

    def install_update(self):
        if self.is_installing:
            return
        self.is_installing = True
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
            old_cfg_backup = target_path.joinpath('locale_config.xml.old')
            if not os.path.isfile(old_cfg_backup) and os.path.isfile(old_cfg):
                shutil.copy(old_cfg, old_cfg_backup)
            with open(old_cfg, 'w', encoding='utf-8') as file:
                file.write(locale_config)
        else:
            with open(target_path.joinpath('locale_config.xml'), 'w', encoding='utf-8') as file:
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
                        self.download_info.set('下载体验增强包——完成')
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
        # 汉化包
        self.install_progress.set('安装汉化包')
        download_src = self.download_source.get()
        nothing_wrong = True
        # remote_version = ''
        # downloaded_mo = ''
        if download_src != 'local':
            self.download_info.set('下载汉化包')
            download_link_base = download_routes['r' if is_release else 'pt'][download_src]['url']
            # Check and Fetch
            downloaded: Tuple = self.check_version_and_fetch_mo(download_link_base, proxies)
            downloaded_mo = downloaded[0]
            remote_version = downloaded[1]
        else:
            downloaded_mo = self.mo_path.get()
            remote_version = 'local'
        if not downloaded_mo.endswith('.mo') or not os.path.isfile(downloaded_mo):
            self.download_info.set('下载汉化包——文件异常')
            nothing_wrong = False
        if nothing_wrong:
            self.download_info.set('下载汉化包——完成')
            mods = self.get_mods()
            downloaded_mo = self.parse_and_apply_mods(downloaded_mo, mods)
            if downloaded_mo == '':
                self.install_progress.set('安装汉化包——文件损坏')
                nothing_wrong = False
        if nothing_wrong:
            self.install_progress.set('安装汉化包——移动文件')
            mo_dir = target_path.joinpath('texts').joinpath('ru').joinpath('LC_MESSAGES')
            mkdir(mo_dir)
            old_mo = mo_dir.joinpath('global.mo')
            old_mo_backup = mo_dir.joinpath('global.mo.old')
            if not is_release:
                if not os.path.isfile(old_mo_backup) and os.path.isfile(old_mo):
                    shutil.copy(old_mo, old_mo_backup)
            shutil.copy(downloaded_mo, old_mo)
            info_path = Path('bin').joinpath(run_dir).joinpath('l10n')
            mkdir(info_path)
            info_file = info_path.joinpath('version.info')
            with open(info_file, 'w', encoding='utf-8') as f:
                f.write(remote_version)
                try:
                    float(remote_version)
                except ValueError:
                    f.write(f'\n{time.time()}')
        self.is_installing = False
        self.install_progress.set('完成！' if nothing_wrong else '失败！')
        self.localization_status.set('汉化版本：' + self.get_local_l10n_version())

    def check_version_and_fetch_mo(self, download_link_base: str, proxies: Dict) -> (str, str):
        remote_version: str = 'latest'
        self.download_info.set('下载汉化包——获取版本')
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
                    self.download_info.set(f'下载汉化包——最新版本为{remote_version}')
        except requests.exceptions.RequestException:
            pass
        valid_version = remote_version != 'latest'
        if not valid_version:
            self.download_info.set(f'下载汉化包——版本获取失败')
        mo_file_name = f'{remote_version}.mo'
        output_file = f'l10n_installer/downloads/{mo_file_name}'
        # Check existed
        if valid_version and self.installed_l10n_version == remote_version:
            if os.path.isfile(output_file):
                try:
                    if polib.mofile(output_file):
                        self.download_info.set(f'下载汉化包——使用已下载文件')
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

    def parse_and_apply_mods(self, downloaded_mo: str, mods: List[str]) -> str:
        try:
            downloaded_mo_instance = polib.mofile(downloaded_mo)
        except Exception:
            return ''
        if len(mods) == 0:
            return downloaded_mo
        self.install_progress.set('安装汉化包——应用模组')
        for mod in mods:
            try:
                process_modification_file(downloaded_mo_instance, mod)
            except Exception:
                pass
        for file in os.listdir('l10n_installer/processed/'):
            try:
                os.remove(file)
            except Exception:
                continue
        modded_file_name = f'l10n_installer/processed/modified_{time.time()}.mo'
        downloaded_mo_instance.save(modded_file_name)
        return modded_file_name

    def get_mods(self) -> List[str]:
        if not self.mod_selection.get() or not os.path.isdir('l10n_installer/mods'):
            return []
        files = os.listdir('l10n_installer/mods')
        return [('l10n_installer/mods/' + file) for file in files if (file.endswith('.po') or file.endswith('.mo'))]

    def get_run_dir(self) -> str:
        if self.run_dir and self.run_dir != '':
            return self.run_dir
        self.parse_game_version()
        return self.run_dir

    def get_human_readable_version(self) -> str:
        if self.human_readable_version and self.human_readable_version != '':
            return self.human_readable_version
        self.parse_game_version()
        return self.human_readable_version

    def parse_game_version(self) -> None:
        if os.path.isfile('game_info.xml'):
            game_info = ET.parse('game_info.xml')
            for version in game_info.findall('.//version'):
                if version.get('name') == 'locale':
                    full_version = str(version.get('installed'))
                    self.run_dir = full_version.split('.')[-1]
                    self.human_readable_version = '.'.join(full_version.split('.')[:-1])
                    return
            self.human_readable_version = '未知'
        self.human_readable_version = '未在运行目录下找到战舰世界客户端！'

    def get_local_l10n_version(self) -> str:
        info_file = Path('bin').joinpath(self.get_run_dir()).joinpath('l10n').joinpath('version.info')
        if not os.path.isfile(info_file):
            self.installed_l10n_version = ''
            return '未安装'
        with open(info_file, 'r', encoding='utf-8') as f:
            parsed_version = f.readline()
            self.installed_l10n_version = parsed_version
            try:
                float(parsed_version)
                return parsed_version
            except ValueError:
                pass
            try:
                inst_time = float(f.readline())
                time_formatted = datetime.fromtimestamp(inst_time).strftime('%Y-%m-%d %H:%M:%S')
                return f'未知，于{time_formatted}安装'
            except ValueError:
                return f'未知，安装时间未知'

    def toggle_install_path(self, *args):
        if self.download_source.get() == 'local':
            self.install_path_entry.grid(row=4, column=0, columnspan=3, pady=12)
            self.install_path_button.grid(row=4, column=3)
            self.download_progress_label.grid_forget()
            self.download_progress_info.grid_forget()
        else:
            self.download_progress_label.grid(row=4, column=0, pady=11, sticky=tk.W)
            self.download_progress_info.grid(row=4, column=1, pady=11, columnspan=2, sticky=tk.W)
            self.install_path_entry.grid_forget()
            self.install_path_button.grid_forget()


def mkdir(t_dir: Any):
    os.makedirs(t_dir, exist_ok=True)


def process_modification_file(source_po, translated_path: str):
    if translated_path.endswith('po'):
        translated = polib.pofile(translated_path)
    else:
        translated = polib.mofile(translated_path)
    source_dict_singular = {entry.msgid: entry.msgstr for entry in source_po}
    translation_dict_singular = {entry.msgid: entry.msgstr for entry in translated if entry.msgid != ''}
    translation_dict_plural: Dict[str, List[str]] = {entry.msgid_plural: entry.msgstr_plural for entry in
                                                     translated if entry.msgid_plural != ''}
    singular_count = len(translation_dict_singular)
    plural_count = len(translation_dict_singular)
    for entry in source_po:
        if singular_count != 0 and entry.msgid and entry.msgid in translation_dict_singular:
            target_str = translation_dict_singular[entry.msgid]
            del translation_dict_singular[entry.msgid]
            singular_count -= 1
            if entry.msgid == 'IDS_RIGHTS_RESERVED':
                continue
            entry.msgstr = target_str
        if plural_count != 0 and entry.msgid_plural and entry.msgid_plural in translation_dict_plural:
            entry.msgstr_plural = translation_dict_plural.get(entry.msgid_plural)
            del translation_dict_plural[entry.msgid_plural]
            plural_count -= 1
    if singular_count > 0 or plural_count > 0:
        for t_entry in translated:
            if t_entry.msgid and t_entry.msgid not in source_dict_singular:
                source_po.append(t_entry)


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

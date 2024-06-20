import sys
import os
import time
import shutil
import re
import importlib as imp
from filecmp import dircmp

CLEAR_SCREEN = False


def get_folder_size(path: str):
    total_size = 0
    for dirPath, dirNames, filenames in os.walk(path):
        for f in filenames:
            fp = os.path.join(dirPath, f)
            if not os.path.islink(fp):
                total_size += os.path.getsize(fp)
    return total_size


class ConsoleStyles:
    BOLD = '\033[1m'
    GRAY = '\033[90m'
    ERROR = '\033[91m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    OKBLUE = '\033[94m'
    HEADER = '\033[95m'
    ENDC = '\033[0m'


anyAlert = False
version = sys.version_info
if (version.major, version.minor) != (3, 11):
    anyAlert = True
    version = sys.version_info
    version = str(version.major) + '.' + str(version.minor)
    version += ' ' * (12 - len(version))
    print('''
+---------------ALERT---------------+
| You are using Python {version} |
| Python 3.11 is recommended        |
| Some features may be unavailable  |
+-----------------------------------+
'''.format(version=version))

shortcuts_enabled = True
try:
    win32com_client = imp.import_module('win32com.client')
except ImportError:
    anyAlert = True
    shortcuts_enabled = False
    print('''| Module win32com not found         |
| Shortcut feature is unavailable   |
+-----------------------------------+
''')

if anyAlert:
    input('Press Enter to continue...')
    print('\n')

print(f'''Welcome to NoitaSaves!\n
 > To make a save, you should save and quit the game
 > You also need to close Noita before loading a save
   {ConsoleStyles.WARNING}Do not load a save during Steam sync! It may corrupt the save!{ConsoleStyles.ENDC}
 > If the selected save has not loaded, just load it one more time
   (It may happen due to steam sync)
 > You can also create a shortcut for NoitaSaves on your start menu or desktop
   (Check github page for more info: {ConsoleStyles.OKBLUE}https://github.com/Sedo-KFM/NoitaSaves{ConsoleStyles.ENDC})''')

FILENAME = str(os.path.basename(__file__).removesuffix('.py'))
INFO_FILE_NAME = '.noita_saves_info'
APPDATA = os.getenv('APPDATA')
APPDATA = APPDATA.replace('Roaming', 'LocalLow')
GAME_DIR = APPDATA + r'\Nolla_Games_Noita'
SAVES_DIR = APPDATA + r'\Nolla_Games_Noita_Saves'
CURRENT_SAVE = GAME_DIR + r'\save00'

if not os.path.exists(SAVES_DIR):
    os.mkdir(SAVES_DIR)

saves = []
error_message = ''
command = ''
first_time = True
while command != 'q':

    if error_message != '':
        input('\n' + ConsoleStyles.ERROR + 'Error: ' + error_message + ConsoleStyles.ENDC + '\n\nPress Enter...')
        error_message = ''

    if CLEAR_SCREEN:
        if first_time:
            first_time = False
        else:
            os.system('cls')

    print('\n\nSaves:')
    saves = [(save, os.path.getctime(SAVES_DIR + '\\' + save)) for save in os.listdir(SAVES_DIR)]
    saves.sort(key=lambda s: s[1])
    current_save_size = get_folder_size(CURRENT_SAVE)
    for (index, (save, creation_time)) in enumerate(saves):
        is_equal_to_current = False
        if INFO_FILE_NAME in os.listdir(SAVES_DIR + '\\' + save):
            with open(SAVES_DIR + '\\' + save + '\\' + INFO_FILE_NAME, 'r') as info_file:
                save_size = int(info_file.readline())
        else:
            save_size = get_folder_size(SAVES_DIR + '\\' + save)
            with open(SAVES_DIR + '\\' + save + '\\' + INFO_FILE_NAME, 'w') as info_file:
                info_file.write(str(save_size))
        if save_size == current_save_size:
            dir_cmp_result = dircmp(CURRENT_SAVE, SAVES_DIR + '\\' + save)
            if len(dir_cmp_result.diff_files) == 0:
                is_equal_to_current = True
        printing_save = save.replace('_', ' ')
        print(ConsoleStyles.BOLD + ConsoleStyles.OKGREEN if is_equal_to_current else '',
              '#', index + 1,
              ' ' if index < 9 else '',
              ' >> ', printing_save,
              '' if is_equal_to_current else ConsoleStyles.ENDC + ConsoleStyles.GRAY,
              '  [',
              ' '.join(
                  time.ctime(creation_time)
                  .replace('  ', ' ')
                  .split(' ')[1:4]),
              ' | ',
              '{:1.2f}'.format(save_size / 1000 / 1000), ' Mb',
              ']',
              ConsoleStyles.ENDC,
              sep='')
    if len(saves) == 0:
        print('<< Nothing >>')
    print()

    parameter = ''
    user_input = input('S (Save) | L (Load) | P (Play) | D (Delete) | Q (Quit) >> ').lower().strip().split(maxsplit=1)
    if len(user_input) > 0:
        command = user_input[0]
    if len(user_input) > 1:
        parameter = user_input[1]

    if command in ('l', 's', 'd'):  # :D

        if not os.path.exists(GAME_DIR):
            error_message = 'Game files not found'
            continue

        if command == 's':
            if not os.path.exists(CURRENT_SAVE):
                error_message = 'Current progress not found\nTry to load any save or start a new game'
                continue
            if parameter == '':
                save_name = input('Input save name >> ')
            else:
                save_name = parameter
            save_name_errors = set(re.findall(r'[^A-Za-zА-Яа-я0-9\- ]', save_name))
            if len(save_name_errors) > 0:
                error_message = 'Incorrect symbols: [' + '], ['.join(save_name_errors) + ']'
                continue
            save_name = save_name.replace(' ', '_')
            if save_name in os.listdir(SAVES_DIR):
                error_message = 'This save is already exists'
                continue
            else:
                print('Saving...')
                dirname = SAVES_DIR + '\\' + save_name
                shutil.copytree(CURRENT_SAVE, dirname)
                with open(dirname + '\\' + INFO_FILE_NAME, 'w', encoding='utf-8') as info_file:
                    info_file.write(str(get_folder_size(dirname)))

        elif command in ('l', 'd'):
            if len(saves) == 0:
                error_message = 'Nothing to ' + ('load' if command == 'l' else 'delete')
                continue
            save_index = None
            if parameter == '':
                parameter = input('Select the save index ' +
                                  ('(or "a/all")' if command == 'd' else '(or l/last)') +
                                  ' >> ')
            if parameter.isdecimal():
                save_index = int(parameter)
                parameter = None
                if not 0 < save_index <= len(saves):
                    error_message = 'Incorrect index'
                    continue
            elif command == 'd' and parameter in ('a', 'all'):
                print(ConsoleStyles.WARNING + 'Are you sure? [Y/N]', ConsoleStyles.ENDC, end=' >> ')
                if input().strip().lower() != 'y':
                    continue
                parameter = 'all'
            elif command == 'l' and parameter in ('l', 'last'):
                save_index = len(saves)
            else:
                error_message = 'Incorrect index'
                continue

            if command == 'l':
                print('Loading...')
                if os.path.exists(CURRENT_SAVE):
                    shutil.rmtree(CURRENT_SAVE)
                shutil.copytree(SAVES_DIR + '\\' + saves[save_index - 1][0], CURRENT_SAVE)
                if INFO_FILE_NAME in os.listdir(CURRENT_SAVE):
                    os.remove(CURRENT_SAVE + '\\' + INFO_FILE_NAME)

            elif command == 'd':
                print('Deleting...')
                if parameter == 'all':
                    for (save, _) in saves:
                        shutil.rmtree(SAVES_DIR + '\\' + save)
                else:
                    if save_index == -1:
                        for (save, _) in saves:
                            shutil.rmtree(SAVES_DIR + '\\' + save)
                    else:
                        shutil.rmtree(SAVES_DIR + '\\' + saves[save_index - 1][0])

        print('\nDone!')

    elif command in ('sda', 'sma', 'sdr', 'smr'):
        if not shortcuts_enabled:
            error_message = 'Shortcut feature is disabled'
            continue

        if 'd' in command:
            shortcut_path = os.getenv('USERPROFILE') + r'\Desktop'
        elif 'm' in command:
            shortcut_path = os.getenv('APPDATA') + r'\Microsoft\Windows\Start Menu\Programs'
        shortcut_path += '\\' + FILENAME + '.lnk'
        shortcut_removed = False
        if os.path.exists(shortcut_path):
            os.remove(shortcut_path)
            shortcut_removed = True

        if 'a' in command:
            shell = win32com_client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut(shortcut_path)
            shortcut.Targetpath = os.getcwd() + '\\' + FILENAME + '.py'
            shortcut.IconLocation = os.getcwd() + '\\' + FILENAME + '.ico'
            shortcut.WorkingDirectory = os.getcwd()
            shortcut.save()
            print('\nShortcut ' + ('updated' if shortcut_removed else 'created') + '!')
        else:
            print('\nShortcut removed!')

    elif command == 'p':
        os.system('start steam://rungameid/881100')

    else:
        if len(command) > 0:
            error_message = 'Incorrect command'

if len(saves) == 0:
    os.rmdir(SAVES_DIR)

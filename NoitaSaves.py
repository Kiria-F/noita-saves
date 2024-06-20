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

filename = os.path.basename(__file__).removesuffix('.py')
info_file_name = '.noita_saves_info'
appdata = os.getenv('APPDATA')
appdata = appdata.replace('Roaming', 'LocalLow')
game_dir = appdata + r'\Nolla_Games_Noita'
saves_dir = appdata + r'\Nolla_Games_Noita_Saves'
current_save = game_dir + r'\save00'

if not os.path.exists(saves_dir):
    os.mkdir(saves_dir)

saves = []
error_message = ''
command = ''
first_time = True
while command != 'e':

    if error_message != '':
        input('\n' + ConsoleStyles.ERROR + 'Error: ' + error_message + ConsoleStyles.ENDC + '\n\nPress Enter...')
        error_message = ''

    if CLEAR_SCREEN:
        if first_time:
            first_time = False
        else:
            os.system('cls')

    print('\n\nSaves:')
    saves = [(save, os.path.getctime(saves_dir + '\\' + save)) for save in os.listdir(saves_dir)]
    saves.sort(key=lambda s: s[1])
    current_save_size = get_folder_size(current_save)
    for (index, (save, creation_time)) in enumerate(saves):
        is_equal_to_current = False
        if info_file_name in os.listdir(saves_dir + '\\' + save):
            with open(saves_dir + '\\' + save + '\\' + info_file_name, 'r') as info_file:
                save_size = int(info_file.readline())
        else:
            save_size = get_folder_size(saves_dir + '\\' + save)
            with open(saves_dir + '\\' + save + '\\' + info_file_name, 'w') as info_file:
                info_file.write(str(save_size))
        if save_size == current_save_size:
            dir_cmp_result = dircmp(current_save, saves_dir + '\\' + save)
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
    user_input = input('S (Save) | L (Load) | P (Play) | D (Delete) | E (Exit) >> ').lower().strip().split(maxsplit=1)
    if len(user_input) > 0:
        command = user_input[0]
    if len(user_input) > 1:
        parameter = user_input[1]

    if command in ('l', 's', 'd'):  # :D

        if not os.path.exists(game_dir):
            error_message = 'Game files not found'
            continue

        if command == 's':
            if not os.path.exists(current_save):
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
            if save_name in os.listdir(saves_dir):
                error_message = 'This save is already exists'
                continue
            else:
                print('Saving...')
                dirname = saves_dir + '\\' + save_name
                shutil.copytree(current_save, dirname)
                with open(dirname + '\\' + info_file_name, 'w', encoding='utf-8') as info_file:
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
                if os.path.exists(current_save):
                    shutil.rmtree(current_save)
                shutil.copytree(saves_dir + '\\' + saves[save_index - 1][0], current_save)
                if info_file_name in os.listdir(current_save):
                    os.remove(current_save + '\\' + info_file_name)

            elif command == 'd':
                print('Deleting...')
                if parameter == 'all':
                    for (save, _) in saves:
                        shutil.rmtree(saves_dir + '\\' + save)
                else:
                    if save_index == -1:
                        for (save, _) in saves:
                            shutil.rmtree(saves_dir + '\\' + save)
                    else:
                        shutil.rmtree(saves_dir + '\\' + saves[save_index - 1][0])

        print('\nDone!')

    elif command in ('cs-d', 'cs-w', 'rs-d', 'rs-w'):
        if not shortcuts_enabled:
            error_message = 'Shortcut feature is disabled'
            continue

        if '-d' in command:
            shortcut_path = os.getenv('USERPROFILE') + r'\Desktop'
        elif '-w' in command:
            shortcut_path = os.getenv('APPDATA') + r'\Microsoft\Windows\Start Menu\Programs'
        shortcut_path += '\\' + filename + '.lnk'
        shortcut_removed = False
        if os.path.exists(shortcut_path):
            os.remove(shortcut_path)
            shortcut_removed = True

        if 'cs' in command:
            shell = win32com_client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut(shortcut_path)
            shortcut.Targetpath = os.getcwd() + '\\' + filename + '.py'
            shortcut.IconLocation = os.getcwd() + '\\' + filename + '.ico'
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
    os.rmdir(saves_dir)

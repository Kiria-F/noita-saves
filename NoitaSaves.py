import sys
import os
import time
import shutil
import re
import importlib as imp


def get_folder_size(path: str):
    total_size = 0
    for dirPath, dirNames, filenames in os.walk(path):
        for f in filenames:
            fp = os.path.join(dirPath, f)
            if not os.path.islink(fp):
                total_size += os.path.getsize(fp)
    return total_size


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

filename = os.path.basename(__file__).removesuffix('.py')
appdata = os.getenv('APPDATA')
appdata = appdata.replace('Roaming', 'LocalLow')
game_dir = appdata + r'\Nolla_Games_Noita'
saves_dir = appdata + r'\Nolla_Games_Noita_Saves'

if not os.path.exists(saves_dir):
    os.mkdir(saves_dir)

saves = []
error_message = ''
scenario = 'Init'
first_time = True
while scenario != 'e':
    if first_time:
        first_time = False
        print('''Welcome to NoitaSaves!\n
 > To make a save, you should save and quit the game
 > You also need to close Noita before loading a save
 > If the selected save has not loaded, just load it one more time
   (It can happen due to steam sync)
 > You can also create a shortcut for NoitaSaves on your start menu or desktop
   (Check github page for more info: https://github.com/Sedo-KFM/NoitaSaves)''')

    if error_message != '':
        input('\nError: ' + error_message + '\nPress Enter...')
        error_message = ''

    print('\n\nSaves:')
    saves = [(save, os.path.getctime(saves_dir + '\\' + save)) for save in os.listdir(saves_dir)]
    saves.sort(key=lambda s: s[1])
    for index, (save, creation_time) in enumerate(saves):
        printing_save = save.replace('_', ' ')
        print('#', index + 1,
              ' ' if index < 9 else '',
              ' >> ', printing_save, '  [',
              ' '.join(
                  time.ctime(creation_time)
                  .replace('  ', ' ')
                  .split(' ')[1:4]),
              ' | ',
              '{:1.2f}'.format(get_folder_size(saves_dir + '\\' + save) / 1000 / 1000), ' Mb',
              ']',
              sep='')
    if len(saves) == 0:
        print('<< Nothing >>')
    print()

    scenario_correct = False
    buffer = ''
    while not scenario_correct:
        scenario = input('Save (S) | Load (L) | Delete (D) | Exit (E) >> ').lower().replace(' ', '')
        if scenario == '':
            continue
        if scenario[0] in ('s', 'l', 'd', 'e'):
            buffer = scenario[1:]
            scenario = scenario[0]
        if scenario in ('s', 'l', 'd', 'e', 'cs-d', 'cs-w', 'rs-d', 'rs-w'):
            scenario_correct = True
        else:
            error_message = 'Incorrect scenario'
            continue

    if scenario in ('s', 'l', 'd'):

        if not os.path.exists(game_dir):
            error_message = 'Game files not found'
            continue

        if scenario == 's':
            if buffer == '':
                save_name = input('Input save name >> ')
            else:
                save_name = buffer
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
                shutil.copytree(game_dir + r'\save00', saves_dir + '\\' + save_name)

        elif scenario in ('l', 'd'):
            save_index = None
            if buffer == '':
                buffer = input('Select the save index ' + ('(or "all")' if scenario == 'd' else '') + ' >> ')
            if buffer.isdecimal():
                save_index = int(buffer)
                buffer = None
                if not 0 < save_index <= len(saves):
                    error_message = 'Incorrect index'
                    continue
            elif scenario == 'd' and buffer in ('a', 'al', 'all'):
                buffer = 'all'
            else:
                error_message = 'Incorrect index'
                continue

            if scenario == 'l':
                if os.path.exists(game_dir + r'\save00'):
                    shutil.rmtree(game_dir + r'\save00')
                shutil.copytree(saves_dir + '\\' + saves[save_index - 1][0], game_dir + r'\save00')

            elif scenario == 'd':
                if buffer == 'all':
                    for (save, _) in saves:
                        shutil.rmtree(saves_dir + '\\' + save)
                else:
                    if save_index == -1:
                        for (save, _) in saves:
                            shutil.rmtree(saves_dir + '\\' + save)
                    else:
                        shutil.rmtree(saves_dir + '\\' + saves[save_index - 1][0])

        print('\nDone!')

    if scenario in ('cs-d', 'cs-w', 'rs-d', 'rs-w'):
        if not shortcuts_enabled:
            error_message = 'Shortcut feature is disabled'
            continue

        if '-d' in scenario:
            shortcut_path = os.getenv('USERPROFILE') + r'\Desktop'
        elif '-w' in scenario:
            shortcut_path = os.getenv('APPDATA') + r'\Microsoft\Windows\Start Menu\Programs'
        shortcut_path += '\\' + filename + '.lnk'
        shortcut_removed = False
        if os.path.exists(shortcut_path):
            os.remove(shortcut_path)
            shortcut_removed = True

        if 'cs' in scenario:
            shell = win32com_client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut(shortcut_path)
            shortcut.Targetpath = os.getcwd() + '\\' + filename + '.py'
            shortcut.IconLocation = os.getcwd() + '\\' + filename + '.ico'
            shortcut.WorkingDirectory = os.getcwd()
            shortcut.save()
            print('\nShortcut ' + ('updated' if shortcut_removed else 'created') + '!\n\n')
        else:
            print('\nShortcut removed!\n\n')

if len(saves) == 0:
    os.rmdir(saves_dir)

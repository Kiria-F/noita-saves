import sys
import os
import shutil
import re
import importlib as imp

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
scenario = 'Init'
first_time = True
while scenario != 'e':
    if first_time:
        first_time = False
        print('Welcome to NoitaSaves!')
    print('\n\nSaves:')
    saves = os.listdir(saves_dir)
    printing_saves = [save.replace('_', ' ') for save in saves]
    for index, save in enumerate(printing_saves):
        print('#', index + 1, ' ' if index < 9 else '', ' >> ', save, sep='')
    if len(saves) == 0:
        print('<< Nothing >>')
    print()

    scenario_correct = False
    buffer = ''
    while not scenario_correct:
        scenario = input('Save (S) | Load (L) | Delete (D) | Exit (E) >> ').lower().replace(' ', '')
        if scenario == '':
            continue
        if scenario[0] in ('s', 'l', 'd'):
            buffer = scenario[1:]
            scenario = scenario[0]
        if scenario in ('s', 'l', 'd', 'e', 'cs-d', 'cs-w', 'rs-d', 'rs-w'):
            scenario_correct = True
        else:
            print('\nError: Incorrect scenario\n\n')
            continue

    if scenario in ('s', 'l', 'd'):

        if not os.path.exists(game_dir):
            print('\nError: Game files not found\n\n')
        else:

            if scenario == 's':
                save_name = input('Input save name >> ')
                save_name_errors = set(re.findall(r'[^A-Za-zА-Яа-я0-9\- ]', save_name))
                if len(save_name_errors) > 0:
                    print('Error: Incorrect symbols: [', end='')
                    print(*save_name_errors, sep='], [', end=']\n')
                else:
                    save_name = save_name.replace(' ', '_')
                    if save_name in os.listdir(saves_dir):
                        print('Error: This save is already exists')
                    else:
                        print('Saving...')
                        shutil.copytree(game_dir + r'\save00', saves_dir + '\\' + save_name)
                        print('Done!')

            elif scenario in ('l', 'd'):
                save_index = None
                if buffer == '':
                    buffer = input('Select the save index ' + ('(or "all")' if scenario == 'd' else '') + ' >> ')
                if buffer.isdecimal():
                    save_index = int(buffer)
                    buffer = None
                    if not 0 < save_index <= len(saves):
                        print('Error: Incorrect index')
                        continue
                elif scenario == 'd' and buffer in ('a', 'al', 'all'):
                    buffer = 'all'
                else:
                    print('Error: Incorrect index')
                    continue

                if scenario == 'l':
                    shutil.rmtree(game_dir + r'\save00')
                    shutil.copytree(saves_dir + '\\' + saves[save_index - 1], game_dir + r'\save00')
                if scenario == 'd':
                    if buffer == 'all':
                        for save in saves:
                            shutil.rmtree(saves_dir + '\\' + save)
                    else:
                        if save_index == -1:
                            for save in saves:
                                shutil.rmtree(saves_dir + '\\' + save)
                        else:
                            shutil.rmtree(saves_dir + '\\' + saves[save_index - 1])

            print('\nDone!\n\n')

    if scenario in ('cs-d', 'cs-w', 'rs-d', 'rs-w'):
        if not shortcuts_enabled:
            print('\nError: Shortcut feature is disabled\n\n')
        else:
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

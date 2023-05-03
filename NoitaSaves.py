import os
import shutil
import re
import win32com.client

filename = os.path.basename(__file__).removesuffix('.py')
appdata = os.getenv('APPDATA')
appdata = appdata.replace('Roaming', 'LocalLow')
game_dir = appdata + r'\Nolla_Games_Noita'
saves_dir = appdata + r'\Nolla_Games_Noita_Saves'

if not os.path.exists(saves_dir):
    os.mkdir(saves_dir)

saves = []
scenario = 'Init'
while scenario != 'e':

    print('Saves:')
    saves = os.listdir(saves_dir)
    printing_saves = [save.replace('_', ' ') for save in saves]
    for index, save in enumerate(printing_saves):
        print('#', index + 1, ' ' if index < 9 else '', ' >> ', save, sep='')
    if len(saves) == 0:
        print('<< Nothing >>')
    print()

    scenario_correct = False
    while not scenario_correct:
        scenario = input('Save (S) | Load (L) | Delete (D) | Exit (E) >> ').lower().strip()
        if scenario in ('s', 'l', 'd', 'e', 'cs-d', 'cs-w', 'rs-d', 'rs-w'):
            scenario_correct = True
        else:
            print('Error: Incorrect scenario')

    if scenario == 's':
        save_name_correct = False
        while not save_name_correct:
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
                    save_name_correct = True
        shutil.copytree(game_dir + r'\save00', saves_dir + '\\' + save_name)

    if scenario in ('l', 'd'):
        save_index_correct = False
        while not save_index_correct:
            save_index = int(input('Select the save index >> '))
            if 0 < save_index <= len(saves):
                save_index_correct = True
            else:
                print('Error: Incorrect index')
        if scenario == 'l':
            shutil.rmtree(game_dir + r'\save00')
            shutil.copytree(saves_dir + '\\' + saves[save_index - 1], game_dir + r'\save00')
        if scenario == 'd':
            shutil.rmtree(saves_dir + '\\' + saves[save_index - 1])

    if scenario in ('s', 'l', 'd'):
        print('\nDone!\n\n')

    if scenario in ('cs-d', 'cs-w', 'rs-d', 'rs-w'):
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
            shell = win32com.client.Dispatch("WScript.Shell")
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

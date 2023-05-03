import os
import shutil
import re

appdata = os.getenv('APPDATA')
appdata = appdata.replace('Roaming', 'LocalLow')
game_dir = appdata + r'\Nolla_Games_Noita'
saves_dir = appdata + r'\Nolla_Games_Noita_Saves'

scenario = 'Init'
while scenario != 'e':

    print('Saves:')
    saves = os.listdir(saves_dir)
    printing_saves = [save.replace('_', ' ') for save in saves]
    for index, save in enumerate(printing_saves):
        print('#', index + 1, ' ' if index < 9 else '', ' >> ', save, sep='')
    print()

    scenario = ''
    scenario_correct = False
    while not scenario_correct:
        scenario = input('Save (S) | Load (L) | Delete (D) | Exit (E) >> ').lower().strip()
        if scenario in ('s', 'l', 'd', 'e'):
            scenario_correct = True
        else:
            print('Error: Incorrect scenario')

    if scenario == 's':
        save_name = ''
        save_name_correct = False
        while not save_name_correct:
            save_name = input('Input save name >> ')
            save_name_errors = set(re.findall(r'[^A-Za-zА-Яа-я\- ]', save_name))
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
        save_index = 0
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
        print('Done!\n\n')
    input('Press Enter to continue...')

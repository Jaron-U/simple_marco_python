import pandas as pd
import pyautogui as pg
import pyperclip as pp
df = pd.read_excel('cmd.xlsx', sheet_name='Sheet2')
# print(df.shape[0])
# value = str(df.loc[0][0])
# print(value)
# print(value == 1)
# print(value.isdigit())
# pd.isnull(value)
# print(pd.isnull(value))
hotkeys = df.loc[0][1]
hotkeys = hotkeys.split('+')
print(len(hotkeys))
# inputcontent = df.loc[1][1]
# pp.copy(inputcontent)
# pg.hotkey(hotkeys[0], hotkeys[1], interval=0.25)
# pg.hotkey('command', 'v', interval=0.25)
# pg.press('enter', interval=1)
# pg.hotkey('tab')
# pg.hotkey('enter')
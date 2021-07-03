import clipboard
import pyautogui 
import pywinauto
import pygetwindow as gw
from more_itertools import sliced
import docx2txt
import os
os.chdir(r'C:\Users\chris\Desktop\Find and replace')


def focus_to_window(window_title=None):

    window = gw.getWindowsWithTitle(window_title)[0]
    if window.isActive == False:
        pywinauto.application.Application().connect(handle=window._hWnd).top_window().set_focus()
 
def fileToList(file):
    with open(file, 'r', encoding="latin-1") as f: 
        file_list = [line.strip() for line in f]
        return file_list
    
        
def find_and_replace_keys():
    clipboard.copy(k)
    focus_to_window(window_title='Word')
    pyautogui.hotkey('ctrl', 'h')
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.hotkey('tab')
    clipboard.copy(v)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('enter')
    pyautogui.hotkey('enter')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('enter')
    pyautogui.hotkey('ctrl', 's')
    
# def docx_current_state(file):
#     text = docx2txt.process(file)
#     replace_dict = {'e.g.': 'por exemplo',
#                 'ex.' : 'por exemplo',
#                 'Exmos.' : 'Exmos',
#                 '* '  : '',   
#                 '\t'  : '',}
#     text = replace_all(text, replace_dict)
#     text = re.sub(r'(\n+)',  '\n', text).replace('..', '.').replace(';.', '.').replace(';.', '.')
#     sentences = sent_tokenizer.tokenize(text)
#     return sentences
    
    
    
orignal_sents = fileToList('ORIGNAL_SENTS.txt')
quill_sents = fileToList('DEEPL.txt')

find_and_replace = dict(zip(orignal_sents, quill_sents))
    

for k, v in find_and_replace.items():
    if len(k) <= 255 and len(v) <= 255:
        find_and_replace_keys()
    elif 255 < len(k) <= 510 or 255 < len(v) <= 510:
        firstpart_k, secondpart_k = list(sliced((k), len(k)//2))[0], list(sliced((k), len(k)//2))[1]
        firstpart_v, secondpart_v = list(sliced((v), len(v)//2))[0], list(sliced((v), len(v)//2))[1]
        halved = {firstpart_k : firstpart_v,
                  secondpart_k : secondpart_v}
        for k, v in halved.items():
            find_and_replace_keys()
    elif 510 < len(k) <= 765 or 510 < len(v) <= 765:
        firstpart_k, secondpart_k, thirdpart_k = list(sliced((k), len(k)//3))[0], list(sliced((k), len(k)//3))[1], list(sliced((k), len(k)//3))[2]
        firstpart_v, secondpart_v, thirdpart_v = list(sliced((v), len(v)//3))[0], list(sliced((v), len(v)//3))[1], list(sliced((v), len(v)//3))[2]
        thirds = {firstpart_k : firstpart_v,
                  secondpart_k : secondpart_v,
                  thirdpart_k : thirdpart_v}
        for k, v in halved.items():
            find_and_replace_keys()

import clipboard
import pyautogui 
import pywinauto
import pygetwindow as gw
from more_itertools import sliced
import nltk
nltk.download('machado')
sent_tokenizer=nltk.data.load('tokenizers/punkt/portuguese.pickle')
import os
os.chdir(r'C:\Users\chris\Desktop\Find and replace')

# get target word document in focus
def focus_to_window(window_title=None):

    window = gw.getWindowsWithTitle(window_title)[0]
    if window.isActive == False:
        pywinauto.application.Application().connect(handle=window._hWnd).top_window().set_focus()
        
        
# convert txt file into list       
def fileToList(file):
    with open(file, 'r', encoding="latin-1") as f: 
        file_list = [line.strip() for line in f]
        return file_list

#  ui automation steps
def find_and_replace_keys():
    clipboard.copy(k)
    focus_to_window(window_title='Word')
    pyautogui.hotkey('ctrl', 'h')
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.hotkey('tab')
    clipboard.copy(v)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('tab', presses=3)
    pyautogui.press('enter', presses=2)
    pyautogui.press('tab', presses=6)
    pyautogui.hotkey('enter')
    pyautogui.hotkey('ctrl', 's')
    
    
    
# find and replace from dict
def replace_all(text, dic):
    for i, j in dic.items():
        text = text.replace(i, j)
    return text    

# return current list of sents for docx being processed
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
# orignal_docx = r'HLA.P.1653_RNT EIA Aquisição Sísmica Offshore Bloco 5-06_Final_20201105.docx'


# finding orignal sents and replacing with quill sents
find_and_replace = dict(zip(orignal_sents, quill_sents))
    

# docx_current_sents = set(docx_current_state(orignal_docx))

# ensuring no string has more than 255 chars using more_itertools slice
for k, v in find_and_replace.items():       
    if k != v: 
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

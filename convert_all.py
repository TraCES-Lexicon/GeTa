import os
import shutil
import re
import pyautogui
import subprocess
import time
from zipfile import ZipFile


def launch_geta():
    si = subprocess.STARTUPINFO()
    si.dwFlags = subprocess.STARTF_USESHOWWINDOW
    si.wShowWindow = 3  # SW_MAXIMIZE
    proc = subprocess.Popen('java -jar GeTa\\Tool\\TaggingToolVer6_16October19.jar',
                            startupinfo=si)
    time.sleep(3)
    return proc


def launch_pepper():
    si = subprocess.STARTUPINFO()
    si.dwFlags = subprocess.STARTF_USESHOWWINDOW
    si.wShowWindow = 3  # SW_MAXIMIZE
    proc = subprocess.Popen('..\\pepper-grinder\\pepper-grinder.exe',
                            startupinfo=si)
    time.sleep(2)
    return proc


def unzip_all(dirname):
    """
    Walk over the folder and its subfolders and unzip all zip archives.
    """
    for root, dirs, files in os.walk(dirname):
        for fname in files:
            if not fname.lower().endswith('.zip'):
                continue
            fname = os.path.join(root, fname)
            with ZipFile(fname, 'r') as zipObj:
                print('Extracting ' + fname + '...')
                zipObj.extractall(path=root)


def geta2tei(dirname):
    """
    Open all GeTa files, export them to TEI XML and to the
    intermediary format required by Pepper Grinder.
    """
    textDirs = []
    for root, dirs, files in os.walk(dirname):
        if root.startswith('Export'):
            continue
        for fname in files:
            if not fname.lower().endswith('.json'):
                continue
            fnameFull = os.path.abspath(os.path.join(root, fname))
            print('Opening ' + fnameFull + ' in GeTa...')
            proc = launch_geta()
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.5)
            pyautogui.write(fnameFull)
            pyautogui.press('enter')
            time.sleep(70)  # GeTa is not exactly a fast application
            print('Exporting ' + fnameFull + ' for Pepper Grinder...')
            pyautogui.hotkey('alt', 'f')
            time.sleep(0.2)
            pyautogui.hotkey('alt', 'e')
            time.sleep(0.2)
            pyautogui.press('enter')
            time.sleep(10)
            for exportDir in os.listdir(root):
                if exportDir.startswith('Export') and os.path.isdir(os.path.join(root, exportDir)):
                    os.rename(os.path.isdir(os.path.join(root, exportDir)),
                              os.path.isdir(os.path.join(root, 'Export')))
                    break
            print('Done.')
            print('Saving ' + fnameFull + ' in TEI XML...')
            pyautogui.hotkey('alt', 'f')
            time.sleep(0.2)
            pyautogui.hotkey('alt', 'e')
            time.sleep(0.2)
            pyautogui.hotkey('alt', 't')
            time.sleep(0.2)
            pyautogui.press('enter')
            time.sleep(10)
            print('Done.')
            proc.terminate()
            textDirs.append(os.path.abspath(root))
    print('GeTa > TEI conversion finished. Here are the text folders:')
    for textDir in textDirs:
        print(textDir)


def relocate_annis(dirName):
    """
    Move exported ANNIS files from the Pepper Grinder's output folder
    to the folder where they should be.
    """
    textName = os.listdir('output')[0]
    annisName = os.listdir(os.path.join('output', textName, 'annis'))[0]
    srcDir = os.path.join('output', textName, 'annis', annisName)
    targetDir = dirName[:-6] + 'annis'
    shutil.copytree(srcDir, targetDir)
    shutil.rmtree(os.path.join('output', textName))


def geta2annis(dirname):
    """
    Open all intermediary GeTa files in Pepper Grinder and transform them into ANNIS format.
    """
    textDirs = []
    for root, dirs, files in os.walk(dirname):
        if not root.endswith('Export'):
            continue
        noJson = True
        for fname in files:
            if not fname.lower().endswith('.json'):
                continue
            noJson = False
            break
        if noJson:
            continue
        dirnameFull = os.path.abspath(root)
        print('Opening ' + dirnameFull + ' in Pepper Grinder...')
        proc = launch_pepper()
        pyautogui.click(40, 131)
        time.sleep(0.5)
        pyautogui.click(40, 131)
        for i in range(3):
            pyautogui.press('tab')
            time.sleep(0.2)
        pyautogui.write(dirnameFull)
        time.sleep(0.5)
        okLocation = pyautogui.locateOnScreen('img/ok_pepper.png')
        okX, okY = pyautogui.center(okLocation)
        pyautogui.click(okX, okY)
        time.sleep(0.5)
        pyautogui.click(40, 162)
        time.sleep(30)
        print('Done.')
        proc.terminate()
        relocate_annis(root)
        textDirs.append(os.path.abspath(root))
    print('GeTa > ANNIS conversion finished. Here are the text folders:')
    for textDir in textDirs:
        print(textDir)


def create_new_dir():
    """
    Create an empty directory for all annotation files.
    """
    if os.path.exists('Annotations_processed'):
        shutil.rmtree('Annotations_processed')
    shutil.copytree('Annotations', 'Annotations_processed')


def clean_dir(dirname):
    """
    Rename folders and files with a whitespace in them.
    """
    for root, dirs, files in os.walk(dirname, topdown=False):
        for fname in files:
            newFileName = fname.replace(' ', '_')
            if newFileName != fname:
                os.rename(os.path.join(root, fname), os.path.join(root, newFileName))
        newDirName = re.sub('[^/\\\\]+$', lambda m: m.group(0).replace(' ', '_'), root)
        if newDirName != root:
            os.rename(root, newDirName)
    for root, dirs, files in os.walk(dirname):
        for fname in files:
            if fname.lower().endswith(('.json', '.ann', '.ind')):
                continue
            os.remove(os.path.join(root, fname))


if __name__ == '__main__':
    create_new_dir()
    unzip_all('Annotations_processed')
    clean_dir('Annotations_processed')
    geta2tei('Annotations_processed')
    geta2annis('Annotations_processed')

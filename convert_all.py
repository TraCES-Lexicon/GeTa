import os
import shutil
import re
import pyautogui
import subprocess
import time
from zipfile import ZipFile


curDir = os.path.abspath('.')
pepperDir = os.path.abspath('../pepper-grinder')


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
        if root.startswith(('Export', 'TEI')):
            continue
        for fname in files:
            if not fname.lower().endswith('.json'):
                continue
            fileSize = os.path.getsize(os.path.join(root, fname))
            durOpen = max(30.0, 10 + fileSize * 5e-06)
            durConvertPepper = 5 + fileSize * 1.2e-06
            durConvertTEI = 5 + fileSize * 5e-07
            fnameFull = os.path.abspath(os.path.join(root, fname))
            print('Opening ' + fnameFull + ' in GeTa (' + str(durOpen) + ' seconds)...')
            proc = launch_geta()
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.5)
            pyautogui.write(fnameFull)
            pyautogui.press('enter')
            time.sleep(durOpen)  # GeTa is not exactly a fast application
            print('Exporting ' + fnameFull + ' for Pepper Grinder (' + str(durConvertPepper) + ' seconds)...')
            pyautogui.hotkey('alt', 'f')
            time.sleep(0.2)
            pyautogui.hotkey('alt', 'e')
            time.sleep(0.2)
            pyautogui.press('enter')
            time.sleep(durConvertPepper)
            for exportDir in os.listdir(root):
                if exportDir.startswith('Export') and os.path.isdir(os.path.join(root, exportDir)):
                    os.rename(os.path.join(root, exportDir),
                              os.path.join(root, 'Export'))
                    break
            print('Done.')
            print('Saving ' + fnameFull + ' in TEI XML (' + str(durConvertTEI) + ' seconds)...')
            pyautogui.hotkey('alt', 'f')
            time.sleep(0.2)
            pyautogui.hotkey('alt', 'e')
            time.sleep(0.2)
            pyautogui.hotkey('alt', 't')
            time.sleep(0.2)
            pyautogui.press('enter')
            time.sleep(durConvertTEI)
            os.makedirs(os.path.join(root, 'TEI'))
            for teiFname in os.listdir(root):
                if teiFname.endswith('.xml'):
                    shutil.move(os.path.join(root, teiFname),
                                os.path.join(root, 'TEI', teiFname))
            print('Done.')
            proc.terminate()
            textDirs.append(os.path.abspath(root))
        # if len(textDirs) > 0:
        #     break
    print('GeTa > TEI conversion finished. Here are the text folders:')
    for textDir in textDirs:
        print(textDir)
    time.sleep(1)


def relocate_annis(dirName):
    """
    Move exported ANNIS files from the Pepper Grinder's output folder
    to the folder where they should be.
    """
    outputDir = os.path.join(pepperDir, 'output')
    textName = os.listdir(outputDir)[0]
    annisName = os.listdir(os.path.join(outputDir, textName, 'annis'))[0]
    srcDir = os.path.join(outputDir, textName, 'annis', annisName)
    targetDir = dirName[:-6] + 'annis'
    shutil.copytree(srcDir, targetDir)
    shutil.rmtree(os.path.join(outputDir, textName))


def geta2annis(dirname):
    """
    Open all intermediary GeTa files in Pepper Grinder and convert them to ANNIS format.
    """
    textDirs = []
    for root, dirs, files in os.walk(dirname, topdown=False):
        if not root.endswith('Export'):
            continue
        noJson = True
        for fname in files:
            if not fname.lower().endswith('.json'):
                continue
            noJson = False
            fileSize = os.path.getsize(os.path.join(root, fname))
            durConvert = max(40.0, 20 + fileSize * 6e-06)
            break
        if noJson:
            continue
        dirnameFull = os.path.abspath(root)
        print('Opening ' + dirnameFull + ' in Pepper Grinder (' + str(durConvert) + ' seconds)...')
        os.chdir(pepperDir)
        if os.path.exists('output'):
            shutil.rmtree('output')
        proc = launch_pepper()
        pyautogui.click(40, 131)
        time.sleep(0.5)
        pyautogui.click(40, 131)
        for i in range(3):
            pyautogui.press('tab')
            time.sleep(0.2)
        pyautogui.write(dirnameFull)
        time.sleep(0.5)
        okLocation = pyautogui.locateOnScreen(os.path.join(curDir, 'img/ok_pepper.png'))
        if okLocation is None:
            # dropdown list can hide the button
            pyautogui.press('esc')
            time.sleep(0.2)
            okLocation = pyautogui.locateOnScreen(os.path.join(curDir, 'img/ok_pepper.png'))
        okX, okY = pyautogui.center(okLocation)
        pyautogui.click(okX, okY)
        time.sleep(0.5)
        pyautogui.click(40, 162)
        time.sleep(durConvert)
        print('Done.')
        proc.terminate()
        time.sleep(0.5)
        os.chdir(curDir)
        relocate_annis(root)
        textDirs.append(os.path.abspath(root))
        shutil.rmtree(root)
        # break
    print('GeTa > ANNIS conversion finished. Here are the text folders:')
    for textDir in textDirs:
        print(textDir)
    time.sleep(1)


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


def relocate_geta(dirname):
    """
    Move GeTa files to separate GeTa subfolders.
    """
    for root, dirs, files in os.walk(dirname):
        if re.search('[/\\\\](?:annis|TEI|Export|GeTa)\\b', root) is not None:
            continue
        os.makedirs(os.path.join(root, 'GeTa'))
        for fname in files:
            if os.path.isdir(os.path.join(root, fname)):
                continue
            shutil.move(os.path.join(root, fname),
                        os.path.join(root, 'GeTa', fname))


if __name__ == '__main__':
    create_new_dir()
    unzip_all('Annotations_processed')
    clean_dir('Annotations_processed')
    geta2tei('Annotations_processed')
    geta2annis('Annotations_processed')
    relocate_geta('Annotations_processed')

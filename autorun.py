from tkinter import filedialog
import os
import pyautogui
import pyperclip
import time
import subprocess
import pickle



try:
    with open(os.path.join(os.path.dirname(__file__), "store.pickle"), 'rb') as f:
        store = pickle.load(f)

    sequatorpath = store[0]
    dir = store[1]
except:
    sequatorpath = "C:\\sequator.exe"
    dir = 'C:\\'

print(f"sequator: {sequatorpath}")

if not os.path.isfile(sequatorpath):
    typ = [('sequator', '*.exe')]
    sequatorpath = filedialog.askopenfilename(title="Sequator実行ファイルを選択", filetypes = typ, initialdir = __file__)


typ = [('image', '*.tif'),('image', '*.jpg')]

dmfiles = filedialog.askopenfilenames(title="星空画像を選択", filetypes = typ, initialdir = dir) 

for f in dmfiles:
    print(f)

if str.lower(os.path.splitext(dmfiles[0])[-1]) in [".jpg",".jpeg"]:
  outformat = ".jpg"
else:
  outformat = ".tif"

typ = [('image', '*.png')]
dir = os.path.dirname(dmfiles[0])
maskfile = filedialog.askopenfilename(title="マスク画像を選択", filetypes = typ, initialdir = dir)

print(maskfile)
projectfolder = filedialog.askdirectory(title="プロジェクトフォルダーを選択", initialdir = dir)

sepfolder = os.path.join(projectfolder,"sep")
resultfolder = os.path.join(projectfolder,"result")

os.makedirs(sepfolder, exist_ok=True)
os.makedirs(resultfolder, exist_ok=True)

with open(os.path.join(os.path.dirname(__file__), "store.pickle"), 'wb') as f:
    pickle.dump([sequatorpath,dir], f)

for i in range(len(dmfiles)):
    if i==0:
        stars=dmfiles[0:5]
        basestar=dmfiles[0]
    elif i==1:
        stars=dmfiles[0:5]
        basestar=dmfiles[1]
    elif i==len(dmfiles)-2:
        stars=dmfiles[-5:]
        basestar=dmfiles[-2]
    elif i==len(dmfiles)-1:
        stars=dmfiles[-5:]
        basestar=dmfiles[-1]
    else:
        stars=dmfiles[i-2:i+3]
        basestar=dmfiles[i]

    xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<SequatorProject version="1.0">
    <BaseImage>
        <AbsolutePath>{basestar}</AbsolutePath>
    </BaseImage>
    <StarImage>
        <AbsolutePath>{stars[0]}</AbsolutePath>
    </StarImage>
    <StarImage>
        <AbsolutePath>{stars[1]}</AbsolutePath>
    </StarImage>
    <StarImage>
        <AbsolutePath>{stars[2]}</AbsolutePath>
    </StarImage>
    <StarImage>
        <AbsolutePath>{stars[3]}</AbsolutePath>
    </StarImage>
    <StarImage>
        <AbsolutePath>{stars[4]}</AbsolutePath>
    </StarImage>
    <Output>
        <AbsolutePath>{os.path.join(resultfolder, os.path.splitext(os.path.basename(basestar))[0]+outformat)}</AbsolutePath>
    </Output>
    <UnifyExposure>false</UnifyExposure>
    <HomogenizeVignetting>false</HomogenizeVignetting>
    <CompositionMode max="1">0</CompositionMode>
    <IntegrationMode max="3">2</IntegrationMode>
    <SigmaIndex max="4">2</SigmaIndex>
    <FreezeGroundSelective>true</FreezeGroundSelective>
    <DumpAsLinear>true</DumpAsLinear>
    <TrailsMotionEffect>false</TrailsMotionEffect>
    <AutoBrightness>false</AutoBrightness>
    <HDR>false</HDR>
    <RemoveHotPixels>false</RemoveHotPixels>
    <MergePixels>false</MergePixels>
    <DistortionCorrection max="2">0</DistortionCorrection>
    <ReducePollution>false</ReducePollution>
    <ReducePollutionMode max="1">1</ReducePollutionMode>
    <ReducePollutionStrength max="4">2</ReducePollutionStrength>
    <AggressiveSuppression>false</AggressiveSuppression>
    <EnhanceStars>false</EnhanceStars>
    <EnhanceStarsStrength max="4">3</EnhanceStarsStrength>
    <StarBound>
        <X1>0.000000</X1>
        <Y1>0.000000</Y1>
        <X2>0.000000</X2>
        <Y2>0.000000</Y2>
    </StarBound>
    <GradientRegion>
        <X1>0.000000</X1>
        <Y1>0.000000</Y1>
        <X2>0.000000</X2>
        <Y2>0.000000</Y2>
        <Height>0.000000</Height>
    </GradientRegion>
    <Mask>
        <AbsolutePath>{maskfile}</AbsolutePath>
    </Mask>
    <SkyRegionMode max="3">3</SkyRegionMode>
    <TimeLapse>false</TimeLapse>
    <TimeLapseFrames max="30">5</TimeLapseFrames>
    <ColorSpace max="2">0</ColorSpace>
</SequatorProject>"""
    xmlpath = os.path.join(sepfolder, os.path.splitext(os.path.basename(basestar))[0]+".sep")
    f = open(xmlpath, 'w', encoding='utf-8')
    f.write(xml)
    f.close()
    if i == 0:
        subprocess.Popen(sequatorpath)
    time.sleep(1)
    pyautogui.press('alt')
    pyautogui.press('p')
    pyautogui.press('o')
    if i == 0:
        pyautogui.hotkey('alt', 'd')
        print(sepfolder)
        pyperclip.copy(sepfolder)
        time.sleep(1)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(1)
        pyautogui.hotkey('alt', 'n')
        time.sleep(1)
    pyperclip.copy(os.path.basename(xmlpath))
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    
    while True:
        position=pyautogui.locateOnScreen(os.path.join(os.path.dirname(__file__), "autorun", "load.png") , confidence=0.9)
        if position==None:
            position=pyautogui.locateOnScreen(os.path.join(os.path.dirname(__file__), "autorun", "start.png") , confidence=0.9)
            x, y = pyautogui.center(position)
            pyautogui.click(x, y)
            break
        else:
            time.sleep(1)
    
    while True:
        position=pyautogui.locateOnScreen(os.path.join(os.path.dirname(__file__), "autorun", "close.png") , confidence=0.9)
        if position==None:
            time.sleep(1)
        else:
            x, y = pyautogui.center(position)
            pyautogui.click(x, y)
            break
    


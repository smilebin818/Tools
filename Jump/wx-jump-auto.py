# -*- coding:utf-8 -*-
#!/usr/bin/env python
import pymouse, time, pyHook, pythoncom, aircv, math, sys

import aircv, pymouse
import pyscreenshot as ImageGrab

global start_pos, end_pos

start_pos = None
end_pos = None

m = pymouse.PyMouse()

def getImg():
    im = ImageGrab.grab()
    im.save('3.png')

def getpoint(): 
    global start_pos

    getImg()

    imsrc = aircv.imread('3.png')
    imobj = aircv.imread('2.png')

    # ac.find_sift()
    # rect = aircv.Rect(left=80, top=10, width=50, height=90)
    getpoint = aircv.find_sift(imsrc, imobj)

    y = getpoint['rectangle'][1][1] - 10
    x = getpoint['rectangle'][0][0] + (getpoint['rectangle'][2][0] - getpoint['rectangle'][0][0])/2
    m.move(x, y)

    start_pos = (x, y)

def main():
    hm = pyHook.HookManager()
    hm.KeyDown = onKeyboardEvent

    try:
        hm.HookKeyboard()
    except Exception as e:
        pass

    pythoncom.PumpMessages()

def onKeyboardEvent(event):

    global end_pos
    global start_pos

    if event.Key == 'Space':
        end_pos = m.position()

        if start_pos and end_pos:
            dis = int(math.sqrt(math.pow(start_pos[0]-end_pos[0],2)+math.pow(start_pos[1]-end_pos[1],2)))
            times = round((dis / 0.3) / 1000 , 3) 
            if dis < 500:
                print '>> ', times
                m.press(end_pos[0], end_pos[1])
                time.sleep(times)
                m.release(end_pos[0], end_pos[1])
                start_pos = None
                end_pos = None

                time.sleep(0.5)
                getpoint()
        else:
            getpoint()
    if event.Key == 'Escape':
        sys.exit()
    
    return True

if __name__ == '__main__':
    main()

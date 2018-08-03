import sys

sleep(5)
#Excel1 -> Excel2 -> Excel1を選んで、Alt+tabで移動できるようにする。
#選択位置はExcel1:コピー開始、Excel2:ペースト開始位置にする

for i in range(0, 58):
    #copy       
    keyDown(Key.CTRL)
    type("c")
    keyUp(Key.CTRL)
    #cursol right
    type(Key.RIGHT)
    #Excel1 -> Excel2
    keyDown(Key.ALT)
    type(Key.TAB)
    keyUp(Key.ALT)
    #paste
    keyDown(Key.CTRL)
    type("v")
    keyUp(Key.CTRL)
    #cursol right
    type(Key.RIGHT)
    #Excel2 -> Excel1
    keyDown(Key.ALT)
    type(Key.TAB)
    keyUp(Key.ALT)
    #wait
    sleep(1)

from pymouse import PyMouse
from pykeyboard import PyKeyboard
import time
import win32api
import win32con

import win32gui


DEBUG = 0
CAPTURE = 0

T = 2 if DEBUG else 45

if CAPTURE:
    while True:
        time.sleep(1)
        a,b, (x,y) = (win32gui.GetCursorInfo())
        if x==0 and y==0:
            exit()
        else:
            print(x,y)

m = PyMouse()
k = PyKeyboard()

# 评论框: 618,660
# 评论发布: 836, 620
# 收藏: 803, 660
# 分享: 841, 660
# 综合 664,110
# 订阅 790, 110
# 水平菜单宽度： 42
# 纵向条目高度 ： 70
# 纵向起始坐标：150
# 水平起始坐标： 

pinglun = 2,2
height = 65 # 文章条目高度"
start_height = 195  # 第一条文  章条目纵向坐标"

def click():
    left_down()
    left_up()

def delay(s):
    time.sleep(s)

def move(x, y):
    win32api.SetCursorPos([x, y])
    click()
    delay(3)

def 评论():
    move(618,660)
    k.type_string('zhichixidada')
    time.sleep(0.5)
    k.tap_key(k.space_key)
    time.sleep(0.5)
    move(836, 620)



def MouseWheel(a):
    win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL,0,0, a)


def left_down():
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)

def left_up():
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)



def 视听学习2个视频():
    主界面()
    视听学习()
    视频1()
    阅读45秒()
    回退()
    
    视频2()
    阅读45秒()
    回退()


def 看百灵视频():
    关注()
    delay(1)
    for i in range(5):
    	move(533+i*43,110)
    	move(666, 180)
    	观看视频()
    	回退()

    	move(666, 420)
    	观看视频()
    	回退()

def 阅读综合3篇文章():
    主界面()
    综合()
    off=0
    通用(0,off)
    阅读45秒()
    评论()
    回退()
    
    通用(1,off)
    阅读45秒()
    回退()

    通用(2,off)
    阅读45秒()
    收藏()
    回退()

def 阅读订阅4篇文章():
    主界面()
    订阅()
    off = 80
    通用(0,off)
    阅读45秒()
    收藏()
    分享()
    回退()
    
    通用(1,off)
    阅读45秒()
    收藏()
    分享()
    回退()

    通用(2,off)
    阅读45秒()
    收藏()
    回退()
    
    通用(3,off)
    阅读45秒()
    回退()

def 阅读分享推荐4篇文章():
    主界面()
    推荐()
    
    推荐0()
    阅读45秒()
    评论()
    收藏()
    分享()
    回退()
    
    推荐1()
    阅读45秒()
    评论()
    收藏()
    分享()
    回退()

    推荐2()
    阅读45秒()
    收藏()
    回退()
    
    推荐3()
    阅读45秒()
    回退()


def 阅读新思想3篇文章():
    主界面()
    新思想()
    
    通用1()
    阅读45秒()
    评论()
    收藏()
    回退()
    
    通用2()
    阅读45秒()
    收藏()
    回退()
    
    通用3()
    阅读45秒()
    收藏()
    回退()
    


def 阅读要闻3篇文章():
    主界面()
    要闻()
    
    通用1()
    阅读45秒()
    评论()
    收藏()
    回退()
    
    通用2()
    阅读45秒()
    回退()
    
    通用3()
    阅读45秒()
    回退()
    


def 收藏():
    move(805, 660)
        


def 分享():
    delay(0.5)
    move(841, 660) #分享
    
    move(548, 468) #分享到学习强国
    
    move(662, 269) #选择人社局党支部
    
    move(785, 434) #发送


def 回退():
    move(522, 73) #回退
        



def 主界面():
    move(681, 657) #主界面
        


def 向右拖曳菜单():
    m.move(623, 105)
    left_down()
    m.move(643, 105)
    delay(0.1)
    m.move(663, 105)
    delay(0.1)
    m.move(683, 105)
    delay(0.1)
    m.move(700, 105)
    delay(1)
    m.move(829, 105)
    left_up()
    delay(0.5)



def 视听学习():
    move(754, 656) #视听学习
        


def 视频1():
    move(690, 233)
        


def 视频2():
    move(690, 479)
        


def 听3首歌():
    主界面()
    视听学习()
    向右拖曳菜单()
    move(766, 104) #听音乐
        
    首页音乐()
    音乐1()
    收起音乐播放界面()
    音乐2()
    收起音乐播放界面()
    音乐3()
    收起音乐播放界面()

def 首页音乐():
    move(643, 200)
        

def 音乐1():
    move(601, 274)
        
    阅读45秒()
    move(524, 74)


def 音乐2():
    move(575, 321)
        
    阅读45秒()
    move(524, 74)


def 音乐3():
    move(613, 365)
        
    阅读45秒()
    move(524, 74)


def 收起音乐播放界面():
    move(522, 77)
    delay(1)

def 订阅():
    move(790, 110)
        


def 要闻():
    move(567, 107)
        


def 新思想():
    move(607, 109)
        

def 综合():
    move(664, 110)
        

def 关注():
    move(611, 656)
        



def 阅读45秒():

    delay(T)
    if not DEBUG:
        MouseWheel(-1)
        delay(4)
        MouseWheel(-1)
        delay(4)
        MouseWheel(-1)
        delay(4)
        MouseWheel(-1)
    delay(T)

def 观看视频():
    if not DEBUG:
        delay(20)
    delay(T)

def 推荐():
    move(530, 107)
        

def 通用(n, offset=0):
    move(640, start_height+offset+height*n)
        

def 通用1():
    move(640, start_height)
        


def 通用2():
    move(640, start_height + height)
        


def 通用3():
    move(640, start_height + height*2)
        


def 推荐0():
    move(643, 200)
        


def 推荐1():
    move(672, 405)
        


def 推荐2():
    move(742, 482)
        


def 推荐3():
    move(691, 541)
        

def main():
    delay(2)
    主界面()
    看百灵视频()
    # 阅读分享推荐4篇文章()
    阅读新思想3篇文章()
    阅读综合3篇文章()
    阅读订阅4篇文章()

    视听学习2个视频()
    

# 通用(0, 80)
main()

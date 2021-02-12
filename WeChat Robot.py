import pyautogui as auto
import time
import pyperclip
import win32api
class WechatRobot():
    '''
    PC微信机器人\n
    start_address:微信.exe路径(例如：'C:\Program Files (x86)\Tencent\WeChat\WeChat.exe')\n
    wechat_locate1_address:定位图标1路径  (搜索框右侧+号)\n
    wechat_locate2_address:定位图标2路径  (聊天气泡右下角)\n
    wechat_locate3_address:定位图标3路径  (登录按钮)\n
    wechat_locate4_address:定位图标4路径  (新消息气泡)\n
    wechat_locate5_address:定位图标5路径  ('群聊名称' 图片)\n
    wechat_locate6_address:定位图标6路径  (置顶图标)\n
    '''
    def __init__(self):
        #微信.exe路径
        self.start_address="Add your wechat.exe address"
        #搜索框定位图标（加号）
        self.wechat_locate1_address="Add your picture1 path"
        #聊天气泡定位图标（右下角）
        self.wechat_locate2_address="Add your picture2 path"
        #登录按钮
        self.wechat_locate3_address="Add your picture3 path"
        #新消息气泡
        self.wechat_locate4_address="Add your picture4 path"
        #群聊名称 图片
        self.wechat_locate5_address="Add your picture5 path"
        #置顶 图标
        self.wechat_locate6_address="Add your picture6 path"
    
    def back_to_desktop(self):
        '''
        功能说明：模拟按键win+d返回桌面
        '''
        auto.hotkey('win', 'd')
    
    def open_app(self,cue=True):
        '''
        根据路径打开app\n
        address为exe路径，例如：'C:\Program Files (x86)\Tencent\WeChat\WeChat.exe'
        '''
        win32api.ShellExecute(0,'open',self.start_address,'','',3)
        #最后一个参数表示窗口属性 0不显示 1正常 2最小化 3最大化
        if cue:
            time.sleep(0.3)#等待微信启动
            
            msgList=["请在手机上登录微信","请检查是否登录"]
            buttonList=["我已登陆","忽略"]
            flag=1
            while flag:
                if auto.locateCenterOnScreen(self.wechat_locate1_address):
                    flag=0
                    #判断是否已经打开微信界面，通过定位图标1来确定
                
                else:
                    #如果没有打开界面，则提示登录
                    if flag==1:
                        try:
                            a=auto.locateCenterOnScreen(self.wechat_locate3_address)
                            auto.click(a.x,a.y)
                        except:
                            pass
                        res=auto.confirm(text=msgList[0],buttons=buttonList)
                    else:
                        res=auto.confirm(text=msgList[1],buttons=buttonList)
                    
                    if res=="忽略":
                        flag=0
                    else:
                        flag+=1
        
        sd_button=auto.locateCenterOnScreen(self.wechat_locate6_address)
        try:
            auto.click(sd_button.x,sd_button.y)
        except:
            pass
    
    def find(self,name="",key=True):
        '''
        功能描述:利用微信搜索框查找人名，可直接切换到聊天框\n
        参数说明:\n
        name:搜索的人名\n
        key:是否直接切换到聊天界面(默认True)
        '''
        wechat_locate1=auto.locateCenterOnScreen(self.wechat_locate1_address)
        auto.click(wechat_locate1.x-40,wechat_locate1.y)#先清空
        #print("清空搜索框")
        auto.click(wechat_locate1.x-60,wechat_locate1.y)#然后获取焦点
        auto.click()
        #print("点击搜索框")
        if name !="":
            pyperclip.copy(name)
            auto.hotkey("ctrl","v")

        #print("等待%f秒"%t)
        if key==True:
            t=0.5#设置等待时间 0.5秒后按下enter可直接切换到首个对象
            time.sleep(t)
            auto.press("enter")
            #print("按下enter")
    
    def send_msg(self,msg,name=""):
        '''
        发送消息\n
        msg:需要发送的消息(str)\n
        name:可选参数，发送的对象(str)
        '''
        if name !="":
            self.find(name)
        time.sleep(0.2)
        pyperclip.copy(msg)
        auto.hotkey("ctrl","v")#粘贴
        auto.hotkey("alt","s")#模拟发送
    
    def recive_msg(self,name="",num=1):
        '''
        接收消息\n
        name:接受的对象\n
        num:接受的数量(默认为1条)\n
        返回值:消息列表(按时间由远到近排列)\n
        '''
        msg_list=[]
        if name!="":
            self.find(name)

        #首先定位以便缩小范围
        d=auto.locateCenterOnScreen(self.wechat_locate1_address)
        
        #如果当前界面内的消息数量少于num
        while len(msg_list)<num:
            print("少于")
            c=auto.locateAllOnScreen(self.wechat_locate2_address,region=(d.x+110,d.y+30,310,760))
            c=list(c)
            
            #当前页面的消息列表
            temp_msg_list=[]
            
            for point in c:
                auto.moveTo(point.left-10,point.top-2)
                auto.click(clicks=2)
                auto.hotkey("ctrl","c")
                temp_msg_list.append(pyperclip.paste())
            #构造返回值列表
            msg_list=temp_msg_list+msg_list
            if len(msg_list)<num:
                try:
                    auto.moveTo(c[0].left,c[0].top)
                except:
                    pass
                auto.scroll(900)
                time.sleep(1.5)
        for i in range(len(msg_list)):
            if msg_list[i]=="":
                msg_list[i]="<表情>"
        return msg_list[-num:]
    
    def get_name(self):
        lc1=auto.locateCenterOnScreen(self.wechat_locate1_address)
        auto.click(x=lc1.x+80,y=lc1.y,clicks=1)
        time.sleep(0.5)
        lc2=auto.locateCenterOnScreen(self.wechat_locate5_address)
        if lc2==None:
            auto.click(lc1.x+650,lc1.y,clicks=1)
            auto.click(lc1.x+700,lc1.y+60,clicks=2)
            auto.hotkey('ctrl','c')
            print(1)
        else:
            #auto.click(x=lc1.x+80,y=lc1.y,clicks=1)
            auto.click(lc2.x-30,lc2.y+35,clicks=3)
            auto.hotkey('ctrl','c')
            print(2)
        return pyperclip.paste()
    
    def acceptNewmsg(self,total=5):
        '''
        接受新消息，返回消息字典 如：{'肖发博': ['q', 'q', '，', '1', '1', '，', '哈哈哈', '嘻嘻'],'小明':['最近过的怎么样？','想死你了']}\n
        total 接受最新的几条消息默认为5
        '''
        newmsg_dic={}#新消息字典 
        new_msg=auto.locateAllOnScreen(self.wechat_locate4_address)
        new_msg=list(new_msg)
        #print("共%d个"%len(new_msg))
        for i in range(len(new_msg)):
            print("正在操作第%d个用户"%(i+1))
            auto.moveTo(new_msg[i].left,new_msg[i].top)
            auto.click()
            name=self.get_name()
            sub=self.recive_msg(num=total)
            newmsg_dic[name]=sub
        return newmsg_dic

    

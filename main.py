# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import time, win32con, win32api, win32gui, ctypes
from pywinauto import clipboard  # 채팅창내용 가져오기 위해
import pandas as pd  # 가져온 채팅내용 DF로 쓸거라서

import firebase_admin
from firebase_admin import credentials, db
import random

# Firebase에 접근하기 위한 ServiceAccountkey 권한 정보가 있어야 함.
cred = credentials.Certificate('./ServiceAccountKey.json')
firebase_admin.initialize_app(cred, {
	'databaseURL' : 'https://what-should-we-eat-today-default-rtdb.firebaseio.com/'
})

# # 카톡창 이름, (활성화 상태의 열려있는 창)
kakao_opentalk_name = '박정현'

PBYTE256 = ctypes.c_ubyte * 256
_user32 = ctypes.WinDLL("user32")
GetKeyboardState = _user32.GetKeyboardState
SetKeyboardState = _user32.SetKeyboardState
PostMessage = win32api.PostMessage
SendMessage = win32gui.SendMessage
FindWindow = win32gui.FindWindow
IsWindow = win32gui.IsWindow
GetCurrentThreadId = win32api.GetCurrentThreadId
GetWindowThreadProcessId = _user32.GetWindowThreadProcessId
AttachThreadInput = _user32.AttachThreadInput

MapVirtualKeyA = _user32.MapVirtualKeyA
MapVirtualKeyW = _user32.MapVirtualKeyW

MakeLong = win32api.MAKELONG
w = win32con

# # 채팅방에 메시지 전송
def kakao_sendtext(chatroom_name, text):
    # # 핸들 _ 채팅방
    hwndMain = win32gui.FindWindow(None, chatroom_name)
    hwndEdit = win32gui.FindWindowEx(hwndMain, None, "RICHEDIT50W", None)

    win32api.SendMessage(hwndEdit, win32con.WM_SETTEXT, 0, text)
    SendReturn(hwndEdit)

# # 채팅내용 가져오기
def copy_chatroom(chatroom_name):
    # # 핸들 _ 채팅방
    hwndMain = win32gui.FindWindow(None, chatroom_name)
    hwndListControl = win32gui.FindWindowEx(hwndMain, None, "EVA_VH_ListControl_Dblclk", None)

    # #조합키, 본문을 클립보드에 복사 ( ctl + c , v )
    PostKeyEx(hwndListControl, ord('A'), [w.VK_CONTROL], False)
    time.sleep(0.5)
    PostKeyEx(hwndListControl, ord('C'), [w.VK_CONTROL], False)
    ctext = clipboard.GetData()
    # print(ctext)
    return ctext

# 조합키 쓰기 위해
def PostKeyEx(hwnd, key, shift, specialkey):
    if IsWindow(hwnd):

        ThreadId = GetWindowThreadProcessId(hwnd, None)

        lparam = MakeLong(0, MapVirtualKeyA(key, 0))
        msg_down = w.WM_KEYDOWN
        msg_up = w.WM_KEYUP

        if specialkey:
            lparam = lparam | 0x1000000

        if len(shift) > 0:
            pKeyBuffers = PBYTE256()
            pKeyBuffers_old = PBYTE256()

            SendMessage(hwnd, w.WM_ACTIVATE, w.WA_ACTIVE, 0)
            AttachThreadInput(GetCurrentThreadId(), ThreadId, True)
            GetKeyboardState(ctypes.byref(pKeyBuffers_old))

            for modkey in shift:
                if modkey == w.VK_MENU:
                    lparam = lparam | 0x20000000
                    msg_down = w.WM_SYSKEYDOWN
                    msg_up = w.WM_SYSKEYUP
                pKeyBuffers[modkey] |= 128

            SetKeyboardState(ctypes.byref(pKeyBuffers))
            time.sleep(0.01)
            PostMessage(hwnd, msg_down, key, lparam)
            time.sleep(0.01)
            PostMessage(hwnd, msg_up, key, lparam | 0xC0000000)
            time.sleep(0.01)
            SetKeyboardState(ctypes.byref(pKeyBuffers_old))
            time.sleep(0.01)
            AttachThreadInput(GetCurrentThreadId(), ThreadId, False)

        else:
            SendMessage(hwnd, msg_down, key, lparam)
            SendMessage(hwnd, msg_up, key, lparam | 0xC0000000)

# # 엔터
def SendReturn(hwnd):
    win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
    time.sleep(0.01)
    win32api.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)

# # 채팅방 열기
def open_chatroom(chatroom_name):
    # # # 채팅방 목록 검색하는 Edit (채팅방이 열려있지 않아도 전송 가능하기 위하여)
    hwndkakao = win32gui.FindWindow(None, "카카오톡")
    time.sleep(0.05)
    hwndkakao_edit1 = win32gui.FindWindowEx(hwndkakao, None, "EVA_ChildWindow", None)
    time.sleep(0.05)
    hwndkakao_edit2_1 = win32gui.FindWindowEx(hwndkakao_edit1, None, "EVA_Window", None)
    time.sleep(0.05)
    hwndkakao_edit2_2 = win32gui.FindWindowEx(hwndkakao_edit1, hwndkakao_edit2_1, "EVA_Window", None)
    time.sleep(0.05)                                            # ㄴ시작핸들을 첫번째 자식 핸들(친구목록) 을 줌(hwndkakao_edit2_1)
    hwndkakao_edit3 = win32gui.FindWindowEx(hwndkakao_edit2_2, None, "Edit", None)
    time.sleep(0.05)
    # # Edit에 검색 _ 입력되어있는 텍스트가 있어도 덮어쓰기됨
    win32api.SendMessage(hwndkakao_edit3, win32con.WM_SETTEXT, 0, chatroom_name)
    time.sleep(0.5)  # 안정성 위해 필요
    SendReturn(hwndkakao_edit3)
    time.sleep(0.5)
# # 채팅내용 초기 저장 _ 마지막 채팅
def chat_last_save():
    open_chatroom(kakao_opentalk_name)  # 채팅방 열기
    ttext = copy_chatroom(kakao_opentalk_name)  # 채팅내용 가져오기
    # if ttext == 'copy_chatroom(kakao_opentalk_name)':
    #     while True:
    #         print("카카오톡 내용을 읽지 못하고 있습니다.")
    #         open_chatroom(kakao_opentalk_name)  # 채팅방 열기
    #         ttext = copy_chatroom(kakao_opentalk_name)
    #         if ttext != 'copy_chatroom(kakao_opentalk_name)':
    #             print("카카오톡 내용을 읽었습니다!")
    #             break

    a = ttext.split('\r\n')  # \r\n 으로 스플릿 __ 대화내용 인용의 경우 \r 때문에 해당안됨
    df = pd.DataFrame(a)  # DF 으로 바꾸기

    df[0] = df[0].str.replace('\[([\S\s]+)\] \[(오전|오후)([0-9:\s]+)\] ', '', regex=True)  # 정규식으로 채팅내용만 남기기
    return df.iloc[-2, 0]

chat_command = ['!저장', '!음식저장', '!기록', '!등록', '!음식등록', '!음식기록', '!추가', '!음식추가']
chat_command2 = ['!추천메뉴', '!음식추천', '!뭐먹지', '!추천해줘']
chat_command3 = ['!메뉴', "!메뉴보기", "!메뉴정보", "!메뉴보기"]
chat_command4 = '!삭제'
chat_command5 = ['!명령어', '!help']

# # 채팅방 커멘드 체크
def chat_chek_command(clst):
    open_chatroom(kakao_opentalk_name)
    ttext = copy_chatroom(kakao_opentalk_name)  # 채팅내용 가져오기
    # if ttext == 'copy_chatroom(kakao_opentalk_name)':
    #     while True:
    #         print("카카오톡 내용을 읽지 못하고 있습니다.")
    #         open_chatroom(kakao_opentalk_name)  # 채팅방 열기
    #         ttext = copy_chatroom(kakao_opentalk_name)
    #         if ttext != 'copy_chatroom(kakao_opentalk_name)':
    #             print("카카오톡 내용을 읽었습니다!")
    #             break

    a = ttext.split('\r\n')  # \r\n 으로 스플릿 __ 대화내용 인용의 경우 \r 때문에 해당안됨
    df = pd.DataFrame(a)  # DF 으로 바꾸기

    df[0] = df[0].str.replace('\[([\S\s]+)\] \[(오전|오후)([0-9:\s]+)\] ', '', regex=True)  # 정규식으로 채팅내용만 남기기
    last_value = df.iloc[-2, 0]
    if last_value == clst:
        print("채팅 없었음..")
        return last_value
    else:
        print("채팅 있었음")

        df1 = last_value  # 최근 채팅내용만 남김
        df2 = df1.split(' ')[0]
        save_data = df1.replace(df2, "")
        save_data = save_data.replace(" ", "")

        if df2 in chat_command: # 음식 메뉴 저장하는 커멘드 확인
            print("-------저장 커멘드 확인!")
            # 새로운 데이터 추가하기
            ref = db.reference('음식 메뉴') # receive from firebase_database
            firebase_db = ref.get()
            check = False
            for i in firebase_db:
                if i == save_data:
                    check = True
                    break
            if check:
                print("이미 등록된 음식!")
                message = "<AI>\n이미 등록된 음식입니다!"
                kakao_sendtext(kakao_opentalk_name, message)
            elif save_data:
                print("새로운 음식 등록!")
                ref.update({len(firebase_db) : save_data}) # Firebase의 DB에 update함
                message = "<AI>\n음식\'{0}\'이(가) DB에 성공적으로 추가되었습니다!".format(save_data)
                kakao_sendtext(kakao_opentalk_name, message)  # 메시지 전송
            else:
                print("음식명 누락!")
                message = "<AI>\n음식명이 누락되어 등록에 실패했어요.\nT^T".format(save_data)
                kakao_sendtext(kakao_opentalk_name, message)  # 메시지 전송

            return last_value
        elif df2 == chat_command4: # 메뉴 삭제 커멘드 확인
            print("-------음식_삭제 커멘드 확인!")
            ref = db.reference('음식 메뉴')  # receive from firebase_database
            firebase_db = ref.get()
            if save_data:
                for i in firebase_db:
                    if i == save_data:
                        print("음식 삭제 완료!")
                        ref.update({firebase_db.index(save_data):None}) # Firebase의 DB에 update함 None을 값으로 취하면 삭제
                        message = "<AI>\n음식\'{0}\'이(가) 성공적으로 삭제되었습니다!".format(save_data)
                        kakao_sendtext(kakao_opentalk_name, message)  # 메시지 전송
                        return last_value
                print("음식 삭제 실패!")
                message = "<AI>\n음식\'{0}\'이(가) 등록이 되어있지 않아 삭제에 실패했습니다.".format(save_data)
                kakao_sendtext(kakao_opentalk_name, message)  # 메시지 전송
                return last_value
            else:
                print("음식 삭제 실패!")
                message = "<AI>\n음식명이 누락되어 삭제에 실패했어요.\nT^T".format(save_data)
                kakao_sendtext(kakao_opentalk_name, message)  # 메시지 전송
                return last_value

        elif df1 in chat_command3: # 메뉴 정보 불러오는 커멘드 확인
            print("-------음식메뉴 커멘드 확인!")
            ref = db.reference('음식 메뉴') # receive from firebase_database
            firebase_db = ref.get()
            text = ""
            for i in range(len(firebase_db)):
                text += str(i)+": "+firebase_db[i]+"\n"
            message = "<AI>\n등록된 음식 메뉴 DB 정보입니다.\n" + text
            print(message)
            kakao_sendtext(kakao_opentalk_name, message)  # 메시지 전송
            return last_value

        elif df1 in chat_command5: # 명령어 정보 확인하는 커멘드 확인
            print("-------명령어_확인 커멘드 확인!")
            f = open("./command.txt", "r", encoding="utf-8")
            text = f.read()
            f.close()
            message = "<AI>\n저를 사용하기 위한 명령어입니다.\n" + text
            print(message)
            kakao_sendtext(kakao_opentalk_name, message)  # 메시지 전송
            return last_value

        elif df1 in chat_command2: # 음식 추천해주는 커멘드 확인
            print("-------음식_추천 커멘드 확인!")
            ref = db.reference('음식 메뉴') # receive from firebase_database
            message = "<AI>\n오늘은 이거닷!\n\'{0}\' 추천해드립니다!".format(random.sample(ref.get(), 1)[0])
            print(message)
            kakao_sendtext(kakao_opentalk_name, message)  # 메시지 전송
            return last_value

        else:
            print("커멘드 미확인")
            return last_value


def main():
    # sched = BackgroundScheduler()
    # sched.start()
    clst = chat_last_save()  # 초기설정 _ 마지막채팅 저장
    time.sleep(1)
    message = "<AI>\n'오늘은 뭐 먹지?' AI가 실행되었습니다.\n명령어를 확인하고 싶을 땐, '!명령어' 또는 '!help'를 입력해 주세요. ^^"
    kakao_sendtext(kakao_opentalk_name, message)  # 메시지 전송

    while True:
        print("실행중.....")
        clst = chat_chek_command(clst) # 커맨드 체크

if __name__ == '__main__':
    main()

import subprocess
import os
import threading
from tkinter import ttk

import win32com.client
import tkinter as tk

'''
点击运行 填写名称和url 
内容格式[hh：MM：ss@content|hh：MM：ss@content|...] 
'''

"""
def read_file():
    file_path = "D:\\onedrive\\桌面\\大三上活动\\2_家教\\数学\\one_url_command.txt"
    all_file_path = "D:\\onedrive\\桌面\\大三上活动\\2_家教\\数学\\all_url_command.txt"
    with open(file_path, "r", encoding="utf-8") as fp:
        lines = fp.readlines()
    # add in all .txt
    # with open(all_file_path, "a", encoding="utf-8") as ff:
    #     for line in lines:
    #         ff.write(line)
    #     ff.write("\n\n")
    '''
       url,name and content
       '''
    name = lines[0][:-1]
    url_path = lines[1][:-1]
    if url_path.__contains__("&"):
        url_path = url_path.replace("&", '"&"')
    content = lines[2]
    return name, url_path, content
"""

AND_JOIN = '&'
OR_JOIN = '|'
AI_JOIN = '@'
CHINESE_COLON_JOIN = '：'

HISTORY_FILE_PATH = './history.txt'

BASE_DIR_PATH = "D:/onedrive/桌面/word_bilibili/"

_WELCOME_MESSAGE = "欢迎来到SmartScreenshot"
_PROCESS_ERROR_MESSAGE = "提交出错，请重试！"


def read_content(content):
    '''
        split the content to two arr with &
    '''
    big_split = OR_JOIN
    small_split = AI_JOIN
    content = content.split(big_split)
    screenshot_times = []
    comment = []
    for c in content:
        c = c.split(small_split)
        screenshot_times.append(c[0])
        comment.append(c[1])
    # make times to screenshot
    times_in_second = []
    for t in screenshot_times:
        cnt_for_maohao = 0
        for i in t:
            if i == '：':
                cnt_for_maohao += 1
        if (cnt_for_maohao == 2):
            # time
            hour = t[:t.find(CHINESE_COLON_JOIN)]
            minute = t[t.find(CHINESE_COLON_JOIN) + 1:t.rfind(CHINESE_COLON_JOIN)]
            second = t[t.rfind(CHINESE_COLON_JOIN) + 1:]

            atime = int(hour) * 3600 + int(minute) * 60 + int(second)
            print(atime)
            times_in_second.append(atime)
        elif cnt_for_maohao == 1:
            minute = t[:t.rfind(CHINESE_COLON_JOIN)]
            second = t[t.rfind(CHINESE_COLON_JOIN) + 1:]

            atime = int(minute) * 60 + int(second)
            times_in_second.append(atime)
        else:
            # a wrong number to show we should not do a screenshot just write sentence
            times_in_second.append(-1)
    return comment, times_in_second


def init_docx(app, name):
    doc = app.Documents.Add()
    # head
    para = doc.Paragraphs.Add()
    para_range = para.Range
    para_range.Text = name
    doc.Content.InsertAfter("\n")
    para_range.Font.Size = 14
    para_range.Bold = 1
    para_range.ParagraphFormat.Alignment = 1
    return doc


def add_text_in_para(doc, curr, max, comment):
    para = doc.Paragraphs.Add()
    para_range = para.Range
    if curr < max - 1: para_range.Text = comment[curr + 1]
    para_range.Bold = 1
    para_range.Font.Size = 10
    doc.Content.InsertAfter("\n")
    return para_range


def exec_you_get(u, p):
    command = "you-get -O " + p + " --format=dash-flv480 " + u
    os.system(command)


def process_sss(video_name, url_path, content, progress_bar, _l_error_msg, base):
    progress = progress_bar["value"]

    # video_name, url_path, content = read_file()
    video_path = BASE_DIR_PATH + video_name
    try:
        # you-get
        exec_you_get(url_path, video_path)
        progress += 20
        progress_bar["value"] = progress
        progress_bar.update()
        video_path += ".mp4"
        comment, _times = read_content(content)
    except:
        progress_bar["value"] = 0
        progress_bar.update()
        _l_error_msg.grid(row=base + 4, column=2)
        return
    # app
    app = win32com.client.Dispatch("Word.Application")
    # app.visible=1
    doc = init_docx(app, video_name)

    progress += 10
    progress_bar["value"] = progress
    progress_bar.update()

    # add first comment
    add_text_in_para(doc, -1, len(comment), comment)
    # do screenshot
    for i in range(len(_times)):
        time_in_seconds = _times[i]
        # just text
        if time_in_seconds == -1:
            # 3）添加新的段落
            add_text_in_para(doc, i, len(comment), comment)
            continue
        # need screenshot
        img_file = video_name + "_" + str(i) + ".png"
        # subprocess.run(['ffmpeg', '-i', video_path, '-ss', str(time_in_seconds), '-vframes', '1', fn])
        os.system("ffmpeg -i " + video_path + " -qscale:v 2 -ss " + str(time_in_seconds) + " -vframes 1 " + img_file)
        # 3）添加新的段落
        para_range = add_text_in_para(doc, i, len(comment), comment)
        img_file = os.path.join(os.path.abspath(os.path.curdir), img_file)
        # 在当前的段落中插入图片
        para_range.InlineShapes.AddPicture(img_file)
        os.remove(img_file)
        if progress <= 89:
            progress += 10
            progress_bar["value"] = progress
            progress_bar.update()
    # 保存文档
    doc.SaveAs(BASE_DIR_PATH + video_name + ".docx")
    # 关闭Word应用程序
    app.Quit()
    # 删除视频
    os.remove(video_path)

    progress_bar["value"] = 100
    progress_bar.update()


def _clear_text(*args):
    for _ in args:
        _.delete(0, tk.END)


def _check_(_e_name, _e_url, _e_content) -> bool:
    if _e_name.get().__eq__("") & \
            _e_url.get().__eq__("") & \
            _e_content.get().__eq__(""):
        return False
    return True


def _save_text(_e_name, _e_url, _e_content):
    if not _check_(_e_name, _e_url, _e_content): return
    list = [_e_name.get(), _e_url.get(), _e_content.get().replace("\n", "")]
    line = AND_JOIN.join(list)
    with open(HISTORY_FILE_PATH, "a", encoding="utf-8") as fp:
        fp.write(line)
        fp.write('\n')


def count_file_lines(filename) -> int:
    cnt = 0
    with open(filename, "r", encoding="utf-8") as fp:
        for _ in fp:
            cnt += 1
    return cnt


def read_file_lines(filename) -> []:
    with open(filename, "r", encoding="utf-8") as fp:
        lines = fp.readlines()
    return lines


def _text_history(filename, _lb_show_history, base_row, offset):
    # show
    _lb_show_history.grid(row=base_row + offset, column=0)
    # compare the lines count
    fls = count_file_lines(filename)
    lbls = _lb_show_history.size()
    if fls <= lbls:
        return
    # lines = read_file_lines(filename)
    # for i in range(len(lines) - lbls):
    #     _lb_show_history.insert(0, lines[i + lbls])
    with open(filename,"r",encoding="utf-8")as fp:
        cnt = 0
        for line in fp:
            cnt += 1
            if cnt > lbls:
                _lb_show_history.insert(0, line)

def _hide_history(_lb_show_history):
    _lb_show_history.grid_remove()
    # _lb_show_history.delete(0, tk.END)


def _choose_text(_lb_show_history, _e_name, _e_url, _e_content):
    # no choose
    if len(_lb_show_history.curselection()) == 0: return

    lb_item = _lb_show_history.get(
        _lb_show_history.curselection()[0])
    lb_item_list = lb_item.split(AND_JOIN)
    if len(lb_item_list) < 3: return

    _e_name.delete(0, tk.END)
    _e_name.insert(0, lb_item_list[0])
    _e_url.delete(0, tk.END)
    _e_url.insert(0, lb_item_list[1])
    _e_content.delete(0, tk.END)
    _e_content.insert(0, lb_item_list[2])


def _init_gui():
    _window = tk.Tk()
    _window.wm_title("smart screenshot @copyright 2023.9-2024 by Demouo")
    base_row = 1
    _l_welcome = tk.Label(_window, text=_WELCOME_MESSAGE)
    _l_welcome.grid(row=base_row - 1, column=1)

    name = tk.StringVar()
    url = tk.StringVar()
    content = tk.StringVar()

    text = "请输入名称"
    _l_name = tk.Label(_window, text=text)
    _e_name = tk.Entry(_window, textvariable=name)
    _l_name.grid(row=base_row, column=0)
    _e_name.grid(row=base_row, column=1)

    text = "请输入网址"
    _l_url = tk.Label(_window, text=text)
    _e_url = tk.Entry(_window, textvariable=url)
    _l_url.grid(row=base_row + 1, column=0)
    _e_url.grid(row=base_row + 1, column=1)

    text = "请输入内容"
    _l_content = tk.Label(_window, text=text)
    _e_content = tk.Entry(_window, textvariable=content)
    _l_content.grid(row=base_row + 2, column=0)
    _e_content.grid(row=base_row + 2, column=1)

    # 创建一个进度条并放在底部
    _progress_bar = ttk.Progressbar(_window, orient="horizontal", length=150, mode="determinate")
    _progress_bar.grid(row=base_row + 3, column=1)

    def start_process():
        if not _check_(_e_name, _e_url, _e_content):
            _l_error_msg.grid(row=base_row + 4, column=2)
            return
        _progress_bar["value"] = 10
        _progress_bar.update()
        # 创建一个守护进程来执行任务，并传递进度条
        threading.Thread(target=process_sss,
                         args=(name.get(), url.get(), content.get(), _progress_bar, _l_error_msg, base_row),
                         daemon=True).start()

    # err msg
    _l_error_msg = tk.Label(_window, text=_PROCESS_ERROR_MESSAGE)
    _l_error_msg.grid_remove()

    _f_b_1 = tk.Frame(_window)
    _f_b_1.grid(row=base_row + 4, column=1)
    _b_commit = tk.Button(_f_b_1, text="提交", command=start_process)
    _b_commit.grid(row=base_row + 4, column=1)
    _b_clear_input = tk.Button(_f_b_1, text="清空", command=lambda: _clear_text(_e_name, _e_url, _e_content))
    _b_clear_input.grid(row=base_row + 4, column=0)

    _f_b_2 = tk.Frame(_window)
    _f_b_2.grid(row=base_row + 4, column=0)
    _b_clear = tk.Button(_f_b_2, text="隐藏", command=lambda: _hide_history(_lb_show_history))
    _b_clear.grid(row=base_row + 4, column=1)
    _b_save = tk.Button(_f_b_2, text="保存", command=lambda: _save_text(_e_name, _e_url, _e_content))
    _b_save.grid(row=base_row + 4, column=2)
    _b_history = tk.Button(_f_b_2, text="历史",
                           command=lambda: _text_history(HISTORY_FILE_PATH, _lb_show_history, base_row, 5))
    _b_history.grid(row=base_row + 4, column=0)

    _lb_show_history = tk.Listbox(_window)

    # 绑定Listbox的选择事件
    _lb_show_history.bind("<<ListboxSelect>>",
                          lambda event: _choose_text(_lb_show_history, _e_name, _e_url, _e_content))
    _window.mainloop()


if __name__ == '__main__':
    _init_gui()

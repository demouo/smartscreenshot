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


def read_content(content):
    '''
        split the content to two arr with &
    '''
    big_split = '|'
    small_split = '@'
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
            # hour
            hour = t[:t.find("：")]
            minute = t[t.find("：") + 1:t.rfind("：")]
            second = t[t.rfind("：") + 1:]

            atime = int(hour) * 3600 + int(minute) * 60 + int(second)
            print(atime)
            times_in_second.append(atime)
        elif cnt_for_maohao == 1:
            minute = t[:t.rfind("：")]
            second = t[t.rfind("：") + 1:]

            atime = int(minute) * 60 + int(second)
            times_in_second.append(atime)
        else:
            # a wrong number to show we should not screenshot
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
    base_dir_path = "D:/onedrive/桌面/word_bilibili/"
    # video_name, url_path, content = read_file()
    video_path = base_dir_path + video_name
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
    doc.SaveAs(base_dir_path + video_name + ".docx")
    # 关闭Word应用程序
    app.Quit()
    # 删除视频
    os.remove(video_path)

    progress_bar["value"] = 100
    progress_bar.update()


def _clear_text_(*args):
    for _ in args:
        _.delete(0,tk.END)


def _init_gui():
    _window = tk.Tk()
    base = 1
    _l_welcome = tk.Label(_window, text="欢迎来到SmartScreenshot")
    _l_welcome.grid(row=base - 1, column=0)

    name = tk.StringVar()
    url = tk.StringVar()
    content = tk.StringVar()

    text = "请输入名称"
    _l_name = tk.Label(_window, text=text)
    _e_name = tk.Entry(_window, textvariable=name)
    _l_name.grid(row=base,column=0)
    _e_name.grid(row=base,column=1)

    text = "请输入网址"
    _l_url = tk.Label(_window, text=text)
    _e_url = tk.Entry(_window, textvariable=url)
    _l_url.grid(row=base+1, column=0)
    _e_url.grid(row=base+1, column=1)

    text = "请输入内容"
    _l_content = tk.Label(_window, text=text)
    _e_content = tk.Entry(_window, textvariable=content)
    _l_content.grid(row=base+2, column=0)
    _e_content.grid(row=base+2, column=1)

    # 创建一个进度条并放在底部
    _progress_bar = ttk.Progressbar(_window, orient="horizontal", length=150, mode="determinate")
    _progress_bar.grid(row=base + 3, column=1)

    def start_process():
        _progress_bar["value"] = 10
        _progress_bar.update()
        # 创建一个守护进程来执行任务，并传递进度条
        threading.Thread(target=process_sss,
                         args=(name.get(), url.get(), content.get(), _progress_bar, _l_error_msg, base),
                         daemon=True).start()

    # err msg
    _l_error_msg = tk.Label(_window, text="提交出错，请重试！")
    _l_error_msg.grid_remove()

    _b_commit = tk.Button(_window, text="提交", command=start_process)
    _b_commit.grid(row=base + 4, column=1)

    _b_clear = tk.Button(_window, text="清空", command=lambda: _clear_text_(_e_name, _e_url, _e_content))
    _b_clear.grid(row=base + 4, column=0)

    _window.mainloop()

if __name__ == '__main__':
    _init_gui()

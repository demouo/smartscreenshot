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
DOLLAR_JOIN = '$'
OR_JOIN = '|'
AI_JOIN = '@'
CHINESE_COLON_JOIN = '：'
ENGLISH_COLON_JOIN = ':'

HISTORY_FILE_PATH = './history.txt'

BASE_DIR_PATH = "D:/onedrive/桌面/大三上活动/2_家教/数学/"

_WELCOME_MESSAGE = "欢迎来到SmartScreenshot"
_PROCESS_ERROR_MESSAGE = "提交出错，请重试！"

base_row = 0
base_column = 0


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
        c[0] = c[0].replace(CHINESE_COLON_JOIN, ENGLISH_COLON_JOIN)
        screenshot_times.append(c[0])
        comment.append(c[1])
    # make times to screenshot
    times_in_second = []
    for t in screenshot_times:
        cnt_colon = 0
        for i in t:
            if i == ENGLISH_COLON_JOIN:
                cnt_colon += 1
        if (cnt_colon == 2):
            # time
            hour = t[:t.find(ENGLISH_COLON_JOIN)]
            minute = t[t.find(ENGLISH_COLON_JOIN) + 1:t.rfind(ENGLISH_COLON_JOIN)]
            second = t[t.rfind(ENGLISH_COLON_JOIN) + 1:]

            atime = int(hour) * 3600 + int(minute) * 60 + int(second)
            print(atime)
            times_in_second.append(atime)
        elif cnt_colon == 1:
            minute = t[:t.rfind(ENGLISH_COLON_JOIN)]
            second = t[t.rfind(ENGLISH_COLON_JOIN) + 1:]

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
    # 进度条 10
    progress = progress_bar["value"]

    # video_name, url_path, content = read_file()
    video_path = BASE_DIR_PATH + video_name
    try:
        # you-get
        exec_you_get(url_path, video_path)
        # you-get 结束 进度条+20 共30
        progress += 20
        progress_bar["value"] = progress
        progress_bar.update()
        video_path += ".mp4"
        comment, _times = read_content(content)
    except:
        progress_bar["value"] = 0
        progress_bar.update()
        _l_error_msg.grid(row=base + 4, column=3)
        return
    # app
    app = win32com.client.Dispatch("Word.Application")
    doc = init_docx(app, video_name)
    # 初始化word结束 进度条+10 共40
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
        if not os.path.exists(img_file):
            print("文件不存在阿！！")
            progress_bar["value"] = 0
            progress_bar.update()
            _l_error_msg.grid(row=base + 4, column=3)
            continue
        # 在当前的段落中插入图片
        para_range.InlineShapes.AddPicture(img_file)
        os.remove(img_file)
        # 进度条更新 每成功一个+5 卡到最后的90就停止
        if progress <= 89:
            progress += 5
            progress_bar["value"] = progress
            progress_bar.update()
    # 保存文档
    output_file_path = BASE_DIR_PATH + video_name + ".docx"
    doc.SaveAs(output_file_path)
    # 关闭Word应用程序
    app.Quit()
    # 删除视频
    os.remove(video_path)
    # 删除you-get的附加弹幕文件
    for fn in os.listdir("./"):
        if fn.endswith(".cmt.xml"):
            os.remove("./" + fn)
    # 结束时 进度条就是100
    progress_bar["value"] = 100
    progress_bar.update()
    _l_error_msg.grid_remove()
    output_file_path = output_file_path.replace("/", "\\")
    subprocess.Popen(r'explorer /select,"{}"'.format(output_file_path))


def _clear_text(*args):
    for k in args:
        if isinstance(k, list):
            for i in k:
                i.delete(0, tk.END)
        else:
            k.delete(0, tk.END)


def _check_(_e_name, _e_url, _e_content) -> bool:
    if _e_name.get().__eq__("") & \
            _e_url.get().__eq__("") & \
            _e_content.get().__eq__(""):
        return False
    return True


def _save(name, url, content):
    line = name + DOLLAR_JOIN + url + DOLLAR_JOIN + content
    with open(HISTORY_FILE_PATH, "a", encoding="utf-8") as fp:
        fp.write(line)
        fp.write('\n')


def _show_history(filename, _lb_show_history, base_row, offset):
    # show
    _lb_show_history.grid(row=base_row + offset, column=0, rowspan=5)
    # compare the lines count
    fls = count_file_lines(filename)
    lbls = _lb_show_history.size()
    if fls <= lbls:
        return
    with open(filename, "r", encoding="utf-8") as fp:
        cnt = 0
        for line in fp:
            cnt += 1
            if cnt > lbls:
                _lb_show_history.insert(0, line)


def _hide_history(_lb_show_history):
    _lb_show_history.grid_remove()
    # _lb_show_history.delete(0, tk.END)


def _choose_listbox_item(_lb_show_history, _e_name, _e_url, _e_time_list,
                         _e_content_list, _window):
    # no choose
    if len(_lb_show_history.curselection()) == 0: return

    lb_item = _lb_show_history.get(
        _lb_show_history.curselection()[0])
    lb_item_list = lb_item.split(DOLLAR_JOIN)
    if len(lb_item_list) < 3: return

    _e_name.delete(0, tk.END)
    _e_name.insert(0, lb_item_list[0])
    _e_url.delete(0, tk.END)
    _e_url.insert(0, lb_item_list[1])
    # for content
    content = lb_item_list[2].split(OR_JOIN)
    c_size = len(content)
    _e_size = len(_e_time_list)
    if c_size > _e_size:
        #     add entry
        for i in range(c_size - _e_size+1):
            _e_time_list.append(tk.Entry(_window))
            _e_content_list.append(tk.Entry(_window))
            # show
            _e_time_list[i + _e_size].grid(row=base_row + _e_size + i + 3, column=base_column + 1)
            _e_content_list[i + _e_size].grid(row=base_row + _e_size + i + 3, column=base_column + 2)
    elif c_size < _e_size:
        # delete more list item
        for i in range(_e_size - c_size-1):
            _e_time_list[_e_size-1-i].destroy()
            _e_content_list[_e_size-1-i].destroy()
            _e_time_list.pop(_e_size-1-i)
            _e_content_list.pop(_e_size-1-i)

    # show
    for i in range(len(content)):
        t_c = content[i].split(AI_JOIN)
        t = t_c[0]
        c = t_c[1]
        _e_time_list[i].delete(0, tk.END)
        _e_time_list[i].insert(0, t)
        _e_content_list[i].delete(0, tk.END)
        _e_content_list[i].insert(0, c)
    # clear the last to input
    _e_time_list[len(_e_time_list)-1].delete(0,tk.END)
    _e_content_list[len(_e_content_list) - 1].delete(0, tk.END)
    # add the listener to the last content
    _e_content_list[len(_e_content_list)-1].bind("<KeyPress>", lambda e: add_one_entry(_window, _e_time_list, _e_content_list))


def concat_time_content(_e_time_list, e_content_list):
    line = ""
    for i in range(len(_e_time_list)):
        c = e_content_list[i].get()
        if c == "":
            continue
        t = _e_time_list[i].get()
        line += ("" if i == 0 else "|") + t + "@" + c
    return line


def _explore_output():
    subprocess.Popen(r'explorer /select,"{}"'.format("D:\onedrive\桌面\word_bilibili\猫猫.docx"))


def add_one_entry(_window, _e_time_list, _e_content_list):
    old_size = len(_e_time_list)
    _e_time_list.append(tk.Entry(_window))
    _e_content_list.append(tk.Entry(_window))
    # show
    _e_time_list[old_size].grid(row=base_row + old_size + 3, column=base_column + 1)
    _e_content_list[old_size].grid(row=base_row + old_size + 3, column=base_column + 2)

    _e_content_list[old_size-1].unbind("<KeyPress>")
    _e_content_list[old_size].bind("<KeyPress>", lambda e:add_one_entry(_window, _e_time_list, _e_content_list))


def _init_ui():
    _window = tk.Tk()
    _window.wm_title("smart screenshot @copyright 2023.9-2024 by Demouo")

    # 创建一个进度条
    _progress_bar = ttk.Progressbar(_window, orient="horizontal", length=150, mode="determinate")
    _progress_bar.grid(row=base_row + 1, column=base_column + 3)
    # 显示输出路径
    _l_save_path = tk.Label(_window, text="> " + BASE_DIR_PATH)
    _l_save_path.grid(row=base_row, column=base_column + 3)

    text = "请输入名称"
    _l_name = tk.Label(_window, text=text)
    _e_name = tk.Entry(_window)
    _l_name.grid(row=base_row, column=base_column)
    _e_name.grid(row=base_row, column=base_column + 1)

    text = "请输入网址"
    _l_url = tk.Label(_window, text=text)
    _e_url = tk.Entry(_window)
    _l_url.grid(row=base_row + 1, column=base_column)
    _e_url.grid(row=base_row + 1, column=base_column + 1)

    text = "请输入内容"
    _l_content = tk.Label(_window, text=text)
    _l_content.grid(row=base_row + 2, column=base_column)
    text = "时间"
    _l_time = tk.Label(_window, text=text)
    _l_time.grid(row=base_row + 2, column=base_column + 1)
    text = "备注"
    _l_time = tk.Label(_window, text=text)
    _l_time.grid(row=base_row + 2, column=base_column + 2)
    # 默认 五个 输入内容
    cnt_entry: int = 3
    # 时间entry列表
    _e_time_list = []
    # 内容entry列表
    _e_content_list = []
    # 生成并组织entry
    for i in range(cnt_entry - len(_e_time_list)):
        _e_time_list.append(tk.Entry(_window))
        _e_content_list.append(tk.Entry(_window))
        # show
        _e_time_list[i].grid(row=base_row + i + 3, column=base_column + 1)
        _e_content_list[i].grid(row=base_row + i + 3, column=base_column + 2)
    _e_content_list[len(_e_time_list)-1].bind("<KeyPress>", lambda e:add_one_entry(_window, _e_time_list, _e_content_list))

    # 提交激活
    def start_process():
        # if not _check_(_e_name, _e_url, _e_content):
        #     _l_error_msg.grid(row=base_row + 4, column=base_column + 3)
        #     return
        _progress_bar["value"] = 10
        _progress_bar.update()
        content = concat_time_content(_e_time_list, _e_content_list)
        # save the input
        _save(_e_name.get(), _e_url.get(), content)
        # 创建一个守护进程来执行任务，并传递进度条
        threading.Thread(target=process_sss,
                         args=(
                             _e_name.get(), _e_url.get(), content,
                             _progress_bar,
                             _l_error_msg, base_row),
                         daemon=True).start()

    # err msg
    _l_error_msg = tk.Label(_window, text=_PROCESS_ERROR_MESSAGE)
    _l_error_msg.grid_remove()
    # 按钮集合1
    _f_b_1 = tk.Frame(_window)
    _f_b_1.grid(row=base_row + 2, column=base_column + 3)
    _b_explore = tk.Button(_f_b_1, text="浏览", command=_explore_output)
    _b_explore.grid(row=base_row + 4, column=base_column + 1)
    _b_commit = tk.Button(_f_b_1, text="提交", command=start_process)
    _b_commit.grid(row=base_row + 4, column=base_column + 2)
    _b_clear_input = tk.Button(_f_b_1, text="清空",
                               command=lambda: _clear_text(_e_name, _e_url, _e_time_list, _e_content_list))
    _b_clear_input.grid(row=base_row + 4, column=base_column)
    # 按钮集合2
    _f_b_2 = tk.Frame(_window)
    _f_b_2.grid(row=base_row + 4, column=base_column)
    _b_clear = tk.Button(_f_b_2, text="隐藏", command=lambda: _hide_history(_lb_show_history))
    _b_clear.grid(row=base_row + 4, column=base_column + 1)
    _b_save = tk.Button(_f_b_2, text="保存",
                        command=lambda: _save(_e_name.get(), _e_url.get(),
                                              concat_time_content(_e_time_list, _e_content_list)))
    _b_save.grid(row=base_row + 4, column=base_column + 2)
    _b_history = tk.Button(_f_b_2, text="历史",
                           command=lambda: _show_history(HISTORY_FILE_PATH, _lb_show_history, base_row, 5))
    _b_history.grid(row=base_row + 4, column=base_column)

    _lb_show_history = tk.Listbox(_window, height=5)

    # 绑定Listbox的选择事件
    _lb_show_history.bind("<<ListboxSelect>>",
                          lambda event: _choose_listbox_item(_lb_show_history, _e_name, _e_url,
                                                             _e_time_list, _e_content_list, _window))
    _window.mainloop()


if __name__ == '__main__':
    _init_ui()

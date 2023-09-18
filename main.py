import subprocess
import os
import win32com.client

# 按间距中的绿色按钮以运行脚本。

'''
去家教文件夹下填写  one_url_command.txt 
默认输出在桌面
'''


def read_file():
    file_path = "D:\\onedrive\\桌面\\大三上活动\\2_家教\\数学\\one_url_command.txt"
    all_file_path = "D:\\onedrive\\桌面\\大三上活动\\2_家教\\数学\\all_url_command.txt"
    with open(file_path, "r", encoding="utf-8") as fp:
        lines = fp.readlines()
    with open(all_file_path, "a", encoding="utf-8") as ff:
        for line in lines:
            ff.write(line)
        ff.write("\n\n")
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
    big_split = '#'
    small_split = '&'
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


def add_text_in_para(doc, curr, max):
    para = doc.Paragraphs.Add()
    para_range = para.Range
    if curr < max - 1: para_range.Text = comment[curr + 1]
    para_range.Bold = 1
    para_range.Font.Size = 10
    doc.Content.InsertAfter("\n")
    return para_range


def exec_you_get(url, path):
    os.system("you-get " + url + " -O " + path)

def func():
    pass

if __name__ == '__main__':
    base_dir_path = "D:/onedrive/桌面/"
    video_name, url_path, content = read_file()
    '''
    you-get the video 
    '''
    video_path = base_dir_path + video_name
    # you-get
    exec_you_get(url_path, video_path)
    video_path += ".mp4"
    comment, _times = read_content(content)
    # app
    app = win32com.client.Dispatch("Word.Application")
    # app.visible=1
    doc = init_docx(app, video_name)
    # add first comment
    add_text_in_para(doc, -1, len(comment))
    # do screenshot
    for i in range(len(_times)):
        time_in_seconds = _times[i]
        # just text
        if time_in_seconds == -1:
            # 3）添加新的段落
            add_text_in_para(doc, i, len(comment))
            continue
        # need screenshot
        img_file = video_name + "_" + str(i) + ".png"
        # subprocess.run(['ffmpeg', '-i', video_path, '-ss', str(time_in_seconds), '-vframes', '1', fn])
        os.system("ffmpeg -i " + video_path + " -ss " + str(time_in_seconds) + " -vframes 1 " + img_file)
        # 3）添加新的段落
        para_range = add_text_in_para(doc, i, len(comment))
        img_file = os.path.join(os.path.abspath(os.path.curdir), img_file)
        # 在当前的段落中插入图片
        para_range.InlineShapes.AddPicture(img_file)
        os.remove(img_file)
    # 保存文档
    doc.SaveAs(base_dir_path + video_name + ".docx")
    # 关闭Word应用程序
    app.Quit()
    # 删除视频
    os.remove(video_path)

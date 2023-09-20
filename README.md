# smartscreenshot
> 平时经常需要找b站的视频进行学习，同时对视频截图做笔记到word文档中，因此萌生了一个自动化截图和写笔记的想法，所以诞生了本项目。
## UI概览
![image](https://github.com/demouo/smartscreenshot/assets/121025366/151a9131-9562-4f6c-9718-6b375afd2fad)

## 项目简介
> 填入文件名称和网址，再根据指定的格式[time@content|time@content|...],对应到视频的哪一秒钟进行截图及备注，生成word文档。

![image](https://github.com/demouo/smartscreenshot/assets/121025366/b50169ba-8c4e-42d7-a5c6-ae91a5a1f8c9)

> 结果展示(备注：原图没有裁剪，已脱敏)

 ![image](https://github.com/demouo/smartscreenshot/assets/121025366/656a2adc-20cb-44bb-b955-dfbcd7adf485)

## 其他功能介绍
> 历史记录选中

![image](https://github.com/demouo/smartscreenshot/assets/121025366/65b4ac8b-09fc-4da6-923d-58759f9b1ecc)

 

## 项目技术
> 主要用到了tkinter做界面，you-get爬取视频,ffmpeg对视频截图，win32com操作word文档

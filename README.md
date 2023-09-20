# smartscreenshot
> 平时经常需要找b站的视频进行学习，同时对视频截图做笔记到word文档中，因此萌生了一个自动化截图和写笔记的想法，所以诞生了本项目。
## UI概览
![image](https://github.com/demouo/smartscreenshot/assets/121025366/4234f7c9-5769-49c5-a6e6-9e19d7a4a1b4)
## 项目简介
> 填入文件名称和网址，再根据指定的格式[time@content|time@content|...],对应到视频的哪一秒钟进行截图及备注，生成word文档。
> ![image](https://github.com/demouo/smartscreenshot/assets/121025366/d5f5d7df-9077-45e6-950b-604f3b120366)
---
## 结果展示(备注：原图没有裁剪，已脱敏)
> ![image](https://github.com/demouo/smartscreenshot/assets/121025366/656a2adc-20cb-44bb-b955-dfbcd7adf485)

## 项目技术
> 主要用到了tkinter做界面，you-get爬取视频,ffmpeg对视频截图，win32com操作word文档

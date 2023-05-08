# PDF Splitter
## description
This is application to split PDF files by specified size or pages. This application uses pypdf, tkinter, tkinterdnd2, ttkbootstrap.<br/>
<p align="center">
<img style="width: 60%;" src="https://user-images.githubusercontent.com/44375258/236114167-cf698d3c-da55-4c52-b9c5-275d8d0f5085.png" />
</p>


## pyinstaller options
When you use pyinstaller to build exe file, you should set additional hook directory using hook-tkinterdnd2.py file included. Please refer to <https://github.com/pmgagne/tkinterdnd2>.
```
pyinstaller --onefile --windowed --add-data "C:/[your working path]/oss_notice.txt;." --additional-hooks-dir "."  "C:/[your working path]/pdfsplitter.py"
```

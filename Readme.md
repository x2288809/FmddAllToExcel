
# 简介
多种文件格式转化为福马多多项目任务库导入excel的表格


# 环境

安装依赖

```
pip install tkinterdnd2 openpyxl requests pillow
pip install tkinterdnd2==0.3.0
pip install openpyxl python-docx tkinterdnd2
```


# 打包

加上 --noconsole --onefile --uac-admin，并排除不必要模块，命令：

```
pyinstaller -F -w -n Mlumo小红书项目表格转换器 --noconsole --exclude-module=ctypes --exclude-module=asyncio imgurlToImgfile.py

pyinstaller -F -w -n Zip压缩包转福马多多素材库Excel导入表格 --noconsole --exclude-module=ctypes --exclude-module=asyncio zipToExcel.py

pyinstaller -F -w -n 视频文件名导入到Excel --noconsole --exclude-module=ctypes --exclude-module=asyncio videoToExcel.py

pyinstaller -F -w -n 202605全能转换Excel --noconsole --exclude-module=ctypes --exclude-module=asyncio allToExcel.py

pyinstaller -F -w -n 全能转换Excel --noconsole --icon=favicon.ico --exclude-module=ctypes --exclude-module=asyncio app.py --additional-hooks-dir=.
```


# -PDF-
一个可以根据EXCEL信息批量给PDF设定密码的简单Python程序，今天给同事写的。

该程序用到了一些库
执行以下命令下载，注意要64位python，不然下载不下来。
```
pip install pikepdf openpyxl
```

代码文件夹下执行以下命令运行,确保代码文件和待处理文件在同一文件夹下。
```
python passkey.py --excel password.xlsx
```

# nmapReport

来源：https://github.com/mrschyte/nmap-converter

使用 Nmap 的 Xml 报告生成 Excel 报告。

这里做了参数部分、输入输出方式、自动处理路径的修改，因在使用原脚本的时候经常要重新打开脚本熟悉其参数处理过程才能运行起来，这里只是对一些不是很友好的操作进行了优化。感谢原作者带来这么强大好用的 Excel 输出脚本。

# 快速开始

### 安装依赖

```bash 
sudo pip install python-libnmap
sudo pip install XlsxWriter
```

或

```bash 
sudo pip install -r requirements.txt
```

### 使用说明

**帮助内容**

```
usage: nmapReport.py [-h] -r REPORTS [REPORTS ...] [-o OUTPUT]

optional arguments:
  -h, --help            show this help message and exit
  -r REPORTS [REPORTS ...], --reports REPORTS [REPORTS ...]
                        xml 文件路径，可以是文件路径也可以是文件夹路径
  -o OUTPUT, --output OUTPUT
                        Path to xlsx output.
```

**使用流程**

1. 首先需要 Nmap 在扫描的时候使用 `-oX` 或 `-oA` 输出 Xml 格式报告，如：`nmap -A -oX scan_result 192.168.0.0/24`
2. 等待 Nmap 扫描完毕将会在执行扫描命令的目录下生成 *scan_result.xml* 文件
3. 使用本工具 `$ python3 nmapReport.py -r scan_result.xml`

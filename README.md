# windows 10 注册

TURING.dll单独复制到\Windows\SysWOW64目录中进行管理员身份注册。

# windows 7 注册

```
regsvr32 TURING.dll
```

把“TURING.dll”插件放在“c:\TURING\”目录中，再手动运行一下“双击注册.bat”文件来注册一下插件。

【注意】【注意】【注意】：Win10系统，请右键“管理员身份打开”运行“双击注册.bat”文件。

可以运行脚本输出版本号进行确认：

Set TURING = CreateObject("TURING.FISR")
TracePrint TURING.Version()
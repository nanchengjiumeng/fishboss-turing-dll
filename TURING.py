#================================================================
#         《我的眼睛 -- 图灵识别》v2.5.8.20190703
#----------------------------------------------------------------
#        【作者】：鱼头之王（FishBOSS）
#        【ＱＱ】：1204400520
#        【Ｑ群】：314452472/866837563
#        【创建】：2019.03.25
#        【更新】：2019.07.03
#        【注意】：插件需要手动注册一次(请使用32位python)
#================================================================
import win32com.client

#创建全局COM对象
global TURING
TURING = win32com.client.Dispatch('TURING.FISR')

#关于
def About():
    return TURING.About()

#----- 句柄 START -----

#句柄_关联窗口（目前仅支持系统的普通窗口"normal","gdi","aero"）
def Link(hwnd, mod = "normal"):
    b = TURING.Link(hwnd, mod)     #后台截图设置项
    return b[0]

#句柄_解除窗口关联
def UnLink():
    TURING.UnLink()
    return 

#句柄_查找窗口句柄
def Window_FindHwnd(iClass=0, iTitle=0):
    r = TURING.Window_FindHwnd(iClass, iTitle)
    return r

#句柄_得到窗口大小
def Window_GetSize(iHwnd):
    r = TURING.Window_GetSize(iHwnd)
    return r

#句柄_获取祖窗口句柄
def Window_GetAncestor(iHwnd):
    h = TURING.Window_GetAncestor(iHwnd)
    return h[0]

#句柄_获取指定坐标下句柄
def Window_GetPointHwnd(x, y):
    r = TURING.Window_GetPointHwnd(x, y)
    return r

#句柄_移动窗口位置
def Window_MoveTo(iHwnd=0, iLeft=0, iTop=0):
    TURING.Window_MoveTo(iHwnd, iLeft, iTop)
    return
    
#----- 句柄 END -----

#----- 来源 START -----

#来源_获取屏幕像素数据
def Pixel_FromScreen(iLeft, iTop, iRight, iBottom):
    TURING.Pixel_FromScreen(iLeft, iTop, iRight, iBottom)
    return

#来源_获取图片像素数据
def Pixel_FromPicture(ImagePath, Mode = 0):
    TURING.Pixel_FromPicture(ImagePath, Mode)
    return

#来源_获取获取鼠标图案数据
def Pixel_FromMouse(iWidth = 32, iHeight = 32):
    TURING.Pixel_FromMouse(iWidth, iHeight)
    return

#来源_获取剪切板图像数据
def Pixel_GetClpImageData():
    TURING.Pixel_GetClpImageData()
    return

#来源_设置剪切板图像数据
def Pixel_SetClpImageData():
    TURING.Pixel_SetClpImageData()
    return

#来源_设置剪切板屏幕像素
def Pixel_SetClpScreenData(left, top, right, bottom):
    TURING.Pixel_SetClpScreenData(left, top, right, bottom)
    return

#来源_设置剪切板图片像素
def Pixel_SetClpPictureData(ImagePath):
    TURING.Pixel_SetClpPictureData(ImagePath)
    return

#来源_剪切板图像数据保存为图片
def Pixel_SaveClpImageData(ImagePath):
    TURING.Pixel_SaveClpImageData(ImagePath)
    return

#来源_拆分图像数据,上限10个
def Pixel_CutImageData(iLeft, iTop, iRight, iBottom, Serial = 1):
    TURING.Pixel_CutImageData(iLeft, iTop, iRight, iBottom, Serial)
    return

#来源_配置图像数据 拆分,上限10个
def Pixel_SetImageData(Serial = 1):
    TURING.Pixel_SetImageData(Serial)
    return

#来源_配置图像数据 切割
def Pixel_SetImageDataCut(Serial = 0):
    TURING.Pixel_SetImageDataCut(Serial)
    return

#来源_设置 图像像素数据为图中图（0场景，1查找图）
def Pixel_SetSceneImageData(Mode = 0):
    TURING.Pixel_SetSceneImageData(Mode)
    return

#来源_配置图层图像数据，处理结果图像为黑白图（黑底白字）
def Pixel_SetLayerImageData(Num = 0):
    TURING.Pixel_SetLayerImageData(Num)
    return

#来源_获取颜色图层像素，处理结果图像为黑白图（黑底白字），返回图层数量
def Pixel_ColorImageData(Interval = 2, Num = "15"):
    n = TURING.Pixel_ColorImageData(Interval, Num)
    return n[0]

#来源_获取通道图像（仅支持RGB通道）
def Pixel_ChannelImageData():
    TURING.Pixel_ChannelImageData()
    return

#来源_图像的像素数据保存为图片
def SaveImageData(SaveImagePath):
    TURING.SaveImageData(SaveImagePath)
    return

#来源_获取图像的像素数据
def GetImageData():
    Pixels = TURING.GetImageData()
    return Pixels

#来源_载入图像的像素数据
def LoadImageData(ImageData):
    TURING.LoadImageData(ImageData)
    return

#来源_图片格式转换 jpg 默认 80% 压缩率
def ImageFormatConverter(ImagePath, ImageSavePath, value = 80):
    TURING.ImageFormatConverter(ImagePath, ImageSavePath, value)
    return

#来源_像素预览
def Pixel_Preview(Mode = 0):
    TURING.Pixel_Preview(Mode)
    return

#----- 来源 END -----

#----- 屏幕 START ------

#屏幕_打印图像
def Screen_PrintImage(x = 0, y = 0):
    TURING.Screen_PrintImage(x, y)
    return

#屏幕_打印文字
def Screen_PrintText(text, x = 0, y = 0, FBcolor = "0000FF|000000", FontNameSize = "宋体|9", Mode = 0):
    TURING.Screen_PrintText(text, x, y, FBcolor, FontNameSize, Mode)
    return

#屏幕_强制刷新
def Screen_Refresh():
    TURING.Screen_Refresh()
    return

#----- 屏幕 END -----

#----- 滤镜 START -----     

#滤镜_二值化 
#色阶阈值（范围：0-255）
#或者：指定颜色串BBGGRR-BDGDRD（"0000FF-000080|00FFFF"）(反色效果："@BBGGRR-DBDGDR")
#或者：通过最大类间方差法[Otsu]取得（"auto"）
def Filter_Binaryzation(value = "128"):
    TURING.Filter_Binaryzation(value)
    return

#滤镜_灰度  模式（默认0:标准，1:Photoshop算法）
def Filter_Gray(Mode = 0):
    TURING.Filter_Gray(Mode)
    return

#滤镜_色调分离  色阶阈值（范围：2~255）
def Filter_Posterization(Value = 3, Interval = 0):
    TURING.Filter_Posterization(Value, Interval)
    return

#滤镜_清除杂点
def Filter_Despeckle(Value = 6, Interval = 0):
    TURING.Filter_Despeckle(Value, Interval)
    return

#滤镜_去掉直线  点数百分比（范围：1~100）
def Filter_EraseLine(Value = 50):
    TURING.Filter_EraseLine(Value)
    return

#滤镜_获取轮廓
def Filter_Outline():
    TURING.Filter_Outline()
    return

#滤镜_提取色块
def Filter_ExtractBlock(iWidth = 3, iHeight = 3, num = 8):
    TURING.Filter_ExtractBlock(iWidth, iHeight, num)
    return

#滤镜_倾斜矫正
def Filter_SlantCorrect():
    TURING.Filter_SlantCorrect()
    return

#滤镜_旋转纠正
def Filter_RotateCorrect(Angle = 45, Value = 1):
    TURING.Filter_RotateCorrect(Angle, Value)
    return

#滤镜_颠倒颜色   效果：白多变黑
def Filter_InverseColor():
    TURING.Filter_InverseColor()
    return

#滤镜_膨胀腐蚀
def Filter_DilationErosion():
    TURING.Filter_DilationErosion()
    return

#滤镜_细化抽骨
def Filter_ThinBone():
    TURING.Filter_ThinBone()
    return

#滤镜_等比缩放
def Filter_Zoom(xTimes = 2, yTimes = 2):
    TURING.Filter_Zoom(xTimes, yTimes)
    return

#滤镜_缩放归一化
def Filter_ZoomOne(iScaleWidth, iScaleHeight):
    TURING.Filter_ZoomOne(iScaleWidth, iScaleHeight)
    return

#滤镜_有效图像
#@param Value:字符串型，可选，裁剪方式（默认空：裁剪黑边，"auto"：四角相同颜色裁剪）
def Filter_ValidCut(Value):
    TURING.Filter_ValidCut(Value)
    return

#滤镜_色选（@不保留选中颜色）
#指定颜色串BBGGRR-BDGDRD（"0000FF-000080|00FFFF"）(反选效果："@BBGGRR-DBDGDR")
def Filter_ColorChoose(value = "000000"):
    TURING.Filter_ColorChoose(value)
    return

#滤镜_固定旋转  旋转的正负度数值，正数顺时针（默认45，范围：正负0~360）
def Filter_Rotate(angle = 45):
    TURING.Filter_Rotate(angle)
    return

#滤镜_固定移位  像素移位特征串|开始行列数（移动数值：正数向左移动，负数向右移动）
def Filter_Shift(Value, Direction = 0):
    TURING.Filter_Shift(Value, Direction)
    return

#滤镜_祛除斑点
def Filter_DispelSpot(Sensitivity = 25, Num = 2):
    TURING.Filter_ColorChoose(Sensitivity, Num)
    return

#滤镜_查找互补色
def Filter_Complementary():
    colors = TURING.Filter_Complementary()
    return colors

#滤镜_差异提取
def Filter_DiffeExtract(ImageData1, ImageData2, Similarity = 1.0):
    n = TURING.Filter_DiffeExtract(ImageData1, ImageData2, Similarity)
    return n[0]

#----- 滤镜 END -----

#----- 切割 START -----

#切割_固定位置
def Incise_FixedLocation(qx, qy, iWidth, iHeight, Interval, num):
    n = TURING.Incise_FixedLocation(qx, qy, iWidth, iHeight, Interval, num)
    return n[0]

#切割_随机方位
def Incise_RandomOrientation(width = 0, height = 0):
    n = TURING.Incise_RandomOrientation()
    return n[0]

#切割_连通区域
def Incise_ConnectedArea(Through, width = 0, height = 0):
    n = TURING.Incise_ConnectedArea(Through)
    return n[0]

#切割_范围投影
def Incise_ScopeAisle(Row = 2, Column = 1, width = 0, height = 0):
    n = TURING.Incise_ScopeAisle(Row, Column)
    return n[0]

#切割_颜色分层
def Incise_ColorLayered(Interval, num, width = 0, height = 0):
    n = TURING.Incise_ColorLayered(Interval, num)
    return n[0]

#切割_自适应矩形（体验版）
def Incise_Adaptive(width = 0, height = 0):
    n = TURING.Incise_Adaptive()
    return n[0]

#切割_修改字符切割图像数据
def Incise_ModifyCharData(num, iLeft = 0, iTop = 0):
    TURING.Incise_ModifyCharData(num, iLeft, iTop)
    return

#切割_清除切割图像数据
def Incise_EraseData():
    TURING.Incise_EraseData()
    return

#切割_追加图像数据为切割数据（传入多个图像）
def Incise_AddCharData(iLeft = 0, iTop = 0):
    n = TURING.Incise_AddCharData(iLeft, iTop)
    return n[0]

#切割_获取切割字符的信息（字符串：左,上,宽,高,点阵|左,上,宽,高,点阵|……）
def Incise_GetCharData():
    s = TURING.Incise_GetCharData()
    return s

#切割_合并切割字符数据(后面参数的字符合并到前面字符，并删除后面的字符)
def Incise_JoinCharData(Num1, Num2):
    n = TURING.Incise_JoinCharData(Num1, Num2)
    return n[0]

#切割_切割字符大小归一化
def Incise_CharSizeOne(iWidth, iHeight):
    TURING.Incise_CharSizeOne(iWidth, iHeight)
    return

#切割_字符预览
def Incise_Preview(num):
    TURING.Incise_Preview(num)
    return

#----- 切割 END -----

#----- 字库 START -----

#字库_生成二进制字符串点阵
def Lib_Generate():
    s = TURING.Lib_Generate()
    return s

#字库_字符切割数据   切割后的每个字符图像  "5,8|01010101001" = 切割后的每个字符图像
def Lib_OneCharData(num):
    s = TURING.Lib_OneCharData(num)
    return s[0]

#字库_分析切割字符的字体与字号   Mode样式，默认"0|0"（格式：0正常,1粗体,2斜体,3粗斜体|0中文字体，1英文字体）
def Lib_AnalyzeFontSize(Text, Mode):
    s = TURING.Lib_AnalyzeFontSize(Text, Mode)
    return s[0]

#字库_存储识别字库
def Lib_Save(Text, Lattice, LibPath):
    TURING.Lib_Save(Text, Lattice, LibPath)
    return

#字库_加载识别字库 拓展
#@param LibStr：字符串型，识别字库的字符串内容（回车换行符用于分割数据）
def Lib_LoadEx(LibStr):
    n = TURING.Lib_LoadEx(LibStr)
    return n[0]

#字库_加载识别字库
def Lib_Load(LibPath):
    n = TURING.Lib_Load(LibPath)
    return n[0]

#字库_创建系统字体识别字库
def Lib_Create(iFont, iSize, iText = ""):
    n = TURING.Lib_Create(iFont, iSize, iText)
    return n[0]

#字库_添加新的识别字库,上限10个
def Lib_Add(Serial = 1):
    TURING.Lib_Add(Serial)
    return

#字库_设置使用哪个识别字库,上限10个
def Lib_Use(Serial = 1):
    TURING.Lib_Use(Serial)
    return

#字库_内部的图像数据追加为识别字库
def Lib_AddImageData(iText):
    TURING.Lib_AddImageData(iText)
    return

#字库_点阵预览
def Lib_Preview(num):
    TURING.Lib_Preview(num)
    return

#字库_获得字库数量
def Lib_UBound():
    num = TURING.Lib_UBound()
    return num

#----- 字库 END -----

#----- 识别 START -----

#识别_点阵比对
def OCR(Similar=0, Mode = 0):
    s = TURING.OCR(Similar, Mode)
    return s[0]

#识别_点阵比对_增强版 （参数1：百分比相似度）
def OCRex(Similar=0):
    s = TURING.OCRex(Similar)
    return s[0]

#识别_多边形识别（体验版）
def FindShape(Distance, Length):
    s = TURING.FindShape(Distance, Length)
    return s[0]

#识别_鼠标形状识别
def FindMouseShape(Mode = 0):
    d = TURING.FindMouseShape(Mode = 0)
    return d[0]

#---- 识别 END -----

#----- 算法（数学统计 and 加密揭秘）START -----

#算法_与众不同
def Different():
    s = TURING.Different()
    return s

#算法_统计差平方
def EvalVariance():
    d = TURING.EvalVariance()
    return d

#算法_取直线上所有坐标
def GetLineAllPos(x1, y1, x2, y2):
    s = TURING.GetLineAllPos(x1, y1, x2, y2)
    return s[0]

#算法_两线交叉坐标（体验版）
def TwoLinesCrossPos(sx1, sy1, sx2, sy2, ex1, ey1, ex2, ey2):
    s = TURING.TwoLinesCrossPos(sx1, sy1, sx2, sy2, ex1, ey1, ex2, ey2)
    return s[0]

#算法_获取所有端点坐标 (X,Y|X,Y|…)
def GetAllPoints(value):
    s = TURING.GetAllPoints(value)
    return s[0]

#算法_抽取端点之间的线段 (X,Y-X,Y-X,Y…)
def GetOneLine(x1, y1, x2, y2):
    s = TURING.GetOneLine(x1, y1, x2, y2)
    return s[0]

#算法_抽取所有端点之间的线段 (X,Y-X,Y-X,Y…|X,Y-X,Y-X,Y…|…)  0任意线，1直线
def GetAllLines(value = 0, Num = 0):
    s = TURING.GetAllLines(value, Num)
    return s[0]

#算法_统计颜色点数量
def CountColorNum(value):
    n = TURING.CountColorNum(value)
    return n[0]

#算法_字符串MD5加密   Code编码，默认3（格式：0:ANSI，1:ANSI-UTF8，2:GB2312，3:GB2312-UTF8）
def Pass_MD5String(Text, Code):
    s = TURING.Pass_MD5String(Text, Code)
    return s[0]

#算法_简单加密（10位秘钥）    明文内容 ,"123,4,56,78,90,1,234,56,78,90" 不超过255
def Pass_Encode(TextString, Password):
    s = TURING.Pass_Encode(TextString, Password)
    return s[0]

#算法_简单解密（10位秘钥）    密文内容 ,"123,4,56,78,90,1,234,56,78,90" 不超过255
def Pass_Uncode(TextString, Password):
    s = TURING.Pass_Uncode(TextString, Password)
    return s[0]

#算法_图片Base64编码  对图片进行Base64编码（支持bmp/png/jpg/gif等），IsHead默认False，True为含包头（格式：“data:image/<后缀名>;base64,”）
def Image_Base64Encode(FilePath, IsHead):
    s = TURING.Image_Base64Encode(FilePath, IsHead)
    return s[0]

#算法_二进制转十六进制字符串
def BITtoHEX(BITString):
    s = TURING.BITtoHEX(BITString)
    return s[0]

#算法_十六进制转二进制字符串
def HEXtoBIT(HEXString):
    s = TURING.HEXtoBIT(HEXString)
    return s[0]

#----- 算法（数学统计 and 加密揭秘） END -----

#----- 图色 START -----

#图色_获取指定位置颜色
def GetPixelColor(x, y, Mode = 0):
    s = TURING.GetPixelColor(x, y, Mode)
    return s[0]

#图色_屏幕区域找色
def FindColor(iLeft, iTop, iRight, iBottom, iColor, Direction, Similarity):
    z = TURING.FindColor(iLeft, iTop, iRight, iBottom, iColor, Direction, Similarity)
    return z[0].split(",")

#图色_屏幕区域找多色多坐标
def FindColorExS(iLeft, iTop, iRight, iBottom, iColorS, Direction, Similarity):
    zs = TURING.FindColorExS(iLeft, iTop, iRight, iBottom, iColorS, Direction, Similarity)
    return zs[0].split(",")

#图色_屏幕区域找图
def FindImage(iLeft, iTop, iRight, iBottom, ImagePath, Similarity):
    z = TURING.FindImage(iLeft, iTop, iRight, iBottom, ImagePath, Similarity)
    return z[0].split(",")

#图色_屏幕区域找所有图
def FindImageS(iLeft, iTop, iRight, iBottom, ImagePath, Similarity):
    zs = TURING.FindImageS(iLeft, iTop, iRight, iBottom, ImagePath, Similarity)
    return zs[0].split(",")

#图色_屏幕区域找其中图
def FindImageEx(iLeft, iTop, iRight, iBottom, ImagePathS, Similarity):
    z = TURING.FindImageEx(iLeft, iTop, iRight, iBottom, ImagePathS, Similarity)
    return z[0].split(",")

#图色_屏幕区域找所有图所有坐标
def FindImageExS(iLeft, iTop, iRight, iBottom, ImagePathS, Similarity):
    z = TURING.FindImageExS(iLeft, iTop, iRight, iBottom, ImagePathS, Similarity)
    return z[0].split(",")

#图色_RGB转HSV
def RGBtoHSV(iColor):
    c = TURING.RGBtoHSV(iColor)
    return c[0].split(",")

#图色_HSV转RGB
def HSVtoRGB(Hue, Saturation, Value):
    c = TURING.HSVtoRGB(Hue, Saturation, Value)
    return c[0]

#----- 图色 END -----
    
#----- 绘图 START

#绘图_创建画布
def Draw_CreateCanvas(iWidth = 256, iHeight = 256):
    TURING.Draw_CreateCanvas(iWidth, iHeight)
    return

#绘图_画点
def Draw_Point(x, y, cR = 255, cG = 0, cB = 0):
    TURING.Draw_Point(x, y, cR, cG, cB)
    return

#绘图_画线
def Draw_Line(x1, y1, x2, y2, cR = 255, cG = 0, cB = 0):
    TURING.Draw_Line(x1, y1, x2, y2, cR, cG, cB)
    return

#绘图_矩形
def Draw_Rectangle(iLeft, iTop, iRight, iBottom, cR = 255, cG = 0, cB = 0):
    TURING.Draw_Rectangle(iLeft, iTop, iRight, iBottom, cR, cG, cB)
    return

#绘图_圆形
def Draw_Circle(x, y, Radius, R = 255, G = 0, B = 0):
    TURING.Draw_Circle(x, y, Radius, R, G, B)
    return

#绘图_文字
def Draw_Text(x, y, text, FontSizeMode = "宋体|9|0", cR = 255, cG = 0, cB = 0):
    TURING.Draw_Text(x, y, text, FontSizeMode, cR, cG, cB)
    return

#绘图_填充[左上右下]
def Draw_Fill(x, y, Through = False, cR = 255, cG = 0, cB = 0):
    s = TURING.Draw_Fill(x, y, Through, cR, cG, cB)
    return s[0]

#绘图_生成验证码   返回验证码中计算的结果
def Draw_CAPTCHA():
    n = TURING.Draw_CAPTCHA()
    return n

#绘图_图像数据备份，上限64个
def Draw_Backups(Serial = 1):
    TURING.Draw_Backups(Serial)
    return

#绘图_图像数据还原，上限64个
def Draw_Recover(Serial = 1):
    TURING.Draw_Recover(Serial)
    return

#绘图_生成gif文件
def FileSaveGIF(LoadName, SaveName, Delay = 100):
    TURING.FileSaveGIF(LoadName, SaveName, Delay)
    return

#----- 绘图 END -----

#----- 鼠标模拟 START -----

#键鼠_键盘按键
def KM_KeyPress(Asck):
    TURING.KM_KeyPress(Asck)
    return

#键鼠_键盘按下
def KM_KeyDown(Asck):
    TURING.KM_KeyDown(Asck)
    return
    
#键鼠_键盘弹起
def KM_KeyUp(Asck):
    TURING.KM_KeyUp(Asck)
    return

#键鼠_键盘输入文字
def KM_SendString(x=0, y=0):
    TURING.KM_SendString(x, y)
    return

#键鼠_鼠标左键单击
def KM_LeftClick(x=0, y=0):
    TURING.KM_LeftClick(x, y)
    return
#键鼠_鼠标左键双击
def KM_LeftDbClick(x=0, y=0):
    TURING.KM_LeftDbClick(x, y)
    return

#键鼠_鼠标左键按下
def KM_LeftDown(x=0, y=0):
    TURING.KM_KeyPress(x, y)
    return
    
#键鼠_鼠标左键弹起
def KM_LeftUp(x=0, y=0):
    TURING.KM_LeftUp(x, y)
    return
    
#键鼠_鼠标中键单击
def KM_MiddleClick(x=0, y=0):
    TURING.KM_MiddleClick(x, y)
    return

#键鼠_鼠标中键按下
def KM_MiddleDown(x=0, y=0):
    TURING.KM_MiddleDown(x, y)
    return

#键鼠_鼠标中键弹起
def KM_MiddleUp(x=0, y=0):
    TURING.KM_MiddleUp(x, y)
    return

#键鼠_鼠标右键单击
def KM_RightClick(x=0, y=0):
    TURING.KM_RightClick(x, y)
    return
    
#键鼠_鼠标右键按下
def KM_RightDown(x=0, y=0):
    TURING.KM_RightDown(x, y)
    return

#键鼠_鼠标右键弹起
def KM_RightUp(x=0, y=0):
    TURING.KM_RightUp(x, y)
    return

#键鼠_鼠标移动到
def KM_MoveTo(x, y):
    TURING.KM_MoveTo(x, y)
    return

#键鼠_得到鼠标当前位置
def KM_GetCursorPos(Asck):
    position = TURING.KM_GetCursorPos(Asck)
    return position

#键鼠_延时
def KM_Delay(ms):
    TURING.KM_Delay(ms)
    return


#----- 鼠标模拟 END -----

#----- 其它 START -----

#其他_清理内存
def Memory_Clear():
    TURING.Memory_Clear()
    return

#其他_总共的物理内存|可用的物理内存|已用的内存比率
def Memory_GetInfo():
    s = TURING.Memory_GetInfo()
    return s

#其他_命令提示符(运行命令行)
def Run(Command):
    s = TURING.Run(Command)
    return s[0]

#版本号
def Version():
    v = TURING.Version()
    return v

#----- 其它 END -----

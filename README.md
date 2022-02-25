# HttpCanarySaverToExcel
## 可将HttpCanary（黄鸟）存档文件转为Excel文件，从而很方便地在Excel中进行查看、查找、编辑等

软件基于.NET Framework v4.7.2 开发，在已有该框架运行库的Windows系统电脑中可运行

## 反馈交流
QQ群：132108644<br/>
由于作者目前能遇到的content-type有限，目前编写了最常用的application/json、text/html、application/x-www-form-urlencoded的解析适配器，如果您有其他的当前未适配的content-type，请联系我并附上抓包存档，我将为此进行适配

## 使用说明
### 存档转为Excel功能
首先，在手机中的HttpCanary，将本次抓包的数据进行保存，然后复制到电脑中<br/>
然后打开本软件的主程序HttpCanarySaverToExcel.exe<br/>
在下图中，先选择HttpCanary所在文件夹，再选择要保存为的Excel路径。下面的其他选项可以根据提示自行设置，最后点击“导出”按钮即可。<br/>
注意：HttpCanary界面中标出的序号，与它导出存档的序号正好相反，所以本转换工具提供了选项，可以选择进行的编号与存档序号相同还是与HttpCanary界面显示相同<br/><br/>
![](https://github.com/zhangqi-ulua/HttpCanarySaverToExcel/blob/main/%E4%BD%BF%E7%94%A8%E8%AF%B4%E6%98%8E/1.png)<br/>
生成的Excel样例如下图所示：<br/>
总览Sheet表列出所有抓包的基础信息以及跳转到每一条抓包详情Sheet表的超链接，如果在HttpCanary存档文件夹中，为某些抓包的子文件夹设置了备注（只需要在子文件夹默认以序号命名的文件名后面追加用中文括号，在中文括号内写的内容即会本软件识别为备注），也会进行显示<br/>
![](https://github.com/zhangqi-ulua/HttpCanarySaverToExcel/blob/main/%E4%BD%BF%E7%94%A8%E8%AF%B4%E6%98%8E/2.png)<br/>
各条抓包的详情页面中，将列举请求、回应中涉及的详细信息<br/>
![](https://github.com/zhangqi-ulua/HttpCanarySaverToExcel/blob/main/%E4%BD%BF%E7%94%A8%E8%AF%B4%E6%98%8E/3.png)<br/>

### 可在Excel的总览页面为各个序号的抓包填写备注信息，然后通过本软件自动为存档子文件夹重命名
![](https://github.com/zhangqi-ulua/HttpCanarySaverToExcel/blob/main/%E4%BD%BF%E7%94%A8%E8%AF%B4%E6%98%8E/4.png)<br/>

## 赞助
如果您觉得软件还不错，并且愿意请作者喝杯咖啡的话，欢迎打赏<br/>
<img src="https://github.com/zhangqi-ulua/FiddlerHeadConvertor/blob/main/%E4%BD%BF%E7%94%A8%E8%AF%B4%E6%98%8E/wechat.jpg" width="40%">
<img src="https://github.com/zhangqi-ulua/FiddlerHeadConvertor/blob/main/%E4%BD%BF%E7%94%A8%E8%AF%B4%E6%98%8E/alipay.jpg" width="40%">

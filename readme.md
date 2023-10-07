Microsoft 365 下载地址：

https://officecdn.microsoft.com/pr/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/zh-CN/O365ProPlusRetail.img

感谢 [项目作者](https://massgrave.dev/) 辛勤付出，感谢 [知彼而知己](https://mp.weixin.qq.com/s/CGzS1KrZd8rXRDpfGxFKfw) 原创。

Office分为零售版和批量版，此前如果想要永久激活，需要Retail密钥或MAK密钥，既可以在线一键激活，也可以电话离线激活。

Microsoft 365已有的激活方法，最正规的是订阅激活。比如通过购买各种订阅计划，或者免费白嫖E5/E3订阅计划。退而求其次，还有安装Mondo 2016证书，将其转化成Mondo 2016后，采用KMS离线激活，虽然是2016的证书，但是基本能解锁Microsoft 365绝大部分的功能（不包括OneDrive等）。

在浏览[全球最大的同性交友网站](https://github.com)时，发现了一个新项目[【ohook】](https://github.com/asdcorp/ohook)，它目前可以在完全离线的情况下，永久激活以下版本：

除了Office 2010以及UWP版本Office以外，全部支持（既包括零售版也包括批量版）。不过，不支持Windows7及以下系统。也就是说Office 2013~2021以及Microsoft 365都被破解了。

下面先介绍手动使用的方法，再简要说一下其中的原理。

## 1、【手动激活方法】

以激活Microsoft 365为例：

步骤1、下载dll文件（注意区分32位和64位，打开Word，文件→账户→关于Word，可看到Office是32位还是64位），名称改为sppc.dll，然后将其复制到目录：

```shell
%ProgramFiles%\Microsoft Office\root\vfs\System
```

![](https://github.com/Vincenzo155/ohook/blob/principalis/Example/Images/01.png)

步骤2、以管理员身份运行下面两条命令：

如果是64位Office：

```shell
mklink "%ProgramFiles%\Microsoft Office\root\vfs\System\sppcs.dll" "%windir%\System32\sppc.dll"
```

如果是32位Office：

```shell
mklink "%ProgramFiles%\Microsoft Office\root\vfs\System\sppcs.dll" "%windir%\SysWOW64\sppc.dll"
```

安装Microsoft 365（O365ProPlusRetail）密钥：

```shell
slmgr -ipk 2N382-D6PKK-QTX4D-2JJYK-M96P2
```

![](https://github.com/Vincenzo155/ohook/blob/principalis/Example/Images/02.png)

上述步骤，所有版本的Office均适用，只是不同版本的密钥不同、以及Office安装路径不同，替换一下即可。

一般情况下，至此大功告成！

如果你激活的是Microsoft 365，最好再执行一步：

对于Microsoft 365有一定概率会向微软服务器发出请求，询问订阅是否到期，一旦检测到则激活失效，所以在hosts中屏蔽一下服务器的检测，或者修改注册表均可。

（可选）步骤3、在C:\Windows\System32\drivers\etc\hosts末尾添加下述内容：

```shell
127.0.0.1 ols.officeapps.live.com
```

或者运行下述命令：

```shell
reg add HKCU\Software\Microsoft\Office\16.0\Common\Licensing\Resiliency /v "TimeOfLastHeartbeatFailure" /t REG_SZ /d "2033-08-18T22:18:45Z" /f
```

打开Office后，可以看到已显示订阅激活，此时并未登录账户。

![](https://github.com/Vincenzo155/ohook/blob/principalis/Example/Images/03.png)

## 2、【激活原理】

我画了两张图帮助大家理解原理。👻

在Office启动的过程中，调用C:\Windows\System32\sppc.dll（系统文件）中的函数SLGetLicensingStatusInformation，用于检查Office的许可证状态，然后Office得到答案：已激活/未激活。

![](https://github.com/Vincenzo155/ohook/blob/principalis/Example/Images/04.png)

通过sppc.dll（破解文件）劫持该函数，让它给Office发出一个假信号，谎报激活状态。所以启动Office后，得到的答案是：已被激活。

![](https://github.com/Vincenzo155/ohook/blob/principalis/Example/Images/05.png)

具体到操作流程上，首先通过mklink命令，将系统文件sppc.dll链接到Office目录下，改名为sppcs.dll，用破解的sppc.dll劫持sppcs.dll，在询问激活状态时，谎报结果即可骗过Office。Office真的以为自己被激活了，但是Windows并没有认为它已经被激活。但这已经不重要了，Office自认为被激活就可以了，所有功能都拿去用吧。在甜言蜜语的糖衣炮弹中，Office迷失了自我👻

![](https://github.com/Vincenzo155/ohook/blob/principalis/Example/Images/06.png)

## 3、【和正版激活有什么区别？】

①是【伪激活】还是【真激活】？

使用本文方法激活后，如果使用命令查询激活状态，依旧显示未激活。如果上面两张图看明白了，就知道原因了——

【真激活】通常指的是采用微软钦定的正宗激活方式，比如：KMS、Retail/MAK密钥。

【伪激活】只是文字上显示激活，实际上未激活，功能该不能用还是不能用。本文的方法，使用命令查询也是未激活状态，但从它的原理可知，是在Office启动过程中欺骗Office，因此它和真激活无区别。一言以蔽之，只要一打开Office，它就是激活的，只欺骗了Office，没有欺骗Windows，

严格来说，本文的方法属于【破解】的范畴，它是真的解锁了Office的功能，并不是【伪激活】。如果微软不采取任何措施的话，此方式就是永久有效的。所以既不是【真激活】，也不是【伪激活】，是【真破解】😻

②功能有无区别？

使用本文方法激活Office后为永久激活，年度发布的版本和使用密钥永久激活的，在功能上没有区别。

大家比较关心的Microsoft 365，和正版订阅激活的区别还是有的，毕竟是离线验证，Microsoft 365的一些云功能，比如OneDrive的1TB空间依旧不能使用，因为OneDrive要向微软服务器请求。

③和安装Mondo 2016证书激活Microsoft 365有么有区别？

Mondo证书虽说能解锁Microsoft 365的绝大部分功能，但是针对较新版本，Mondo证书和365的证书，策略是有所区别的，也就是说Mondo证书除了不能解锁云功能（OneDrive）外，其他某些功能上也有所区别，具体是哪些功能，很难梳理，总之是不完全相同的。

但是使用本文的方法激活后，除了云功能不能用外（OneDrive等），其他功能与订阅激活的没有任何区别，因为许可证没变。

④是否安全？

从前面介绍的原理可以看到，并未对系统文件或组件进行修改和破坏，安全性、稳定性不用担心。

⑤Office是否可以更新？

可以。（如果以后微软封杀此方式，则不要更新）

## 4、【遥遥领先】

大家如果对其中的具体原理感兴趣，可以参见该项目源代码：

https://github.com/asdcorp/ohook

该项目作者@asdcorp将此项目命名为【ohook】，那么以后Office离线永久激活就叫做ohook。

近十多年，Office的永久激活一直都是通过密钥激活。自Office 2010开始，没有其他完美的离线激活方法，被微软卡脖子了，此项目没有破坏系统文件，突破关键技术，遥遥领先

严格来说，ohook已经不是【激活】，只能算是【破解】，微软只需一个系统更新(或Office更新)就能干掉它，大家觉得微软以后会下手吗？

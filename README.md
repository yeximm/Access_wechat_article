## 微信文章获取（Access_wechat_article）

本项目是基于Python语言的爬虫程序，支持对微信公众号文章内容获取

目前支持 Windows / Linux 开箱即用，做的比较粗糙，望见谅！

## 更新内容

1. **2024.9.12**更新
   - 优化重要参数的获取方式
   - 更新具体功能展示效果图
2. **2024.8.29**更新
   - 绕过微信公众号文章用代码访问时产生的验证提示（反爬虫机制）
   - 优化文章列表与内容获取逻辑

## 主要功能介绍

1. 获取**微信公众号文章**的网页文本数据
2. 获取**微信公众号**下所有历史文章，以**excel文件**形式保存
3. 获取微信公众号文章的**所有信息**，如浏览量、点赞数、评论等信息。

## 下载 / Download

- [Github / Download](https://github.com/yeximm/Access_wechat_article/archive/refs/heads/master.zip)

👆👆👆以上为本项目文件，直接clone该项目，或下载此链接均可。

建议使用虚拟环境运行项目

[requirements.txt](https://github.com/yeximm/Access_wechat_article/blob/master/requirements.txt)中包含所需python包文件名称

使用`pip install -r requirements.txt`批量安装python包文件

## 项目所需环境及工具

1. 系统环境：Windows 10 ×64
2. 程序运行环境：python 3.12
3. 涉及应用：微信**PC版**，当前项目适配的微信版本为3.9.11.25
4. 使用工具：fiddler

## 运行参数 Windows/Linux

1. 项目主文件为：`main.py`，另外几个文件为功能文件，为主文件服务
   项目存储路径为：`./data/`（程序会自动创建）
2. 运行命令：

​		进入项目目录后运行：`python main.py`

## 功能详情

**save_content.py**

1.获取文章文本内容 SaveContent

- 完成网页验证
- 获取单个文章的网页文本数据
- 保存单个文章的网页为pdf格式(**待实现**)

2.获取文章列表 GetList

- 获取公众号下所有历史文章
- 获取公众号下最新的N页历史文章(一页15篇)
- 保存列表到文件
- 保存文章内容到文件

**get_detail.py**

- 获取文章全部内容 SaveAllDetail
- 获取单个文章的网页文本数据
- 获取该文章的 浏览量，点赞数，评论等信息

**实现代理(待实现)**

- 使用Python代理电脑，监听微信获取关键字值
- 通过截取到的信息对目标文章进行下载

## 功能截图

**功能1：**

![1724053373108.png](https://m.dyeddie.top/?explorer/share/file&hash=a32dol-HaPvvzDtpRYS8iVZn6YKc9Zx9YUQ8BAXBR9XsBKg2gK9dnNKXNSpk3vwTKA)

![](https://m.dyeddie.top/?explorer/share/file&hash=91393jkW0LAJbt5gk0az_OGvTz8gkUw7PWgCSNIwSTIDzayZ64aR_dm7wdMKCVRJfw)

**功能2：**

![](https://m.dyeddie.top/?explorer/share/file&hash=122fBMYK3xEUgjJ5L2QrFI8z_cgKEda4hy6DBeHfJkhhGAwlZvIDJf3PDgmuQWjwNw)

![](https://m.dyeddie.top/?explorer/share/file&hash=d8d7GmcrMLiIEDP8yomoH6AqNScZ-CZyaFDDSgEj736PyIcA8M6zCQXB4Ji93VgChw)

**功能3：**

![](https://m.dyeddie.top/?explorer/share/file&hash=2bb0u1ImC8e4xoYnGiUhG5Ix_kMybAwAUCDJkM7Fyh-ZOauR5grYPACOqJXHQtAD-g)

![](https://m.dyeddie.top/?explorer/share/file&hash=34d5YvkarXpFJxLX7IlIJ35NYvno19I_CAgkUi6ODtUOwvYoH4Q3_j9_6HORFKA6hw)

**功能4：**

![](https://m.dyeddie.top/?explorer/share/file&hash=c1a2GMLCwqr1xUpc4O0NQasQ2OOIGpbEkmF0mXqk7UlBp197Z4ZHEiVCgE8ZnyNGGA)

## 免责声明

**本项目仅供技术研究，请勿用于任何商业用途，请勿用于非法用途！如有任何人凭此做何非法事情，均于作者无关，特此声明。**

**对于使用本项目产生的额外问题，如账户封禁被盗等，作者不对此负责，请谨慎使用。**

**如有不当之处，请联系本人，联系方式：**

<p align = "center">    
<img  src="https://m.dyeddie.top/?explorer/share/file&hash=92ddotJ8TUT7AviXIknm8ey8EjCCxzxsZoIb-Ohk_rej6n7RRpVEtrRpykqiaU2emg" width="200" />
</p>


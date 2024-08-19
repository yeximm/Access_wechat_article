# Access_wechat_article

# 微信文章获取

基于Python语言的爬虫程序，支持对微信公众号文章内容获取

目前支持 Windows / Linux 开箱即用

> 注：本项目仅供学习，请不要用于商业、违法途径，本人不对此源码造成的违法负责！

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

![1724053373108.png](https://img.picui.cn/free/2024/08/19/66c2f77d6ee93.png)

## 免责声明

**本项目仅供技术研究，请勿用于任何商业用途，请勿用于非法用途！如有任何人凭此做何非法事情，均于作者无关，特此声明。**

**对于使用本项目产生的额外问题，如账户封禁被盗等，作者不对此负责，请谨慎使用。**

**如有不当之处，请联系本人，联系方式：**

<p align = "center">    
<img  src="https://imgos.cn/2024/08/14/66bc5aa192094.jpg" width="200" />
</p>
# PythonCreatePPT
## <自动化办公> 两秒完成250页豆瓣电影PPT
> PPT并不好用, 但还是得用它, 这里借用豆瓣Top250的电影信息, 利用`python-pptx (0.6.7)`自动生成250张PPT, 希望通过实例, 给常年整理PPT报表的上班族, 一个解放生产力的新思路
> ![Python 自动生成 豆瓣电影PPT](https://upload-images.jianshu.io/upload_images/3203841-ca4e6ae9b0fc56e3.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

# 最终效果展示
> ![PPT动图](https://upload-images.jianshu.io/upload_images/3203841-840c476f81264cc2.gif?imageMogr2/auto-orient/strip)

## 数据哪里来的?
> 爬虫抓的!
## 不懂爬虫怎么办?
> 看这里[《进击的虫师》爬取豆瓣电影海报(Top250)](https://www.jianshu.com/p/2e287b534772)

## 自动化制作PPT 的 一二三
#### 先制作PPT模板

> ![插入占位符](https://upload-images.jianshu.io/upload_images/3203841-70c7ef2e1b65461d.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)
> ![占位符种类](https://upload-images.jianshu.io/upload_images/3203841-ad7e60ae023c7521.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

> 制作模板的过程, 就是插入占位符的过程, 可以根据自己的需求插入各种占位符, 比如,豆瓣电影Top250的需求是, 插入图片和文本内容, 那就从占位符中选择, `内容`, `图片`, 插入模板就好, 然后再对模板中的`内容`样式和`图片`位置进行调整, 就能得到符合需求的模板了 
> ![模板长这样](https://upload-images.jianshu.io/upload_images/3203841-0809eb2c14a00615.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)
## 准备数据:
> 我直接把原来写过的,python爬取豆瓣电影的脚本, 运行了一遍, 图片和文本数据就都齐了[《进击的虫师》爬取豆瓣电影海报(Top250)](https://www.jianshu.com/p/2e287b534772)
> ![准备资源](https://upload-images.jianshu.io/upload_images/3203841-53ccae79c5cd9d39.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

## Python编程(将数据按照模板填空, 导出到最终的ppt中)
## 源码如下(注释详尽):
```
from pptx import Presentation
from pptx.util import Inches

# 获取豆瓣电影信息
def getInfo():
    movies_info = []

    with open('./douban_movie_top250.txt') as f:
        for line in f.readlines():
            line_list = line.split("\'")
            one_movie_info = {}
            one_movie_info['index'] = line_list[1]
            one_movie_info['title'] = line_list[3]
            one_movie_info['score'] = line_list[5]

            try:
                one_movie_info['desc'] = line_list[7]
            except:
                one_movie_info['desc'] = ''

            one_movie_info['image_path'] = "./Top250_movie_images/"+ str(line_list[1]) + '_' + line_list[3] + ".jpg"

            movies_info.append(one_movie_info)

    return movies_info


# 创建ppt
def createPpt(movies_info):
    prs = Presentation('model.pptx')
    for movie_info in movies_info:
        # 获取模板个数
        templateStyleNum = len(prs.slide_layouts)
        # 按照第一个模板创建 一张幻灯片
        oneSlide = prs.slides.add_slide(prs.slide_layouts[0])
        # 获取模板可填充的所有位置
        body_shapes = oneSlide.shapes.placeholders
        for index, body_shape in enumerate(body_shapes):
            if index == 0:
                body_shape.text = movie_info['index']+movie_info['title']
            elif index == 1:
                img_path = movie_info['image_path']
                body_shape.insert_picture(img_path)
            elif index == 2:
                body_shape.text = movie_info['desc']
            elif index == 3:
                body_shape.text = movie_info['score']
    # 对ppt的修改  
    prs.save('豆瓣Top250推荐.pptx')


def main():
    # 获取豆瓣电影信息
    movies_info = getInfo()
    createPpt(movies_info)

if __name__ == '__main__':
    main()
```

## Python生成图表(豆瓣电影Top20的评分为例)
> ![Top20报表](https://upload-images.jianshu.io/upload_images/3203841-c5d25baf9ca0da5d.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)
```python
# encoding: utf-8
from pptx import Presentation
from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches

# 创建幻灯片
prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])

# 定义图表数据
chart_data = ChartData()
chart_data.categories = ['肖申克的救赎', '霸王别姬', '这个杀手不太冷', '阿甘正传', '美丽人生', '千与千寻', '泰坦尼克号', '辛德勒的名单', '盗梦空间', '机器人总动员', '海上钢琴师', '三傻大闹宝莱坞', '忠犬八公的故事', '放牛班的春天', '大话西游之大圣娶亲', '楚门的世界', '教父', '龙猫', '熔炉', '乱世佳人']
chart_data.add_series('豆瓣电影', (9.6, 9.5, 9.4, 9.4, 9.5, 9.2, 9.2, 9.4, 9.3, 9.3, 9.2, 9.2, 9.2, 9.2, 9.2, 9.1, 9.2, 9.1, 9.2, 9.2))
# 将图表添加到幻灯片
x, y, cx, cy = Inches(0), Inches(0), Inches(10), Inches(8)
slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data)
prs.save('豆瓣 Top20 评分图.pptx')
```

>  关于数据图形化: Python有很多优秀的图形库, 比如matplotlab, 以及Google推出的在线编程工具colabratory, 都可以方便的实现数据可视化, 掌握了Python图形库的使用, 基本可以和PPT图表说拜拜了...
![colabratory](https://upload-images.jianshu.io/upload_images/3203841-4d0d10432f2e9bde.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)


> 教程涉及到的资源我都通过百度网盘分享给大家,为了便于大家的下载,资源整合到了一张独立的帖子里,链接如下:
[http://www.jianshu.com/p/4f28e1ae08b1](https://www.jianshu.com/p/4f28e1ae08b1)

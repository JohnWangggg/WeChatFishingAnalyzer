import csv
import jieba
from wordcloud import WordCloud
import matplotlib.pyplot as plt
from collections import Counter

# 定义停用词集合
stop_words = set([
    # 原有的停用词
    '的', '了', '和', '是', '就', '都', '而', '及', '与', '着',
    '或', '一个', '没有', '这个', '那个', '这样', '那样', '这些',
    '那些', '在', '我', '你', '他', '她', '它', '们', '可以',
    '这', '那', '不', '也', '很', '但', '还', '到', '对', '说',
    '被', '让', '给', '从', '向', '再', '有', '个', '然后', '因为',
    '已经', '于是', '这么', '那么', '什么', '谁', '为什么',
    
    # 新增的停用词
    '我们', '你们', '他们', '不是', '就是', '现在', '还是', '自己',
    '但是', '怎么', '今天', '不能', '知道', '应该', '还有', '问题',
    '一下', '一直', '一定', '一样', '一些', '时候', '出来', '觉得',
    '可能', '如果', '因为', '所以', '只是', '需要', '不要', '不会',
    '可以', '这样', '那样', '时间', '没有', '什么', '为什么', '怎么样',
    '如何', '哪里', '谁', '啊', '吧', '呢', '吗', '啦', '呀', '哦',
    '一般', '一起', '不过', '之前', '之后', '以前', '以后', '其实',
    '大家', '每个', '每天', '有点', '比较', '完全', '真的', '确实',
    '总是', '一直', '已经', '马上', '立刻', '曾经', '赶紧', '只要',
    '必须', '可是', '或者', '要是', '其他', '别人', '反正', '无法'
    # 新增的停用词
    '看看', '是不是', '感觉', '不如', '直接', '起来',
    '看到', '看着', '看了', '觉得', '感到', '认为',
    '下去', '上来', '进来', '出去', '过来', '回去',
    # 新增的停用词
    "好像","看看","多少","这种","有没有","开始","不行","群里","旺柴","捂脸",
    "东西","估计","有人","一点","很多","肯定","为啥","不到","不用","的话",
    "10","两个","不了","只有","不好","只能","一年","几个"
])
# 存储所有消息
messages = []

# 读取CSV文件
with open("v2ex2.csv", "r", encoding="utf-8") as f:
    reader = csv.reader(f)
    for row in reader:
        nickName = row[10]
        msg = row[7]
        # if nickName == "具体nickName":
        #     if "<" in msg or ">" in msg:
        #         continue
        #     if msg == "":
        #         continue
        #     if "拍了拍" in msg:
        #         continue
        #     if "@" in msg:
        #         continue
            
        #     messages.append(msg)
        if "<" in msg or ">" in msg:
                continue
        if msg == "":
            continue
        # if "拍了拍" in msg:
        #     continue
        # if "@" in msg:
        #     continue
        if "撤回" in msg:
            continue
            
        messages.append(msg)

# 将所有消息合并成一个字符串
text = " ".join(messages)

# 使用jieba进行分词，并过滤停用词
word_list = [word for word in jieba.cut(text) if word not in stop_words and len(word) > 1]
word_space = " ".join(word_list)

# 创建词云对象
wordcloud = WordCloud(
    font_path="C:\\Windows\\Fonts\\msyh.ttc",  # 使用微软雅黑字体
    width=800,
    height=400,
    background_color="white",
    min_font_size=10,
    max_font_size=150
)

# 生成词云
wordcloud.generate(word_space)

# 显示词云图
plt.figure(figsize=(10, 5))
plt.imshow(wordcloud, interpolation="bilinear")
plt.axis("off")
plt.show()

# 保存词云图片（可选）
wordcloud.to_file("wordcloud.png")

# 输出词频统计（可选）
word_counts = Counter(word_list).most_common(20)
print("\n最常见的20个词：")
for word, count in word_counts:
    print(f"{word}: {count}")


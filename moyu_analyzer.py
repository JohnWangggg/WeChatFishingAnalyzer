import csv
import datetime
from collections import Counter
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
import jieba
from wordcloud import WordCloud
import re
import numpy as np
import pandas as pd
from matplotlib.ticker import FuncFormatter
import os
import base64
from io import BytesIO
from matplotlib.patches import Patch

# 重点用户列表（可配置）
key_users = ["用户A","用户B"]  # 请替换为您想要重点关注的用户昵称

# 创建结果目录
result_dir = "moyu_results"
if not os.path.exists(result_dir):
    os.makedirs(result_dir)

# 处理昵称中的表情符号
def clean_nickname(nickname):
    # 移除表情符号和特殊字符
    return re.sub(r'[^\u4e00-\u9fa5a-zA-Z0-9.,，。、？！]', '', nickname)

# 检查消息是否为有效的文本消息（非系统消息、非XML内容）
def is_valid_message(msg):
    """
    检查消息是否为有效的文本消息（非系统消息、非XML内容）
    """
    # 过滤空消息
    if not msg or msg.strip() == "":
        return False
    
    # 过滤包含XML标签的消息
    if "<" in msg or ">" in msg:
        return False
    
    # 过滤撤回消息
    if "撤回" in msg:
        return False
    
    # 过滤包含常见XML属性的消息
    xml_keywords = ["xml", "cdn", "aeskey", "thumburl", "imgurl", "signature", 
                   "platform", "version", "length", "md5", "encryver", "hdwidth", 
                   "hdheight", "thumbwidth", "thumbheight"]
    
    for keyword in xml_keywords:
        if keyword in msg.lower():
            return False
    
    return True

# 定义工作时间（周一至周五，上午9点到下午6点）
def is_work_time(timestamp, nickname=None):
    try:
        # 如果是重点用户，也统计非工作时间的消息
        if nickname in key_users:
            return True
            
        # 转换Unix时间戳为datetime对象
        dt = datetime.datetime.fromtimestamp(int(timestamp))
        # 检查是否为工作日（0=周一，6=周日）
        if dt.weekday() >= 5:  # 周末
            return False
        # 检查是否在工作时间内（9:00-18:00）
        if 9 <= dt.hour < 18:
            return True
        return False
    except Exception as e:
        print(f"时间解析错误: {e}, 时间戳: {timestamp}")
        return False

# 将图表保存为base64编码，用于嵌入HTML
def fig_to_base64(fig):
    buf = BytesIO()
    fig.savefig(buf, format='png', bbox_inches='tight')
    buf.seek(0)
    img_str = base64.b64encode(buf.read()).decode('utf-8')
    buf.close()
    return img_str

# 定义停用词集合
stop_words = set([
    '的', '了', '和', '是', '就', '都', '而', '及', '与', '着',
    '或', '一个', '没有', '这个', '那个', '这样', '那样', '这些',
    '那些', '在', '我', '你', '他', '她', '它', '们', '可以',
    '这', '那', '不', '也', '很', '但', '还', '到', '对', '说',
    '被', '让', '给', '从', '向', '再', '有', '个', '然后', '因为',
    '已经', '于是', '这么', '那么', '什么', '谁', '为什么',
    '我们', '你们', '他们', '不是', '就是', '现在', '还是', '自己',
    '但是', '怎么', '今天', '不能', '知道', '应该', '还有', '问题',
    '一下', '一直', '一定', '一样', '一些', '时候', '出来', '觉得',
    '可能', '如果', '因为', '所以', '只是', '需要', '不要', '不会',
    '可以', '这样', '那样', '时间', '没有', '什么', '为什么', '怎么样',
    '如何', '哪里', '谁', '啊', '吧', '呢', '吗', '啦', '呀', '哦',
    '一般', '一起', '不过', '之前', '之后', '以前', '以后', '其实',
    '大家', '每个', '每天', '有点', '比较', '完全', '真的', '确实',
    '总是', '一直', '已经', '马上', '立刻', '曾经', '赶紧', '只要',
    '必须', '可是', '或者', '要是', '其他', '别人', '反正', '无法',
    '看看', '是不是', '感觉', '不如', '直接', '起来',
    '看到', '看着', '看了', '觉得', '感到', '认为',
    '下去', '上来', '进来', '出去', '过来', '回去',
    '好像', '看看', '多少', '这种', '有没有', '开始', '不行', '群里', 
    '东西', '估计', '有人', '一点', '很多', '肯定', '为啥', '不到', '不用', '的话',
    '10', '两个', '不了', '只有', '不好', '只能', '一年', '几个'
])

# 初始化HTML内容
html_content = """
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>微信群摸鱼排行榜分析报告</title>
    <style>
        body {
            font-family: 'Microsoft YaHei', Arial, sans-serif;
            margin: 0;
            padding: 20px;
            background-color: #f5f5f5;
            color: #333;
        }
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background-color: white;
            padding: 20px;
            box-shadow: 0 0 10px rgba(0,0,0,0.1);
            border-radius: 5px;
        }
        h1, h2, h3 {
            color: #2c3e50;
        }
        h1 {
            text-align: center;
            padding-bottom: 10px;
            border-bottom: 2px solid #3498db;
            margin-bottom: 30px;
        }
        .section {
            margin-bottom: 40px;
            padding-bottom: 20px;
            border-bottom: 1px solid #eee;
        }
        .chart {
            text-align: center;
            margin: 20px 0;
        }
        .chart img {
            max-width: 100%;
            height: auto;
            border: 1px solid #ddd;
            border-radius: 4px;
            padding: 5px;
            background-color: white;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
        }
        th, td {
            padding: 12px 15px;
            text-align: left;
            border-bottom: 1px solid #ddd;
        }
        th {
            background-color: #3498db;
            color: white;
        }
        tr:nth-child(even) {
            background-color: #f2f2f2;
        }
        .highlight {
            background-color: #ffffcc;
            font-weight: bold;
        }
        .footer {
            text-align: center;
            margin-top: 30px;
            color: #7f8c8d;
            font-size: 0.9em;
        }
        .key-user {
            background-color: #e8f8f5;
            font-weight: bold;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>微信群摸鱼排行榜分析报告</h1>
"""

# 存储用户摸鱼数据
moyu_counter = Counter()
user_hour_stats = {}  # 用户在不同时间段的发言数量
user_weekday_stats = {}  # 用户在不同工作日的发言数量
user_messages = {}  # 存储用户的消息内容
work_days = set()  # 记录工作日天数
hour_counter = Counter()  # 时间段分布
weekday_counter = Counter()  # 工作日分布
moyu_content = []  # 所有摸鱼内容

# 读取CSV文件
try:
    with open("v2ex3.csv", "r", encoding="utf-8") as f:
        reader = csv.reader(f)
        header = next(reader, None)  # 假设第一行是表头
        
        # 确定列索引（根据您的CSV文件结构）
        time_index = 5  # CreateTime列
        nickname_index = 10  # NickName列
        msg_index = 7  # StrContent列
        
        for row in reader:
            if len(row) <= max(time_index, nickname_index, msg_index):
                continue  # 跳过格式不正确的行
                
            timestamp = row[time_index]
            nickname = clean_nickname(row[nickname_index])
            msg = row[msg_index]
            
            # 使用更严格的消息过滤条件
            if not is_valid_message(msg):
                continue
            
            try:
                dt = datetime.datetime.fromtimestamp(int(timestamp))
                
                # 记录工作日
                if dt.weekday() < 5:  # 工作日
                    work_days.add(dt.strftime("%Y-%m-%d"))
                    
                    # 统计工作日分布
                    weekday_counter[dt.weekday()] += 1
                
                # 统计时间分布
                hour_counter[dt.hour] += 1
                
                # 检查是否在工作时间发言
                if is_work_time(timestamp, nickname):
                    moyu_counter[nickname] += 1
                    moyu_content.append(msg)
                    
                    # 初始化用户数据结构
                    if nickname not in user_hour_stats:
                        user_hour_stats[nickname] = Counter()
                    if nickname not in user_weekday_stats:
                        user_weekday_stats[nickname] = Counter()
                    if nickname not in user_messages:
                        user_messages[nickname] = []
                    
                    # 统计用户在不同时间段的发言
                    user_hour_stats[nickname][dt.hour] += 1
                    
                    # 统计用户在不同工作日的发言
                    user_weekday_stats[nickname][dt.weekday()] += 1
                    
                    # 存储用户消息
                    user_messages[nickname].append(msg)
            except:
                continue
except Exception as e:
    print(f"读取CSV文件时出错: {e}")
    exit(1)

# 计算总工作日数
total_work_days = len(work_days)
print(f"\n数据集中共有 {total_work_days} 个工作日")

# 获取摸鱼排行前10名
top_moyu = moyu_counter.most_common(10)
# 打印结果
print("\n摸鱼排行榜前10名:")
for i, (name, count) in enumerate(top_moyu, 1):
    avg_per_day = count / total_work_days if total_work_days > 0 else 0
    print(f"{i}. {name}: {count}条工作时间消息 (平均每天 {avg_per_day:.2f}条)")

# 处理重点用户
# 获取前10名用户名列表
top10_names = [name for name, _ in top_moyu]

# 创建扩展排行榜，包含前10名和不在前10名的重点用户
extended_ranking = list(top_moyu)
key_users_not_in_top10 = []

# 添加不在前10名的重点用户
for user in key_users:
    if user in moyu_counter and user not in top10_names:
        key_users_not_in_top10.append((user, moyu_counter[user]))

# 将不在前10名的重点用户添加到扩展排行榜
extended_ranking.extend(key_users_not_in_top10)

# 打印重点用户信息
if key_users_not_in_top10:
    print("\n不在前10名的重点用户:")
    for name, count in key_users_not_in_top10:
        avg_per_day = count / total_work_days if total_work_days > 0 else 0
        print(f"{name}: {count}条工作时间消息 (平均每天 {avg_per_day:.2f}条)")

# 添加基本信息到HTML
html_content += f"""
    <div class="section">
        <h2>基本统计信息</h2>
        <p>数据集中共有 <span class="highlight">{total_work_days}</span> 个工作日</p>
        
        <h3>摸鱼排行榜</h3>
        <table>
            <tr>
                <th>排名</th>
                <th>昵称</th>
                <th>工作时间消息数</th>
                <th>平均每天消息数</th>
                <th>备注</th>
            </tr>
"""

# 添加排行榜到HTML
for i, (name, count) in enumerate(extended_ranking, 1):
    avg_per_day = count / total_work_days if total_work_days > 0 else 0
    is_key_user = name in key_users
    row_class = 'class="key-user"' if is_key_user else ''
    note = "重点关注用户" if is_key_user else ""
    
    html_content += f"""
            <tr {row_class}>
                <td>{i}</td>
                <td>{name}</td>
                <td>{count}</td>
                <td>{avg_per_day:.2f}</td>
                <td>{note}</td>
            </tr>
    """

html_content += """
        </table>
    </div>
"""

# 设置字体
font = FontProperties(fname="C:\\Windows\\Fonts\\msyh.ttc")  # 使用微软雅黑字体

# 设置全局字体
plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']
plt.rcParams['axes.unicode_minus'] = False

# 1. 可视化摸鱼排行榜
plt.figure(figsize=(12, 6))
names = [item[0] for item in extended_ranking]
counts = [item[1] for item in extended_ranking]

# 创建颜色列表，重点用户使用不同颜色
colors = ['skyblue' if name not in key_users else 'lightgreen' for name in names]

# 创建横向条形图
plt.barh(range(len(names)), counts, color=colors)
plt.yticks(range(len(names)), names)
plt.xlabel('工作时间消息数量')
plt.title('微信群摸鱼排行榜')

# 在条形图上显示具体数值
for i, v in enumerate(counts):
    plt.text(v + 1, i, str(v), va='center')

# 添加图例
legend_elements = [
    Patch(facecolor='skyblue', label='普通用户'),
    Patch(facecolor='lightgreen', label='重点关注用户')
]
plt.legend(handles=legend_elements, loc='lower right')

plt.tight_layout()
ranking_img = fig_to_base64(plt.gcf())
plt.savefig(os.path.join(result_dir, "moyu_ranking.png"))
plt.close()

# 添加排行榜图表到HTML
html_content += """
    <div class="section">
        <h2>摸鱼排行榜</h2>
        <div class="chart">
            <img src="data:image/png;base64,""" + ranking_img + """" alt="摸鱼排行榜">
        </div>
    </div>
"""

# 2. 可视化时间分布
plt.figure(figsize=(12, 6))
hours = list(range(24))
counts = [hour_counter[hour] for hour in hours]

plt.bar(hours, counts, color='lightgreen')
plt.xlabel('小时')
plt.ylabel('消息数量')
plt.title('24小时消息分布')
plt.xticks(hours)

# 标记工作时间区域
plt.axvspan(9, 18, alpha=0.2, color='red')
plt.text(13.5, max(counts)*0.9, '工作时间', ha='center')

plt.tight_layout()
time_dist_img = fig_to_base64(plt.gcf())
plt.savefig(os.path.join(result_dir, "moyu_time_distribution.png"))
plt.close()

# 添加时间分布到HTML
html_content += """
    <div class="section">
        <h2>时间分布分析</h2>
        <div class="chart">
            <img src="data:image/png;base64,""" + time_dist_img + """" alt="24小时消息分布">
        </div>
        
        <h3>工作时间内的消息分布</h3>
        <table>
            <tr>
                <th>时间段</th>
                <th>消息数量</th>
            </tr>
"""

for hour in range(9, 18):
    html_content += f"""
            <tr>
                <td>{hour}点-{hour+1}点</td>
                <td>{hour_counter[hour]}</td>
            </tr>
    """
    print(f"{hour}点-{hour+1}点: {hour_counter[hour]}条消息")

html_content += """
        </table>
    </div>
"""

# 3. 创建详细分析用户列表（前三名 + 重点用户）
detailed_users = [name for name, _ in top_moyu[:3]]  # 取前三名用户

# 添加不在前三名的重点用户
for user in key_users:
    if user in moyu_counter and user not in detailed_users:
        detailed_users.append(user)

# 分析详细用户的摸鱼时间分布
print("\n===== 摸鱼达人时间分析 =====")

plt.figure(figsize=(14, 8))
bar_width = 0.25 / (len(detailed_users) if len(detailed_users) > 0 else 1)  # 调整柱状图宽度
index = range(9)  # 9小时工作时间

for i, user in enumerate(detailed_users):
    if user in user_hour_stats:
        user_data = [user_hour_stats[user][hour] for hour in range(9, 18)]
        plt.bar([x + i * bar_width for x in index], user_data, bar_width, 
                label=user, alpha=0.7)

plt.xlabel('工作时间')
plt.ylabel('消息数量')
plt.title('摸鱼达人时间分布对比')
plt.xticks([x + bar_width * len(detailed_users) / 2 for x in index], [f"{h}点" for h in range(9, 18)])
plt.legend()
plt.tight_layout()
detailed_users_time_img = fig_to_base64(plt.gcf())
plt.savefig(os.path.join(result_dir, "detailed_users_time_distribution.png"))
plt.close()

# 添加摸鱼达人时间分析到HTML
html_content += """
    <div class="section">
        <h2>摸鱼达人时间分析</h2>
        <div class="chart">
            <img src="data:image/png;base64,""" + detailed_users_time_img + """" alt="摸鱼达人时间分布对比">
        </div>
    </div>
"""

# 4. 计算并显示每人每天平均摸鱼消息数
if total_work_days > 0:
    print("\n每人每天平均摸鱼消息数:")
    for i, (name, count) in enumerate(extended_ranking, 1):
        avg_per_day = count / total_work_days
        print(f"{i}. {name}: {avg_per_day:.2f}条/天")

    # 可视化平均摸鱼效率
    plt.figure(figsize=(12, 6))
    names = [item[0] for item in extended_ranking]
    avg_counts = [item[1]/total_work_days for item in extended_ranking]
    
    # 创建颜色列表，重点用户使用不同颜色
    colors = ['orange' if name not in key_users else 'lightgreen' for name in names]
    
    plt.barh(range(len(names)), avg_counts, color=colors)
    plt.yticks(range(len(names)), names)
    plt.xlabel('平均每天摸鱼消息数')
    plt.title('微信群日均摸鱼效率排行')
    
    for i, v in enumerate(avg_counts):
        plt.text(v + 0.1, i, f"{v:.2f}", va='center')
    
    # 添加图例
    legend_elements = [
        Patch(facecolor='orange', label='普通用户'),
        Patch(facecolor='lightgreen', label='重点关注用户')
    ]
    plt.legend(handles=legend_elements, loc='lower right')
    
    plt.tight_layout()
    efficiency_img = fig_to_base64(plt.gcf())
    plt.savefig(os.path.join(result_dir, "moyu_efficiency.png"))
    plt.close()
    
    # 添加摸鱼效率到HTML
    html_content += """
        <div class="section">
            <h2>摸鱼效率分析</h2>
            <div class="chart">
                <img src="data:image/png;base64,""" + efficiency_img + """" alt="微信群日均摸鱼效率排行">
            </div>
        </div>
    """

# 5. 可视化周一至周五的摸鱼趋势
plt.figure(figsize=(10, 6))
weekdays = ['周一', '周二', '周三', '周四', '周五']
counts = [weekday_counter[day] for day in range(5)]

plt.bar(weekdays, counts, color='purple')
plt.xlabel('工作日')
plt.ylabel('消息数量')
plt.title('工作日摸鱼趋势')

for i, v in enumerate(counts):
    plt.text(i, v + 100, str(v), ha='center')

plt.tight_layout()
weekday_trend_img = fig_to_base64(plt.gcf())
plt.savefig(os.path.join(result_dir, "weekday_trend.png"))
plt.close()

# 添加工作日趋势到HTML
html_content += """
    <div class="section">
        <h2>工作日摸鱼趋势</h2>
        <div class="chart">
            <img src="data:image/png;base64,""" + weekday_trend_img + """" alt="工作日摸鱼趋势">
        </div>
    </div>
"""

# 6. 详细用户的工作日摸鱼对比
plt.figure(figsize=(12, 6))
bar_width = 0.25 / (len(detailed_users) if len(detailed_users) > 0 else 1)  # 调整柱状图宽度
index = range(5)  # 5个工作日

for i, user in enumerate(detailed_users):
    if user in user_weekday_stats:
        user_data = [user_weekday_stats[user][day] for day in range(5)]
        plt.bar([x + i * bar_width for x in index], user_data, bar_width, 
                label=user, alpha=0.7)

plt.xlabel('工作日')
plt.ylabel('消息数量')
plt.title('摸鱼达人工作日分布对比')
plt.xticks([x + bar_width * len(detailed_users) / 2 for x in index], weekdays)
plt.legend()
plt.tight_layout()
detailed_users_weekday_img = fig_to_base64(plt.gcf())
plt.savefig(os.path.join(result_dir, "detailed_users_weekday_distribution.png"))
plt.close()

# 添加摸鱼达人工作日分布到HTML
html_content += """
    <div class="section">
        <h2>摸鱼达人工作日分布</h2>
        <div class="chart">
            <img src="data:image/png;base64,""" + detailed_users_weekday_img + """" alt="摸鱼达人工作日分布对比">
        </div>
    </div>
"""

# 7. 摸鱼内容分析
print("\n===== 摸鱼内容分析 =====")

# 将所有消息合并成一个字符串
text = " ".join(moyu_content)

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
plt.title('摸鱼内容词云')
plt.tight_layout()
wordcloud_img = fig_to_base64(plt.gcf())
plt.savefig(os.path.join(result_dir, "moyu_content_wordcloud.png"))
plt.close()

# 输出词频统计
word_counts = Counter(word_list).most_common(20)
print("\n摸鱼内容中最常见的20个词：")
for word, count in word_counts:
    print(f"{word}: {count}次")

# 添加摸鱼内容分析到HTML
html_content += """
    <div class="section">
        <h2>摸鱼内容分析</h2>
        <div class="chart">
            <img src="data:image/png;base64,""" + wordcloud_img + """" alt="摸鱼内容词云">
        </div>
        
        <h3>摸鱼内容中最常见的20个词</h3>
        <table>
            <tr>
                <th>词语</th>
                <th>出现次数</th>
            </tr>
"""

for word, count in word_counts:
    html_content += f"""
            <tr>
                <td>{word}</td>
                <td>{count}</td>
            </tr>
    """

html_content += """
        </table>
    </div>
"""

# 8. 重点用户的摸鱼内容分析
print("\n===== 重点用户摸鱼内容分析 =====")

for user in key_users:
    if user in user_messages and len(user_messages[user]) > 0:
        print(f"\n{user}的摸鱼内容分析:")
        
        # 将用户消息合并成一个字符串
        user_text = " ".join(user_messages[user])
        
        # 使用jieba进行分词，并过滤停用词
        user_word_list = [word for word in jieba.cut(user_text) if word not in stop_words and len(word) > 1]
        user_word_space = " ".join(user_word_list)
        
        # 创建词云对象
        user_wordcloud = WordCloud(
            font_path="C:\\Windows\\Fonts\\msyh.ttc",  # 使用微软雅黑字体
            width=800,
            height=400,
            background_color="white",
            min_font_size=10,
            max_font_size=150
        )
        
        # 生成词云
        if user_word_space.strip():  # 确保有内容再生成词云
            user_wordcloud.generate(user_word_space)
            
            # 显示词云图
            plt.figure(figsize=(10, 5))
            plt.imshow(user_wordcloud, interpolation="bilinear")
            plt.axis("off")
            plt.title(f'{user}的摸鱼内容词云')
            plt.tight_layout()
            user_wordcloud_img = fig_to_base64(plt.gcf())
            plt.savefig(os.path.join(result_dir, f"{user}_wordcloud.png"))
            plt.close()
            
            # 输出词频统计
            user_word_counts = Counter(user_word_list).most_common(10)
            print(f"\n{user}摸鱼内容中最常见的10个词：")
            for word, count in user_word_counts:
                print(f"{word}: {count}次")
            
            # 添加用户摸鱼内容分析到HTML
            html_content += f"""
                <div class="section">
                    <h2>{user}的摸鱼内容分析</h2>
                    <div class="chart">
                        <img src="data:image/png;base64,{user_wordcloud_img}" alt="{user}的摸鱼内容词云">
                    </div>
                    
                    <h3>{user}摸鱼内容中最常见的10个词</h3>
                    <table>
                        <tr>
                            <th>词语</th>
                            <th>出现次数</th>
                        </tr>
            """
            
            for word, count in user_word_counts:
                html_content += f"""
                        <tr>
                            <td>{word}</td>
                            <td>{count}</td>
                        </tr>
                """
            
            html_content += """
                    </table>
                </div>
            """
        else:
            print(f"{user}的摸鱼内容不足以生成词云")

# 添加页脚到HTML
html_content += """
        <div class="footer">
            <p>微信群摸鱼排行榜分析报告 - 生成时间: """ + datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S") + """</p>
        </div>
    </div>
</body>
</html>
"""

# 保存HTML报告
with open(os.path.join(result_dir, "moyu_report.html"), "w", encoding="utf-8") as f:
    f.write(html_content)

print(f"\n分析报告已保存到 {os.path.join(result_dir, 'moyu_report.html')}")
markdown
# 淘宝评论批量采集系统

> 基于 DrissionPage 的淘宝/天猫商品评论采集工具，支持自动滚动、验证码检测、断点续传

## 技术栈

- Python 3.9+
- DrissionPage（浏览器自动化）
- Pandas（数据处理）
- 正则表达式 + JSON 解析

## 核心功能

- 📦 **浏览器自动化**：自动打开商品页，点击评价入口
- 🔄 **动态加载**：模拟滚动，触发评论加载
- 📡 **网络监听**：拦截评论接口数据包（`rate.detaillist.get`、`rate.tmall`）
- 🧩 **断点续传**：已爬商品自动跳过，中断可恢复
- 🔐 **验证码处理**：自动检测滑块验证码，暂停等待手动通过
- 📊 **数据导出**：追加写入 CSV，支持内存去重

## 快速开始

### 1. 安装依赖

```bash
pip install -r requirements.txt
2. 准备 Excel 文件
创建 商品链接.xlsx，第一列放商品链接：

A
https://item.taobao.com/item.htm?id=123456
https://detail.tmall.com/item.htm?id=789012
3. 运行
bash
python batch_crawl_reviews.py
4. 扫码登录
程序会自动打开淘宝登录页，扫码登录后按回车开始。

输出结果
生成 苹果汇总_含商品信息.csv，包含字段：

字段	说明
商品ID	从 URL 提取
商品名称	页面标题
活动价/原价	价格信息
用户/内容/时间	评论核心数据
SKU	商品规格
图片	评论图片链接
追评	追评内容
项目结构
text
├── batch_crawl_reviews.py   # 主程序
├── 商品链接.xlsx             # 输入文件（需自己准备）
├── 苹果汇总_含商品信息.csv   # 输出文件
└── requirements.txt         # 依赖
演示效果
单商品可抓取 800-1000 条评论，日均采集 5000+ 条，数据完整率 95% 以上。

注意事项
首次运行需要手动扫码登录

遇到验证码会暂停，手动通过后自动继续

建议使用稳定的网络环境，避免频繁验证码

作者
GitHub: tainliu746-dotcom

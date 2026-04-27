import time
import json
import random
import re
import pandas as pd
import os
from urllib.parse import urlparse, parse_qs
from DrissionPage import ChromiumPage, ChromiumOptions

# ================= 配置区 =================
# 输入文件 (你的原始Excel文件)
INPUT_FILE = '商品链接.xlsx'

# 输出文件 (结果会自动追加到这里)
OUTPUT_FILE = '苹果汇总_含商品信息.csv'

# 每个商品最大抓取滚动次数 (40次滚动约等于800-1000条评论)
MAX_SCROLL_PER_ITEM = 60


# =========================================

def extract_id_from_url(url):
    """从链接中提取ID，用于日志显示"""
    try:
        parsed = urlparse(url)
        params = parse_qs(parsed.query)
        if 'id' in params: return params['id'][0]
    except:
        pass
    return "未知ID"


def clean_mtop_jsonp(body):
    """解析淘宝复杂的JSONP格式数据"""
    try:
        if isinstance(body, dict): return body
        if isinstance(body, str):
            # 提取 mtopjsonp1(...) 中的内容
            match = re.search(r'mtopjsonp\d+\((.*)\)', body, re.DOTALL)
            if match: return json.loads(match.group(1))
            # 备用提取
            if '(' in body and body.strip().endswith(')'):
                s = body.find('(') + 1
                e = body.rfind(')')
                return json.loads(body[s:e])
    except:
        pass
    return None


def scroll_by_dragging(page):
    """
    通过拖拽最后一条评论元素来触发加载
    这是解决'容器滚动失效'最稳妥的方法
    """
    try:
        cards = page.eles('xpath://div[contains(@class, "Comment--")]')
        if cards:
            last_card = cards[-1]
            # 强制将最后一个元素顶到视口顶部
            page.run_js("arguments[0].scrollIntoView({behavior: 'auto', block: 'start'});", last_card)
            return True
        return False
    except:
        return False


def handle_captcha(page):
    """
    检测是否有滑块验证码
    如果有，暂停脚本并发出警告，直到用户手动划过
    """
    is_captcha = False
    # 检测常见的滑块元素特征
    if page.ele('xpath://iframe[contains(@src, "baxia")]', timeout=0.5) or \
            page.ele('text:向右滑动', timeout=0.5) or \
            page.ele('@id:nc_1_n1z', timeout=0.5) or \
            "验证" in page.title:

        is_captcha = True
        print("\n" + "!" * 50)
        print("【警告】检测到滑块验证码！脚本已暂停！")
        print("请在浏览器中手动完成滑块验证...")
        print("验证成功后，脚本会自动检测并继续。")
        print("!" * 50 + "\a")  # 发出系统提示音

        # 死循环等待验证消失
        while True:
            # 再次检查滑块是否存在
            if not (page.ele('xpath://iframe[contains(@src, "baxia")]', timeout=0.2) or \
                    page.ele('text:向右滑动', timeout=0.2) or \
                    page.ele('@id:nc_1_n1z', timeout=0.2)):
                print(">>> 验证已通过，恢复抓取...")
                time.sleep(2)  # 等待页面刷新或数据加载
                break
            time.sleep(1)

    return is_captcha


def get_product_basic_info(page):
    """
    提取商品标题、现价、原价
    已适配：政府补贴页、天猫新版、淘宝旧版
    """
    title = "未获取"
    current_price = "未获取"
    original_price = "未获取"

    # === 1. 提取标题 ===
    try:
        title_ele = None
        selectors = [
            'xpath://span[contains(@class, "mainTitle")]',  # 补贴页/新版SSR
            'xpath://div[contains(@class, "MainTitle")]',  # 补贴页变体
            'xpath://h1',  # 天猫通用
            'xpath://div[@id="J_Title"]/h3',  # 淘宝通用
            'xpath://h3[contains(@class, "tb-main-title")]',  # 淘宝旧版
        ]
        for sel in selectors:
            title_ele = page.ele(sel, timeout=0.1)
            if title_ele: break

        if title_ele:
            title = title_ele.text.strip()
    except:
        pass

    # === 2. 提取现价 (大字价格) ===
    try:
        price_ele = None
        selectors = [
            'xpath://div[contains(@class, "highlightPrice")]//span[contains(@class, "text")]',  # 补贴页
            'xpath://span[contains(@class, "Price--priceText")]',  # 天猫新版
            'xpath://div[contains(@class, "Price--current")]',
            'xpath://em[@id="J_PromoPriceNum"]',  # 旧版促销
            'css:.tm-price',
        ]
        for sel in selectors:
            eles = page.eles(sel)
            if eles:
                # 取最后一个通常是纯数字 (避开人民币符号)
                price_ele = eles[-1]
                break

        if price_ele:
            current_price = price_ele.text.strip().replace('¥', '').replace('￥', '')
    except:
        pass

    # === 3. 提取原价 (划线价) ===
    try:
        org_ele = None
        selectors = [
            'xpath://div[contains(@class, "subPrice")]//span[contains(@class, "text")]',  # 补贴页
            'xpath://span[contains(@class, "Price--extraPrice")]',  # 天猫新版划线
            'xpath://span[contains(@class, "originalPrice")]',
            'xpath://dl[@id="J_StrPriceMod"]//span[contains(@class, "tm-price")]',  # 旧版原价
        ]
        for sel in selectors:
            eles = page.eles(sel)
            if eles:
                org_ele = eles[-1]
                break

        if org_ele:
            original_price = org_ele.text.strip().replace('¥', '').replace('￥', '')
    except:
        pass

    print(f"    [商品信息] {title[:15]}... | 现价:{current_price} | 原价:{original_price}")
    return title, current_price, original_price


def crawl_one_url(page, url):
    """单个商品的完整抓取流程"""
    item_id = extract_id_from_url(url)
    print(f"\n>>> [商品ID: {item_id}] 正在打开链接...")

    try:
        page.get(url)
    except Exception as e:
        print(f"    打开链接失败: {e}")
        return []

    time.sleep(3)  # 等待DOM加载

    # 1. 先抓取商品基础信息
    p_title, p_price, p_origin_price = get_product_basic_info(page)

    # 2. 开启网络监听
    page.listen.start(targets=['rate.detaillist.get', 'rate.tmall'])

    # 3. 寻找并点击“全部评价”
    try:
        btn = page.ele('text:查看全部评价', timeout=2)
        if not btn: btn = page.ele('text:全部评价', timeout=1)
        if btn:
            page.run_js("arguments[0].scrollIntoView();", btn)
            time.sleep(0.5)
            btn.click(by_js=True)
            print("    已点击评价入口")
            time.sleep(2)
    except:
        print("    未找到显式评价入口，尝试直接滚动")

    item_reviews = []
    no_data_count = 0

    # 鼠标聚焦到页面中间，防止滚轮失效
    try:
        page.actions.move_to(x=500, y=500).click()
    except:
        pass

    # 4. 循环滚动
    for i in range(MAX_SCROLL_PER_ITEM):
        print(f"    -> 滚动 {i + 1}/{MAX_SCROLL_PER_ITEM}", end="", flush=True)

        # A. 检查验证码 (如果有验证码，这里会卡住等待)
        was_captcha = handle_captcha(page)
        if was_captcha:
            no_data_count = 0  # 验证回来后重置计数器
            page.listen.clear()  # 清空积压数据

        # B. 执行滚动
        scrolled = scroll_by_dragging(page)
        if not scrolled:
            page.actions.scroll(-600)  # 备用方案

        # C. 等待数据包
        res = page.listen.wait(count=1, timeout=3)

        got_new = False
        if res:
            data = clean_mtop_jsonp(res.response.body)
            if data:
                rate_list = []
                # 兼容多种数据结构
                if 'data' in data and isinstance(data['data'], dict):
                    rate_list = data['data'].get('rateList', [])
                elif 'rateList' in data:
                    rate_list = data['rateList']
                elif 'items' in data:
                    rate_list = data['items']

                if rate_list:
                    new_cnt = 0
                    for item in rate_list:
                        content = item.get('feedback', '') or item.get('rateContent', '') or item.get('content', '')
                        if not content: content = "无文本"

                        # 图片提取
                        pic_list = item.get('feedPicPathList', []) or [p.get('url', '') for p in item.get('pics', [])]
                        pic_list = [('https:' + p if p.startswith('//') else p) for p in pic_list]

                        row = {
                            '商品ID': item_id,
                            '商品名称': p_title,
                            '活动价': p_price,
                            '原价': p_origin_price,
                            '用户': item.get('userNick', item.get('displayUserNick', '匿名')),
                            '内容': content,
                            '时间': item.get('createTime', item.get('rateDate', '')),
                            'SKU': item.get('skuValueStr', item.get('auctionSku', '')),
                            '图片': str(pic_list),
                            '追评': item.get('append', {}).get('content', '') if item.get('append') else '',
                            '原始链接': url
                        }

                        # 内存简单去重
                        key = f"{row['用户']}|{row['内容'][:15]}"
                        if not any(f"{x['用户']}|{x['内容'][:15]}" == key for x in item_reviews):
                            item_reviews.append(row)
                            new_cnt += 1

                    if new_cnt > 0:
                        print(f" [获取 {new_cnt} 条]")
                        got_new = True
                        no_data_count = 0
                    else:
                        print(" [重复]")
        else:
            print(" [等待]")

        # D. 智能停止 (连续无新数据)
        if not got_new:
            no_data_count += 1
            if no_data_count >= 1:
                print("\n    连续无新数据，本商品结束。")
                break

        time.sleep(random.uniform(1.2, 1.5))

    return item_reviews


def run_main():
    edge_path = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    # 1. 读取链接
    print(f"正在读取 {INPUT_FILE} ...")
    try:
        # header=None 默认读取第一列
        df = pd.read_excel(INPUT_FILE, header=None)
        url_list = df.iloc[:, 0].dropna().astype(str).tolist()
        url_list = [u for u in url_list if u.startswith('http')]
        print(f"成功加载 {len(url_list)} 个有效链接。")
    except Exception as e:
        print(f"读取 Excel 失败: {e}")
        return

    # 2. 启动浏览器
    co = ChromiumOptions()
    co.set_browser_path(edge_path)  # 告诉 DrissionPage 用 Edge
    page = ChromiumPage(addr_or_opts=co)

    # 登录页
    page.get("https://login.taobao.com/member/login.jhtml")
    print("=" * 50)
    print("【请手动扫码登录】")
    input(">>> 登录成功后，请在控制台按回车开始: ")
    print("=" * 50)

    # 3. 断点续传 (检查已完成的链接)
    processed_urls = []
    if os.path.exists(OUTPUT_FILE):
        try:
            df_exist = pd.read_csv(OUTPUT_FILE, usecols=['原始链接'])
            processed_urls = df_exist['原始链接'].astype(str).unique().tolist()
            print(f"检测到历史文件，将跳过 {len(processed_urls)} 个已爬链接。")
        except:
            pass

    # 4. 遍历链接
    for url in url_list:
        url = url.strip()+"&mi_id=0000fnNeh4UB38c14E-nsXNt5fnBuAMng-c5b4swKBFCTVU&priceTId=213e0a5017682144090011268e1918&pvid=8491b9f0-1921-46d3-ae49-e869600f5a58&scm=1007.56401.415219.0&spm=tbpc.pc_sem_alimama%2Fa.201876.d2.6e912a89KKzHla&xxc=home_recommend"

        if url in processed_urls:
            continue

        page.listen.clear()

        reviews = crawl_one_url(page, url)

        if reviews:
            df_new = pd.DataFrame(reviews)
            # 调整列顺序
            cols = ['商品ID', '商品名称', '活动价', '原价', '用户', '内容', '时间', 'SKU', '追评', '图片', '原始链接']
            # reindex 确保列存在
            df_new = df_new.reindex(columns=cols)

            # 追加写入
            hdr = not os.path.exists(OUTPUT_FILE)
            df_new.to_csv(OUTPUT_FILE, mode='a', header=hdr, index=False, encoding='utf-8-sig')
            print(f"    -> 已保存 {len(reviews)} 条评论。")
        else:
            print("    -> 未抓取到有效数据。")

        # 商品间随机休眠，防风控
        sleep_t = random.uniform(3, 5)
        print(f"    休息 {sleep_t:.1f} 秒...")
        time.sleep(sleep_t)

    print("\n所有链接处理完毕！")


if __name__ == "__main__":
    run_main() 
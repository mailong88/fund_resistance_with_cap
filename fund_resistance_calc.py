"""
历史概念流通市值与资金阻力计算

任务：
a. 通过akshare获取同花顺的概念列表
b. 用iwencai获取每个概念的成分股数据(查询语句：XXX概念；上市时间早于2024年1月1日)
   保存在pool文件夹中，以BK+板块编号.json命名
c. 用成分股的流通市值计算每个概念的数据提取日的流通市值
d. 用baostock获取每个概念的历史数据并计算每个交易日的概念流通市值
e. 用市场资金流向分析_east.py中的逻辑计算每个概念的n日资金流向,进而计算每个概念每天的资金阻力
"""

import os
import json
import time
import akshare as ak
import pandas as pd
import baostock as bs
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading
import requests
import re
import execjs
from lxml import etree

# ============== 配置 ==============
POOL_FOLDER = "pools"
OUTPUT_FOLDER = "output"
CUTOFF_DATE = "2024-01-01"  # 上市时间早于此日期

# ============== Step A: 获取同花顺概念列表 ==============
def get_ths_concept_list():
    """通过akshare获取同花顺概念列表"""
    print("=" * 50)
    print("Step A: 获取同花顺概念列表...")
    
    try:
        # 获取所有概念板块
        df = ak.stock_board_concept_name_em()
        print(f"成功获取 {len(df)} 个概念板块")
        
        # 保存概念列表
        concept_list = []
        for _, row in df.iterrows():
            concept_list.append({
                'code': row['代码'],
                'name': row['名称']
            })
        
        return concept_list
    except Exception as e:
        print(f"获取概念列表失败: {e}")
        return []

# ============== Step B: iwencai获取成分股数据 ==============
class WenCai:
    """同花顺问财数据获取类"""
    def __init__(self, proxy=None):
        self.cookies = 'cid=f339b3b06f20f9a3f57511efc61f69561732676413; cid=f339b3b06f20f9a3f57511efc61f69561732676413; ComputerID=f339b3b06f20f9a3f57511efc61f69561732676413; WafStatus=0; other_uid=Ths_iwencai_Xuangu_dhfyapc0j8s63crisy4bluu2kqf1zxxu; ta_random_userid=ofrg6m3r2p; u_ukey=A10702B8689642C6BE607730E11E6E4A; u_uver=1.0.0; u_dpass=8vbH2AZwsYKOAQB90LnApBF6tFP6foVqAiLa0I6LV%2BGq52j91%2FiAPfxZW3mxg421Hi80LrSsTFH9a%2B6rtRvqGg%3D%3D; u_did=A962F087FEBC4DCBB0DFB4AFD8A929A4; u_ttype=WEB; ttype=WEB; user=MDptb180NjE5NzM3MTA6Ok5vbmU6NTAwOjQ3MTk3MzcxMDo3LDExMTExMTExMTExLDQwOzQ0LDExLDQwOzYsMSw0MDs1LDEsNDA7MSwxMDEsNDA7MiwxLDQwOzMsMSw0MDs1LDEsNDA7OCwwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMSw0MDsxMDIsMSw0MDoyNDo6OjQ2MTk3MzcxMDoxNzYzMjA5NTgxOjo6MTUzNjI0MTgwMDo2MDQ4MDA6MDoxMGI4ZDIzNTBiZDg3YzVmOTIxZjU5OTQyNjcyZWNjNzI6ZGVmYXVsdF81OjE%3D; userid=461973710; u_name=mo_461973710; escapename=mo_461973710; ticket=5ac58fb6a7ab2594ed9adaf1ea804370; user_status=0; utk=3f126086ef468b4c1000cbe2f463587f; sess_tk=eyJ0eXAiOiJKV1QiLCJhbGciOiJFUzI1NiIsImtpZCI6InNlc3NfdGtfMSIsImJ0eSI6InNlc3NfdGsifQ.eyJqdGkiOiI3MmNjMmU2NzQyOTlmNTIxZjljNTg3YmQ1MDIzOGQwYjEiLCJpYXQiOjE3NjMyMDk1ODEsImV4cCI6MTc2MzgxNDM4MSwic3ViIjoiNDYxOTczNzEwIiwiaXNzIjoidXBhc3MuaXdlbmNhaS5jb20iLCJhdWQiOiIyMDIwMTExODUyODg5MDcyIiwiYWN0Ijoib2ZjIiwiY3VocyI6IjA0NGFjNWU5YTQwZTRhNTMxOGJlMGU1Njg5NmZhYWEyOTVhYzhmNDE3MDdlOTNmODk2ZmVhZDFkOGJmNzRhNDAifQ._ZW4ZCilVpVdpW6p7hE1t0cB09F5GGoFEL8l0w6EIsI5Yj-Hq-vAfGauJZ0XXPXnb5OSeBY3u-pNuJgDHZ6m5A; cuc=4mt9un9omypt; v=A6K-n632P-XnTyM5bSrpuKa98yMB86Q-WPaaMew4zOXBa0yd1IP2HSiH6mO_'
        self.session = requests.session()
        lst_cookies = self.cookies.split(';')
        for cookie in lst_cookies:
            cookie_parts = cookie.split('=')
            if len(cookie_parts) >= 2:
                self.session.cookies.set(cookie_parts[0].strip(), cookie_parts[1].strip())
        
        headers = {
            "Accept": "application/json, text/plain, */*",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
            "Cache-control": "no-cache",
            "Connection": "keep-alive",
            "Content-Type": "application/x-www-form-urlencoded",
            "Origin": "https://www.iwencai.com",
            "Pragma": "no-cache",
            "Referer": "https://www.iwencai.com/unifiedwap/result?w=%E4%BA%92%E8%81%94%E7%BD%91%E9%87%91%E8%9E%8D%E6%A6%82%E5%BF%B5%EF%BC%9B2024-09-19%E8%87%B32024-09-25%E6%B6%A8%E5%B9%85%E5%A4%A7%E4%BA%8E10%25%EF%BC%9B2024-09-19%E8%87%B32024-09-25%E8%BF%9E%E6%9D%BF%E5%A4%A9%E6%95%B0%EF%BC%9B2024-09-25%E6%88%90%E4%BA%A4%E9%87%91%E9%A2%9D%EF%BC%9B2024-09-26%E5%9D%87%E4%BB%B7%EF%BC%9B2024-09-25%E5%9D%87%E4%BB%B7&querytype=stock&addSign=1727709784457",
            "Sec-Fetch-Dest": "empty",
            "Sec-Fetch-Mode": "cors",
            "Sec-Fetch-Site": "same-origin",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36 Edg/129.0.0.0",
            "sec-ch-ua": "\"Microsoft Edge\";v=\"129\", \"Not=A?Brand\";v=\"8\", \"Chromium\";v=\"129\"",
            "sec-ch-ua-mobile": "?0",
            "sec-ch-ua-platform": "\"Windows\""
        }
        self.session.headers = headers
        self.hexin_v = self.get_hexin_v(self.get_server_time())

    def get_server_time(self, proxy=None):
        url = 'http://www.iwencai.com/unifiedwap/home/index'
        try:
            if proxy:
                resp = self.session.get(url, proxies=proxy, timeout=10)
            else:
                resp = self.session.get(url, timeout=10)
            resp_text = resp.text
            tree = etree.HTML(resp_text)
            js_url = "http:" + tree.xpath("//script[1]/@src")[0]
            resp.close()
            js_resp = self.session.get(js_url, timeout=10)
            js_text = js_resp.text
            obj = re.compile(r'var TOKEN_SERVER_TIME=(?P<time>.*?);var n=function')
            server_time = obj.search(js_text).group('time')
            return server_time
        except Exception as e:
            print(f"获取服务器时间失败: {e}")
            return str(int(time.time() * 1000))

    def get_hexin_v(self, time_val):
        try:
            with open("kou.js", "r", encoding='utf-8') as f:
                js_content = f.read()
                js_content = 'var TOKEN_SERVER_TIME=' + str(time_val) + ";\n" + js_content
                js = execjs.compile(js_content)
                v = js.call("rt.updata")
                return v
        except Exception as e:
            print(f"获取hexin_v失败: {e}")
            return ""

    def get_answer(self, question, secondary_intent="stock", proxy_manager=None):
        url = 'http://www.iwencai.com/customized/chart/get-robot-data'
        data = {
            'add_info': "{\"urp\":{\"scene\":1,\"company\":1,\"business\":1},\"contentType\":\"json\",\"searchInfo\":true}",
            'block_list': "",
            'log_info': "{\"input_type\":\"typewrite\"}",
            'page': 1,
            'perpage': 100,
            'query_area': "",
            'question': question,
            'rsh': "Ths_iwencai_Xuangu_y1wgpofrs18ie6hdpf0dvhkzn2myx8yq",
            'secondary_intent': secondary_intent,
            'source': "Ths_iwencai_Xuangu",
            'version': "2.0"
        }
        
        result = []
        try:
            self.session.headers['hexin-v'] = self.hexin_v
            self.session.headers['Content-Type'] = 'application/json'
            
            if proxy_manager:
                proxy = proxy_manager.get_current_proxy()
                resp = self.session.post(url, data=json.dumps(data), proxies=proxy, timeout=30)
            else:
                resp = self.session.post(url, data=json.dumps(data), timeout=30)
            
            resp_data = resp.json()
            resp.close()
            
            if 'data' in resp_data.keys():
                result = resp_data['data']['answer'][0]['txt'][0]['content']['components'][0]['data']['datas']
                
                if len(result) == 100:
                    result.extend(self.get_page_data(resp_data, question, secondary_intent))
        except Exception as e:
            print(f"获取数据失败: {e}")
        
        return result

    def get_page_data(self, f_resp_data, question, secondary_intent):
        try:
            comp_id = f_resp_data["data"]["answer"][0]["txt"][0]["content"]["components"][0]["cid"]
            uuid = f_resp_data["data"]["answer"][0]["txt"][0]["content"]["components"][0]["data"]["meta"]["uuids"]
            self.session.headers['Content-Type'] = 'application/x-www-form-urlencoded'
            url = "https://www.iwencai.com/gateway/urp/v7/landing/getDataList"
            page = 2
            result = []
            
            while 1:
                data = {
                    "query": question,
                    "urp_sort_way": "desc",
                    "urp_sort_index": "最新涨跌幅",
                    "page": page,
                    "perpage": "100",
                    "addheaderindexes": "",
                    "condition": "",
                    "codelist": "",
                    "indexnamelimit": "",
                    "logid": "",
                    "ret": "json_all",
                    "sessionid": "",
                    "source": "Ths_iwencai_Xuangu",
                    "date_range[0]": "",
                    "date_range[1]": "",
                    "iwc_token": "",
                    "urp_use_sort": "",
                    "user_id": "",
                    "uuids[0]": uuid,
                    "query_type": secondary_intent,
                    "comp_id": comp_id,
                    "business_cat": "",
                    "uuid": uuid
                }
                resp = self.session.post(url, data=data, timeout=30)
                resp_data = resp.json()
                resp.close()
                result_temp = resp_data["answer"]["components"][0]["data"]["datas"]
                print(f"  页 {page}: 获取 {len(result_temp)} 条")
                result.extend(result_temp)
                if len(result_temp) < 100:
                    break
                page += 1
        except Exception as e:
            print(f"获取分页数据失败: {e}")
        
        return result


def fetch_concept_stocks(concept_info, wencai, proxy_manager=None):
    """获取单个概念的成分股数据"""
    code = concept_info['code']
    name = concept_info['name']
    
    # 构建查询语句：XXX概念；上市时间早于2024年1月1日
    question = f"{name}概念；上市时间早于2024年1月1日"
    
    try:
        stocks_data = wencai.get_answer(question, proxy_manager=proxy_manager)
        
        stocks = []
        for stock in stocks_data:
            try:
                # 尝试提取代码、名称和流通市值
                stock_info = {
                    '代码': stock.get('代码', ''),
                    '名称': stock.get('名称', ''),
                    '流通市值': stock.get('流通市值', 0)
                }
                if stock_info['代码']:
                    stocks.append(stock_info)
            except Exception as e:
                continue
        
        return {
            'code': code,
            'name': name,
            'update_time': datetime.now().strftime('%Y-%m-%d'),
            'stocks': stocks
        }
    except Exception as e:
        print(f"获取 {name} 成分股失败: {e}")
        return {
            'code': code,
            'name': name,
            'update_time': datetime.now().strftime('%Y-%m-%d'),
            'stocks': []
        }


def save_pool_file(concept_data, pool_folder=POOL_FOLDER):
    """保存成分股数据到pool文件夹"""
    os.makedirs(pool_folder, exist_ok=True)
    filename = os.path.join(pool_folder, f"BK{concept_data['code']}.json")
    
    with open(filename, 'w', encoding='utf-8') as f:
        json.dump(concept_data, f, ensure_ascii=False, indent=2)
    
    return filename


# ============== Step C: 计算概念流通市值 ==============
def calculate_concept_circulating_market_cap(pool_folder=POOL_FOLDER):
    """用成分股的流通市值计算每个概念的数据提取日的流通市值"""
    print("=" * 50)
    print("Step C: 计算每个概念的流通市值...")
    
    concept_caps = []
    pool_files = [f for f in os.listdir(pool_folder) if f.startswith('BK') and f.endswith('.json')]
    
    for filename in pool_files:
        filepath = os.path.join(pool_folder, filename)
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            code = data.get('code', '')
            name = data.get('name', '')
            stocks = data.get('stocks', [])
            update_time = data.get('update_time', '')
            
            # 计算总流通市值
            total_cap = sum(stock.get('流通市值', 0) for stock in stocks)
            
            concept_caps.append({
                'code': code,
                'name': name,
                'update_time': update_time,
                'stock_count': len(stocks),
                'circulating_market_cap': total_cap
            })
        except Exception as e:
            print(f"处理 {filename} 失败: {e}")
    
    return pd.DataFrame(concept_caps)


# ============== Step D: 获取概念历史数据并计算每日流通市值 ==============
def get_baostock_concept_history(code, name, start_date, end_date):
    """使用baostock获取概念的历史K线数据"""
    try:
        # 格式转换：BKxxx -> bk_xxx
        bs_code = f"bk_{code}"
        
        rs = bs.query_history_k_data_plus(
            bs_code,
            "date,code,open,high,low,close,volume,amount",
            start_date=start_date,
            end_date=end_date,
            frequency="d",
            adjusttype="1"
        )
        
        data_list = []
        while rs.error_code == '0' and rs.next():
            row = rs.get_row_data()
            data_list.append({
                '日期': row[0],
                '板块代码': row[1],
                '板块名称': name,
                '开盘价': float(row[2]) if row[2] else 0,
                '最高价': float(row[3]) if row[3] else 0,
                '最低价': float(row[4]) if row[4] else 0,
                '收盘价': float(row[5]) if row[5] else 0,
                '成交量': float(row[6]) if row[6] else 0,
                '成交额': float(row[7]) if row[7] else 0
            })
        
        return pd.DataFrame(data_list)
    except Exception as e:
        print(f"获取 {name} 历史数据失败: {e}")
        return pd.DataFrame()


def calculate_daily_concept_market_cap(concept_df, concept_circulating_cap_df, start_date, end_date):
    """计算每个概念的历史每日流通市值"""
    print("=" * 50)
    print("Step D: 计算每个概念每日流通市值...")
    
    # 初始化baostock
    bs.login()
    
    all_data = []
    data_lock = threading.Lock()
    
    def fetch_and_process(concept_info):
        code = concept_info['code']
        name = concept_info['name']
        
        # 获取该概念的历史数据
        history_df = get_baostock_concept_history(code, name, start_date, end_date)
        
        if not history_df.empty:
            # 获取概念的基础流通市值
            cap_row = concept_circulating_cap_df[concept_circulating_cap_df['code'] == code]
            if not cap_row.empty:
                base_cap = cap_row.iloc[0]['circulating_market_cap']
            else:
                base_cap = 0
            
            # 计算每日流通市值 = 基础流通市值 * (收盘价 / 首日收盘价)
            first_close = history_df.iloc[0]['收盘价'] if len(history_df) > 0 and history_df.iloc[0]['收盘价'] > 0 else 1
            history_df['概念流通市值'] = history_df['收盘价'].apply(
                lambda x: base_cap * (x / first_close) if first_close > 0 else 0
            )
            
            with data_lock:
                all_data.append(history_df)
            
            return name, True
        return name, False
    
    # 使用线程池并行获取
    with ThreadPoolExecutor(max_workers=10) as executor:
        futures = [executor.submit(fetch_and_process, row.to_dict()) 
                   for _, row in concept_circulating_cap_df.iterrows()]
        
        success_count = 0
        for future in as_completed(futures):
            name, success = future.result()
            if success:
                success_count += 1
                print(f"  完成: {name}")
    
    bs.logout()
    
    if all_data:
        return pd.concat(all_data, ignore_index=True)
    return pd.DataFrame()


# ============== Step E: 计算资金流向和资金阻力 ==============
def calculate_capital_flow_and_resistance(sector_df, period, count, weight=(0.5, 0.5)):
    """
    计算n日资金流向和资金阻力
    资金阻力公式（来自市场资金流向分析_east.py）:
    资金阻力 = 成交额变化/流通值/(涨跌幅+30) * 10e7 ± 涨跌幅/100
    """
    if sector_df is None or sector_df.empty:
        return None, None
    
    dates = sector_df['日期'].unique()
    if len(dates) < period:
        print(f"数据不足{period}个交易日，无法计算")
        return None, None
    
    # 确保按板块和日期排序
    sector_df = sector_df.sort_values(['板块名称', '日期'])
    
    # 计算涨跌幅
    sector_df['涨跌幅'] = sector_df.groupby('板块名称')['收盘价'].pct_change() * 100
    sector_df['涨跌幅'] = sector_df['涨跌幅'].fillna(0)
    
    # 计算成交额变化量
    sector_df['成交额变化量'] = sector_df.groupby('板块名称')['成交额'].diff()
    sector_df['成交额变化量'] = sector_df['成交额变化量'].fillna(0)
    
    result_list = []
    
    # 滑动窗口计算
    for i in range(len(dates) - (period - 1)):
        window_dates = dates[i:i + period]
        window_df = sector_df[sector_df['日期'].isin(window_dates)]
        
        # 计算n日累计涨跌幅
        window_ret = window_df.groupby('板块名称')['涨跌幅'].sum().reset_index()
        window_ret.columns = ['板块名称', f'{period}日涨跌幅']
        
        # 计算n日成交额变化总量
        window_vol = window_df.groupby('板块名称')['成交额变化量'].sum().reset_index()
        window_vol.columns = ['板块名称', f'{period}日成交额变化']
        
        # 合并结果
        window_result = pd.merge(window_ret, window_vol, on='板块名称')
        
        # 获取概念流通市值（使用窗口最后一日）
        window_market_cap = window_df.groupby('板块名称')['概念流通市值'].last().reset_index()
        window_result = pd.merge(window_result, window_market_cap, on='板块名称')
        
        # 计算资金阻力（来自市场资金流向分析_east.py的核心公式）
        def calc_resistance(row):
            change = row[f'{period}日成交额变化']
            cap = row['概念流通市值']
            ret = row[f'{period}日涨跌幅']
            
            if change >= 0:
                resistance = int((change / cap / (ret + 30)) * 10e7) + round(ret / 100, 2)
            else:
                resistance = int((change / cap / (ret + 30)) * 10e7) - round(ret / 100, 2)
            return resistance
        
        window_result['资金阻力'] = window_result.apply(calc_resistance, axis=1)
        
        # 计算排名
        window_result[f'{period}日涨跌幅排名'] = window_result[f'{period}日涨跌幅'].rank(
            ascending=False, method='min')
        window_result[f'{period}日资金流向排名'] = window_result[f'{period}日成交额变化'].rank(
            ascending=False, method='min')
        
        # 计算强度
        window_result['强度'] = round(
            ((count - window_result[f'{period}日涨跌幅排名']) * weight[0] +
             (count - window_result[f'{period}日资金流向排名']) * weight[1]) * 100 / count, 2)
        
        # 添加窗口结束日期
        window_result['窗口结束日期'] = window_dates[-1]
        
        result_list.append(window_result)
    
    if not result_list:
        return None, None
    
    result_df = pd.concat(result_list, ignore_index=True)
    
    # 转换为宽格式
    wide_df = result_df.pivot_table(
        index='板块名称',
        columns='窗口结束日期',
        values=['强度'],
        aggfunc='first'
    )
    wide_df.columns = [f"{col[1]}" for col in wide_df.columns]
    wide_df = wide_df.reset_index()
    
    wide_resistance = result_df.pivot_table(
        index='板块名称',
        columns='窗口结束日期',
        values=['资金阻力'],
        aggfunc='first'
    )
    wide_resistance.columns = [f"{col[1]}" for col in wide_resistance.columns]
    wide_resistance = wide_resistance.reset_index()
    
    return wide_df, wide_resistance


def save_results(result_nd, result_3d, resistance_nd, resistance_3d, output_file):
    """保存结果到Excel"""
    os.makedirs(os.path.dirname(output_file) if os.path.dirname(output_file) else '.', exist_ok=True)
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # 保存n日板块强度
        result_nd.to_excel(writer, index=False, sheet_name=f'{PERIOD}日板块强度')
        
        # 保存3日板块强度
        result_3d.to_excel(writer, index=False, sheet_name='3日板块强度')
        
        # 保存n日资金阻力
        resistance_nd.to_excel(writer, index=False, sheet_name=f'{PERIOD}日资金阻力')
        
        # 保存3日资金阻力
        resistance_3d.to_excel(writer, index=False, sheet_name='3日资金阻力')
    
    print(f"结果已保存到: {output_file}")


# ============== 主流程 ==============
def main():
    print("=" * 60)
    print("历史概念流通市值与资金阻力计算")
    print("=" * 60)
    
    # Step A: 获取概念列表
    concept_list = get_ths_concept_list()
    if not concept_list:
        print("获取概念列表失败，程序退出")
        return
    
    print(f"\n共获取 {len(concept_list)} 个概念")
    
    # Step B: 获取成分股数据并保存到pool文件夹
    print("\n" + "=" * 50)
    print("Step B: 获取成分股数据...")
    
    os.makedirs(POOL_FOLDER, exist_ok=True)
    
    # 初始化问财
    wencai = WenCai()
    
    # 统计已存在的文件
    existing_codes = set()
    if os.path.exists(POOL_FOLDER):
        for f in os.listdir(POOL_FOLDER):
            if f.startswith('BK') and f.endswith('.json'):
                code = f.replace('BK', '').replace('.json', '')
                existing_codes.add(code)
    
    print(f"已有 {len(existing_codes)} 个概念的数据")
    
    # 获取所有概念的成分股数据
    for i, concept in enumerate(concept_list):
        code = concept['code']
        name = concept['name']
        
        if code in existing_codes:
            print(f"[{i+1}/{len(concept_list)}] 跳过 {name} (已存在)")
            continue
        
        print(f"[{i+1}/{len(concept_list)}] 获取 {name} 成分股...")
        
        concept_data = fetch_concept_stocks(concept, wencai)
        save_pool_file(concept_data)
        
        print(f"  -> 保存了 {len(concept_data['stocks'])} 只股票")
        
        # 避免请求过快
        time.sleep(1)
    
    # Step C: 计算概念流通市值
    concept_cap_df = calculate_concept_circulating_market_cap()
    print(f"\n共计算了 {len(concept_cap_df)} 个概念的流通市值")
    
    # Step D: 获取历史数据并计算每日流通市值
    end_date = datetime.now().strftime('%Y-%m-%d')
    start_date = (datetime.now() - timedelta(days=100)).strftime('%Y-%m-%d')
    
    sector_df = calculate_daily_concept_market_cap(
        concept_df=None,
        concept_circulating_cap_df=concept_cap_df,
        start_date=start_date,
        end_date=end_date
    )
    
    if sector_df.empty:
        print("获取历史数据失败")
        return
    
    # Step E: 计算资金流向和资金阻力
    print("\n" + "=" * 50)
    print("Step E: 计算资金流向和资金阻力...")
    
    count = len(concept_cap_df)
    
    result_20d, resistance_20d = calculate_capital_flow_and_resistance(
        sector_df, PERIOD, count, weight=(0.5, 0.5))
    result_3d, resistance_3d = calculate_capital_flow_and_resistance(
        sector_df, 3, count, weight=(0.5, 0.5))
    
    if result_20d is not None and result_3d is not None:
        output_file = os.path.join(
            OUTPUT_FOLDER, 
            f"概念资金阻力_{start_date}_{end_date}.xlsx"
        )
        save_results(result_20d, result_3d, resistance_20d, resistance_3d, output_file)
    else:
        print("计算资金阻力失败")


# ============== 配置参数 ==============
PERIOD = 20  # n日窗口，默认20日


if __name__ == '__main__':
    main()

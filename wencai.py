import json
import os
import re
import execjs
import vcr
from lxml import etree
import requests


class WenCai:
    def __init__(self,proxy=None):
        self.cookies = 'cid=f339b3b06f20f9a3f57511efc61f69561732676413; cid=f339b3b06f20f9a3f57511efc61f69561732676413; ComputerID=f339b3b06f20f9a3f57511efc61f69561732676413; WafStatus=0; other_uid=Ths_iwencai_Xuangu_dhfyapc0j8s63crisy4bluu2kqf1zxxu; ta_random_userid=ofrg6m3r2p; u_ukey=A10702B8689642C6BE607730E11E6E4A; u_uver=1.0.0; u_dpass=8vbH2AZwsYKOAQB90LnApBF6tFP6foVqAiLa0I6LV%2BGq52j91%2FiAPfxZW3mxg421Hi80LrSsTFH9a%2B6rtRvqGg%3D%3D; u_did=A962F087FEBC4DCBB0DFB4AFD8A929A4; u_ttype=WEB; ttype=WEB; user=MDptb180NjE5NzM3MTA6Ok5vbmU6NTAwOjQ3MTk3MzcxMDo3LDExMTExMTExMTExLDQwOzQ0LDExLDQwOzYsMSw0MDs1LDEsNDA7MSwxMDEsNDA7MiwxLDQwOzMsMSw0MDs1LDEsNDA7OCwwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMSw0MDsxMDIsMSw0MDoyNDo6OjQ2MTk3MzcxMDoxNzYzMjA5NTgxOjo6MTUzNjI0MTgwMDo2MDQ4MDA6MDoxMGI4ZDIzNTBiZDg3YzVmOTIxZjU5OTQyNjcyZWNjNzI6ZGVmYXVsdF81OjE%3D; userid=461973710; u_name=mo_461973710; escapename=mo_461973710; ticket=5ac58fb6a7ab2594ed9adaf1ea804370; user_status=0; utk=3f126086ef468b4c1000cbe2f463587f; sess_tk=eyJ0eXAiOiJKV1QiLCJhbGciOiJFUzI1NiIsImtpZCI6InNlc3NfdGtfMSIsImJ0eSI6InNlc3NfdGsifQ.eyJqdGkiOiI3MmNjMmU2NzQyOTlmNTIxZjljNTg3YmQ1MDIzOGQwYjEiLCJpYXQiOjE3NjMyMDk1ODEsImV4cCI6MTc2MzgxNDM4MSwic3ViIjoiNDYxOTczNzEwIiwiaXNzIjoidXBhc3MuaXdlbmNhaS5jb20iLCJhdWQiOiIyMDIwMTExODUyODg5MDcyIiwiYWN0Ijoib2ZjIiwiY3VocyI6IjA0NGFjNWU5YTQwZTRhNDMxOGJlMGU1Njg5NmZhYWEyOTVhYzhmNDE3MDdlOTNmODk2ZmVhZDFkOGJmNzRhNDAifQ._ZW4ZCilVpVdpW6p7hE1t0cB09F5GGoFEL8l0w6EIsI5Yj-Hq-vAfGauJZ0XXPXnb5OSeBY3u-pNuJgDHZ6m5A; cuc=4mt9un9omypt; v=A6K-n632P-XnTyM5bSrpuKa98yMB86Q-WPaaMew4zOXBa0yd1IP2HSiH6mO_'
        self.session = requests.session()
        lst_cookies = self.cookies.split(';')
        for cookie in lst_cookies:
            self.session.cookies.set(cookie.split('=')[0], cookie.split('=')[1])
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
        if proxy:
            self.hexin_v = self.get_hexin_v(self.get_server_time(proxy))
        else:
            self.hexin_v = self.get_hexin_v(self.get_server_time())

    def get_server_time(self,proxy=None):
        url = 'http://www.iwencai.com/unifiedwap/home/index'
        if proxy:
            resp = self.session.get(url,proxies=proxy)
        else:
            resp = self.session.get(url)
        resp_text = resp.text
        tree = etree.HTML(resp_text)
        js_url = "http:" + tree.xpath("//script[1]/@src")[0]
        resp.close()
        js_resp = self.session.get(js_url)
        js_text = js_resp.text
        obj = re.compile(r'var TOKEN_SERVER_TIME=(?P<time>.*?);var n=function')
        server_time = obj.search(js_text).group('time')
        return server_time

    def get_answer(self,question, secondary_intent,proxy_manager=None):
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
        # filename = f'D:/cassette/{question}.yml'
        # print(filename)
        result = []
        try:
            # with vcr.use_cassette(filename):
            self.session.headers['hexin-v'] = self.hexin_v
            self.session.headers['Content-Type'] = 'application/json'
            if proxy_manager:
                proxy = proxy_manager.get_current_proxy()
                resp = self.session.post(url, data=json.dumps(data),proxies=proxy)
            else:
                resp = self.session.post(url, data=json.dumps(data))
            resp_data = resp.json()
            resp.close()
            if 'data' in resp_data.keys():
                result = resp_data['data']['answer'][0]['txt'][0]['content']['components'][0]['data']['datas']
                pass
                if len(result) == 100:
                    result.extend(self.get_page_data(resp_data, question, secondary_intent))
        except Exception as e:
            print(e)
        return result

    def get_page_data(self,f_resp_data, question, secondary_intent):
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
            resp = self.session.post(url, data=data)
            resp_data = resp.json()
            resp.close()
            result_temp = resp_data["answer"]["components"][0]["data"]["datas"]
            print(page)
            result.extend(result_temp)
            if len(result_temp) < 100:
                break
            page += 1
        return result
    def get_hexin_v(self,time):
        with open("kou.js", "r", encoding='utf-8') as f:
            js_content = f.read()
            js_content = 'var TOKEN_SERVER_TIME=' + str(time) + ";\n" + js_content
            js = execjs.compile(js_content)
            v = js.call("rt.updata")
            return v

if __name__ == '__main__':
    wc = WenCai()

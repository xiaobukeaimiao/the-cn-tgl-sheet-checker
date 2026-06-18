import re
import os
import json
import httpx
import random
import asyncio
import datetime
from collections import Counter
from bilibili_api.utils.network import HEADERS 
from bilibili_api import video, user, ResponseCodeException, Credential, favorite_list

# === 基础配置 ===
HEADERS["User-Agent"] = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
HEADERS["Referer"] = "https://www.bilibili.com/"

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_FILE = os.path.join(CURRENT_DIR, 'all_hyperlinks.txt')
OUTPUT_FILE = os.path.join(CURRENT_DIR, '已检查的链接.txt')
CONFIG_FILE = os.path.join(CURRENT_DIR, 'account_data.json')

with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
    config = json.load(f)
SESSDATA = config.get("SESSDATA", "")
BILI_JCT = config.get("BILI_JCT", "")

ENABLE_AUTO_FAVORITE = False

# 策略配置
MIN_SLEEP = 3
MAX_SLEEP = 7
MAX_CYCLES = 10  # 最大循环轮数（防止因为某个死链接一直报错导致死循环）

credential = None
if ENABLE_AUTO_FAVORITE:
    credential = credential = Credential(sessdata=SESSDATA, bili_jct=BILI_JCT)
NAME_PATTERN = re.compile(r"原B站名'([^']*)'")




# ==========================================
#              核心功能函数
# ==========================================

    


def append_result(result_line):
    """写入结果到文件"""
    with open(OUTPUT_FILE, 'a', encoding='utf-8') as f:
        f.write(result_line + "\n")

def get_remaining_tasks():
    """
    对比输入文件和结果文件，返回还未完成的任务列表
    """
    # 1. 读取已完成的历史记录（存入集合，方便快速查找）
    finished_urls = set()
    if os.path.exists(OUTPUT_FILE):
        with open(OUTPUT_FILE, 'r', encoding='utf-8') as f:
            for line in f:
                # 提取URL：假设格式为 [状态] URL | ...
                # 简单粗暴的方法：只要行里包含这个URL字符串，就算做过
                # 为了更精准，我们分割一下
                parts = line.split('|')
                if len(parts) > 0:
                    # 去掉开头的 [状态] 标记，尝试提取URL
                    # 这里简化处理：将整行内容作为判断依据太宽泛
                    # 我们假设 URL 是每行的第二个“单词”（以空格分隔）
                    # "[✅ 有效] https://..."
                    content = line.strip()
                    finished_urls.add(content) # 这里存整行不合适，需要提取特征
                    
    # 为了更精准的对比，我们重新读取一遍Output文件，提取其中的URL部分
    checked_url_set = set()
    if os.path.exists(OUTPUT_FILE):
        with open(OUTPUT_FILE, 'r', encoding='utf-8') as f:
            content = f.read()
            # 这里采用“如果在文件内容里能找到这个URL字符串”就算已完成
            # 这种方法虽然粗糙，但对于防止重复检查是有效的
            checked_url_set_content = content 
    else:
        checked_url_set_content = ""

    tasks = []
    current_section = None
    
    if not os.path.exists(INPUT_FILE):
        return []

    with open(INPUT_FILE, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if line == "=== 玩家链接 ===":
                current_section = 'player'
                continue
            elif line == "=== 单元格链接 ===":
                current_section = 'cell'
                continue
            
            if not line or current_section is None: continue
            
            parts = line.split(' ', 1)
            url = parts[0]
            extra = parts[1] if len(parts) > 1 else ""
            
            # === 核心过滤逻辑 ===
            # 如果 URL 已经在结果文件的内容里了，就跳过
            if url in checked_url_set_content:
                continue
            
            tasks.append({
                "url": url,
                "section": current_section,
                "extra": extra,
                "raw_line": line
            })
    
    return tasks

# === 新增：正则匹配定义 ===
# 匹配 BV 号 (支持 BV/bv, 忽略大小写, 10位字符)
BV_PATTERN = re.compile(r'(BV[a-zA-Z0-9]{10})', re.IGNORECASE)
# 匹配 AV 号 (支持 av/AV + 数字)
AV_PATTERN = re.compile(r'av(\d+)', re.IGNORECASE)
# 匹配 UID (space.bilibili.com/数字)
UID_PATTERN = re.compile(r'space\.bilibili\.com/(\d+)')

# === 新增：解析短链的辅助函数 ===
async def get_real_url(original_url):
    """
    模拟浏览器访问 b23.tv 链接，获取重定向后的真实地址。
    """
    try:
        # 使用 HEAD 请求节省流量，follow_redirects=True 会自动跟踪跳转
        async with httpx.AsyncClient(headers=HEADERS, follow_redirects=True, timeout=10.0) as client:
            resp = await client.get(original_url)
            
            # 如果短链本身失效（404），直接报错
            if resp.status_code == 404:
                return None, "HTTP 404 (短链失效)"
            
            # 返回跳转后的最终 URL
            return str(resp.url), None
    except httpx.ConnectTimeout:
        return None, "连接超时"
    except httpx.ConnectError:
        return None, "无法连接服务器"
    except Exception as e:
        return None, str(e)

# === 修改：核心检查函数 ===
async def check_item(url, section, extra_info):
    """
    逻辑：
    1. 针对 b23.tv 短链：必须发起网络请求验证跳转，确保原链接未写错。
    2. 针对普通长链：直接正则提取 ID，不请求网页，直接调 API 查存活。
    """
    try:
        target_url = url # 默认目标就是原链接

        # 1. 特殊处理 b23.tv 短链
        if "b23.tv" in url:
            resolved_url, error_msg = await get_real_url(url)
            if not resolved_url:
                return "❌ 链接错误", f"短链无法访问: {error_msg}"
            target_url = resolved_url

        # 2. 玩家检查 (Space)
        if section == 'player':
            # 宽松检查：只要最终地址包含 space.bilibili.com 即可
            if "space.bilibili.com" not in target_url:
                return "⛔ 格式错误", "未指向space域名"
            
            match = UID_PATTERN.search(target_url)
            if not match: return "⛔ 格式错误", "无法提取UID"
            
            u = user.User(uid=int(match.group(1)), credential=credential)
            info = await u.get_user_info()
            current_name = info.get('name', '未知')
            
            name_match = NAME_PATTERN.search(extra_info)
            if name_match:
                original = name_match.group(1)
                if current_name != original:
                    return "⚠️ 不一致", f"原名 '{original}' -> 现名 '{current_name}'"
            return "✅ 有效", f"现名: {current_name}"

        # 3. 视频检查 (Video)
        elif section == 'cell':
            # 在 target_url (可能是跳转后的) 中找 ID
            bv_match = BV_PATTERN.search(target_url)
            av_match = AV_PATTERN.search(target_url)
            
            v = None
            if bv_match:
                # 兼容大小写，强转标准格式 BV...
                bvid = "BV" + bv_match.group(1)[2:]
                v = video.Video(bvid=bvid, credential=credential)
            elif av_match:
                aid = int(av_match.group(1))
                v = video.Video(aid=aid, credential=credential)
            else:
                return "⛔ 格式错误", "链接中未包含有效BV/av号"

            # 调用 API 检查视频内容
            info = await v.get_info()

            archive_msg = ""
            if ENABLE_AUTO_FAVORITE:
                # 只有当视频有效时才执行归档
                archive_msg = "    "+await process_archive_action(v, info)
                # 稍微增加一点延迟，防止连续写操作触发验证码
                if "已收藏" in archive_msg or "已点赞" in archive_msg:
                    await asyncio.sleep(random.uniform(1, 3))
            
            return "✅ 有效", f"标题: {info.get('title', '')}{archive_msg}"

    except ResponseCodeException as e:
        # 404: 找不到; 62002: 仅自己可见/审核中; 62012: 视频不见了
        if e.code in [-404, 404, 62002, 62012]: 
            return "❌ 失效", f"视频已删除/不可见 (Code: {e.code})"
        if e.code == -412: 
            raise RuntimeError("IP_BANNED") 
        return "⚠️ API报错", f"Code: {e.code}"
    except Exception as e:
        return "⚠️ 异常", str(e)
    
    return "❓ 未知", "未覆盖逻辑"

def generate_final_report():
    """
    全部完成后调用的输出函数
    统计并在控制台打印最终报告
    """
    print("\n" + "="*50)
    print("📋 最终检查报告")
    print("="*50)
    
    if not os.path.exists(OUTPUT_FILE):
        print("未找到结果文件。")
        return

    stats = Counter()
    issues = []

    with open(OUTPUT_FILE, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if not line: continue
            
            # 统计状态，例如提取 [✅ 有效] 中的 "✅ 有效"
            # 假设格式：[状态] URL | ...
            if line.startswith("["):
                end_index = line.find("]")
                if end_index != -1:
                    status = line[1:end_index]
                    stats[status] += 1
                    
                    # 收集非完美结果（除去有效和格式错误外的其他问题）
                    if "✅" not in status:
                        issues.append(line)

    # 打印统计数据
    print(f"总检查数量: {sum(stats.values())}")
    for status, count in stats.items():
        print(f"{status}: {count}")
    
    print("-" * 50)
    if issues:
        print(f"⚠️ 发现 {len(issues)} 个问题：")
        for i, issue in enumerate(issues):
            print(f'[{i+1}] {issue}')
    else:
        print("🎉 完美！")
    print("="*50)
    print(f"结果文件路径: {OUTPUT_FILE}")

# ==========================================
#              主循环逻辑
# ==========================================

async def main_loop():
    cycle = 0
    
    while cycle < MAX_CYCLES:
        cycle += 1
        print(f"\n🔄 --- 开始第 {cycle} 轮扫描 ---")
        
        # 1. 获取剩余任务
        tasks = get_remaining_tasks()
        
        if not tasks:
            print("✅ 所有任务已全部完成！")
            break
        
        print(f"本轮剩余任务数: {len(tasks)}")
        
        consecutive_errors = 0
        
        # 2. 遍历执行
        for i, item in enumerate(tasks):
            print(f"[{cycle}轮][{i+1}/{len(tasks)}] 检查: {item['url']}")
            
            try:
                status, msg = await check_item(item['url'], item['section'], item['extra'])
                output_line = f"[{status}] {item['url']} | {msg} | {item['extra']}"
                
                # 只有明确结果才保存
                if "✅" in status or "❌" in status or "⛔" in status or "⚠️ 不一致" in status:
                    append_result(output_line)
                    consecutive_errors = 0
                    print(f"   -> 已保存: {status} {msg}")
                else:
                    # 临时错误，跳过保存，留给下一轮
                    print(f"   ⚠️ 跳过保存(进入下一轮): {status} {msg}")
                    consecutive_errors += 1
                
                await asyncio.sleep(random.uniform(MIN_SLEEP, MAX_SLEEP))

            except RuntimeError as e:
                if str(e) == "IP_BANNED":
                    print("\n🔴 检测到IP被封锁，停止脚本。")
                    return # 直接结束程序
            except Exception as e:
                print(f"   ⚠️ 程序异常: {e}")
                consecutive_errors += 1
            
            # 熔断
            if consecutive_errors >= 10:
                print("⚠️ 本轮连续错误过多，提前结束本轮，休息后重试...")
                break
        
        # 3. 轮次间休息
        if cycle < MAX_CYCLES:
            print(f"\n⏳ 第 {cycle} 轮结束。等待 5 秒后检查是否有遗漏...")
            await asyncio.sleep(5)
    
    if cycle >= MAX_CYCLES:
        print(f"\n⚠️ 已达到最大循环轮数 ({MAX_CYCLES})，停止重试。可能部分链接持续报错。")

    # 4. 全部结束后，进入输出函数
    generate_final_report()


def clean_failed_results():
    """
    清理功能：
    读取结果文件，删除包含 ⚠️(异常/API错误)、⛔(格式错误)、❌(失效) 的行。
    保留 ✅(有效) 的行。
    这样下次运行程序时，被删除的链接会被视为“未完成任务”重新检查。
    """
    if not os.path.exists(OUTPUT_FILE):
        return

    print(f"🧹 正在清理 {OUTPUT_FILE} 中的失败记录...")

    with open(OUTPUT_FILE, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    original_count = len(lines)
    new_lines = []

    targets = ["⚠️", "⛔", "❌"]

    for line in lines:
        # 如果行内包含任意一个目标符号，就跳过（删除）
        if any(symbol in line for symbol in targets):
            continue
        new_lines.append(line)

    removed_count = original_count - len(new_lines)

    # 重新写入文件
    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        f.writelines(new_lines)

    print(f"✅ 清理完成！")



# ==========================================
#             自动归档与点赞功能
# ==========================================

# 全局缓存：存储 { "2019年5-8月": 12345678 } 这样的映射，避免每次都请求API查收藏夹
FOLDER_CACHE = {} 
IS_CACHE_INITED = False

def get_folder_name_by_date(timestamp):
    
    #根据时间戳生成收藏夹名称
    #策略：每4个月一个周期 (1~4, 5~8, 9~12)
    
    dt = datetime.datetime.fromtimestamp(timestamp)
    year = dt.year
    month = dt.month
    
    if 1 <= month <= 4:
        suffix = "1~4"
    elif 5 <= month <= 8:
        suffix = "5~8"
    else:
        suffix = "9~12"
        
    return f"Top Golden {year}.{suffix}"

async def fetch_created_fav_lists(uid):
    url = "https://api.bilibili.com/x/v3/fav/folder/created/list-all"
    params = {"up_mid": uid}
    cookies = {"SESSDATA": SESSDATA}
    async with httpx.AsyncClient(headers=HEADERS, timeout=10.0) as client:
        resp = await client.get(url, params=params, cookies=cookies)
        data = resp.json()
    if data["code"] == 0 and "list" in data["data"]:
        return data["data"]["list"]
    return []

# --- 2. 直连API: 创建收藏夹 ---
async def create_fav_folder_api(title):
    url = "https://api.bilibili.com/x/v3/fav/folder/add"
    data = {
        "title": title,
        "intro": "Auto-generated",
        "privacy": 0,
        "csrf": BILI_JCT
    }
    cookies = {"SESSDATA": SESSDATA}
    async with httpx.AsyncClient(headers=HEADERS, timeout=10.0) as client:
        resp = await client.post(url, data=data, cookies=cookies)
        res = resp.json()
    if res["code"] == 0: return res["data"]["id"]
    return None

# --- 3. 直连API: 添加视频到收藏夹 ---
async def add_video_to_fav_api(aid, folder_id):
    
    #rid: 视频的 av 号 (不是 BV 号)
    #type: 2 (代表视频)
    #add_media_ids: 目标收藏夹 ID

    url = "https://api.bilibili.com/x/v3/fav/resource/deal"
    data = {
        "rid": aid,
        "type": 2,
        "add_media_ids": folder_id,
        "csrf": BILI_JCT
    }
    cookies = {"SESSDATA": SESSDATA}
    async with httpx.AsyncClient(headers=HEADERS, timeout=10.0) as client:
        resp = await client.post(url, data=data, cookies=cookies)
        return resp.json()

# --- 4. 直连API: 点赞视频 ---
async def like_video_api(aid):
    url = "https://api.bilibili.com/x/web-interface/archive/like"
    data = {
        "aid": aid,
        "like": 1, # 1=点赞, 2=取消
        "csrf": BILI_JCT
    }
    cookies = {"SESSDATA": SESSDATA}
    async with httpx.AsyncClient(headers=HEADERS, timeout=10.0) as client:
        resp = await client.post(url, data=data, cookies=cookies)
        return resp.json()

# --- 缓存初始化逻辑 ---
async def init_folder_cache(uid):
    global FOLDER_CACHE, IS_CACHE_INITED
    if IS_CACHE_INITED: return
    try:
        print("📂 正在拉取现有收藏夹列表(API)...")
        folders = await fetch_created_fav_lists(uid)
        for folder in folders:
            FOLDER_CACHE[folder['title']] = folder['id']
        IS_CACHE_INITED = True
        print(f"📂 收藏夹列表缓存完毕，共 {len(FOLDER_CACHE)} 个")
    except Exception as e:
        print(f"⚠️ 初始化缓存失败: {e}")

async def get_or_create_folder(folder_name):
    if folder_name in FOLDER_CACHE: return FOLDER_CACHE[folder_name]
    try:
        print(f"🆕 尝试创建新收藏夹: [{folder_name}]")
        new_id = await create_fav_folder_api(folder_name)
        if new_id:
            FOLDER_CACHE[folder_name] = new_id
            return new_id
    except Exception as e:
        print(f"❌ 创建失败: {e}")
    return None

# --- 主逻辑函数 (更新版) ---
async def process_archive_action(video_obj, video_info):
    
    #执行点赞和收藏的核心逻辑 (完全脱离 Video 对象方法)
    
    if not ENABLE_AUTO_FAVORITE or not credential: return "未开启归档"

    try:
        # 1. 获取基础信息
        pubdate = video_info.get('pubdate')
        aid = video_info.get('aid') # 必须获取 av号
        if not pubdate or not aid: return "无法获取AID/时间"
        
        target_folder_name = get_folder_name_by_date(pubdate)
        
        # 2. 确保缓存初始化
        if not IS_CACHE_INITED:
            # 获取自己UID
            self_info = await user.get_self_info(credential)
            await init_folder_cache(self_info['mid'])

        # 3. 获取/创建文件夹
        folder_id = await get_or_create_folder(target_folder_name)
        if not folder_id: return "收藏夹获取失败"

        # 4. 执行收藏 (使用直连API)
        fav_result = ""
        try:
            res = await add_video_to_fav_api(aid, folder_id)
            code = res.get('code')
            if code == 0:
                fav_result = f"已收藏[{target_folder_name}]"
            elif code == 11201: # 错误码：已经收藏过了
                fav_result = "已跳过(重复收藏)"
            else:
                fav_result = f"收藏错({code})"
        except Exception as e:
            fav_result = "收藏异常"

        # ⚠️ 模拟人类操作间隔 (防止风控)
        await asyncio.sleep(random.uniform(1.5, 3.0))

        # 5. 执行点赞 (使用直连API)
        like_result = ""
        try:
            res = await like_video_api(aid)
            code = res.get('code')
            if code == 0:
                like_result = "已点赞"
            elif code == 65006: # 错误码：已赞过
                like_result = "已赞过"
            else:
                like_result = f"点赞错({code})"
        except Exception as e:
            like_result = "点赞异常"
        
        return f"{fav_result} & {like_result}"

    except Exception as e:
        return f"归档异常: {str(e)}"





if __name__ == '__main__':
    # clean_failed_results()

    try:
        if not os.path.exists(INPUT_FILE):
            print(f"❌ 找不到输入文件: {INPUT_FILE}")
        else:
            asyncio.run(main_loop())
    except KeyboardInterrupt:
        print("\n🛑 用户强制停止")
        # 即使强制停止，也尝试生成当前进度的报告
        generate_final_report()
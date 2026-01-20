#!/usr/bin/env python3
"""
小红书评论抓取脚本 - 浏览器辅助模式
使用 Playwright 自动化浏览器抓取评论数据
"""

import argparse
import json
import time
import re
from datetime import datetime
from pathlib import Path

def extract_comments(url: str, output_path: str, max_scroll: int = 50, headless: bool = False):
    """
    从小红书帖子抓取评论
    
    Args:
        url: 小红书帖子链接
        output_path: 输出 JSON 文件路径
        max_scroll: 最大滚动次数
        headless: 是否无头模式
    """
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("错误: 请先安装 playwright: pip install playwright && playwright install chromium")
        return None
    
    comments = []
    
    with sync_playwright() as p:
        # 启动浏览器 - 使用用户数据目录保持登录状态
        user_data_dir = Path.home() / ".xhs-browser-data"
        user_data_dir.mkdir(exist_ok=True)
        
        context = p.chromium.launch_persistent_context(
            user_data_dir=str(user_data_dir),
            headless=headless,
            viewport={"width": 1280, "height": 800},
            locale="zh-CN"
        )
        
        page = context.pages[0] if context.pages else context.new_page()
        
        print(f"正在访问: {url}")
        page.goto(url, wait_until="networkidle", timeout=60000)
        
        # 等待页面加载
        time.sleep(3)
        
        # 检查是否需要登录
        if "login" in page.url.lower():
            print("\n⚠️ 需要登录！请在打开的浏览器中登录小红书账号...")
            print("登录完成后，脚本将自动继续。")
            page.wait_for_url(lambda u: "login" not in u.lower(), timeout=300000)
            page.goto(url, wait_until="networkidle")
            time.sleep(3)
        
        # 获取帖子标题
        title = ""
        try:
            title_elem = page.query_selector(".title, .note-title, h1")
            if title_elem:
                title = title_elem.inner_text().strip()
        except:
            pass
        
        print(f"帖子标题: {title}")
        print("开始抓取评论...")
        
        # 点击展开评论区（如果需要）
        try:
            expand_btn = page.query_selector('[class*="comment"] button, .show-more')
            if expand_btn:
                expand_btn.click()
                time.sleep(1)
        except:
            pass
        
        # 滚动加载更多评论
        last_count = 0
        no_new_count = 0
        
        for i in range(max_scroll):
            # 滚动到页面底部
            page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            time.sleep(1.5)
            
            # 尝试点击"查看更多评论"按钮
            try:
                more_btn = page.query_selector('[class*="more"], .load-more, [class*="展开"]')
                if more_btn and more_btn.is_visible():
                    more_btn.click()
                    time.sleep(1)
            except:
                pass
            
            # 获取当前评论数量
            comment_elems = page.query_selector_all('[class*="comment-item"], [class*="commentItem"], .comment')
            current_count = len(comment_elems)
            
            print(f"  滚动 {i+1}/{max_scroll}, 已发现 {current_count} 条评论")
            
            if current_count == last_count:
                no_new_count += 1
                if no_new_count >= 5:
                    print("  连续5次无新评论，停止滚动")
                    break
            else:
                no_new_count = 0
            
            last_count = current_count
        
        # 提取评论数据
        print("\n正在提取评论数据...")
        
        # 尝试多种选择器
        selectors = [
            '[class*="comment-item"]',
            '[class*="commentItem"]',
            '.comment-inner',
            '[class*="comment"] > div'
        ]
        
        comment_elems = []
        for selector in selectors:
            comment_elems = page.query_selector_all(selector)
            if comment_elems:
                break
        
        for idx, elem in enumerate(comment_elems):
            try:
                comment_data = extract_single_comment(elem, idx + 1)
                if comment_data and comment_data.get("content"):
                    comments.append(comment_data)
            except Exception as e:
                print(f"  提取评论 {idx+1} 失败: {e}")
        
        context.close()
    
    # 保存结果
    result = {
        "url": url,
        "title": title,
        "crawl_time": datetime.now().isoformat(),
        "total_comments": len(comments),
        "comments": comments
    }
    
    output_file = Path(output_path)
    output_file.parent.mkdir(parents=True, exist_ok=True)
    
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    
    print(f"\n✅ 抓取完成！共 {len(comments)} 条评论")
    print(f"📄 保存至: {output_file}")
    
    return result


def extract_single_comment(elem, index: int) -> dict:
    """提取单条评论数据"""
    
    # 评论内容
    content = ""
    content_selectors = ['[class*="content"]', '.text', 'p', 'span']
    for sel in content_selectors:
        try:
            content_elem = elem.query_selector(sel)
            if content_elem:
                content = content_elem.inner_text().strip()
                if content and len(content) > 2:
                    break
        except:
            pass
    
    # 用户昵称
    nickname = ""
    nick_selectors = ['[class*="nickname"]', '[class*="name"]', '.user', 'a']
    for sel in nick_selectors:
        try:
            nick_elem = elem.query_selector(sel)
            if nick_elem:
                nickname = nick_elem.inner_text().strip()
                if nickname and len(nickname) < 30:
                    break
        except:
            pass
    
    # 点赞数
    likes = 0
    like_selectors = ['[class*="like"]', '[class*="count"]', '.likes']
    for sel in like_selectors:
        try:
            like_elem = elem.query_selector(sel)
            if like_elem:
                like_text = like_elem.inner_text().strip()
                # 提取数字
                nums = re.findall(r'\d+', like_text)
                if nums:
                    likes = int(nums[0])
                    break
        except:
            pass
    
    # 发布时间
    pub_time = ""
    time_selectors = ['[class*="time"]', '[class*="date"]', 'time']
    for sel in time_selectors:
        try:
            time_elem = elem.query_selector(sel)
            if time_elem:
                pub_time = time_elem.inner_text().strip()
                if pub_time:
                    break
        except:
            pass
    
    # 是否作者回复
    is_author = False
    try:
        author_badge = elem.query_selector('[class*="author"], [class*="作者"]')
        is_author = author_badge is not None
    except:
        pass
    
    # 子评论
    sub_comments = []
    try:
        sub_elems = elem.query_selector_all('[class*="reply"], [class*="sub-comment"]')
        for sub in sub_elems[:5]:  # 最多取5条子评论
            sub_content = sub.inner_text().strip()
            if sub_content:
                sub_comments.append(sub_content[:200])
    except:
        pass
    
    return {
        "index": index,
        "nickname": nickname,
        "content": content[:500],  # 限制长度
        "likes": likes,
        "time": pub_time,
        "is_author_reply": is_author,
        "sub_comments": sub_comments
    }


def main():
    parser = argparse.ArgumentParser(description="小红书评论抓取工具")
    parser.add_argument("url", help="小红书帖子链接")
    parser.add_argument("--output", "-o", default="comments.json", help="输出文件路径")
    parser.add_argument("--max-scroll", type=int, default=50, help="最大滚动次数")
    parser.add_argument("--headless", action="store_true", help="无头模式（不显示浏览器）")
    
    args = parser.parse_args()
    extract_comments(args.url, args.output, args.max_scroll, args.headless)


if __name__ == "__main__":
    main()

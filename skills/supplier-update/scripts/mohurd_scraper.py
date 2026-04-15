#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
住建部安全生产许可证查询爬虫
使用 Playwright 实现自动化查询
"""

import sys
import time
import json
import logging
from datetime import datetime, timedelta
from typing import Dict, Optional, List

logger = logging.getLogger(__name__)

# 尝试导入 patchright
try:
    from patchright.sync_api import sync_playwright
except ImportError:
    logger.warning("patchright 未安装，正在安装...")
    import subprocess
    subprocess.run([sys.executable, "-m", "pip", "install", "patchright", "--break-system-packages", "-q"])
    from patchright.sync_api import sync_playwright


class MohurdScraper:
    """住建部安全生产许可证查询爬虫"""
    
    SEARCH_URL = "https://zlaq.mohurd.gov.cn/fwmh/bjxcjgl/fwmh/pages/construction_safety/qyaqscxkz/qyaqscxkz"
    
    def __init__(self, headless: bool = True):
        """
        初始化爬虫
        
        Args:
            headless: 是否使用无头模式（不显示浏览器窗口）
        """
        self.headless = headless
        self.browser = None
        self.playwright = None
        self.page = None
        
    def __enter__(self):
        """上下文管理器入口"""
        self.start()
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """上下文管理器出口"""
        self.close()
        
    def start(self):
        """启动浏览器"""
        if self.playwright is None:
            self.playwright = sync_playwright().start()
            self.browser = self.playwright.chromium.launch(headless=self.headless)
            self.page = self.browser.new_page(viewport={"width": 1920, "height": 1080})
        
    def close(self):
        """关闭浏览器"""
        if self.browser:
            self.browser.close()
            self.browser = None
        if self.playwright:
            self.playwright.stop()
            self.playwright = None
        self.page = None
    
    def _wait_for_page_load(self, timeout: int = 60):
        """等待页面加载"""
        try:
            self.page.goto(self.SEARCH_URL, wait_until="networkidle", timeout=timeout * 1000)
            time.sleep(2)
            return True
        except Exception as e:
            logger.error(f"页面加载失败: {e}")
            return False
    
    def _expand_search_panel(self):
        """展开搜索面板"""
        try:
            self.page.evaluate("""
                var els = document.querySelectorAll('a.btn.show');
                if (els.length) els[0].click();
            """)
            time.sleep(1)
            return True
        except Exception as e:
            logger.warning(f"展开搜索面板失败: {e}")
            return False
    
    def _fill_company_name(self, company_name: str) -> bool:
        """填写公司名称"""
        try:
            self.page.evaluate(f"""
                (function() {{
                    var input = document.querySelector('#qymc');
                    if (input) {{
                        input.value = '{company_name}';
                        input.dispatchEvent(new Event('input', {{ bubbles: true }}));
                        return true;
                    }}
                    return false;
                }})()
            """)
            time.sleep(0.5)
            return True
        except Exception as e:
            logger.warning(f"填写公司名称失败: {e}")
            return False
    
    def _click_search(self) -> bool:
        """点击查询按钮"""
        try:
            self.page.evaluate("""
                (function() {{
                    var els = document.querySelectorAll('a.btn.search');
                    if (els.length) {{
                        els[0].click();
                        return true;
                    }}
                    return false;
                }})()
            """)
            time.sleep(1)
            return True
        except Exception as e:
            logger.warning(f"点击查询按钮失败: {e}")
            return False
    
    def _wait_for_results(self, timeout: int = 10):
        """等待搜索结果"""
        time.sleep(timeout)
        
    def _extract_result_data(self) -> Optional[Dict]:
        """
        提取搜索结果数据
        
        Returns:
            包含有效期结束日期的字典，如无数据返回 None
        """
        try:
            # 页面使用 common-table 结构
            # 查找包含日期的 TR，然后提取该行的日期
            data = self.page.evaluate("""
                (function() {
                    // 查找所有包含日期格式文本的 td
                    var datePattern = /^\\d{4}-\\d{2}-\\d{2}$/;
                    var allTds = document.querySelectorAll('td');
                    
                    for (var td of allTds) {
                        var text = td.textContent.trim();
                        if (datePattern.test(text)) {
                            // 找到了日期 TD，向上查找 TR
                            var tr = td;
                            while (tr && tr.tagName !== 'TR') {
                                tr = tr.parentElement;
                            }
                            if (tr) {
                                var tds = tr.querySelectorAll('td');
                                if (tds.length >= 9) {
                                    return {
                                        found: true,
                                        expiry_date: tds[8].textContent.trim()
                                    };
                                }
                            }
                        }
                    }
                    return {found: false};
                })()
            """)
            
            return data
            
        except Exception as e:
            logger.warning(f"提取结果数据失败: {e}")
            return None
    
    def query(self, company_name: str, timeout: int = 60) -> Optional[str]:
        """
        查询公司的安全生产许可证有效期
        
        Args:
            company_name: 公司名称
            timeout: 超时时间（秒）
            
        Returns:
            有效期结束日期字符串 (YYYY-MM-DD)，如未找到返回 None
        """
        try:
            # 确保浏览器已启动
            if self.page is None:
                self.start()
            
            # 加载页面
            if not self._wait_for_page_load(timeout):
                return None
            
            # 展开搜索面板
            self._expand_search_panel()
            
            # 填写公司名称
            if not self._fill_company_name(company_name):
                return None
            
            # 点击查询
            self._click_search()
            
            # 等待结果
            self._wait_for_results(timeout=10)
            
            # 提取数据
            result = self._extract_result_data()
            
            if result and result.get('expiry_date'):
                logger.info(f"查询成功: {company_name} -> {result['expiry_date']}")
                return result['expiry_date']
            
            logger.info(f"未找到数据: {company_name}")
            return None
            
        except Exception as e:
            logger.error(f"查询失败: {company_name}, 错误: {e}")
            return None
    
    def batch_query(self, companies: List[str], timeout: int = 60) -> Dict[str, str]:
        """
        批量查询多个公司
        
        Args:
            companies: 公司名称列表
            timeout: 每个公司的超时时间（秒）
            
        Returns:
            字典 {公司名: 有效期结束日期}
        """
        results = {}
        
        for i, company in enumerate(companies):
            if not company or not isinstance(company, str):
                continue
                
            logger.info(f"[{i+1}/{len(companies)}] 查询: {company}")
            
            expiry_date = self.query(company.strip(), timeout)
            
            if expiry_date:
                results[company] = {
                    'date': expiry_date,
                    'source': '住建部官网'
                }
        
        return results


def query_mohurd_single(company_name: str) -> Optional[str]:
    """
    快捷函数：查询单家公司
    
    Args:
        company_name: 公司名称
        
    Returns:
        有效期结束日期或 None
    """
    with MohurdScraper(headless=True) as scraper:
        return scraper.query(company_name)


def query_mohurd_batch(companies: List[str]) -> Dict[str, Dict]:
    """
    快捷函数：批量查询
    
    Args:
        companies: 公司名称列表
        
    Returns:
        字典 {公司名: {'date': 日期, 'source': '住建部官网'}}
    """
    with MohurdScraper(headless=True) as scraper:
        return scraper.batch_query(companies)


if __name__ == "__main__":
    # 测试
    test_company = "江苏八达路桥有限公司"
    print(f"测试查询: {test_company}")
    
    result = query_mohurd_single(test_company)
    print(f"结果: {result}")

import tkinter as tk
from tkinter import ttk, messagebox
import json
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import time
import requests
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
 
# 删除之前的代理环境变量设置
if 'HTTP_PROXY' in os.environ:
    del os.environ['HTTP_PROXY']
if 'HTTPS_PROXY' in os.environ:
    del os.environ['HTTPS_PROXY']

# 设置 webdriver_manager 的缓存路径
os.environ['WDM_LOCAL'] = '1'  # 启用本地缓存
os.environ['WDM_SSL_VERIFY'] = '0'  # 禁用 SSL 验证

# 如果需要代理，取消下面的注释并填入代理地址
# os.environ['HTTPS_PROXY'] = 'http://your-proxy:port'
# os.environ['HTTP_PROXY'] = 'http://your-proxy:port'

def download_chromedriver():
    """手动下载 ChromeDriver"""
    try:
        # 创建下载目录
        download_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'drivers')
        if not os.path.exists(download_dir):
            os.makedirs(download_dir)
        
        # ChromeDriver 下载地址（这里使用的是 114.0.5735.90 版本，你可能需要根据你的 Chrome 版本调整）
        driver_url = "https://chromedriver.storage.googleapis.com/114.0.5735.90/chromedriver_win32.zip"
        driver_path = os.path.join(download_dir, 'chromedriver.exe')
        
        if not os.path.exists(driver_path):
            print(f"正在下载 ChromeDriver 到 {driver_path}")
            response = requests.get(driver_url, verify=False)
            
            # 保存 zip 文件
            zip_path = os.path.join(download_dir, 'chromedriver.zip')
            with open(zip_path, 'wb') as f:
                f.write(response.content)
            
            # 解压文件
            import zipfile
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(download_dir)
            
            # 删除 zip 文件
            os.remove(zip_path)
            
            print("ChromeDriver 下载完成")
        else:
            print("ChromeDriver 已存在")
        
        return driver_path
        
    except Exception as e:
        print(f"下载 ChromeDriver 失败: {e}")
        raise

def save_companies(companies):
    """保存公司列表到文件"""
    try:
        with open('company_config.json', 'w', encoding='utf-8') as f:
            json.dump(companies, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"保存公司配置失败: {e}")

def load_companies():
    """从文件加载公司列表"""
    try:
        if os.path.exists('company_config.json'):
            with open('company_config.json', 'r', encoding='utf-8') as f:
                return json.load(f)
    except Exception as e:
        print(f"加载公司配置失败: {e}")
    return ["华润广东医药有限公司"]  # 默认公司

class ConfigWindow:
    def __init__(self):
        self.window = tk.Toplevel()
        self.window.title("查询配置")
        self.window.geometry("500x450")  # 增加窗口高度
        
        # 设置窗口图标（如果有的话）
        try:
            self.window.iconbitmap('icon.ico')  # 如果有图标文件的话
        except:
            pass
        
        # 创建样式
        style = ttk.Style()
        style.configure('TLabel', padding=5, font=('微软雅黑', 9))
        style.configure('TEntry', padding=5)
        style.configure('TCombobox', padding=5)
        style.configure('TButton', padding=5)
        style.configure('Custom.TLabelframe', padding=10)
        style.configure('Custom.TLabelframe.Label', font=('微软雅黑', 9, 'bold'))
        
        # 创建主框架
        main_frame = ttk.Frame(self.window, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置grid权重，使窗口可调整大小
        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # 公司列表框架
        company_frame = ttk.LabelFrame(
            main_frame, 
            text="公司列表", 
            padding="10",
            style='Custom.TLabelframe'
        )
        company_frame.grid(
            row=0, 
            column=0, 
            columnspan=2, 
            sticky=(tk.W, tk.E, tk.N, tk.S), 
            pady=5
        )
        
        # 公司列表
        self.companies = load_companies()
        self.company_listbox = tk.Listbox(
            company_frame,
            selectmode=tk.MULTIPLE,
            height=10,  # 增加列表高度
            font=('微软雅黑', 9),
            activestyle='dotbox',  # 选中项的显示样式
            selectbackground='#0078D7',  # 选中项的背景色
            selectforeground='white'  # 选中项的文字颜色
        )
        self.company_listbox.grid(
            row=0, 
            column=0, 
            columnspan=2, 
            sticky=(tk.W, tk.E, tk.N, tk.S),
            padx=5
        )
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(
            company_frame,
            orient=tk.VERTICAL,
            command=self.company_listbox.yview
        )
        scrollbar.grid(row=0, column=2, sticky=(tk.N, tk.S))
        self.company_listbox.configure(yscrollcommand=scrollbar.set)
        
        # 更新公司列表并默认选择第一个
        self.update_company_list()
        if self.companies:
            self.company_listbox.select_set(0)  # 默认选择第一个公司
        
        # 新公司输入框和按钮的框架
        input_frame = ttk.Frame(company_frame)
        input_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        # 新公司输入框
        self.new_company_var = tk.StringVar()
        ttk.Entry(
            input_frame,
            textvariable=self.new_company_var,
            width=40,
            font=('微软雅黑', 9)
        ).grid(row=0, column=0, sticky=(tk.W, tk.E), padx=5)
        
        # 添加和删除按钮框架
        btn_frame = ttk.Frame(input_frame)
        btn_frame.grid(row=0, column=1, sticky=(tk.W, tk.E))
        
        # 添加和删除按钮
        ttk.Button(
            btn_frame,
            text="添加",
            command=self.add_company,
            style='TButton',
            width=8
        ).grid(row=0, column=0, padx=2)
        
        ttk.Button(
            btn_frame,
            text="删除",
            command=self.delete_company,
            style='TButton',
            width=8
        ).grid(row=0, column=1, padx=2)
        
        # 菜单选择框架
        menu_frame = ttk.LabelFrame(
            main_frame,
            text="菜单选择",
            padding="10",
            style='Custom.TLabelframe'
        )
        menu_frame.grid(
            row=1,
            column=0,
            columnspan=2,
            sticky=(tk.W, tk.E),
            pady=10
        )
        
        # 菜单选择下拉框
        self.menu_options = [
            "流向查询-采购",
            "流向查询-库存",
            "流向查询-销售",
            "流向查询-广东-零售配送",
            "广东-回款记录表"
        ]
        self.menu_var = tk.StringVar(value=self.menu_options[1])
        self.menu_combo = ttk.Combobox(
            menu_frame,
            textvariable=self.menu_var,
            values=self.menu_options,
            state="readonly",
            width=35,
            font=('微软雅黑', 9)
        )
        self.menu_combo.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=5)
        
        # 确定按钮
        ttk.Button(
            main_frame,
            text="确定",
            command=self.confirm,
            style='TButton',
            width=15
        ).grid(row=2, column=0, columnspan=2, pady=15)
        
        # 设置窗口居中
        self.center_window()
        
        # 初始化配置信息
        self.config_info = None
        
        # 设置模态窗口
        self.window.transient(self.window.master)
        self.window.grab_set()
        
        # 绑定回车键到确定按钮
        self.window.bind('<Return>', lambda e: self.confirm())
        
        # 设置最小窗口大小
        self.window.minsize(500, 450)
        
        self.window.wait_window()
    
    def center_window(self):
        """将窗口居中显示"""
        self.window.update_idletasks()
        width = self.window.winfo_width()
        height = self.window.winfo_height()
        x = (self.window.winfo_screenwidth() // 2) - (width // 2)
        y = (self.window.winfo_screenheight() // 2) - (height // 2)
        self.window.geometry('{}x{}+{}+{}'.format(width, height, x, y))
    
    def update_company_list(self):
        """更新公司列表显示"""
        self.company_listbox.delete(0, tk.END)
        for company in self.companies:
            self.company_listbox.insert(tk.END, company)
    
    def add_company(self):
        """添加新公司"""
        new_company = self.new_company_var.get().strip()
        if new_company and new_company not in self.companies:
            self.companies.append(new_company)
            self.update_company_list()
            save_companies(self.companies)
            self.new_company_var.set("")
    
    def delete_company(self):
        """删除选中的公司"""
        selected = self.company_listbox.curselection()
        if selected:
            # 从后往前删除，避免索引变化
            for index in reversed(selected):
                del self.companies[index]
            self.update_company_list()
            save_companies(self.companies)
    
    def confirm(self):
        """确认配置"""
        selected = self.company_listbox.curselection()
        if not selected:
            messagebox.showwarning("警告", "请至少选择一个公司！")
            return
        
        selected_companies = [self.companies[i] for i in selected]
        self.config_info = {
            'companies': selected_companies,
            'menu_option': self.menu_var.get()
        }
        self.window.destroy()
    
    def get_config(self):
        """获取配置信息"""
        return self.config_info

class LoginWindow:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("登录")
        self.window.geometry("300x200")
        
        # 创建样式
        style = ttk.Style()
        style.configure('TLabel', padding=5)
        style.configure('TEntry', padding=5)
        style.configure('TButton', padding=5)
        
        # 创建框架
        main_frame = ttk.Frame(self.window, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 用户名
        ttk.Label(main_frame, text="用户名:").grid(row=0, column=0, sticky=tk.W)
        self.username_var = tk.StringVar()
        self.username_entry = ttk.Entry(main_frame, textvariable=self.username_var)
        self.username_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5)
        
        # 密码
        ttk.Label(main_frame, text="密码:").grid(row=1, column=0, sticky=tk.W)
        self.password_var = tk.StringVar()
        self.password_entry = ttk.Entry(main_frame, textvariable=self.password_var, show="*")
        self.password_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5)
        
        # 记住密码复选框
        self.remember_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(main_frame, text="记住密码", variable=self.remember_var).grid(row=2, column=0, columnspan=2, pady=5)
        
        # 登录按钮
        ttk.Button(main_frame, text="登录", command=self.login).grid(row=3, column=0, columnspan=2, pady=10)
        
        # 加载保存的登录信息
        self.load_saved_login()
        
        # 设置窗口居中
        self.center_window()
        
        self.window.mainloop()
    
    def center_window(self):
        """将窗口居中显示"""
        self.window.update_idletasks()
        width = self.window.winfo_width()
        height = self.window.winfo_height()
        x = (self.window.winfo_screenwidth() // 2) - (width // 2)
        y = (self.window.winfo_screenheight() // 2) - (height // 2)
        self.window.geometry('{}x{}+{}+{}'.format(width, height, x, y))
    
    def load_saved_login(self):
        """加载保存的登录信息"""
        if os.path.exists('login_config.json'):
            try:
                with open('login_config.json', 'r') as f:
                    config = json.load(f)
                    self.username_var.set(config.get('username', ''))
                    self.password_var.set(config.get('password', ''))
            except Exception as e:
                print(f"加载配置文件失败: {e}")
    
    def save_login_info(self):
        """保存登录信息"""
        if self.remember_var.get():
            config = {
                'username': self.username_var.get(),
                'password': self.password_var.get()
            }
            try:
                with open('login_config.json', 'w') as f:
                    json.dump(config, f)
            except Exception as e:
                print(f"保存配置文件失败: {e}")
    
    def login(self):
        """登录处理"""
        username = self.username_var.get()
        password = self.password_var.get()
        
        if not username or not password:
            messagebox.showerror("错误", "用户名和密码不能为空！")
            return
        
        # 保存登录信息
        self.save_login_info()
        
        # 返回登录信息
        self.login_info = {
            'username': username,
            'password': password
        }
        self.window.destroy()

    def get_login_info(self):
        """获取登录信息"""
        if hasattr(self, 'login_info'):
            return self.login_info
        return None

class WebScraper:
    def __init__(self):
        self.driver = None
        self.wait = None
        
    def setup_driver(self):
        """设置浏览器驱动"""
        try:
            options = webdriver.ChromeOptions()
            
            # 性能优化配置
            options.add_argument('--disable-gpu')
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument('--disable-software-rasterizer')
            
            # 添加性能相关参数
            options.add_argument('--disable-extensions')
            options.add_argument('--disable-logging')
            options.add_argument('--disable-notifications')
            options.add_argument('--disable-default-apps')
            options.add_argument('--disable-popup-blocking')
            options.add_argument('--disable-web-security')
            
            # 设置更激进的页面加载策略
            options.page_load_strategy = 'none'
            
            # 添加实验性选项
            options.add_experimental_option('excludeSwitches', ['enable-logging'])
            options.add_experimental_option('useAutomationExtension', False)
            
            # 禁用图片加载以提升速度
            prefs = {
                'profile.managed_default_content_settings.images': 2,
                'profile.default_content_setting_values.notifications': 2
            }
            options.add_experimental_option('prefs', prefs)
            
            # 手动下载并获取 ChromeDriver 路径
            driver_path = download_chromedriver()
            service = Service(driver_path)
            
            self.driver = webdriver.Chrome(service=service, options=options)
            
            # 设置更短的超时时间
            self.driver.set_page_load_timeout(30)
            self.driver.set_script_timeout(30)
            self.wait = WebDriverWait(self.driver, 15)
            
            print("浏览器驱动初始化成功")
            
        except Exception as e:
            print(f"设置浏览器驱动失败: {e}")
            raise
    
    def login(self, username, password):
        """执行登录操作"""
        try:
            print("开始访问登录页面...")
            
            # 简化缓存清理
            print("清除浏览器缓存...")
            try:
                self.driver.delete_all_cookies()
                print("缓存清除完成")
            except Exception as e:
                print(f"清除缓存时出错: {str(e)}")
            
            # 添加重试机制并缩短等待时间
            max_retries = 3
            retry_delay = 2
            
            for attempt in range(max_retries):
                try:
                    self.driver.get("https://lxdata.crpcg.com/login")
                    # 等待登录页面加载的关键元素
                    self.wait.until(
                        EC.presence_of_element_located((By.ID, "loginId"))
                    )
                    # 确保页面完全加载
                    self.wait.until(
                        lambda driver: driver.execute_script("return document.readyState") == "complete"
                    )
                    time.sleep(2)  # 额外等待以确保页面完全加载
                    break
                except Exception as e:
                    if attempt == max_retries - 1:
                        raise
                    print(f"第{attempt + 1}次尝试访问登录页面失败，正在重试...")
                    time.sleep(retry_delay)
            
            print("等待登录页面加载...")
            
            # 使用更可靠的元素定位方式，并确保元素可交互
            username_input = self.wait.until(
                EC.element_to_be_clickable((By.ID, "loginId"))
            )
            username_input.clear()
            time.sleep(0.5)
            username_input.send_keys(str(username))
            time.sleep(0.5)
            
            password_input = self.wait.until(
                EC.element_to_be_clickable((By.ID, "password"))
            )
            password_input.clear()
            time.sleep(0.5)
            password_input.send_keys(str(password))
            time.sleep(0.5)
            
            # 使用更可靠的方式点击登录按钮
            try:
                # 首先尝试常规点击
                login_button = self.wait.until(
                    EC.element_to_be_clickable((By.ID, "loginBtn"))
                )
                login_button.click()
            except Exception as e:
                print(f"常规点击失败，尝试JavaScript点击: {str(e)}")
                # 如果常规点击失败，尝试JavaScript点击
                login_button = self.wait.until(
                    EC.presence_of_element_located((By.ID, "loginBtn"))
                )
                self.driver.execute_script("arguments[0].click();", login_button)
            
            # 等待登录结果
            try:
                WebDriverWait(self.driver, 10).until(
                    lambda driver: "auth/index" in driver.current_url
                )
                print("登录成功")
                return True
            except TimeoutException:
                print("登录失败：超时")
                # 保存错误截图和页面源码
                self.driver.save_screenshot("login_timeout.png")
                with open("login_timeout.html", "w", encoding="utf-8") as f:
                    f.write(self.driver.page_source)
                return False
                
        except Exception as e:
            print(f"登录过程中发生错误: {str(e)}")
            # 保存错误截图和页面源码
            self.driver.save_screenshot("login_error.png")
            with open("login_error.html", "w", encoding="utf-8") as f:
                f.write(self.driver.page_source)
            return False
    
    def fetch_page_data(self, url):
        """获取页面数据"""
        try:
            self.driver.get(url)
            
            # 等待页面加载完成（这里需要根据实际页面情况调整）
            time.sleep(2)
            
            # 获取页面内容（这里需要根据实际页面结构调整）
            content = self.driver.page_source
            print("成功获取页面内容")
            return content
            
        except Exception as e:
            print(f"获取页面数据失败: {e}")
            return None
    
    def close(self):
        """关闭浏览器"""
        if self.driver:
            self.driver.quit()
    
    def perform_search(self, company_name, menu_option):
        """执行流向查询的搜索操作并导出数据"""
        try:
            print("开始执行查询操作...")
            
            # 确保已经登录成功
            if "auth/index" not in self.driver.current_url:
                print("未处于登录状态，尝试刷新页面...")
                self.driver.refresh()
                time.sleep(3)
                
                if "auth/index" not in self.driver.current_url:
                    raise Exception("未处于登录状态，无法执行查询")
            
            # 等待页面完全加载
            time.sleep(3)
            
            # 1. 点击左侧菜单
            print(f"查找并点击{menu_option}菜单...")
            try:
                # 等待菜单可见和可点击
                menu = self.wait.until(
                    EC.presence_of_element_located((
                        By.XPATH, 
                        f"//span[contains(text(), '{menu_option}')]"
                    ))
                )
                # 确保菜单元素在视图中
                self.driver.execute_script("arguments[0].scrollIntoView(true);", menu)
                time.sleep(1)
                
                # 尝试不同的点击方法
                try:
                    # 方法1：直接点击
                    menu.click()
                except Exception:
                    try:
                        # 方法2：JavaScript点击
                        self.driver.execute_script("arguments[0].click();", menu)
                    except Exception:
                        # 方法3：Actions链点击
                        actions = webdriver.ActionChains(self.driver)
                        actions.move_to_element(menu).click().perform()
                
                print("已点击菜单")
                time.sleep(2)
                
            except Exception as e:
                print(f"点击菜单时出错: {str(e)}")
                # 保存当前页面状态以供调试
                self.driver.save_screenshot("menu_error.png")
                with open("menu_error.html", "w", encoding="utf-8") as f:
                    f.write(self.driver.page_source)
                # 打印所有可见的菜单项
                try:
                    menus = self.driver.find_elements(By.XPATH, "//span[contains(@class, 'ant-menu-title-content')]")
                    print("\n可见的菜单项:")
                    for i, m in enumerate(menus):
                        print(f"菜单 {i+1}: {m.text}")
                except Exception as e:
                    print(f"获取菜单信息时出错: {str(e)}")
                raise

            # 2. 点击公司选择框
            print("查找并点击'请选择'按钮...")
            company_select = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//span[text()='请选择']"))
            )
            company_select.click()
            print("已点击'请选择'按钮")
            time.sleep(2)
            
            # 3. 在弹出的下拉框中输入公司名称
            print("输入公司名称...")
            try:
                # 通过搜索图标定位输入框
                search_icon = self.wait.until(
                    EC.presence_of_element_located((
                        By.XPATH, 
                        "//span[@role='img' and @aria-label='search']"
                    ))
                )
                # 从搜索图标找到相关的输入框
                company_input = search_icon.find_element(
                    By.XPATH,
                    "./following::input[@class='ant-input']"
                )
                
                if not company_input.is_displayed() or not company_input.is_enabled():
                    raise Exception("输入框不可见或不可交互")
                
                # 输入公司名称
                company_input.clear()
                company_input.send_keys(company_name)
                time.sleep(1)
                print("公司名称输入完成")
                
            except Exception as e:
                print(f"输入公司名称时出错: {str(e)}")
                self.driver.save_screenshot("input_error.png")
                with open("input_error.html", "w", encoding="utf-8") as f:
                    f.write(self.driver.page_source)
                raise
            
            # 4. 选择下拉选项
            print("选择公司...")
            try:
                # 等待并选择指定的公司
                company_name = "华润广东医药有限公司"
                # 定位复选框
                checkbox_input = self.wait.until(
                    EC.presence_of_element_located((
                        By.CSS_SELECTOR, 
                        "label.ant-checkbox-wrapper.gd-ellipsis input.ant-checkbox-input"
                    ))
                )
                print(f"找到公司复选框")
                
                # 使用JavaScript点击复选框
                self.driver.execute_script("arguments[0].click();", checkbox_input)
                print("已选择公司")
                time.sleep(1)

                # 点击确定按钮
                print("点击确定按钮...")
                confirm_button = self.wait.until(
                    EC.element_to_be_clickable((
                        By.XPATH,
                        "//div[contains(@class, 'row-flex')]//button[contains(@class, 'ant-btn-primary')]//span[text()='确定']/parent::button"
                    ))
                )
                print("找到确定按钮")
                
                # 使用 JavaScript 点击确定按钮
                self.driver.execute_script("arguments[0].click();", confirm_button)
                print("已点击确定按钮")
                time.sleep(1)

            except Exception as e:
                print(f"选择公司或点击确定按钮时出错: {str(e)}")
                self.driver.save_screenshot("select_company_error.png")
                with open("select_company_error.html", "w", encoding="utf-8") as f:
                    f.write(self.driver.page_source)
                print("尝试查找确定按钮...")
                try:
                    # 打印所有按钮的信息
                    buttons = self.driver.find_elements(
                        By.XPATH, 
                        "//div[contains(@class, 'row-flex')]//button"
                    )
                    print(f"找到 {len(buttons)} 个按钮")
                    for i, button in enumerate(buttons):
                        print(f"按钮 {i+1}:")
                        print(f"- 文本: {button.text}")
                        print(f"- 类名: {button.get_attribute('class')}")
                        print(f"- HTML: {button.get_attribute('outerHTML')}")
                except Exception as e:
                    print(f"获取按钮信息时出错: {str(e)}")
                raise
            
            # 5. 点击查询按钮
            print("点击查询按钮...")
            try:
                search_button = self.wait.until(
                    EC.element_to_be_clickable((
                        By.XPATH,
                        "//button[contains(@class, 'ant-btn-primary') and contains(@class, 'fp-btn')]//span[text()='查询']/parent::button"
                    ))
                )
                print("找到查询按钮")
                
                # 使用 JavaScript 点击查询按钮
                self.driver.execute_script("arguments[0].click();", search_button)
                print("已点击查询按钮")
                time.sleep(2)  # 等待查询结果
                
            except Exception as e:
                print(f"点击查询按钮时出错: {str(e)}")
                self.driver.save_screenshot("search_button_error.png")
                with open("search_button_error.html", "w", encoding="utf-8") as f:
                    f.write(self.driver.page_source)
                # 打印按钮信息
                try:
                    buttons = self.driver.find_elements(
                        By.XPATH, 
                        "//button[contains(@class, 'ant-btn')]"
                    )
                    print(f"找到 {len(buttons)} 个按钮")
                    for i, button in enumerate(buttons):
                        print(f"按钮 {i+1}:")
                        print(f"- 文本: {button.text}")
                        print(f"- 类名: {button.get_attribute('class')}")
                        print(f"- HTML: {button.get_attribute('outerHTML')}")
                except Exception as e:
                    print(f"获取按钮信息时出错: {str(e)}")
                raise
            
            # 6. 等待数据加载
            print("等待数据加载...")
            try:
                # 等待表格加载完成
                self.wait.until(
                    EC.presence_of_element_located((By.CLASS_NAME, "fast-table"))
                )
                time.sleep(3)  # 等待数据完全加载
                
                # 获取数据容器
                data_container = self.wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "div.fast-sub-table.BR"))
                )
                
                # 先进行水平和垂直滚动，确保所有数据都加载
                print("开始滚动加载所有数据...")
                
                # 获取表头容器进行水平滚动
                header_container = self.wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "div.fast-sub-table.TR"))
                )
                
                try:
                    # 水平滚动表头到最右
                    print("开始水平滚动...")
                    max_attempts = 20  # 最大尝试次数
                    attempts = 0
                    last_scroll = -1
                    
                    while attempts < max_attempts:
                        # 获取当前滚动位置和最大滚动宽度
                        scroll_width = self.driver.execute_script(
                            "return arguments[0].scrollWidth - arguments[0].clientWidth;",
                            header_container
                        )
                        current_scroll = self.driver.execute_script(
                            "return arguments[0].scrollLeft;",
                            header_container
                        )
                        
                        print(f"当前水平滚动位置: {current_scroll}/{scroll_width}")
                        
                        if current_scroll >= scroll_width or current_scroll == last_scroll:
                            break
                        
                        # 滚动步长设置为200像素
                        self.driver.execute_script(
                            "arguments[0].scrollLeft += 200;",
                            header_container
                        )
                        time.sleep(1)  # 增加等待时间
                        
                        last_scroll = current_scroll
                        attempts += 1
                    
                    # 等待一下确保数据加载
                    time.sleep(2)
                    
                    # 滚动回最左侧
                    self.driver.execute_script(
                        "arguments[0].scrollLeft = 0;",
                        header_container
                    )
                    time.sleep(1)
                    
                    print("水平滚动完成")
                    
                    # 垂直滚动到底部
                    print("开始垂直滚动...")
                    attempts = 0
                    last_height = -1
                    
                    while attempts < max_attempts:
                        # 获取当前高度
                        current_height = self.driver.execute_script(
                            "return arguments[0].scrollTop;",
                            data_container
                        )
                        
                        # 获取可滚动的总高度
                        scroll_height = self.driver.execute_script(
                            "return arguments[0].scrollHeight - arguments[0].clientHeight;",
                            data_container
                        )
                        
                        print(f"当前垂直滚动位置: {current_height}/{scroll_height}")
                        
                        if current_height >= scroll_height or current_height == last_height:
                            break
                        
                        # 滚动一个页面高度
                        self.driver.execute_script(
                            "arguments[0].scrollTop += arguments[0].clientHeight;",
                            data_container
                        )
                        time.sleep(1)
                        
                        last_height = current_height
                        attempts += 1
                    
                    # 等待一下确保数据加载
                    time.sleep(2)
                    
                    # 滚动回顶部
                    self.driver.execute_script(
                        "arguments[0].scrollTop = 0;",
                        data_container
                    )
                    time.sleep(1)
                    
                    print("垂直滚动完成")
                    
                except Exception as e:
                    print(f"滚动过程中出错: {e}")
                    self.driver.save_screenshot("scroll_error.png")
                
                print("所有数据加载完成，开始提取数据...")
                
                # 获取完整表头
                headers = []
                print("开始获取表头...")
                
                # 重新滚动表头以获取所有列
                current_scroll = 0
                header_dict = {}
                
                # 先滚动到最右边获取总宽度
                self.driver.execute_script(
                    "arguments[0].scrollLeft = arguments[0].scrollWidth;",
                    header_container
                )
                time.sleep(1)
                
                total_width = self.driver.execute_script(
                    "return arguments[0].scrollWidth;",
                    header_container
                )
                
                # 滚动回最左边
                self.driver.execute_script(
                    "arguments[0].scrollLeft = 0;",
                    header_container
                )
                time.sleep(1)
                
                # 使用较小的步长获取表头
                scroll_step = 100
                while current_scroll <= total_width:
                    # 滚动表头
                    self.driver.execute_script(
                        "arguments[0].scrollLeft = arguments[1];",
                        header_container,
                        current_scroll
                    )
                    time.sleep(0.5)
                    
                    # 获取当前可见的表头
                    header_cells = self.driver.find_elements(
                        By.CSS_SELECTOR,
                        "div.fast-sub-table.TR .cell.tr"
                    )
                    
                    for cell in header_cells:
                        try:
                            col_index = int(cell.get_attribute('data-col'))
                            if col_index not in header_dict:
                                header_text = cell.find_element(By.CSS_SELECTOR, "span.fast-table-v").text.strip()
                                if header_text:  # 只保存非空表头
                                    header_dict[col_index] = header_text
                                    print(f"找到表头: {header_text} (列 {col_index})")
                        except Exception as e:
                            print(f"处理表头单元格时出错: {e}")
                            continue
                    
                    current_scroll += scroll_step
                
                # 按列索引顺序添加表头，确保不会遗漏
                max_col_index = max(header_dict.keys())
                headers = []
                for i in range(max_col_index + 1):
                    header_text = header_dict.get(i, f'Column_{i}')
                    headers.append(header_text)
                
                print(f"找到所有表头: {headers}")
                print(f"表头总数: {len(headers)}")

                # 获取所有数据
                row_data_dict = {}
                print("开始获取表格数据...")
                
                try:
                    # 获取表格总高度和宽度
                    total_height = self.driver.execute_script(
                        "return arguments[0].scrollHeight;",
                        data_container
                    )
                    total_width = self.driver.execute_script(
                        "return arguments[0].scrollWidth;",
                        data_container
                    )
                    
                    # 获取视口大小
                    viewport_height = data_container.size['height']
                    viewport_width = data_container.size['width']
                    
                    # 增加滚动步长以提高速度
                    vertical_step = viewport_height - 30
                    horizontal_step = viewport_width - 30
                    
                    print(f"表格尺寸 - 高度: {total_height}, 宽度: {total_width}")
                    
                    # 优化的数据获取逻辑
                    for v_scroll in range(0, total_height, vertical_step):
                        self.driver.execute_script(
                            "arguments[0].scrollTop = arguments[1];",
                            data_container,
                            v_scroll
                        )
                        
                        for h_scroll in range(0, total_width, horizontal_step):
                            self.driver.execute_script(
                                "arguments[0].scrollLeft = arguments[1];",
                                data_container,
                                h_scroll
                            )
                            time.sleep(0.1)  # 减少等待时间
                            
                            # 使用JavaScript批量获取单元格数据
                            cells_data = self.driver.execute_script("""
                                const cells = arguments[0].querySelectorAll('.cell.br');
                                return Array.from(cells).map(cell => {
                                    const span = cell.querySelector('span.fast-table-v');
                                    return {
                                        row: parseInt(cell.getAttribute('data-row')),
                                        col: parseInt(cell.getAttribute('data-col')),
                                        text: span ? span.textContent.trim() : ''
                                    };
                                });
                            """, data_container)
                            
                            # 批量处理数据
                            for cell_data in cells_data:
                                if cell_data['text']:
                                    row_idx = cell_data['row']
                                    col_idx = cell_data['col']
                                    if row_idx not in row_data_dict:
                                        row_data_dict[row_idx] = {}
                                    row_data_dict[row_idx][col_idx] = cell_data['text']
                            
                            print(f"滚动位置 - 垂直: {v_scroll}/{total_height}, 水平: {h_scroll}/{total_width}")
                            print(f"当前已获取 {len(row_data_dict)} 行数据")
                    
                except Exception as e:
                    print(f"获取数据时出错: {e}")
                    self.driver.save_screenshot("data_error.png")
                
                # 数据处理
                print("开始处理数据...")
                table_data = []
                
                # 直接使用原始数据，不处理日期格式
                table_data = [
                    [row_data_dict[row_idx].get(col_idx, '') for col_idx in range(len(headers))]
                    for row_idx in sorted(row_data_dict.keys())
                    if any(row_data_dict[row_idx].values())
                ]
                
                print(f"总共找到 {len(table_data)} 行数据")
                
                # 导出到CSV
                if table_data:
                    import csv
                    from datetime import datetime
                    
                    filename = f"库存数据_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
                    
                    with open(filename, 'w', newline='', encoding='utf-8-sig') as f:
                        writer = csv.writer(f)
                        writer.writerow(headers)
                        writer.writerows(table_data)
                    
                    print(f"数据已保存到文件: {filename}")
                    return True
                else:
                    print("未找到有效的表格数据")
                    return False
                
            except Exception as e:
                print(f"处理数据时出错: {str(e)}")
                # 保存错误现场
                self.driver.save_screenshot("data_error.png")
                return False
            
        except Exception as e:
            print(f"执行查询操作时发生错误: {e}")
            self.driver.save_screenshot("search_error.png")
            return False

def main():
    # 创建登录窗口获取用户信息
    login_window = LoginWindow()
    login_info = login_window.get_login_info()
    
    if not login_info:
        print("未获取到登录信息")
        return
    
    # 创建配置窗口获取查询配置
    config_window = ConfigWindow()
    config_info = config_window.get_config()
    
    if not config_info:
        print("未获取到配置信息")
        return
    
    # 创建网页抓取器
    scraper = WebScraper()
    
    try:
        # 设置浏览器驱动
        scraper.setup_driver()
        
        # 对每个选中的公司执行查询操作
        for i, company in enumerate(config_info['companies']):
            print(f"\n开始处理公司: {company} ({i+1}/{len(config_info['companies'])})")
            
            # 检查登录状态，必要时重新登录
            if "auth/index" not in scraper.driver.current_url:
                print("需要重新登录...")
                if not scraper.login(login_info['username'], login_info['password']):
                    print(f"重新登录失败，跳过公司: {company}")
                    continue
                print("重新登录成功")
            
            # 执行查询操作
            try:
                if scraper.perform_search(company, config_info['menu_option']):
                    print(f"{company} 数据查询和导出成功")
                else:
                    print(f"{company} 数据查询或导出失败")
            except Exception as e:
                print(f"处理公司 {company} 时出错: {e}")
                # 保存错误截图
                scraper.driver.save_screenshot(f"error_{company.replace(' ', '_')}.png")
                continue
            
            # 在每次查询后等待一段时间
            time.sleep(3)
            
    except Exception as e:
        print(f"发生错误: {e}")
        messagebox.showerror("错误", str(e))
        
    finally:
        # 关闭浏览器
        scraper.close()

if __name__ == "__main__":
    main()

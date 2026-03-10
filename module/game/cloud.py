import atexit
import os
import json
import psutil
import platform
import sys
import base64
import time
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, SessionNotCreatedException
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.chromium.options import ChromiumOptions
from selenium.webdriver.chromium.service import ChromiumService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.selenium_manager import SeleniumManager
from selenium.webdriver.remote.command import Command
from selenium.common.exceptions import WebDriverException

from module.config import Config
from module.game.base import GameControllerBase
from module.logger import Logger

from utils.console import is_docker_started


class CloudGameController(GameControllerBase):
    MOONLIGHT_BASE_URL = "http://192.168.1.120:23456"   # Moonlight Web 基础地址
    MOONLIGHT_HOST_ID = "3866122460"                     # Sunshine Host ID（可配置）
    MOONLIGHT_APP_ID = "454747157"                       # Moonlight App ID
    GAME_URL = f"{MOONLIGHT_BASE_URL}/stream.html?hostId={MOONLIGHT_HOST_ID}&appId={MOONLIGHT_APP_ID}"  # Moonlight Web 完整地址
    BROWSER_TAG = "--march-7th-assistant-sr-cloud-game"  # 自定义浏览器参数作为标识，用于识别哪些浏览器进程属于三月七小助手
    BROWSER_INSTALL_PATH = os.path.join(os.getcwd(), "3rdparty", "WebBrowser")  # 浏览器安装路径
    INTEGRATED_BROWSER_VERSION = "140.0.7339.207"      # 浏览器版本

    @staticmethod
    def _get_platform_dir() -> str:
        """获取当前平台对应的目录名称"""
        system = platform.system()
        machine = platform.machine()

        if system == "Windows":
            if machine in ("AMD64", "x86_64"):
                return "win64"
            elif machine in ("ARM64", "aarch64"):
                return "win-arm64"  # 未验证
            else:
                return "win64"
        elif system == "Darwin":
            if machine == "arm64":
                return "mac-arm64"
            else:
                return "mac64"  # 未验证
        elif system == "Linux":
            if machine in ("ARM64", "aarch64"):
                return "linux-arm64"  # 未验证
            else:
                return "linux64"  # 未验证
        else:
            return "win64"

    @staticmethod
    def _get_integrated_browser_path() -> str:
        """获取内置浏览器路径"""
        platform_dir = CloudGameController._get_platform_dir()
        browser_install_path = CloudGameController.BROWSER_INSTALL_PATH
        browser_version = CloudGameController.INTEGRATED_BROWSER_VERSION

        if platform.system() == "Darwin":
            return os.path.join(browser_install_path, "chrome", platform_dir, browser_version, "Google Chrome for Testing.app", "Contents", "MacOS", "Google Chrome for Testing")
        elif platform.system() == "Windows":
            return os.path.join(browser_install_path, "chrome", platform_dir, browser_version, "chrome.exe")
        else:  # Linux
            return os.path.join(browser_install_path, "chrome", platform_dir, browser_version, "chrome")  # 未验证

    @staticmethod
    def _get_integrated_driver_path() -> str:
        """获取内置驱动路径"""
        platform_dir = CloudGameController._get_platform_dir()
        browser_install_path = CloudGameController.BROWSER_INSTALL_PATH
        browser_version = CloudGameController.INTEGRATED_BROWSER_VERSION

        if platform.system() == "Darwin":  # macOS
            return os.path.join(browser_install_path, "chromedriver", platform_dir, browser_version, "chromedriver")
        elif platform.system() == "Windows":
            return os.path.join(browser_install_path, "chromedriver", platform_dir, browser_version, "chromedriver.exe")
        else:  # Linux
            return os.path.join(browser_install_path, "chromedriver", platform_dir, browser_version, "chromedriver")  # 未验证
    MAX_RETRIES = 3  # 网页加载重试次数，0=不重试
    PERFERENCES = {
        "profile": {
            "content_settings": {
                "exceptions": {
                    "keyboard_lock": {  # 允许 keyboard_lock 权限
                        "http://192.168.1.120:23456,*": {"setting": 1}
                    },
                    "clipboard": {   # 允许剪贴板读取权限
                        "http://192.168.1.120:23456,*": {"setting": 1}
                    }
                }
            }
        }
    }

    def __init__(self, cfg: Config, logger: Logger):
        super().__init__(script_path=cfg.script_path, logger=logger)
        self.driver = None
        self.cfg = cfg
        self.logger = logger

        atexit.register(self._clean_at_exit)

    def _wait_page_loaded(self, timeout=15) -> None:
        """等待 Moonlight Web 页面加载完成"""
        if not self.driver:
            return
        for retry in range(self.MAX_RETRIES + 1):
            if retry > 0:
                self.log_warning(f"页面加载超时，正在刷新重试... ({retry}/{self.MAX_RETRIES})")
                self.driver.refresh()
            try:
                WebDriverWait(self.driver, timeout).until(
                    lambda d: d.execute_script("return document.readyState") == "complete"
                )
                return
            except TimeoutException:
                pass

        raise Exception("页面加载失败，多次刷新无效。")

    def _confirm_viewport_resolution(self) -> None:
        """
        设置网页分辨率大小
        """
        self.driver.execute_cdp_cmd("Emulation.setDeviceMetricsOverride", {
            "width": 1920,
            "height": 1080,
            "deviceScaleFactor": 1,
            "mobile": False
        })

    def _prepare_browser_and_driver(self, browser_type: str, integrated: bool) -> tuple[str, str]:
        self.user_profile_path = os.path.join(self.BROWSER_INSTALL_PATH, "UserProfile", self.cfg.browser_type.capitalize())
        # 判断环境变量 MARCH7TH_BROWSER_PATH 和 MARCH7TH_DRIVER_PATH，同时存在时优先使用
        env_browser_path = os.environ.get("MARCH7TH_BROWSER_PATH")
        env_driver_path = os.environ.get("MARCH7TH_DRIVER_PATH")
        if env_browser_path and env_driver_path:
            self.log_debug("检测到环境变量 MARCH7TH_BROWSER_PATH 和 MARCH7TH_DRIVER_PATH，优先使用指定路径")
            self.log_debug(f"browser_path = {env_browser_path}")
            self.log_debug(f"driver_path = {env_driver_path}")
            return env_browser_path, env_driver_path

        # 输出平台信息
        platform_dir = self._get_platform_dir()
        self.log_debug(f"检测到系统平台: {platform.system()} {platform.machine()}, 使用目录: {platform_dir}")

        if integrated:
            browser_path = self._get_integrated_browser_path()
            driver_path = self._get_integrated_driver_path()
            if not os.path.exists(browser_path) or not os.path.exists(driver_path):
                args = ["--browser", browser_type,
                        "--cache-path", self.BROWSER_INSTALL_PATH,
                        "--browser-version", self.INTEGRATED_BROWSER_VERSION,
                        "--force-browser-download",
                        "--skip-driver-in-path",
                        "--skip-browser-in-path"]
                if self.cfg.browser_download_use_mirror:
                    self.log_info(f"正在使用镜像源，浏览器镜像源：{self.cfg.browser_mirror_urls['chrome']}，"
                                  f"驱动镜像源：{self.cfg.browser_mirror_urls['chromedriver']}")
                    args.extend([
                        "--browser-mirror-url", self.cfg.browser_mirror_urls["chrome"],
                        "--driver-mirror-url", self.cfg.browser_mirror_urls["chromedriver"],
                    ])
                try:
                    self.log_info("正在下载浏览器和驱动...")
                    SeleniumManager().binary_paths(args)
                except WebDriverException as e:
                    raise Exception(f"浏览器和驱动下载失败：{e}")
        else:
            # 尝试在本地查找浏览器
            args = ["--browser", browser_type,
                    "--cache-path", self.BROWSER_INSTALL_PATH,
                    "--avoid-browser-download",
                    "--skip-driver-in-path"]
            if self.cfg.browser_download_use_mirror:
                if browser_type == "chrome":
                    args.extend([
                        "--driver-mirror-url", self.cfg.browser_mirror_urls["chromedriver"]
                    ])
                elif browser_type == "edge":
                    args.extend([
                        "--driver-mirror-url", self.cfg.browser_mirror_urls["edgedriver"]
                    ])
            try:
                result = SeleniumManager().binary_paths(args)
            except WebDriverException as e:
                raise Exception(f"查找 {browser_type} 浏览器出错：{e}")
            browser_path = result["browser_path"]
            driver_path = result["driver_path"]
        self.log_debug(f"browser_path = {browser_path}")
        self.log_debug(f"driver_path = {driver_path}")
        return browser_path, driver_path

    def _get_browser_arguments(self, headless) -> list[str]:
        args = [
            self.BROWSER_TAG,   # 标记浏览器是由脚本启动
            "--disable-infobars",   # 去掉提示 "Chrome测试版仅适用于自动测试。" 和 "浏览器正由自动测试软件控制。"
            "--lang=zh-CN",     # 浏览器语言中文
            "--log-level=3",    # 浏览器日志等级为error
            f"--force-device-scale-factor={float(self.cfg.browser_scale_factor)}",  # 设置缩放
            f"--app={self.GAME_URL}",   # 以应用模式启动
            "--disable-blink-features=AutomationControlled",  # 去除自动化痕迹，防止被人机验证
            f"--remote-debugging-port={self.cfg.browser_debug_port}",   # 调试端口，可用于复用浏览器
        ]
        if self.cfg.browser_persistent_enable:
            args += [
                f"--user-data-dir={self.user_profile_path}",   # UserProfile 路径
                "--profile-directory=Default",            # UserProfile 名称
            ]
        if headless:
            args += [
                "--headless=new",  # 无窗口模式
                "--mute-audio",    # 后台静音
            ]
            if is_docker_started():
                # Docker 环境下需要额外参数
                args.append("--no-sandbox")
        if self.cfg.cloud_game_fullscreen_enable and not headless:
            args.append("--start-fullscreen")  # 全屏启动
        args.extend(self.cfg.browser_launch_argument)  # 用户自定义参数
        return args

    def _connect_or_create_browser(self, headless=False) -> None:
        """尝试连接到现有的（由小助手启动的）浏览器，如果没有，那就创建一个"""
        browser_type = "chrome" if self.cfg.browser_type in ["integrated", "chrome"] else "edge" if self.cfg.browser_type == "edge" else "chromium"
        integrated = self.cfg.browser_type == "integrated"
        first_run = False
        browser_path, driver_path = self._prepare_browser_and_driver(browser_type, integrated)

        if not os.path.exists(self.user_profile_path):
            first_run = True

        if browser_type == "chrome":
            options = ChromeOptions()
            service = ChromeService(executable_path=driver_path, log_path=os.devnull)
            webdriver_type = webdriver.Chrome
        elif browser_type == "edge":
            options = EdgeOptions()
            service = EdgeService(executable_path=driver_path, log_path=os.devnull)
            webdriver_type = webdriver.Edge
        else:  # chromium
            options = ChromiumOptions()
            service = ChromiumService(executable_path=driver_path, log_path=os.devnull)
            webdriver_type = webdriver.Chrome
        # 记录 driver 可执行路径和 service，以便后续清理 chromedriver 进程
        self.driver_path = driver_path
        self._webdriver_service = service

        # 关掉 headless 不匹配的浏览器，防止端口冲突
        if self.close_all_m7a_browser(headless=not headless):
            self.log_info(f"已关闭正在运行的{'前台' if headless else '后台'}浏览器")
        if self.get_m7a_browsers(headless=headless):
            # 如果发现已经有浏览器，尝试直接连接
            try:
                options.debugger_address = f"127.0.0.1:{self.cfg.browser_debug_port}"
                self.driver = webdriver_type(service=service, options=options)
                self.log_info("已连接到现有浏览器")
                return  # 连接成功，直接返回
            except Exception:
                self.log_info(f"连接现有浏览器失败")
                self.close_all_m7a_browser()  # 连接失败，关闭所有浏览器
                options = None

        self.log_info(f"正在启动 {browser_type} 浏览器")
        options.binary_location = browser_path
        options.add_experimental_option("prefs", self.PERFERENCES)  # 允许云游戏权限权限

        self.log_debug(f"启动参数: {self._get_browser_arguments(headless=headless)}")
        # 设置浏览器启动参数
        for arg in self._get_browser_arguments(headless=headless):
            options.add_argument(arg)

        # 清理失效的断链 (Broken Symlinks) 防止浏览器无法启动
        if is_docker_started():
            singleton_files = ["SingletonCookie", "SingletonLock", "SingletonSocket"]
            for filename in singleton_files:
                file_path = os.path.join(self.user_profile_path, filename)
                try:
                    # 逻辑：是一个链接，但指向的目标不存在
                    if os.path.islink(file_path) and not os.path.exists(file_path):
                        os.remove(file_path)
                        self.log_debug(f"已清理断开的软链接: {file_path}")
                except Exception as e:
                    self.log_warning(f"处理残留链接失败: {file_path}, 错误: {e}")

        try:
            self.log_debug("启动浏览器中...")
            self.driver = webdriver_type(service=service, options=options)
            self.log_debug("浏览器启动成功")
        except SessionNotCreatedException as e:
            self.log_error(f"浏览器启动失败: {e}")
            # 清理残留文件，防止浏览器无法启动
            if is_docker_started():
                singleton_files = ["SingletonCookie", "SingletonLock", "SingletonSocket"]
                for filename in singleton_files:
                    file_path = os.path.join(self.user_profile_path, filename)
                    try:
                        if os.path.lexists(file_path):
                            os.remove(file_path)
                            self.log_debug(f"已删除残留文件: {file_path}")
                    except Exception as e:
                        self.log_warning(f"删除残留文件失败: {file_path}, 错误: {e}")
            self.log_error("如果设置了浏览器启动参数，请去掉所有浏览器启动参数后重试")
            self.log_error("如果仍然存在问题，请更换浏览器重试")
            raise Exception("浏览器启动失败")
        except Exception as e:
            self.log_error(f"浏览器启动失败: {e}")
            raise Exception("浏览器启动失败")

        if not self.cfg.cloud_game_fullscreen_enable:
            self.driver.set_window_size(1920, 1120)

    def _restart_browser(self, headless=False) -> None:
        """重启浏览器"""
        self.stop_game()
        self._connect_or_create_browser(headless=headless)

    def _clean_at_exit(self) -> None:
        """当脚本退出时，关闭所有 headless 浏览器"""
        if self.close_all_m7a_browser(headless=True):
            self.log_info("已关闭所有后台浏览器")

    def download_intergrated_browser(self) -> bool:
        self._prepare_browser_and_driver(browser_type="chrome", integrated=True)

    def is_integrated_browser_downloaded(self) -> bool:
        """当前是否已经下载内置浏览器"""
        return os.path.exists(self._get_integrated_browser_path()) and os.path.exists(self._get_integrated_driver_path())

    def get_m7a_browsers(self, headless=None) -> list[psutil.Process]:
        """
        获取由小助手打开的浏览器
        headless: None 所有，True 仅 headless 无窗口浏览器，False 仅有窗口浏览器

        return 浏览器的 Process
        """
        browsers: list[psutil.Process] = []

        browser_names = {'chrome.exe', 'msedge.exe'}
        browser_tag = self.BROWSER_TAG

        for proc in psutil.process_iter(['pid', 'name']):
            name = proc.info.get('name')
            if name not in browser_names:
                continue

            try:
                cmdline = proc.cmdline()
            except psutil.Error:
                continue

            if browser_tag not in cmdline:
                continue

            if headless is not None:
                is_headless = "--headless=new" in cmdline
                if headless != is_headless:
                    continue

            browsers.append(proc)

        return browsers

    def close_all_m7a_browser(self, headless=None) -> list[psutil.Process]:
        """
        关闭所有由小助手打开的浏览器
        headless: None 所有，True 仅 headless 无窗口浏览器，False 仅 headful 有窗口浏览器

        return 被关闭浏览器的 Process
        """
        closed_proc = []
        for proc in self.get_m7a_browsers(headless=headless) or []:
            try:
                proc.terminate()
                closed_proc.append(proc)
            except Exception:
                pass

        # 也尝试关闭与当前 driver_path 对应的 chromedriver 进程
        closed = self._terminate_chromedriver_processes()
        if closed:
            closed_proc.extend(closed)
        return closed_proc

    def _terminate_chromedriver_processes(self) -> list[psutil.Process]:
        """单独清理 chromedriver 进程并返回被关闭的进程列表"""
        closed: list[psutil.Process] = []

        try:
            chromedrivers: list[psutil.Process] = []

            # 只获取轻量字段，避免 ppid / exe 带来的性能问题
            for proc in psutil.process_iter(['pid', 'name']):
                name = proc.info.get('name')
                if name and name.lower() == 'chromedriver.exe':
                    chromedrivers.append(proc)

            current_pid = os.getpid()
            driver_path_norm = None
            if chromedrivers:
                if hasattr(self, 'driver_path'):
                    driver_path = self.driver_path
                else:
                    driver_path = None
                driver_path_norm = os.path.normcase(driver_path) if driver_path else None

            for proc in chromedrivers:
                try:
                    # 优先通过 exe 路径精确匹配
                    if driver_path_norm:
                        try:
                            exe_path = proc.exe()
                        except psutil.Error:
                            exe_path = None

                        if exe_path and os.path.normcase(exe_path) == driver_path_norm:
                            proc.terminate()
                            closed.append(proc)
                            # 已通过 exe 路径精确匹配并终止进程，无需再执行后续的父进程检查
                            continue

                    # 否则仅终止父进程是当前进程的 chromedriver
                    try:
                        if proc.ppid() == current_pid:
                            proc.terminate()
                            closed.append(proc)
                    except psutil.Error:
                        pass

                except psutil.Error:
                    continue
        except psutil.Error as e:
            self.log_error(f"清理 chromedriver 进程时发生 psutil 错误：{e}")

        return closed

    def try_dump_page(self, dump_dir="logs/webdump") -> None:
        if self.driver:
            os.makedirs(dump_dir, exist_ok=True)
            from datetime import datetime
            ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

            png_path = os.path.join(dump_dir, f"{ts}.png")
            self.driver.save_screenshot(png_path)

            html_path = os.path.join(dump_dir, f"{ts}.html")
            with open(html_path, "w", encoding="utf-8") as f:
                f.write(self.driver.page_source)

            self.log_error(f"相关页面和截图已经保存到：{dump_dir}")

    def start_game_process(self, headless=None) -> bool:
        """启动浏览器进程"""
        try:
            if headless is None:
                headless = self.cfg.browser_headless_enable
            self._connect_or_create_browser(headless=headless)
            self._confirm_viewport_resolution()
            return True
        except Exception as e:
            self.log_error(f"启动或连接浏览器失败 {e}")
            return False

    def is_in_game(self) -> bool:
        """这里无法判断是否在云游戏内，如果返回 False 会导致小助手开始检查云游戏语言设置，而这个功能在 Moonlight Web 场景下不存在，导致卡死"""
        try:
            self._wait_page_loaded()
            self._wait_page_ready()
            self._click_lock_mouse_button()
            self._confirm_viewport_resolution()
            self.log_info("进入云游戏成功（Moonlight Web）")
            return True
        except Exception as e:
            self.try_dump_page()
            self.log_error(f"进入云游戏失败: {e}")
            return False

    def _wait_page_ready(self, timeout=30) -> None:
        """等待 Moonlight Web 串流界面就绪（以 sidebar-button 出现为标志）"""
        if not self.driver:
            return
        self.log_info("等待串流界面就绪...")
        try:
            WebDriverWait(self.driver, timeout).until(
                lambda d: d.execute_script(
                    "return document.getElementById('sidebar-button') !== null"
                )
            )
            self.log_info("串流界面已就绪（sidebar-button 已出现）")
        except TimeoutException:
            raise Exception(f"串流界面在 {timeout}s 内未就绪（sidebar-button 未出现）")

    def enter_cloud_game(self) -> bool:
        """进入云游戏（Moonlight Web 直连，无需登录/排队）"""
        return True

    def _click_lock_mouse_button(self, timeout=10) -> None:
        """先打开侧边菜单，再通过 CDP 模拟真实鼠标点击 Lock Mouse 按钮以激活 pointer lock"""
        if not self.driver:
            return
        try:
            import time
            from selenium.webdriver.support.ui import WebDriverWait

            time.sleep(2)  # 等待页面稳定，避免元素尚未渲染导致的点击失败

            # 1. 点击侧边栏按钮打开菜单
            sidebar_rect = WebDriverWait(self.driver, timeout).until(
                lambda d: d.execute_script("""
                    var btn = document.getElementById('sidebar-button');
                    if (!btn) return null;
                    var r = btn.getBoundingClientRect();
                    return {x: r.x + r.width / 2, y: r.y + r.height / 2};
                """)
            )
            sx, sy = sidebar_rect["x"], sidebar_rect["y"]
            self.driver.execute_cdp_cmd("Input.dispatchMouseEvent", {
                "type": "mousePressed", "x": sx, "y": sy,
                "button": "left", "buttons": 1, "clickCount": 1, "pointerType": "mouse"
            })
            self.driver.execute_cdp_cmd("Input.dispatchMouseEvent", {
                "type": "mouseReleased", "x": sx, "y": sy,
                "button": "left", "buttons": 0, "clickCount": 1, "pointerType": "mouse"
            })
            time.sleep(0.5)

            # 2. 将 Mouse Mode 切换为 Point and Drag
            WebDriverWait(self.driver, timeout).until(
                lambda d: d.execute_script(
                    "return document.getElementById('mouseMode') !== null"
                )
            )
            self.driver.execute_script("""
                var sel = document.getElementById('mouseMode');
                // 使用原生 setter 以兼容 React 等框架
                var nativeSetter = Object.getOwnPropertyDescriptor(
                    HTMLSelectElement.prototype, 'value'
                ).set;
                nativeSetter.call(sel, 'pointAndDrag');
                sel.dispatchEvent(new Event('input', {bubbles: true}));
                sel.dispatchEvent(new Event('change', {bubbles: true}));
            """)
            self.log_info("已将 Mouse Mode 切换为 Point and Drag")
            time.sleep(0.3)

            # 3. 等待 Lock Mouse 按钮在可视区域内出现
            lock_rect = WebDriverWait(self.driver, timeout).until(
                lambda d: d.execute_script("""
                    var buttons = document.querySelectorAll('button');
                    var vw = window.innerWidth, vh = window.innerHeight;
                    for (var btn of buttons) {
                        if (btn.innerText.trim() === 'Lock Mouse') {
                            var r = btn.getBoundingClientRect();
                            if (r.width > 0 && r.height > 0
                                && r.left >= 0 && r.top >= 0
                                && r.right <= vw && r.bottom <= vh)
                                return {x: r.x + r.width / 2, y: r.y + r.height / 2};
                        }
                    }
                    return null;
                """)
            )

            # 4. CDP 点击 Lock Mouse 按钮
            lx, ly = lock_rect["x"], lock_rect["y"]
            self.driver.execute_cdp_cmd("Input.dispatchMouseEvent", {
                "type": "mousePressed", "x": lx, "y": ly,
                "button": "left", "buttons": 1, "clickCount": 1, "pointerType": "mouse"
            })
            self.driver.execute_cdp_cmd("Input.dispatchMouseEvent", {
                "type": "mouseReleased", "x": lx, "y": ly,
                "button": "left", "buttons": 0, "clickCount": 1, "pointerType": "mouse"
            })
            time.sleep(0.5)
            self.log_info("已点击 Lock Mouse 按钮")
        except Exception as e:
            self.log_warning(f"点击 Lock Mouse 按钮失败: {e}")

    def take_screenshot(self) -> bytes:
        """浏览器内截图"""
        if not self.driver:
            return None
        # 仅在 macOS 非 headless 模式下使用 CDP 截图，避免浏览器被切换到前台
        if not self.cfg.browser_headless_enable and platform.system() == "Darwin":
            # Chrome/Chromium 在非 headless 模式下调用 get_screenshot_as_png() 时，
            # 会先确保窗口“可见且未被遮挡”，否则截图内容可能为空或全黑。
            # macOS 的窗口管理要求被截取的 NSWindow 处于前台/可见状态，
            # Chromium 的实现会自动把窗口置前。
            # 改用 CDP 截图接口可以避免这个问题。
            try:
                result = self.driver.execute_cdp_cmd("Page.captureScreenshot", {"format": "png"})
                data = result.get("data") if result else None
                if data:
                    return base64.b64decode(data)
            except Exception as e:
                self.log_warning(f"CDP 截图失败，回退 WebDriver 截图: {e}")
        return self.driver.get_screenshot_as_png()

    def execute_cdp_cmd(self, cmd: str, cmd_args: dict):
        return self.driver.execute_cdp_cmd(cmd, cmd_args)

    def get_window_handle(self) -> int:
        if sys.platform != "win32":
            self.log_warning("当前平台不支持获取云游戏窗口句柄，将返回 None")
            return None
        import win32gui
        return win32gui.FindWindow(None, "云·星穹铁道")

    def switch_to_game(self) -> bool:
        if self.cfg.browser_headless_enable:
            self.log_warning("游戏切换至前台失败：当前为无窗口模式")
            return False
        else:
            return super().switch_to_game()

    def get_input_handler(self):
        from module.automation.cdp_input import CdpInput
        return CdpInput(cloud_game=self, logger=self.logger)

    def copy(self, text):
        self.driver.execute_script("""
            (function copy(text) {
                const ta = document.createElement('textarea');
                ta.value = text;
                ta.style.position = 'fixed';
                ta.style.opacity = '0';
                document.body.appendChild(ta);

                ta.focus();
                ta.select();
                document.execCommand('copy');

                document.body.removeChild(ta);
            })(arguments[0]);
        """, text)

    def _cancel_sunshine_session(self) -> None:
        """通过浏览器内 fetch 通知 Sunshine 关闭当前串流 session（携带浏览器 cookies）"""
        if not self.driver:
            self.log_warning("浏览器未启动，跳过关闭 Sunshine session")
            return
        cancel_url = f"{self.MOONLIGHT_BASE_URL}/api/host/cancel"
        try:
            result = self.driver.execute_script("""
                var resp = await fetch(arguments[0], {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({host_id: arguments[1]}),
                    credentials: 'include'
                });
                return {status: resp.status, text: await resp.text()};
            """, cancel_url, int(self.MOONLIGHT_HOST_ID))
            if result and result.get("status") == 200:
                self.log_info("已通知 Sunshine 关闭串流 session")
            else:
                self.log_warning(f"关闭 Sunshine session 失败: HTTP {result}")
        except Exception as e:
            self.log_warning(f"关闭 Sunshine session 请求异常: {e}")

    def stop_game(self) -> bool:
        """退出游戏，关闭浏览器，并关闭 Sunshine session"""
        # 先关闭 Sunshine 串流 session
        self._cancel_sunshine_session()

        if self.driver:
            try:
                self.driver.execute(Command.CLOSE)
                self.log_info("关闭浏览器成功")
            except Exception:
                pass
            self.driver.quit()
            self.driver = None

        # 清理所有未正常退出的浏览器
        try:
            if self.close_all_m7a_browser():
                self.log_info("检测到由小助手启动的浏览器，已成功关闭")
        except Exception as e:
            self.log_warning(f"检测到由小助手启动的浏览器，关闭失败: {e}")
            return False

        return True

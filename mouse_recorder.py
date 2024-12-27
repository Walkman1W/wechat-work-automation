import pyautogui
import keyboard
import json
import time
from typing import List, Dict
import logging
from datetime import datetime
import os
from PIL import Image

class MouseRecorder:
    def __init__(self, logger=None):
        self.coordinates = []
        self._setup_logging(logger)
        # 创建templates文件夹
        if not os.path.exists('templates'):
            os.makedirs('templates')
        
    def _setup_logging(self, logger):
        """设置日志"""
        self.logger = logger or logging.getLogger('mouse_recorder')
        
    def _capture_template(self, x: int, y: int, step: int) -> str:
        """捕获点击位置周围的模板图片
        Args:
            x: 点击位置的x坐标
            y: 点击位置的y坐标
            step: 步骤编号
        Returns:
            模板图片的文件路径
        """
        try:
            # 计算截图区域，确保不超出屏幕范围
            left = max(0, x - 40)  # 向左40像素
            top = max(0, y - 25)   # 向上25像素
            width = 80   # 总宽度80像素
            height = 50  # 总高度50像素
            
            # 截取图片
            screenshot = pyautogui.screenshot(region=(left, top, width, height))
            
            # 保存模板图片
            template_path = f'templates/step{step}_template.png'
            screenshot.save(template_path)
            
            self.logger.info(f"已保存模板图片: {template_path}")
            return template_path
            
        except Exception as e:
            self.logger.error(f"捕获模板图片失败: {e}")
            return ""
        
    def record(self, total_steps: int = None) -> List[Dict]:
        """记录鼠标坐标
        Args:
            total_steps: 需要记录的总步骤数，达到后自动完成
        Returns:
            记录的坐标列表
        """
        self.logger.info(f"开始记录坐标，计划记录{total_steps if total_steps else '不限'}个坐标点")
        print("开始记录坐标，按住Capslock并点击鼠标左键记录位置，按Ctrl+C结束")
        if total_steps:
            print(f"将在记录{total_steps}个坐标后自动完成")
        
        step = 1
        capslock_was_pressed = False
        mouse_clicked = False
        self.coordinates = []  # 清空之前的记录
        
        try:
            while True:
                try:
                    # 获取当前鼠标位置并打印
                    current_x, current_y = pyautogui.position()
                    print(f"\r当前鼠标位置: ({current_x}, {current_y})", end='')
                    
                    # 检测Capslock和鼠标左键状态
                    if keyboard.is_pressed('capslock'):
                        if not capslock_was_pressed:  # 只在第一次按下时记录状态
                            capslock_was_pressed = True
                    else:
                        capslock_was_pressed = False
                        mouse_clicked = False
                        
                    # 当Capslock按下时，检测鼠标点击
                    if capslock_was_pressed and not mouse_clicked:
                        try:
                            import win32api
                            import win32con
                            if win32api.GetAsyncKeyState(win32con.VK_LBUTTON) < 0:  # 鼠标左键被按下
                                mouse_clicked = True
                                x, y = current_x, current_y
                                
                                # 捕获模板图片
                                template_path = self._capture_template(x, y, step)
                                
                                # 记录坐标和模板路径
                                coord = {
                                    f"step{step}": {
                                        "x": x,
                                        "y": y,
                                        "template": template_path
                                    }
                                }
                                self.coordinates.append(coord)
                                print(f"\n记录位置 step{step}: ({x}, {y})")
                                print(f"当前已记录的所有坐标: {self.coordinates}")
                                self.logger.info(f"记录坐标点 step{step}: ({x}, {y})")
                                step += 1
                                time.sleep(0.5)
                                
                                # 检查是否达到指定步骤数
                                if total_steps and step > total_steps:
                                    print(f"\n已达到指定的{total_steps}个坐标点")
                                    break
                        except ImportError:
                            print("\n需要安装pywin32库来检测鼠标点击")
                            print("请运行: pip install pywin32")
                            break
                    
                    if keyboard.is_pressed('ctrl') and keyboard.is_pressed('c'):
                        break
                    
                    # 添加短暂延迟，减少CPU使用
                    time.sleep(0.1)
                    
                except Exception as e:
                    if not isinstance(e, KeyboardInterrupt):
                        print(f"\n记录过程出错: {e}")
                        self.logger.error(f"记录过程出错: {e}")
                        print(f"错误类型: {type(e)}")
                    break
                    
        except Exception as e:
            self.logger.error(f"记录坐标时出错: {e}")
            print(f"错误类型: {type(e)}")
            
        # 保存并返回结果
        if self.coordinates:
            print("\n记录完成")
            print(f"共记录 {len(self.coordinates)} 个坐标点:")
            for coord in self.coordinates:
                print(coord)
            self.logger.info(f"完成记录，共记录 {len(self.coordinates)} 个坐标点")
        else:
            print("\n没有记录到任何坐标")
            self.logger.warning("没有记录到任何坐标")
            
        return self.coordinates
    
    def save_to_file(self, filename: str = 'coordinates.json'):
        """保存坐标到JSON文件"""
        try:
            with open(filename, 'w') as f:
                json.dump(self.coordinates, f, indent=2)
            self.logger.info(f"坐标已保存到文件: {filename}")
            return True
        except Exception as e:
            self.logger.error(f"保存坐标文件失败: {e}")
            return False
    
    def load_from_file(self, filename: str = 'coordinates.json') -> List[Dict]:
        """从JSON文件加载坐标"""
        try:
            with open(filename, 'r') as f:
                self.coordinates = json.load(f)
            self.logger.info(f"从文件加载了 {len(self.coordinates)} 个坐标点")
            return self.coordinates
        except FileNotFoundError:
            self.logger.warning(f"未找到坐标文件: {filename}")
            return []
        except Exception as e:
            self.logger.error(f"加载坐标文件失败: {e}")
            return []

# 测试代码
if __name__ == "__main__":
    # 设置日志
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )
    logger = logging.getLogger('mouse_recorder_test')
    
    # 创建MouseRecorder实例
    recorder = MouseRecorder(logger)
    
    print("开始测试坐标记录功能...")
    print("请输入需要记录的坐标数量（直接回车则不限制数量）: ")
    steps = input().strip()
    total_steps = int(steps) if steps else None
    
    # 记录坐标
    coordinates = recorder.record(total_steps)
    
    # 如果记录到了坐标，保存到文件
    if coordinates:
        recorder.save_to_file()
        print("\n坐标已保存到coordinates.json文件") 
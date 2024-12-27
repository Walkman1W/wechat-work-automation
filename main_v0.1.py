import pyautogui
import keyboard
import time
import pandas as pd
import pyperclip
from datetime import datetime
import logging
from typing import List, Dict
import os
from mouse_recorder import MouseRecorder

class MouseAutomation:
    def __init__(self):
        self.running = False
        self.paused = False
        self._setup_logging()
        self.mouse_recorder = MouseRecorder(self.logger)
        
    def _setup_logging(self):
        """设置日志"""
        # 创建logger
        self.logger = logging.getLogger('mouse_automation')
        self.logger.setLevel(logging.INFO)
        
        # 创建按日期命名的日志文件
        log_filename = f'automation_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log'
        file_handler = logging.FileHandler(log_filename, encoding='utf-8')
        
        # 设置日志格式
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        file_handler.setFormatter(formatter)
        
        # 添加处理器
        self.logger.addHandler(file_handler)
        
        self.logger.info("程序启动")

    def record_coordinates(self, total_steps: int = None):
        """记录鼠标坐标的模块"""
        coordinates = self.mouse_recorder.record(total_steps)
        if coordinates:
            self.mouse_recorder.save_to_file()
        return True

    def _is_valid_phone(self, phone: str) -> bool:
        """验证手机号是否合法"""
        return len(str(phone)) == 11 and str(phone).isdigit()

    def automate_process(self):
        """自动化处理模块"""
        print("\n开始自动化处理...")
        self.logger.info("开始自动化处理")
        # 加载Excel数据
        try:
            print("正在加载Excel文件...")
            df = pd.read_excel('phone.xlsx')
            print(f"成功加载Excel文件，共 {len(df)} 条记录")
            self.logger.info(f"成功加载Excel文件，共 {len(df)} 条记录")
        except Exception as e:
            error_msg = f"加载Excel文件失败: {e}"
            print(error_msg)
            self.logger.error(error_msg)
            return True

        print("正在加载坐标文件...")
        coordinates = self.mouse_recorder.load_from_file()
        
        if not coordinates:
            self.logger.warning("未找到坐标文件或坐标文件为空")
            print("未找到坐标文件或坐标文件为空，请先记录坐标")
            return True
            
        # 检查坐标点数量是否足够
        if len(coordinates) < 4:  # 需要4个坐标点
            error_msg = f"坐标点数量不足，需要4个坐标点，当前只有{len(coordinates)}个"
            print(error_msg)
            self.logger.error(error_msg)
            return True
        
        print(f"成功加载坐标文件，共 {len(coordinates)} 个坐标点")
        self.logger.info(f"成功加载坐标文件，共 {len(coordinates)} 个坐标点")
        
        # 注册快捷键
        keyboard.add_hotkey('ctrl+f1', self._toggle_pause)
        keyboard.add_hotkey('ctrl+f2', self._stop)
        
        self.running = True
        print("\n开始自动化处理，按Ctrl+F1暂停/继续，按Ctrl+F2结束")
        print("=" * 50)

        try:
            for index, row in df.iterrows():
                if not self.running:
                    print("\n检测到停止信号，结束处理")
                    self.logger.info("检测到停止信号，结束处理")
                    break

                # 如果状态不为空，跳过
                if pd.notna(row['状态']):
                    print(f"跳过已处理的记录: {row['手机号']}")
                    self.logger.info(f"跳过已处理的记录: {row['手机号']}")
                    continue

                # 验证手机号
                if not self._is_valid_phone(row['手机号']):
                    print(f"无效的手机号: {row['手机号']}")
                    self.logger.warning(f"无效的手机号: {row['手机号']}")
                    df.at[index, '状态'] = '无效手机号'
                    continue

                while self.paused:
                    time.sleep(0.1)

                try:
                    print(f"\n正在处理第 {index + 1} 条记录，手机号: {row['手机号']}")
                    self.logger.info(f"开始处理第 {index + 1} 条记录，手机号: {row['手机号']}")
                    
                    # Step 1: 点击添加
                    print("步骤1: 点击添加按钮")
                    self.logger.debug("执行步骤1: 点击添加按钮")
                    pyautogui.click(x=coordinates[0]['step1']['x'], y=coordinates[0]['step1']['y'])
                    time.sleep(2)

                    # Step 2: 点击输入框并输入手机号
                    print("步骤2: 点击输入框并输入手机号")
                    self.logger.debug("执行步骤2: 点击输入框并输入手机号")
                    pyautogui.click(x=coordinates[1]['step2']['x'], y=coordinates[1]['step2']['y'])
                    time.sleep(1)
                    pyperclip.copy(str(row['手机号']))
                    pyautogui.hotkey('ctrl', 'v')
                    time.sleep(1)
                    pyautogui.press('enter')
                    time.sleep(2)

                    # Step 3: 点击添加按钮
                    print("步骤3: 点击添加按钮")
                    self.logger.debug("执行步骤3: 点击添加按钮")
                    pyautogui.click(x=coordinates[2]['step3']['x'], y=coordinates[2]['step3']['y'])
                    time.sleep(2)

                    # Step 4: 点击发送邀请（测试模式）
                    print("步骤4: [测试模式] 模拟点击发送邀请")
                    self.logger.debug("执行步骤4: [测试模式] 模拟点击发送邀请")
                    pyautogui.click(x=coordinates[3]['step4']['x'], y=coordinates[3]['step4']['y'])  # 正式环境取消注释
                    print(f"[测试] 手机号 {row['手机号']} 处理完成，坐标: ({coordinates[3]['step4']['x']}, {coordinates[3]['step4']['y']})")
                    self.logger.info(f"[测试] 手机号 {row['手机号']} 处理完成")
                    df.at[index, '状态'] = '测试完成'
                    time.sleep(2)

                except Exception as e:
                    error_msg = f"处理手机号 {row['手机号']} 时出错: {e}"
                    print(error_msg)
                    self.logger.error(error_msg)
                    df.at[index, '状态'] = f'错误: {str(e)}'

                # 保存进度
                df.to_excel('phone.xlsx', index=False)
                
                # 设置列宽
                try:
                    from openpyxl import load_workbook
                    wb = load_workbook('phone.xlsx')
                    ws = wb.active
                    # 设置列宽：第1,2列为9，第3,4列为15
                    ws.column_dimensions['A'].width = 9
                    ws.column_dimensions['B'].width = 9
                    ws.column_dimensions['C'].width = 15
                    ws.column_dimensions['D'].width = 15
                    wb.save('phone.xlsx')
                except Exception as e:
                    self.logger.warning(f"设置Excel列宽失败: {e}")
                
                print("进度已保存到Excel文件")
                self.logger.info("进度已保存到Excel文件")
                print("-" * 50)

        except Exception as e:
            error_msg = f"自动化处理出错: {e}"
            print(error_msg)
            self.logger.error(error_msg)
        finally:
            # 清理快捷键
            keyboard.unhook_all()
            print("\n自动化处理完成")
            self.logger.info("自动化处理完成")
            print("=" * 50)
            return True

    def _toggle_pause(self):
        """暂停/继续自动化处理"""
        self.paused = not self.paused
        status = "已暂停" if self.paused else "继续运行"
        print(status)
        self.logger.info(status)

    def _stop(self):
        """停止自动化处理"""
        self.running = False
        self.logger.info("程序已停止")
        print("程序已停止")

def main():
    automation = MouseAutomation()
    while True:
        print("\n请选择模式:")
        print("1: 记录坐标")
        print("2: 自动化处理")
        print("3: 退出程序")
        
        mode = input("请输入选择 (1/2/3): ")
        automation.logger.info(f"用户选择模式: {mode}")
        
        if mode == "1":
            steps = input("请输入需要记录的坐标数量（直接回车则不限制数量）: ")
            total_steps = int(steps) if steps.strip() else None
            automation.record_coordinates(total_steps)
        elif mode == "2":
            automation.automate_process()
        elif mode == "3":
            automation.logger.info("程序退出")
            print("程序已退出")
            break
        else:
            print("无效的选择，请重新输入")
            automation.logger.warning(f"无效的模式选择: {mode}")

if __name__ == "__main__":
    main() 
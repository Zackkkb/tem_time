import os
import re
import matplotlib

matplotlib.use('Agg')  # 使用非交互式后端
import matplotlib.pyplot as plt
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

# 配置中文显示
plt.rcParams["font.family"] = ["SimHei", "Arial Unicode MS"]
plt.rcParams["axes.unicode_minus"] = False  # 正确显示负号
plt.rcParams["font.size"] = 9  # 全局字体大小


def create_directories():
    """创建数据和图表文件夹（如果不存在）"""
    try:
        if not os.path.exists('data/data'):
            os.makedirs('data/data')
        if not os.path.exists('data/charts'):
            os.makedirs('data/charts')
        return True
    except Exception as e:
        print(f"创建文件夹失败: {e}")
        return False


def validate_filename(filename):
    """验证并清理文件名"""
    invalid_chars = r'[\\/*?:"<>|]'
    cleaned = re.sub(invalid_chars, '', filename)
    if not cleaned:
        return "temp_profile"
    return cleaned[:50]


def get_user_input():
    """获取用户输入的参数（时间单位改为小时）"""
    params = {}
    print("请输入以下参数（输入后按回车确认）：")
    try:
        # 温度参数保持不变
        params['initial_temp'] = float(input("初始温度（℃）: "))
        params['initial_time'] = float(input("初始温度持续时间（h）: "))  # 改为小时
        params['recovery_temp'] = float(input("回温温度（℃）: "))
        params['recovery_time'] = float(input("回温温度持续时间（h）: "))  # 改为小时
        params['high_temp'] = float(input("高温温度（℃）: "))
        params['high_tolerance'] = float(input("高温允差（℃）: "))
        params['low_temp'] = float(input("低温温度（℃）: "))
        params['low_tolerance'] = float(input("低温允差（℃）: "))

        # 时间参数全部改为小时
        params['first_high_time'] = float(input("首循环高温保持时间（h）: "))
        params['first_low_time'] = float(input("首循环低温保持时间（h）: "))
        params['last_high_time'] = float(input("末循环高温保持时间（h）: "))
        params['last_low_time'] = float(input("末循环低温保持时间（h）: "))
        params['middle_high_time'] = float(input("中间循环高温保持时间（h）: "))
        params['middle_low_time'] = float(input("中间循环低温保持时间（h）: "))

        # 速率单位从℃/min转换为℃/h（乘以60）
        heat_rate_per_min = float(input("升温速率（℃/min）: "))
        params['heat_rate'] = heat_rate_per_min * 60  # 转换为℃/h
        cool_rate_per_min = float(input("降温速率（℃/min）: "))
        params['cool_rate'] = cool_rate_per_min * 60  # 转换为℃/h

        params['cycles'] = int(input("循环次数（次）: "))

        # 验证输入
        for key, value in params.items():
            if value < 0 and key not in ['initial_temp', 'recovery_temp', 'high_temp', 'low_temp']:
                raise ValueError(f"{key} 不能为负数")
        return params
    except ValueError as e:
        print(f"输入错误: {e}，请重新运行程序")
        return None


def calculate_temperature_profile(params):
    """计算温度曲线数据（内部计算使用小时单位）"""
    try:
        profile = []
        key_points = []
        current_time = 0.0  # 时间单位：小时
        current_temp = params['initial_temp']

        # 1. 初始温度阶段
        profile.append((current_time, current_temp, "初始温度开始"))
        key_points.append(len(profile) - 1)
        current_time += params['initial_time']  # 直接累加小时
        profile.append((current_time, current_temp, "初始温度结束"))
        key_points.append(len(profile) - 1)

        # 2. 循环过程
        for cycle in range(params['cycles']):
            if cycle == 0:
                high_time = params['first_high_time']
                low_time = params['first_low_time']
                cycle_type = "首循环"
            elif cycle == params['cycles'] - 1:
                high_time = params['last_high_time']
                low_time = params['last_low_time']
                cycle_type = "末循环"
            else:
                high_time = params['middle_high_time']
                low_time = params['middle_low_time']
                cycle_type = f"中间循环{cycle}"

            # 升温到高温（速率单位为℃/h，直接计算）
            temp_diff = params['high_temp'] - current_temp
            if temp_diff != 0:
                # 时间变化 = 温差 / 速率（已转换为℃/h）
                time_change = abs(temp_diff) / (params['heat_rate'] if temp_diff > 0 else params['cool_rate'])
                current_time += time_change
                profile.append((current_time, params['high_temp'], f"{cycle_type}高温达到"))
                key_points.append(len(profile) - 1)
                current_temp = params['high_temp']

            # 高温保持（小时单位直接累加）
            current_time += high_time
            profile.append((current_time, current_temp, f"{cycle_type}高温结束"))
            key_points.append(len(profile) - 1)

            # 降温到低温（非最后一轮）
            if cycle < params['cycles'] - 1:
                temp_diff = params['low_temp'] - current_temp
                if temp_diff != 0:
                    time_change = abs(temp_diff) / (params['heat_rate'] if temp_diff > 0 else params['cool_rate'])
                    current_time += time_change
                    profile.append((current_time, params['low_temp'], f"{cycle_type}低温达到"))
                    key_points.append(len(profile) - 1)
                    current_temp = params['low_temp']

                # 低温保持
                current_time += low_time
                profile.append((current_time, current_temp, f"{cycle_type}低温结束"))
                key_points.append(len(profile) - 1)

        # 3. 回温阶段
        temp_diff = params['recovery_temp'] - current_temp
        if temp_diff != 0:
            time_change = abs(temp_diff) / (params['heat_rate'] if temp_diff > 0 else params['cool_rate'])
            current_time += time_change
            profile.append((current_time, params['recovery_temp'], "回温温度达到"))
            key_points.append(len(profile) - 1)
            current_temp = params['recovery_temp']

        current_time += params['recovery_time']
        profile.append((current_time, current_temp, "回温结束"))
        key_points.append(len(profile) - 1)

        return profile, key_points
    except Exception as e:
        print(f"计算出错: {e}")
        return None, None


def create_excel_file(params, profile, filename):
    """创建Excel数据文件（显示小时单位）"""
    try:
        wb = Workbook()
        ws_params = wb.active
        ws_params.title = "输入参数"

        # 参数表（单位改为小时）
        ws_params['A1'], ws_params['B1'], ws_params['C1'] = "参数名称", "数值", "单位"
        for cell in ['A1', 'B1', 'C1']:
            ws_params[cell].font = Font(bold=True)

        # 注意：速率显示时转换回℃/min
        param_list = [
            ("初始温度", params['initial_temp'], "℃"),
            ("初始温度持续时间", params['initial_time'], "h"),
            ("回温温度", params['recovery_temp'], "℃"),
            ("回温温度持续时间", params['recovery_time'], "h"),
            ("高温温度", params['high_temp'], "℃"),
            ("高温允差", params['high_tolerance'], "℃"),
            ("低温温度", params['low_temp'], "℃"),
            ("低温允差", params['low_tolerance'], "℃"),
            ("首循环高温保持时间", params['first_high_time'], "h"),
            ("首循环低温保持时间", params['first_low_time'], "h"),
            ("末循环高温保持时间", params['last_high_time'], "h"),
            ("末循环低温保持时间", params['last_low_time'], "h"),
            ("中间循环高温保持时间", params['middle_high_time'], "h"),
            ("中间循环低温保持时间", params['middle_low_time'], "h"),
            ("升温速率", params['heat_rate'] / 60, "℃/min"),  # 转换回原单位显示
            ("降温速率", params['cool_rate'] / 60, "℃/min"),  # 转换回原单位显示
            ("循环次数", params['cycles'], "次")
        ]

        for i, (name, value, unit) in enumerate(param_list, start=2):
            ws_params[f'A{i}'] = name
            ws_params[f'B{i}'] = round(value, 2) if isinstance(value, float) else value
            ws_params[f'C{i}'] = unit

        # 数据表（时间单位显示为小时）
        ws_data = wb.create_sheet(title="温度随时间变化")
        ws_data['A1'], ws_data['B1'], ws_data['C1'] = "时间 (h)", "温度 (℃)", "说明"
        for cell in ['A1', 'B1', 'C1']:
            ws_data[cell].font = Font(bold=True)

        for i, (time, temp, desc) in enumerate(profile, start=2):
            ws_data[f'A{i}'] = round(time, 2)
            ws_data[f'B{i}'] = round(temp, 2)
            ws_data[f'C{i}'] = desc

        wb.save(os.path.join('data/data', filename))
        print(f"数据文件已保存至: {os.path.join('data/data', filename)}")
        return True
    except Exception as e:
        print(f"Excel创建失败: {e}")
        return False


def create_chart_png(profile, key_points, filename):
    """生成PNG图表（显示小时单位）"""
    try:
        times = [p[0] for p in profile]
        temps = [p[1] for p in profile]
        descriptions = [p[2] for p in profile]

        # 创建图表
        plt.figure(figsize=(14, 9))

        # 绘制温度曲线
        plt.plot(times, temps, 'b-', linewidth=2, label='温度曲线')

        # 处理关键节点
        key_indices = list(set(key_points))
        key_indices.sort()
        used_positions = []

        # 为每个关键节点添加标记和标注
        for idx in key_indices:
            time = times[idx]
            temp = temps[idx]
            desc = descriptions[idx]

            # 绘制关键节点标记
            plt.scatter(time, temp, color='red', s=80, marker='d', zorder=5)

            # 绘制垂直和水平虚线
            plt.axvline(x=time, color='gray', linestyle='--', linewidth=1, alpha=0.7)
            plt.axhline(y=temp, color='gray', linestyle='--', linewidth=1, alpha=0.7)

            # 动态计算x轴标注位置（小时单位）
            x_pos = time
            y_text_pos = plt.ylim()[0] + (plt.ylim()[1] - plt.ylim()[0]) * 0.02
            label_offset = 0.03 * (plt.xlim()[1] - plt.xlim()[0])

            # 检查重叠
            overlap = any(abs(x_pos - used_x) < label_offset for (used_x, _) in used_positions)
            if overlap:
                y_text_pos = plt.ylim()[0] + (plt.ylim()[1] - plt.ylim()[0]) * 0.06
            else:
                used_positions.append((x_pos, y_text_pos))

            # x轴时间标注（显示小时）
            plt.text(x_pos, y_text_pos, f'{time:.2f}h',
                     horizontalalignment='center',
                     verticalalignment='bottom',
                     color='darkred',
                     bbox=dict(facecolor='white', alpha=0.9, boxstyle='round,pad=0.2'))

            # 动态计算y轴标注位置
            y_pos = temp
            x_text_pos = plt.xlim()[0] + (plt.xlim()[1] - plt.xlim()[0]) * 0.02

            # 检查y轴标注重叠
            y_overlap = any(
                abs(y_pos - used_y) < 0.03 * (plt.ylim()[1] - plt.ylim()[0]) for (_, used_y) in used_positions)
            if y_overlap:
                x_text_pos = plt.xlim()[0] + (plt.xlim()[1] - plt.xlim()[0]) * 0.06
            else:
                used_positions.append((x_text_pos, y_pos))

            # y轴温度标注
            plt.text(x_text_pos, y_pos, f'{temp:.1f}℃',
                     horizontalalignment='left',
                     verticalalignment='center',
                     color='darkgreen',
                     bbox=dict(facecolor='white', alpha=0.9, boxstyle='round,pad=0.2'))

            # 动态调整说明文字位置
            chart_center_x = (plt.xlim()[0] + plt.xlim()[1]) / 2
            chart_center_y = (plt.ylim()[0] + plt.ylim()[1]) / 2

            if time > chart_center_x and temp > chart_center_y:
                text_x, text_y = -10, -10
            elif time < chart_center_x and temp > chart_center_y:
                text_x, text_y = 10, -10
            elif time > chart_center_x and temp < chart_center_y:
                text_x, text_y = -10, 10
            else:
                text_x, text_y = 10, 10


        # 坐标轴设置（小时单位）
        plt.xlim(0, max(times) * 1.15)
        min_temp = min(temps)
        plt.ylim(min(0, min_temp * 1.15), max(temps) * 1.15)

        plt.xlabel('时间 (h)', fontsize=11)  # 改为小时
        plt.ylabel('温度 (℃)', fontsize=11)
        plt.title('温度随时间变化曲线', fontsize=13)
        plt.grid(True, linestyle='--', alpha=0.7)
        plt.legend()

        # 保存图表
        file_path = os.path.join('data/charts', filename)
        plt.savefig(file_path, dpi=250, bbox_inches='tight')
        print(f"图表已保存至: {file_path}")

        plt.close()
        return True
    except Exception as e:
        print(f"图表创建失败: {e}")
        return False


def main():
    print("===== 温度曲线生成器（小时单位） =====")
    if not create_directories():
        return

    params = get_user_input()
    if not params:
        return

    profile, key_points = calculate_temperature_profile(params)
    if not profile or not key_points:
        return

    filename_input = input("请输入文件名（无需扩展名）: ")
    cleaned_filename = validate_filename(filename_input)

    if not create_excel_file(params, profile, f"{cleaned_filename}.xlsx"):
        return

    if not create_chart_png(profile, key_points, f"{cleaned_filename}.png"):
        return

    print("操作完成！")


if __name__ == "__main__":
    main()

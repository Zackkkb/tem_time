import streamlit as st
import os
import re
import matplotlib

matplotlib.use('Agg')  # 非交互式后端，避免显示问题
import matplotlib.pyplot as plt
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from io import BytesIO  # 用于在内存中处理文件

# ==================== 配置与工具函数 ====================
# 配置matplotlib中文显示
plt.rcParams["font.family"] = ["SimHei", "Arial Unicode MS"]
plt.rcParams["axes.unicode_minus"] = False
plt.rcParams["font.size"] = 9


# 创建临时目录
def create_temp_dirs():
    for dir_name in ['data', 'charts']:
        if not os.path.exists(dir_name):
            os.makedirs(dir_name)


# 清理文件名
def validate_filename(filename):
    invalid_chars = r'[\\/*?:"<>|]'
    cleaned = re.sub(invalid_chars, '', filename)
    return cleaned if cleaned else "temp_profile"


# ==================== 核心计算函数 ====================
def calculate_temperature_profile(params):
    """计算温度曲线数据（时间单位：小时）"""
    try:
        profile = []
        key_points = []
        current_time = 0.0  # 小时
        current_temp = params['initial_temp']

        # 初始温度阶段
        profile.append((current_time, current_temp, "初始温度开始"))
        key_points.append(len(profile) - 1)
        current_time += params['initial_time']
        profile.append((current_time, current_temp, "初始温度结束"))
        key_points.append(len(profile) - 1)

        # 循环过程
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

            # 升温到高温
            temp_diff = params['high_temp'] - current_temp
            if temp_diff != 0:
                time_change = abs(temp_diff) / (params['heat_rate'] if temp_diff > 0 else params['cool_rate'])
                current_time += time_change
                profile.append((current_time, params['high_temp'], f"{cycle_type}高温达到"))
                key_points.append(len(profile) - 1)
                current_temp = params['high_temp']

            # 高温保持
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

        # 回温阶段
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
        st.error(f"计算出错: {e}")
        return None, None


# ==================== 文件生成函数 ====================
def create_excel_file(params, profile):
    """在内存中生成Excel文件，返回字节流用于下载"""
    try:
        wb = Workbook()
        ws_params = wb.active
        ws_params.title = "输入参数"

        # 参数表
        ws_params['A1'], ws_params['B1'], ws_params['C1'] = "参数名称", "数值", "单位"
        for cell in ['A1', 'B1', 'C1']:
            ws_params[cell].font = Font(bold=True)

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
            ("升温速率", params['heat_rate'] / 60, "℃/min"),
            ("降温速率", params['cool_rate'] / 60, "℃/min"),
            ("循环次数", params['cycles'], "次")
        ]

        for i, (name, value, unit) in enumerate(param_list, start=2):
            ws_params[f'A{i}'] = name
            ws_params[f'B{i}'] = round(value, 2) if isinstance(value, float) else value
            ws_params[f'C{i}'] = unit

        # 数据表
        ws_data = wb.create_sheet(title="温度随时间变化")
        ws_data['A1'], ws_data['B1'], ws_data['C1'] = "时间 (h)", "温度 (℃)", "说明"
        for cell in ['A1', 'B1', 'C1']:
            ws_data[cell].font = Font(bold=True)

        for i, (time, temp, desc) in enumerate(profile, start=2):
            ws_data[f'A{i}'] = round(time, 2)
            ws_data[f'B{i}'] = round(temp, 2)
            ws_data[f'C{i}'] = desc

        # 保存到内存字节流
        excel_buffer = BytesIO()
        wb.save(excel_buffer)
        excel_buffer.seek(0)
        return excel_buffer
    except Exception as e:
        st.error(f"Excel创建失败: {e}")
        return None


def create_chart_image(profile, key_points):
    """生成图表，返回图片字节流用于显示和下载"""
    try:
        times = [p[0] for p in profile]
        temps = [p[1] for p in profile]
        descriptions = [p[2] for p in profile]

        plt.figure(figsize=(14, 8))
        plt.plot(times, temps, 'b-', linewidth=2, label='温度曲线')

        # 处理关键节点
        key_indices = list(set(key_points))
        key_indices.sort()
        used_positions = []

        for idx in key_indices:
            time = times[idx]
            temp = temps[idx]
            desc = descriptions[idx]

            plt.scatter(time, temp, color='red', s=80, marker='d', zorder=5)
            plt.axvline(x=time, color='gray', linestyle='--', linewidth=1, alpha=0.7)
            plt.axhline(y=temp, color='gray', linestyle='--', linewidth=1, alpha=0.7)

            # 动态计算标注位置（避免重叠）
            x_pos = time
            y_text_pos = plt.ylim()[0] + (plt.ylim()[1] - plt.ylim()[0]) * 0.02
            label_offset = 0.03 * (plt.xlim()[1] - plt.xlim()[0])

            overlap = any(abs(x_pos - used_x) < label_offset for (used_x, _) in used_positions)
            if overlap:
                y_text_pos = plt.ylim()[0] + (plt.ylim()[1] - plt.ylim()[0]) * 0.06
            else:
                used_positions.append((x_pos, y_text_pos))

            # 时间标注
            plt.text(x_pos, y_text_pos, f'{time:.2f}h',
                     horizontalalignment='center',
                     verticalalignment='bottom',
                     color='darkred',
                     bbox=dict(facecolor='white', alpha=0.9, boxstyle='round,pad=0.2'))

            # 温度标注
            y_pos = temp
            x_text_pos = plt.xlim()[0] + (plt.xlim()[1] - plt.xlim()[0]) * 0.02
            y_overlap = any(
                abs(y_pos - used_y) < 0.03 * (plt.ylim()[1] - plt.ylim()[0]) for (_, used_y) in used_positions)

            if y_overlap:
                x_text_pos = plt.xlim()[0] + (plt.xlim()[1] - plt.xlim()[0]) * 0.06
            else:
                used_positions.append((x_text_pos, y_pos))

            plt.text(x_text_pos, y_pos, f'{temp:.1f}℃',
                     horizontalalignment='left',
                     verticalalignment='center',
                     color='darkgreen',
                     bbox=dict(facecolor='white', alpha=0.9, boxstyle='round,pad=0.2'))


        # 保存到内存字节流
        chart_buffer = BytesIO()
        plt.savefig(chart_buffer, dpi=300, bbox_inches='tight', format='png')  # 提高分辨率到300dpi
        chart_buffer.seek(0)
        plt.close()
        return chart_buffer
    except Exception as e:
        st.error(f"图表创建失败: {e}")
        return None


# ==================== Streamlit页面布局与交互 ====================
def main():
    # 页面配置
    st.set_page_config(
        page_title="温度曲线生成器",
        page_icon="📈",
        layout="wide"
    )

    # 页面标题
    st.title("📈 温度曲线生成器")
    st.markdown("输入参数后生成温度随时间变化的曲线及数据文件")

    # 创建临时目录
    create_temp_dirs()

    # 用表单组织输入
    with st.form(key="temp_profile_form"):
        # 分两列显示输入框
        col1, col2 = st.columns(2)

        with col1:
            # 基础温度参数
            initial_temp = st.number_input("初始温度（℃）", value=70.0, step=0.1)
            initial_time = st.number_input("初始温度持续时间（h）", value=0.5, step=0.1, min_value=0.0)
            recovery_temp = st.number_input("回温温度（℃）", value=40.0, step=0.1)
            recovery_time = st.number_input("回温持续时间（h）", value=0.5, step=0.1, min_value=0.0)

            # 高低温参数
            high_temp = st.number_input("高温温度（℃）", value=100.0, step=0.1)
            high_tolerance = st.number_input("高温允差（℃）", value=2.0, step=0.1, min_value=0.0)
            low_temp = st.number_input("低温温度（℃）", value=-20.0, step=0.1)
            low_tolerance = st.number_input("低温允差（℃）", value=2.0, step=0.1, min_value=0.0)

        with col2:
            # 循环时间参数
            first_high_time = st.number_input("首循环高温保持时间（h）", value=0.25, step=0.1, min_value=0.0)
            first_low_time = st.number_input("首循环低温保持时间（h）", value=2.0, step=0.1, min_value=0.0)
            last_high_time = st.number_input("末循环高温保持时间（h）", value=0.25, step=0.1, min_value=0.0)
            last_low_time = st.number_input("末循环低温保持时间（h）", value=2.0, step=0.1, min_value=0.0)
            middle_high_time = st.number_input("中间循环高温保持时间（h）", value=1.0, step=0.1, min_value=0.0)
            middle_low_time = st.number_input("中间循环低温保持时间（h）", value=1.0, step=0.1, min_value=0.0)

            # 速率和循环次数
            heat_rate_per_min = st.number_input("升温速率（℃/min）", value=3.0, step=0.1, min_value=0.1)
            cool_rate_per_min = st.number_input("降温速率（℃/min）", value=4.0, step=0.1, min_value=0.1)
            cycles = st.number_input("循环次数（次）", value=3, step=1, min_value=1)

        # 文件名输入
        filename = st.text_input("输出文件名", value="温度曲线数据")
        cleaned_filename = validate_filename(filename)

        # 提交按钮
        submit_btn = st.form_submit_button("生成温度曲线", use_container_width=True)

    # 处理表单提交
    if submit_btn:
        # 组装参数
        params = {
            'initial_temp': initial_temp,
            'initial_time': initial_time,
            'recovery_temp': recovery_temp,
            'recovery_time': recovery_time,
            'high_temp': high_temp,
            'high_tolerance': high_tolerance,
            'low_temp': low_temp,
            'low_tolerance': low_tolerance,
            'first_high_time': first_high_time,
            'first_low_time': first_low_time,
            'last_high_time': last_high_time,
            'last_low_time': last_low_time,
            'middle_high_time': middle_high_time,
            'middle_low_time': middle_low_time,
            'heat_rate': heat_rate_per_min * 60,  # 转换为℃/h
            'cool_rate': cool_rate_per_min * 60,  # 转换为℃/h
            'cycles': int(cycles)
        }

        # 显示加载状态
        with st.spinner("正在计算并生成结果..."):
            # 计算温度曲线
            profile, key_points = calculate_temperature_profile(params)
            if not profile or not key_points:
                st.error("未能生成温度曲线，请检查输入参数")
                return

            # 生成Excel和图表
            excel_buffer = create_excel_file(params, profile)
            chart_buffer = create_chart_image(profile, key_points)

            if excel_buffer and chart_buffer:
                # 显示图表
                st.subheader("温度曲线结果")
                st.image(chart_buffer, caption="温度随时间变化曲线")

                # 创建下载按钮区域（分两列显示）
                download_cols = st.columns(2)

                with download_cols[0]:
                    # Excel下载按钮
                    st.download_button(
                        label="下载Excel数据",
                        data=excel_buffer,
                        file_name=f"{cleaned_filename}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

                with download_cols[1]:
                    # 图表下载按钮（新增功能）
                    st.download_button(
                        label="下载图表（PNG）",
                        data=chart_buffer,
                        file_name=f"{cleaned_filename}.png",
                        mime="image/png",
                        use_container_width=True
                    )

                st.success("生成成功！您可以下载Excel数据和PNG图表")


# 启动应用
if __name__ == "__main__":
    main()

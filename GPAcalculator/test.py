import logging
from openpyxl import load_workbook

def save_to_excel(data, course_name, template_path="SEMESTER 1 MARKINGS.xlsx", output_path="Updated_Markings.xlsx"):
    # 根据课程名称设置列和行的映射
    course_mapping = {
        "OPS": {
            "columns": {
                "Lab": "O",
                "Quiz": "J",
                "Midterm": "C",
                "Final": "C"
            },
            "lab_start_row": 63,
            "quiz_start_row": 63,
            "midterm_row": 66,
            "final_row": 67
        },
        "CPR": {
            "columns": {
                "Activity": "J",
                "Quiz": "O",
                "ICT": "D",
                "Final": "D"
            },
            "activity_start_row": 27,
            "quiz_start_row": 27,
            "ICT_row": 30,
            "final_row": 31
        },
        # 其他课程的映射可以在这里添加
    }

    try:
        # 打开模板文件
        workbook = load_workbook(template_path)
        sheet = workbook["TOTAL"]  # 假设所有数据填充到 "TOTAL" 表

        # 获取当前课程对应的列映射
        course_config = course_mapping.get(course_name, {})
        columns = course_config.get("columns", {})

        # 提取数据
        activity_data = []
        quiz_data = []
        midterm_grade = None
        final_grade = None
        ict_grade = None

        # 直接使用传入的小数格式数据
        for item_name, grade in data:
            # 根据课程名称分类
            if course_name == "OPS":
                if "Lab" in item_name:
                    activity_data.append((item_name, grade))
                elif "Quiz" in item_name:
                    quiz_data.append((item_name, grade))
                elif "Midterm" in item_name:
                    midterm_grade = grade
                elif "Final" in item_name:
                    final_grade = grade

            elif course_name == "CPR":
                if "Activity" in item_name:
                    activity_data.append((item_name, grade))
                elif "Quiz" in item_name:
                    quiz_data.append((item_name, grade))
                elif "Final" in item_name:
                    final_grade = grade
                elif "ICT" in item_name:  # 处理 ICT News
                    ict_grade = grade  # 保存 ICT News 的分数

        # 排序数据
        activity_data.sort(key=lambda x: int(''.join(filter(str.isdigit, x[0]))))  # 提取数字并排序
        quiz_data.sort(key=lambda x: int(''.join(filter(str.isdigit, x[0]))))  # 提取数字并排序

        # 填充数据
        if course_name == "OPS":
            # 填充 Activity 数据
            for index, (item_name, grade) in enumerate(activity_data):
                row = course_config["lab_start_row"] + index  # 从起始行开始填充
                sheet[f"{columns['Lab']}{row}"] = grade

            # 填充 Quiz 数据
            for index, (item_name, grade) in enumerate(quiz_data):
                row = course_config["quiz_start_row"] + index  # 从起始行开始填充
                sheet[f"{columns['Quiz']}{row}"] = grade

            # 填充 Midterm
            if midterm_grade is not None:
                sheet[f"{columns['Midterm']}{course_config['midterm_row']}"] = midterm_grade

            # 填充 Final
            if final_grade is not None:
                sheet[f"{columns['Final']}{course_config['final_row']}"] = final_grade

        elif course_name == "CPR":
            # 填充 Activity 数据
            for index, (item_name, grade) in enumerate(activity_data):
                row = course_config["activity_start_row"] + index  # 从起始行开始填充
                sheet[f"{columns['Activity']}{row}"] = grade

            # 填充 Quiz 数据
            for index, (item_name, grade) in enumerate(quiz_data):
                row = course_config["quiz_start_row"] + index  # 从起始行开始填充
                sheet[f"{columns['Quiz']}{row}"] = grade

            # 填充 ICT 数据
            if ict_grade is not None:
                sheet[f"{columns['ICT']}{course_config['ICT_row']}"] = ict_grade

            # 填充 Final
            if final_grade is not None:
                sheet[f"{columns['Final']}{course_config['final_row']}"] = final_grade

        # 保存新文件
        workbook.save(output_path)
        logging.info(f"Grades for {course_name} saved to {output_path} successfully!")
    except Exception as e:
        logging.error(f"Failed to save data for {course_name} to Excel: {e}")

def main():
    # OPS 数据（假设已经是小数格式）
    data1 = [
        ("Lab #5", 1.0),  # 3 / 3
        ("Lab #2", 0.6667),  # 2 / 3
        ("Lab #4", 1.0),  # 3 / 3
        ("Lab #3", 1.0),  # 3 / 3
        ("Lab #1", 0.3333),  # 1 / 3
        ("Quiz 2", 0.4),  # 4 / 10
        ("Quiz 1", 0.8),  # 8 / 10
        ("Midterm-Fall2024", 0.8967),  # 54.75 / 61
        # Final没有数据，留白
    ]
    save_to_excel(data1, "OPS")

    # CPR 数据（假设已经是小数格式）
    data2 = [
        ("Activity 03", 0.96),  # 96 / 100
        ("Activity 04", 0.97),  # 97 / 100
        ("Activity 05", 0.98),  # 98 / 100
        ("Activity 06", 0.95),  # 95 / 100
        ("Activity 07", 0.95),  # 95 / 100
        ("Activity 08", 0.98),  # 98 / 100
        ("Activity 09", 0.95),  # 95 / 100
        ("Quiz 04", 0.8),  # 80 / 100
        ("Activity 01", 0.9),  # 90 / 100
        ("Activity 02", 0.96),  # 96 / 100
        ("Activity 10", 1.0),  # 100 / 100
        ("Quiz 01", 0.8),  # 80 / 100
        ("Quiz 02", 1.0),  # 100 / 100
        ("Quiz 03", 1.0),  # 100 / 100
        ("Quiz 05", 0.9),  # 90 / 100
        ("Quiz 06", 1.0),  # 100 / 100
        ("Quiz 07", 0.6),  # 60 / 100
        ("Quiz 08", 1.0),  # 100 / 100
        ("Quiz 09", 1.0),  # 100 / 100
        ("Quiz 10", 0.8),  # 80 / 100
        ("ICT News", 0.9),  # 90 / 100
    ]
    save_to_excel(data2, "CPR")

if __name__ == "__main__":
    main()
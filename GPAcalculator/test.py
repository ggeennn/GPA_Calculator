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
        # 其他课程的映射可以在这里添加
    }

    try:
        # 打开模板文件
        workbook = load_workbook(template_path)
        sheet = workbook["TOTAL"]  # 假设所有数据填充到 "TOTAL" 表

        # 获取当前课程对应的列映射
        course_config = course_mapping.get(course_name, {})
        columns = course_config.get("columns", {})

        # 提取 Lab 和 Quiz 数据
        lab_data = []
        quiz_data = []
        midterm_grade = None
        final_grade = None

        for item_name, grade in data:
            # 将分数格式转换为百分比
            score, total = map(float, grade.split('/'))
            formatted_grade = (score / total) # 使用小数

            # 根据item_name分类
            if "Lab" in item_name:
                lab_data.append((item_name, formatted_grade))
            elif "Quiz" in item_name:
                quiz_data.append((item_name, formatted_grade))
            elif "Midterm" in item_name:
                midterm_grade = formatted_grade
            elif "Final" in item_name:
                final_grade = formatted_grade

        # 排序 Lab 和 Quiz 数据
        lab_data.sort(key=lambda x: int(''.join(filter(str.isdigit, x[0]))))  # 提取数字并排序
        quiz_data.sort(key=lambda x: int(''.join(filter(str.isdigit, x[0]))))  # 提取数字并排序

        # 填充 Lab 数据
        for index, (item_name, grade) in enumerate(lab_data):
            row = course_config["lab_start_row"] + index  # 从起始行开始填充
            sheet[f"{columns['Lab']}{row}"] = grade

        # 填充 Quiz 数据
        for index, (item_name, grade) in enumerate(quiz_data):
            row = course_config["quiz_start_row"] + index  # 从起始行开始填充
            sheet[f"{columns['Quiz']}{row}"] = grade

        # 填充 Midterm
        if midterm_grade is not None:
            sheet[f"{columns['Midterm']}{course_config['midterm_row']}"] = midterm_grade

        # 仅在包含 Final 数据时填充
        if final_grade is not None:
            sheet[f"{columns['Final']}{course_config['final_row']}"] = final_grade

        # 保存新文件
        workbook.save(output_path)
        logging.info(f"Grades for {course_name} saved to {output_path} successfully!")
    except Exception as e:
        logging.error(f"Failed to save data for {course_name} to Excel: {e}")

def main():
    data = [
        ("Lab #5", "3 / 3"),
        ("Lab #2", "2 / 3"),
        ("Lab #4", "3 / 3"),
        ("Lab #3", "3 / 3"),
        ("Lab #1", "1/ 3"),
        ("Quiz 2", "4 / 10"),
        ("Quiz 1", "8/10"),
        ("Midterm-Fall2024", "54.75 / 61"),
        # Final没有数据，留白
    ]
    save_to_excel(data, "OPS")

if __name__ == "__main__":
    main()
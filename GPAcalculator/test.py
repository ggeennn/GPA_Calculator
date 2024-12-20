import logging
from openpyxl import load_workbook

def normalize_name(name):
    return ''.join(filter(str.isalnum, name)).lower()  # 去掉非字母数字字符，转换为小写

def save_to_excel(workbook, data, course_name):
    course_mapping = {
        "OPS": {
            "columns": {
                "Lab": "O",
                "Quiz": "J",
                "Midterm": "C",
                "Final": "C"
            },
            "rows": {
                "lab_start_row": 63,   # Lab 动态行号起始行
                "quiz_start_row": 63,  # Quiz 动态行号起始行
                "midterm_row": 66,     # Midterm 固定行号
                "final_row": 67        # Final 固定行号
            }
        },
        "CPR": {
            "columns": {
                "Activity": "J",
                "Quiz": "O",
                "ICT": "D",
                "Final": "D"
            },
            "rows": {
                "activity_start_row": 27,  # Activity 动态行号起始行
                "quiz_start_row": 27,      # Quiz 动态行号起始行
                "ict_row": 30,             # ICT 固定行号
                "final_row": 31            # Final 固定行号
            }
        },
    }

    # 获取课程的配置
    course_config = course_mapping.get(course_name, {})
    columns = course_config.get("columns", {})
    rows = course_config.get("rows", {})

    # 填充数据
    for item_name, grade in data:
        logging.debug(f"Processing item: {item_name}, grade: {grade}")
        matched = False
        normalized_item = normalize_name(item_name)

        for key, column in columns.items():
            normalized_key = normalize_name(key)
            if normalized_key in normalized_item:
                matched = True
                if key in ["Lab", "Quiz", "Activity"]:  # 动态计算行号
                    item_number = int(''.join(filter(str.isdigit, item_name)) or '0')
                    row = rows[f"{key.lower()}_start_row"] + item_number - 1
                else:  # 固定行
                    row = rows[f"{key.lower()}_row"]
                
                logging.debug(f"Writing to column: {column}, row: {row}, grade: {grade}")
                workbook["TOTAL"][f"{column}{row}"] = grade  # 填充数据
                break  # 找到匹配后跳出内层循环

        if not matched:
            logging.warning(f"No matching key found for item: {item_name}")

def main():
    # 打开模板文件
    template_path = "SEMESTER 1 MARKINGS.xlsx"
    output_path = "Updated_Markings.xlsx"
    workbook = load_workbook(template_path)

    # OPS 数据（假设已经是小数格式）
    data1 = [
        ("Lab #1", 0.3333),  # 1 / 3
        ("Lab #2", 0.6667),  # 2 / 3
        ("Lab #3", 1.0),     # 3 / 3
        ("Lab #4", 1.0),     # 3 / 3
        ("Lab #5", 1.0),     # 3 / 3
        ("Quiz 1", 0.8),     # 8 / 10
        ("Quiz 2", 0.4),     # 4 / 10
        ("Midterm-Fall2024", 0.8967),  # 54.75 / 61
        # Final没有数据，留白
    ]
    save_to_excel(workbook, data1, "OPS")

    # CPR 数据（假设已经是小数格式）
    data2 = [
        ("Activity 01", 0.9),  # 90 / 100
        ("Activity 02", 0.96),  # 96 / 100
        ("Activity 03", 0.96),  # 96 / 100
        ("Activity 04", 0.97),  # 97 / 100
        ("Activity 05", 0.98),  # 98 / 100
        ("Activity 06", 0.95),  # 95 / 100
        ("Activity 07", 0.95),  # 95 / 100
        ("Activity 08", 0.98),  # 98 / 100
        ("Activity 09", 0.95),  # 95 / 100
        ("Activity 10", 1.0),    # 100 / 100
        ("Quiz 01", 0.8),       # 80 / 100
        ("Quiz 02", 1.0),       # 100 / 100
        ("Quiz 03", 1.0),       # 100 / 100
        ("Quiz 04", 0.8),       # 80 / 100
        ("Quiz 05", 0.9),       # 90 / 100
        ("Quiz 06", 1.0),       # 100 / 100
        ("Quiz 07", 0.6),       # 60 / 100
        ("Quiz 08", 1.0),       # 100 / 100
        ("Quiz 09", 1.0),       # 100 / 100
        ("Quiz 10", 0.8),       # 80 / 100
        ("ICT News", 0.9),      # 90 / 100
    ]
    save_to_excel(workbook, data2, "CPR")

    # 保存新文件
    workbook.save(output_path)
    logging.info(f"All grades saved to {output_path} successfully!")

if __name__ == "__main__":
    logging.basicConfig(level=logging.DEBUG)  # 设置日志级别为 DEBUG
    main()
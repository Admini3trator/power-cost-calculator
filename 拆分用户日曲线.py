import pandas as pd
import os

# 拆分后的存放路径
file_paths = [f'D:/当日用电量详情 ({i + 1}).xlsx' for i in range(*)] #*为当月总天数

company_column = '用户名称'

for file_index, file_path in enumerate(file_paths):
    try:
        excel_file = pd.ExcelFile(file_path)
        sheet_names = excel_file.sheet_names

        for sheet_name in sheet_names:
            df = excel_file.parse(sheet_name)

            # 获取各企业名称
            companies = df[company_column].unique()

            # 遍历企业
            for company in companies:
                # 筛选数据
                company_data = df[df[company_column] == company]

                company_folder = f'D:/{company}'  # 创建文件夹的路径
                if not os.path.exists(company_folder):
                    os.makedirs(company_folder)

                # 保存数据
                output_file_path = f'{company_folder}/当日用电量详情 ({file_index + 1}).xlsx'
                company_data.head(2).to_excel(output_file_path, index=False)
    except Exception as e:
        print(f'处理文件 {file_path} 时出现错误: {e}')
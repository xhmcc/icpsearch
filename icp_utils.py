from re import match
from pandas import DataFrame, ExcelWriter, read_excel, isna
from math import ceil
from os import path, makedirs
from openpyxl.styles import Font, Alignment

def is_valid_domain(domain):
    domain = domain.strip()
    ip_pattern = r'^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$'
    if match(ip_pattern, domain):
        parts = domain.split('.')
        for part in parts:
            if int(part) > 255:
                return True
        return False
    domain_pattern = r'^[a-zA-Z0-9][a-zA-Z0-9-]{0,61}[a-zA-Z0-9](\.[a-zA-Z]{2,})+$'
    if match(domain_pattern, domain):
        return True
    if '。' in domain or '．' in domain:
        return True
    return True

def save_results(results, filename, show_message=True):
    try:
        existing_df = DataFrame()
        if path.exists(filename):
            try:
                existing_df = read_excel(filename)
            except:
                if show_message:
                    print(f"读取现有文件失败，将创建新文件: {filename}")
        new_df = DataFrame(results)
        if not existing_df.empty:
            columns = ['企业名称', '备案域名', '备案IP', '备案微信小程序', '备案微信公众号', '备案APP']
            existing_df = existing_df.reindex(columns=columns, fill_value='')
            new_df = new_df.reindex(columns=columns, fill_value='')
            for _, row in new_df.iterrows():
                company_name = row['企业名称']
                mask = existing_df['企业名称'] == company_name
                if mask.any():
                    existing_df.loc[mask] = row.values
                else:
                    existing_df = DataFrame.concat([existing_df, DataFrame([row])], ignore_index=True)
            result_df = existing_df
        else:
            result_df = new_df
        with ExcelWriter(filename, engine='openpyxl') as writer:
            result_df.to_excel(writer, index=False)
            worksheet = writer.sheets['Sheet1']
            max_widths = {}
            for col in worksheet.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    if cell.value:
                        lines = str(cell.value).split('\n')
                        for line in lines:
                            length = sum(2 if '\u4e00' <= char <= '\u9fff' else 1 for char in line)
                            max_length = max(max_length, length)
                adjusted_width = min(max_length * 1.2 + 2, 50)
                worksheet.column_dimensions[column].width = adjusted_width
                max_widths[column] = adjusted_width
            for cell in worksheet[1]:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            for row in worksheet.iter_rows(min_row=2):
                max_lines = 1
                max_chars_per_line = 0
                for cell in row:
                    if cell.value:
                        cell.alignment = Alignment(wrap_text=True, vertical='top')
                        lines = str(cell.value).split('\n')
                        for line in lines:
                            chars = sum(2 if '\u4e00' <= char <= '\u9fff' else 1 for char in line)
                            max_chars_per_line = max(max_chars_per_line, chars)
                        column_width = max_widths[cell.column_letter]
                        lines_needed = max(1, ceil(max_chars_per_line / (column_width * 0.8)))
                        max_lines = max(max_lines, lines_needed)
                worksheet.row_dimensions[row[0].row].height = max_lines * 15 + 5
        if show_message:
            print(f"已保存结果到 {filename}")
        return True
    except Exception as e:
        print(f"保存结果失败: {str(e)}")
        return False 
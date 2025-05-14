# icpsearch
批量根据企业名称查询企业备案域名，目前支持爱企查


## 使用命令： ##
python icpsearch.py
python icpsearch.py -proxy http://127.0.0.1:8080

## 使用步骤： ##
1、在config.yaml文件中填写爱企查cookie

![image](https://github.com/user-attachments/assets/64e6f062-ef59-4b5d-bcce-a71c47eb1688)

2、在同文件下表格（company_name.xlsx）内填入企业名称，脚本会根据表格内企业名称依次查询备案

![image](https://github.com/user-attachments/assets/5b8111a3-d5a1-4ad2-b89d-f591be715007)

3、最后直接打开终端运行即可
python icpsearch.py

![image](https://github.com/user-attachments/assets/d63e5b06-9eeb-49d3-8861-d1fa40545be9)

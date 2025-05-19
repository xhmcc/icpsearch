# icpsearch
批量根据企业名称查询企业备案域名、IP、微信小程序、微信公众号、app，目前支持爱企查aqc、icp

icp备案查询可使用代理池查询（须走http代理）
## 使用命令（icpsearch_aqc.py）： ##
python icpsearch_aqc.py -f input.xlsx -o output.xlsx

python icpsearch_aqc.py -f input.xlsx -o output.xlsx -proxy http://127.0.0.1:8080

python icpsearch_aqc.py -f input.xlsx -o output.xlsx -d 1  # 设置请求间隔为1秒

options:

-  -f FILE       指定输入Excel文件路径，默认为company_name.xlsx
  
-  -o OUTPUT     指定输出Excel文件路径，默认为company_domains_result.xlsx
  
-  -proxy PROXY  设置代理服务器，例如: http://127.0.0.1:8080
  
-  -d DELAY      设置请求间隔时间（秒），默认为0秒
  
## 使用命令（icpsearch_icp.py）： ##
python icpsearch_icp.py -f input.xlsx -o output.xlsx

python icpsearch_icp.py -f input.xlsx -o output.xlsx -d 1  # 设置请求间隔为1秒

python icpsearch_icp.py -f input.xlsx -o output.xlsx -proxy http://127.0.0.1:8008  # 使用代理

python icpsearch_icp.py -f input.xlsx -o output.xlsx -proxy proxypool.txt  # 使用代理池

options:

-  -f FILE       指定输入Excel文件路径，默认为company_name.xlsx
  
-  -o OUTPUT     指定输出Excel文件路径，默认为company_domains_result.xlsx
                        
-  -d DELAY      置请求间隔时间（秒），默认为0秒
                        
-  -proxy PROXY  设置代理服务器或代理池文件路径
## 使用步骤（aqc和icp使用步骤相同）： ##
1、在config.yaml文件中填写爱企查cookie

![image](https://github.com/user-attachments/assets/64e6f062-ef59-4b5d-bcce-a71c47eb1688)

2、在同文件夹下表格（company_name.xlsx）内、或自行指定表格填入企业名称，脚本会根据表格内企业名称依次查询备案

![image](https://github.com/user-attachments/assets/5b8111a3-d5a1-4ad2-b89d-f591be715007)

3、最后直接打开终端运行即可

python icpsearch.py -f input.xlsx -o output.xlsx

![image](https://github.com/user-attachments/assets/a82b1139-71fd-4b9a-a763-24f6a4356d6b)

4、运行结束后会自动生成一个表格（company_domains_result.xlsx），或-o自行指定输出文件

可在此表格中看到企业备案域名及IP
![image](https://github.com/user-attachments/assets/0fe1678c-8765-4740-9e81-ec06027ed1af)


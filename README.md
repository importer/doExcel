# doExcel
一个python操作excel的实例

# 基于 win32com.client 的操作方案
> excel_formula.py
优点：可保留excel模版内的格式（特别是公式），相当于模拟人工点击
缺点：速度慢，性能是硬伤
excel_formula.py

# 基于 xlrd,xlutils,xlwt 也可使用openpyxl （可以读写）
> db2_conn.py
优点：快
缺点：单元格格式会丢失


### 功能描述
> 从数据库取出数据，按照单元格所处位置填入数据
### 运行环境
> python2
### quick start
> 需安装包
xlrd,xlutils,xlwt,ibm_db,pywin32

### 运行
```
python excel_formula.py 201901(具体日期)
```

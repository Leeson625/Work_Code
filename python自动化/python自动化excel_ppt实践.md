# python自动化excel/ppt实践

# 自动化提效概览

![image.png](https://alidocs.oss-cn-zhangjiakou.aliyuncs.com/res/5VLqXLb7LagJxqX1/img/3bdea7bf-4e35-421e-845f-c99fc42b371c.png)

*   比如在老客月会ppt汇报材料中，每周需要刷新到最新数据然后替换PPT中的图片，每次都需要2-3h的工作时间：
    
    *   1.刷新底表数据和图片 -- 1.5h~2h
        
    *   2.每个图片对应粘贴进ppt中，调整格式 -- 0.5h
        
*   尝试用python代码绘图并进行自动替换后，每次仅需改一下业务日期（全局变量引入），需要约20分钟。
    

## 1.1 前期准备

### 1.1.1 下载miniconda及vscode

*   因anaconda在公司使用存在版权问题，建议安装miniconda
    

[https://www.anaconda.com/download/success: https://www.anaconda.com/download/success](https://www.anaconda.com/download/success)

*   VScode安装地址及[官方使用教程](https://vscode.js.cn/docs/getstarted/getting-started)
    

[https://code.visualstudio.com/: https://code.visualstudio.com/](https://code.visualstudio.com/)

*   conda配合vscode进行安装配置
    

[https://zhuanlan.zhihu.com/p/1960979391359197505: https://zhuanlan.zhihu.com/p/1960979391359197505](https://zhuanlan.zhihu.com/p/1960979391359197505)

### 1.1.2 安装相应的库（此处仅列举教程需要的库）

```Shell
pip install numpy pandas matplotlib 

pip install pywin32
pip install pyodps
pip install python-pptx
```
:::
#### odps 

`odps`  （常指阿里云 MaxCompute 的 Python SDK，包名通常为 pyodps）用于在 Python 中与 MaxCompute 交互。

作用：与阿里云 MaxCompute（ODPS）交互，执行 SQL、提交任务、读取结果到 Pandas。

主要功能：提交 SQL 查询、读取表、任务管理、权限与项目绑定。

优点：直接对接大数据仓库，处理海量数据，高吞吐；返回 Pandas 友好，便于下游分析。

参考文档：[Python编程接口开发MaxCompute作业-PyODPS-云原生大数据计算服务 MaxCompute-阿里云](https://help.aliyun.com/zh/maxcompute/user-guide/pyodps-3/?spm=a2c4g.11186623.help-menu-27797.d_2_5_3.56eb31eavkbYe3)
:::
:::
#### pywin32

`win32com` 是 Python 在 Windows 平台上一种用于访问 COM（Component Object Model） 对象的库／模块。它隶属于第三方扩展包 pywin32的一部分。

借助 win32com，Python 脚本可以与 Windows 系统的底层组件或支持 COM 接口的应用程序通信，并操控它们——比如常见的办公软件（如 Excel、Word）、以及其他支持 COM 的软件。这样就能通过 Python 自动化许多原本依赖鼠标/键盘/手动操作的流程。

优点：覆盖 Excel 原生 99% 能力（透视表、刷新、宏、复杂格式）——这是 openpyxl 做不到的，适合基于“现有模板”的自动化，不重建报表，只替换数据 + 刷新 + 导出。

参考文档：[Python 3中的win32com使用教程+示例 - 知乎](https://zhuanlan.zhihu.com/p/1982926811475231366)

```python
from win32com.client import Dispatch

excel = Dispatch("Excel.Application")
excel.Visible = True          # 可见 Excel 窗口（方便调试／查看效果）

wb = excel.Workbooks.Open(r"D:\python自动化实践\excel样例\demo1.xlsx")  #打开工作簿
sheet = workbook.Worksheets(1)        #获取第一个工作表

sheet.Cells(1, 1).Value = "Hello"         #向若干个单元格写入数据
sheet.Cells(1, 2).Value = "World"
sheet.Cells(2, 1).Value = "Python"
sheet.Cells(2, 2).Value = "win32com"

workbook.SaveAs(r"C:\path\to\your\file.xlsx")   #保存为 .xlsx 文件
workbook.Close(False)
excel.Quit()   #关闭 workbook 和 Excel 应用程序
```
:::
:::
#### python-pptx

`python-pptx`是 Python 用来读写/生成 PowerPoint（.pptx） 的第三方库。它通过操作 Office Open XML（PPTX 本质是压缩包里的 XML）来完成创建、修改幻灯片内容。

借助python-pptx，我们可以通过编程方式创建、修改PowerPoint (.pptx) 演示文稿。这个库非常适合自动化处理PPT文档，比如新建/打开 PPTX，新增、删除、复制幻灯片；插入和编辑文本框、标题、段落、项目符号；插入图片、形状、表格、图表；调整常见样式：字体、字号、颜色、对齐、位置、大小、背景等

优点：跨平台（Windows/macOS/Linux 都能用）；适合“批量生成/批量替换内容”的场景：日报、周报、月报、一键出图进 PPT

参考文档：[Python-pptX — Python-pptX 1.0.0 文档](https://python-pptx.readthedocs.io/en/latest/)、[Python自动化办公-PPT操作篇 - 知乎](https://zhuanlan.zhihu.com/p/690653470)
:::

### 1.1.3 申请阿里云AccessKey，配置本地变量

*   申请流程：钉钉-> OA审批 -> 阿里云资源申请 -> 填写类别和申请原因进行申请
    

*   把申请得到的AccessKey配置到本地环境变量中，可参考[在Linux、macOS和Windows环境变量中配置阿里云AccessKey-阿里云SDK-阿里云](https://help.aliyun.com/zh/sdk/developer-reference/configure-the-alibaba-cloud-accesskey-environment-variable-on-linux-macos-and-windows-systems?spm=a2c4g.11186623.0.0.6c456845G1s82f#da8301dcdaoyk)，不想配置的话也可以直接手动传入AccessKey。
    
    ![image.png](https://alidocs.oss-cn-zhangjiakou.aliyuncs.com/res/5VLqXLb7LagJxqX1/img/82b629e4-cfa9-4d35-85e4-7026b1398630.png)
    
*   以上步骤配置好后，尝试在本地运行一段sql查询，如果正常运行出结果，说明odps配置成功。其他本地操作可参考[在本地使用PyODPS执行SQL查询及操作DataFrame-云原生大数据计算服务 MaxCompute-阿里云](https://help.aliyun.com/zh/maxcompute/user-guide/use-pyodps-in-local-environment?spm=a2c4g.11186623.help-menu-27797.d_2_5_3_2_1.ee547824mPPM1l&scm=20140722.H_470450._.OR_help-T_cn~zh-V_1)
    

```python
import os
from odps import ODPS
import numpy as np
import pandas as pd

##initialize odps
o = ODPS(
    # （推荐）确保已设置环境变量。
    # 确保ALIBABA_CLOUD_ACCESS_KEY_ID环境变量设置为用户 Access Key ID。
    access_id=os.getenv('ALIBABA_CLOUD_ACCESS_KEY_ID'),
    # access_id='YOUR_ACCESS_KEY_ID',
    
    # 确保ALIBABA_CLOUD_ACCESS_KEY_SECRET环境变量设置为用户Access Key Secret。
    secret_access_key=os.getenv('ALIBABA_CLOUD_ACCESS_KEY_SECRET'),
    # secret_access_key='YOUR_ACCESS_KEY_SECRET',
    project='xyf_jingying_dev',
    endpoint='https://service.cn-beijing.maxcompute.aliyun.com/api',
)
your_select = '''


'''
result=o.execute_sql(your_select)
loacal_dataframe=result.open_reader().to_pandas()
```


## 1.2 excel与PPT提效主要思路

### 1.2.1 节省数据下载与复制粘贴时间，同时可以使用str.format方法设置全局传入参数

日常工作更新周报月报中，我们常需要在最新的时间窗口上重跑代码，将数据进行更新与替换，而有了odps接口，我们可以节省数据下载与复制的时间，同时可以结合python中一些字符串方法，实现周期性的数据重刷与替换工作

> **str.format** 方法是 Python 中用于字符串格式化的强大工具。它通过在字符串中使用花括号 {} 作为占位符，然后将其替换为传递给 format 方法的参数来工作

例如，在本例中，对于周报中需要批量更新时间的sql代码，可自定义全局变量begin\_date和end\_date,从而可快速实现批量替换

```python
SELECT  first_order_number
       ,order_number
       ,user_no
       ,cust_no
       ,period
       ,loan_amt
       ,loan_time
       ,first_order_time
       ,asset_type_flag
       ,fee_rate    
FROM xyf_dws.dws_inloan_user_order_df
WHERE 1 = 1
AND pt = '${bizdate}'   
AND app IN ('xyf01', 'fxk') --xyf01和fxk贷款 
AND business_line IN ('APP', '小程序端') --定义APP贷款 
AND loan_status = 'success'
AND loan_flag IN ('复贷', '加贷')  --复贷 
AND loan_time IS NOT NULL
AND DATE(loan_time) >= '2024-11-01'
AND DATE(loan_time) <= '2025-11-20'   --每次更新一下日期条件
```
```python
## 自定义bizdate 与数据统计区间
bizdate = '20251218'
begin_date = '2025-01-01'   
end_date = '2025-12-18'

query1='''
	SELECT  first_order_number
	       ,order_number
	       ,user_no
	       ,cust_no
	       ,period
	       ,loan_amt
	       ,loan_time
	       ,first_order_time
           ,asset_type_flag
	       ,fee_rate    --0.3599 
	FROM xyf_dws.dws_inloan_user_order_df
	WHERE 1 = 1
	AND pt = '{bizdate}'
	AND app IN ('xyf01', 'fxk') --xyf01和fxk贷款 
	AND business_line IN ('APP', '小程序端') --定义APP贷款 
	AND loan_status = 'success'
	AND loan_flag IN ('复贷', '加贷')  --复贷 
	AND loan_time IS NOT NULL
	AND DATE(loan_time) >= '{begin_date}'
	AND DATE(loan_time) <= '{end_date}'   
'''.format(bizdate = bizdate, begin_date = begin_date, end_date = end_date)

query2='''
WITH apr AS
(
	SELECT  order_number
	       ,SUM(initial_principal)                                                                 AS initial_principal
	       ,SUM(nvl(initial_interest,0)+nvl(initial_after_loan_fee,0)+nvl(initial_platform_fee,0)) AS initial_interest_fee
	FROM xyf_dwd.dwd_repay_loan_repay_plan_df
	WHERE pt = '{bizdate}'
	GROUP BY  order_number
), 

order_apr AS(
SELECT  a.cust_no
       ,a.user_no AS app_user_id
       ,a.loan_time
       ,a.loan_amt
       ,b.initial_principal
       ,b.initial_interest_fee
FROM xyf_dws.dws_inloan_user_order_df a
INNER JOIN apr b
ON a.order_number = b.order_number AND a.pt = '{bizdate}' 
AND DATE(a.loan_time) >= '{begin_date}' AND a.app IN ('xyf01', 'fxk') AND a.loan_status = 'success' 
AND a.loan_flag IN ('复贷','加贷') AND a.business_line IN ('APP', '小程序端') AND a.loan_time IS NOT NULL 
)

SELECT  substr(loan_time,1,7)     AS loan_mth
       ,SUM(loan_amt)             AS loan_amt
       ,SUM(initial_principal)    AS principal
       ,SUM(initial_interest_fee) AS fee
FROM order_apr
GROUP BY  substr(loan_time,1,7)
ORDER BY  substr(loan_time,1,7);
'''.format(bizdate = bizdate, begin_date = begin_date)
```

### 1.2.2 可在本地利用python的多元化库进行数据处理和绘制图片

python有各种衍生的库能够进行数据分析和探索，且当前AI工具较为成熟，能帮助实现各种个性化需求

需要额外注意一点，ODPS 用 open\_reader().to\_pandas() 拉回的金额列有时是 decimal.Decimal（在 pandas 中为 object），与 float 直接混算会报类型错误或产生精度问题。建议在运行结束及时对数据转换类型

```python
#放款统计
result = o.execute_sql(query)
data_stats = result.open_reader().to_pandas()

# 字符串字段
str_cols = ['放款月', '申请月','首复贷flag']
for col in str_cols:
    if col in data_stats.columns:
        # data_stats[col] = data_stats[col].astype('string').fillna('')  # 填充空值为空字符串
        data_stats[col] = data_stats[col].astype('string')

# 日期字段 
date_cols = [ '放款日期' ]
for col in date_cols:
    if col in data_stats.columns:
        data_stats[col] = pd.to_datetime(data_stats[col]).dt.strftime('%Y-%m-%d')

# 整数字段
int_cols = ['放款人数','订单数']
for col in int_cols:
    if col in data_stats.columns:
        data_stats[col] = pd.to_numeric(data_stats[col], errors='coerce').astype('Int64')  # 使用Pandas 的 Int64（可空整数）/ NumPy 的 int64（普通整数）


# 浮点数字段 - 保留2位小数
float_cols = ['放款金额']
for col in float_cols:
    if col in data_stats.columns:
        data_stats[col] = pd.to_numeric(data_stats[col], errors='coerce').astype('float64') # “清洗 + 转类型”，对混有字符串数字（如 "12.3"）、空字符串、非法字符等也能处理

```

# excel刷新底表，一键更新数据透视表

## 2.1 更新数据透视表常见问题

在日常分析工作中，我们常需要更新数据透视表底表数据，会遇到以下三类场景：

a.行维度：数据透视表底层数据需要扩充更新或者删除

b.列维度：列字段可能需要增加或删除

c.计算字段：需要批量添加大量计算字段

### 2.1.1 名称管理器-在Excel中创建动态数据透视表以方便刷新扩展数据

*   **名称管理器****：**选中【数据源】-【公式】-【名称管理器】
    

![image.png](https://alidocs.oss-cn-zhangjiakou.aliyuncs.com/res/5VLqXLb7LagJxqX1/img/953973c9-152d-4f55-8172-96cfb8836bf5.png)

*   **编辑名称，输入公式  =OFFSET(放款统计!$A$1,0,0,COUNTA(放款统计!$A:$A),COUNTA(放款统计!$1:$1))**，其核心思想就是构建了一个可动态变化更新的动态引用区域。
    
    ![image.png](https://alidocs.oss-cn-zhangjiakou.aliyuncs.com/res/5VLqXLb7LagJxqX1/img/08a4ed8c-b953-408d-b529-f8bcc2b74a76.png)
    
*   **建立数据透视表**
    

![image.png](https://alidocs.oss-cn-zhangjiakou.aliyuncs.com/res/5VLqXLb7LagJxqX1/img/511952ce-287f-4181-b373-27d4002d6113.png)

*   此外超级表更新也可实现同样的效果，可参考[EXCEL新增数据时,数据透视表/图表自动更新 - 知乎](https://zhuanlan.zhihu.com/p/701237685)
    

### 2.1.2 替换底表数据并自动刷新数据透视表

```python
from win32com.client import Dispatch

def write_dataframe_to_excel_com(file_path, dataframes_dict, start_row=1, start_col=1, include_header=True):
    """
    Args:
        file_path: Excel文件路径
        dataframes_dict: {sheet_name: dataframe} 字典
        start_row: 从第几行开始写入（默认1）
        start_col: 从第几列开始写入（默认1）
        include_header: 是否写入表头（默认True）
    """
    # 启动Excel应用
    excel_app = Dispatch("Excel.Application")
    excel_app.Visible = False               # 后台运行
    excel_app.DisplayAlerts = False         # 禁用警告对话框

    try:
        # 打开工作簿
        workbook = excel_app.Workbooks.Open(file_path)

        for sheet_name, df in dataframes_dict.items():
            try:
                # 获取工作表
                worksheet = workbook.Worksheets(sheet_name)

                # 连表头/旧内容一起清空
                used = worksheet.UsedRange
                if used is not None and used.Rows.Count > 0 and used.Columns.Count > 0:
                    used.ClearContents()

                if df is not None and not df.empty:
                    # Excel/COM 识别 None（写成空单元格），不识别 Pandas 的 pd.NA，写入前：把 Pandas 的 NA/NaN/inf 转成 Excel 可接受的 None
                    df_to_write = df.copy()
                    df_to_write = df_to_write.replace([np.inf, -np.inf], np.nan)
                    df_to_write = df_to_write.astype('object')       # 防止 NAType 保留在扩展dtype里
                    df_to_write = df_to_write.where(pd.notna(df_to_write), None)

                    # 要写入的数据：可选表头 + 数据
                    if include_header:
                        values = [df_to_write.columns.tolist()] + df_to_write.values.tolist()
                    else:
                        values = df_to_write.values.tolist()

                    n_rows = len(values)
                    n_cols = len(values[0]) if n_rows > 0 else 0

                    end_row = start_row + n_rows - 1
                    end_col = start_col + n_cols - 1

                    write_range = worksheet.Range(
                        worksheet.Cells(start_row, start_col),
                        worksheet.Cells(end_row, end_col),
                    )
                    write_range.Value = values

                print(f"成功写入工作表: {sheet_name}")

            except Exception as e:
                print(f"写入工作表 {sheet_name} 时出错: {e}")
                continue
        
        workbook.RefreshAll()                       # 刷新所有数据连接和透视表
        excel_app.Calculate()                       # 强制重新计算公式
        excel_app.CalculateUntilAsyncQueriesDone()  # 等待所有操作完成
        
        # 保存并关闭
        workbook.Save()
        workbook.Close()
        print(f"文件已保存: {file_path}")

    except Exception as e:
        print(f"操作Excel文件时出错: {e}")

    finally:
        excel_app.Quit()        # 退出Excel应用

```

*   在2.1.1设置动态数据透视表基础上，使用win32com更新底表数据，便可实现自动化刷新数据透视表。
    

![image.png](https://alidocs.oss-cn-zhangjiakou.aliyuncs.com/res/5VLqXLb7LagJxqX1/img/fa364ed1-b07c-47f6-9497-c5348b45e635.png)

### 2.1.3 批量添加计算字段

```python
def add_pivot_calculated_fields_com(
    file_path: str,
    sheet_name: str,
    pivot_name: str | None = None,
    pivot_index: int = 1,
    fields: dict | None = None,
    refresh: bool = True,
    add_to_values: bool = True,          # 是否自动加入“值”区域
    skip_if_exists: bool = True,         # 同名计算字段已存在时：True=跳过；False=尝试删除后重建（不稳定）
    verbose: bool = True,                # 是否打印过程日志
):
    """
    在已有 PivotTable 上批量添加 Calculated Fields（计算字段）
    - 可选：自动加入值区域
    - 若遇到同名计算字段/异常：记录并跳过，继续执行剩余字段
    - 执行结束：print 未成功字段清单

    fields 格式：
    {
      "金额加权平均定价2": {"formula": "='金额*定价(%)'/放款金额", "number_format": "0.00", "caption": "金额加权平均定价2"},
      ...
    }
    """
    if not fields:
        raise ValueError("fields 不能为空，例如：{'字段名': {'formula': '=<expr>', 'number_format': '0.00%'}}")

    failed = []  

    excel_app = Dispatch("Excel.Application")
    excel_app.Visible = False
    excel_app.DisplayAlerts = False

    # 常量
    xlDataField = 4  # PivotField 作为“值”区域

    try:
        wb = excel_app.Workbooks.Open(file_path)
        ws = wb.Worksheets(sheet_name)

        # 选取 PivotTable（建议用 pivot_name）
        pvt = ws.PivotTables(pivot_name) if pivot_name else ws.PivotTables(pivot_index)

        if verbose:
            try:
                print(f"[Pivot] sheet={sheet_name} pivot={pvt.Name} cache_index={pvt.PivotCache().Index}")
            except Exception:
                print(f"[Pivot] sheet={sheet_name}")

        for field_name, cfg in fields.items():
            cfg = cfg or {}
            formula = cfg.get("formula")
            if not formula:
                failed.append({"field": field_name, "reason": "missing_formula", "detail": "cfg 中缺少 formula"})
                continue

            number_format = cfg.get("number_format", "General")
            caption = cfg.get("caption", field_name)

            # 1) 检查同名计算字段是否已存在
            exists = False
            try:
                _ = pvt.CalculatedFields(field_name)  # 存在则不报错
                exists = True
            except Exception:
                exists = False

            if exists and skip_if_exists:
                failed.append({"field": field_name, "reason": "already_exists", "detail": "计算字段已存在，按配置跳过"})
                if verbose:
                    print(f"[Skip] {field_name} 已存在，跳过")
                continue

            # 2) （可选）尝试删除旧计算字段（注意：删除常因被其他透视表引用而失败）
            if exists and not skip_if_exists:
                try:
                    pvt.CalculatedFields(field_name).Delete()
                except Exception as e:
                    failed.append({"field": field_name, "reason": "delete_failed", "detail": str(e)})
                    if verbose:
                        print(f"[Fail] 删除旧计算字段失败: {field_name} | {e}")
                    continue

            # 3) 新增计算字段
            try:
                pvt.CalculatedFields().Add(Name=field_name, Formula=formula)
                if verbose:
                    print(f"[OK] Add CalculatedField: {field_name} | {formula}")
            except Exception as e:
                failed.append({"field": field_name, "reason": "add_failed", "detail": str(e)})
                if verbose:
                    print(f"[Fail] Add CalculatedField 失败: {field_name} | {e}")
                continue

            # 4) 是否加入“值区域”
            if add_to_values:
                try:
                    # 更直接：把 PivotField 放到值区域（兼容性更高）
                    pf = pvt.PivotFields(field_name)
                    pf.Orientation = xlDataField
                    pf.NumberFormat = number_format
                    # 如果你想控制显示名，可在 Excel 里再改；COM 里改 Name 有时会触发异常
                    if verbose:
                        print(f"[OK] Add to Values: {field_name} | format={number_format}")
                except Exception as e:
                    # 加值区域失败不影响“计算字段已创建”，但按你的要求也算未成功设置
                    failed.append({"field": field_name, "reason": "add_to_values_failed", "detail": str(e)})
                    if verbose:
                        print(f"[Fail] 加入值区域失败: {field_name} | {e}")
                    continue

        # 5) 刷新（尽量只刷新当前透视表，避免 RefreshAll 造成锁/异步）
        if refresh:
            try:
                pvt.RefreshTable()
            except Exception:
                try:
                    wb.RefreshAll()
                    excel_app.Calculate()
                    try:
                        excel_app.CalculateUntilAsyncQueriesDone()
                    except Exception:
                        pass
                except Exception:
                    pass

        wb.Save()
        wb.Close()

    finally:
        excel_app.Quit()

    # 汇总输出
    if failed:
        print("\n===== 未成功设置的计算字段（含跳过项）=====")
        for item in failed:
            print(f"- {item.get('field')} | {item.get('reason')} | {item.get('detail')}")
        print("===== 结束 =====\n")
    else:
        print("\n全部计算字段设置成功。\n")

    return failed
```

### 2.1.4 VBA批量添加计算字段 

```vb
Dim pf As PivotField
    On Error Resume Next 

    With ActiveSheet.PivotTables(1)
      
        .CalculatedFields.Add "金额加权平均定价更新", "= '金额*定价(%)' /放款金额"
        Set pf = .PivotFields("金额加权平均定价更新")
        pf.Orientation = xlDataField
        pf.NumberFormat = "0.00"

        .CalculatedFields.Add "金额加权平均期限更新", "='金额*期限'/放款金额"
        Set pf = .PivotFields("金额加权平均期限更新")
        pf.Orientation = xlDataField
        pf.NumberFormat = "0.00"
    End With

End Sub
```

*   **在你的 Excel 文件（包含透视表那个），按 Alt + F11，打开 “VBA 编辑器”。**菜单点：插入 (Insert) → 模块 (Module)。右侧出现代码窗口，把上面代码整段粘贴进去。
    
    ![image.png](https://alidocs.oss-cn-zhangjiakou.aliyuncs.com/res/5VLqXLb7LagJxqX1/img/bd19c21f-94ea-440c-a81a-bf3ca5e9a7d2.png)
    
*   **点击左上角绿色运行符号，自定义宏名**
    

![image.png](https://alidocs.oss-cn-zhangjiakou.aliyuncs.com/res/5VLqXLb7LagJxqX1/img/0a2c78e4-496b-42d2-85f0-93715e167554.png)

*   **把上面的代码剪切进下面宏内部，并把上面代码删除**
    
    ![image.png](https://alidocs.oss-cn-zhangjiakou.aliyuncs.com/res/5VLqXLb7LagJxqX1/img/fcfd6d40-e751-4cd5-b98a-69be9fdff11a.png)
    
*   **继续点击绿色运行按钮，可以看到excel透视表中已经出现计算字段**
    

![image.png](https://alidocs.oss-cn-zhangjiakou.aliyuncs.com/res/5VLqXLb7LagJxqX1/img/d5c38d99-e632-4d8f-b6f2-5a991bfa094d.png)

![image.png](https://alidocs.oss-cn-zhangjiakou.aliyuncs.com/res/5VLqXLb7LagJxqX1/img/6f2e909b-551a-4d07-a719-e870e9d47071.png)

*   VBA详细介绍可参考：[EXCEL高级技巧-VBA从入门到放弃 - loaferWP - 博客园](https://www.cnblogs.com/loaferwp/articles/17594447.html)
    

# ppt替换更新图片

## 3.1 数据清洗与常用的堆积柱状图、折线图绘制

### 3.1.1 数据透视表操作python实现


```python
data = o.execute_sql(query,  
    hints={
        'odps.sql.submit.mode': 'script',                     # 多语句
        'odps.sql.validate.orderby.limit': 'false'            # 关闭 ORDER BY 无 LIMIT 校验
    }).open_reader().to_pandas()

data["apply_month"] = data["apply_month"].astype(str)
data['apply_model_score_category'] = data['apply_model_score_category'].astype(str)
data["order_cnt"] = data["order_cnt"].astype(float)
data["risk_pass_order_cnt"] = data["risk_pass_order_cnt"].astype(float)
data["risk_pass_rate"] = data["risk_pass_rate"].astype(float)
data["fund_pass_rate"] = data["fund_pass_rate"].astype(float)
data["fund_2h_pass_rate"] = data["fund_2h_pass_rate"].astype(float)
# 丢掉apply_model_score_category为空的数据
clean_data = data[data['apply_model_score_category'] != '空']

pivot_order_cnt = (
    clean_data.pivot_table(
        index="apply_month",
        columns='apply_model_score_category',
        values='order_cnt',
        aggfunc="sum",
        fill_value=0.0,
    )
    .sort_index()
) 
pivot_order_cnt ["总计"] = pivot_order_cnt.sum(axis=1)

pivot_risk_pass_order_cnt = (
    clean_data.pivot_table(
        index="apply_month",
        columns="apply_model_score_category",
        values="risk_pass_order_cnt",
        aggfunc="sum",
        fill_value=0.0,
    )
    .sort_index()
)
pivot_risk_pass_order_cnt["总计"] = pivot_risk_pass_order_cnt.sum(axis=1)
```

### 3.1.2 堆积柱状图与折线图实现


*   平时分析材料中经常会有堆积柱状图和折线图的绘制，考虑在python中封装函数方法，传入上述的数据透视表来更快地绘制相关图形，同时可以统一化字体大小、线条、柱体颜色等。同时将自动把生成图片存放在save\_dir工作目录下的 imgs，图片命名为“图表标题.png”格式，以实现后续按照名字匹配的方式进行一一替换
    

```python
# 堆积柱状图函数
def plot_stacked_bar_from_pivot(
    pivot: pd.DataFrame,
    cols: list,
    title: str,
    save_dir: str = None,
    suffix: str = ".png",
    colors: dict = None,
    width: float = 0.55,
    p: float = 1.0,
    fontsize: int = 14,
    top_label_fmt: str = None,
    segment_label_fmt: str = "{p:.1f}%",
    y_label: str = None,
    x_rotation: int = 0,
    dpi: int = 300,
    bbox_inches: str = "tight",
    hide_y_ticks: bool = True,          # y轴不显示刻度线
    hide_y_tick_labels: bool = True,   # 仅隐藏刻度线时一般够用；想更“干净”可设 True
    spine_off: tuple = ("top", "right", "left"),
):
    """
    从 pivot_table 画堆积柱状图，并在每段显示占比、柱顶显示总数。

    参数
    - pivot: DataFrame，index为x轴（如月份），columns为分组，values为数值
    - cols: 你希望的绘制列顺序（缺列自动补0）
    - title: 图标题，同时用于保存文件名（title + suffix）
    - save_dir: 保存目录；None 则默认当前工作目录下的 imgs
    - colors: dict，{col: color}。不传则用默认配色方案（预置6个）
    - width: 柱宽
    - p: 段内百分比展示阈值（单位：百分比，比如0.5表示<0.5%不标）
    - fontsize: 标签字号（段内 & 顶部）
    - top_label_fmt: 顶部总数格式化字符串，如 "{t:.1f}"；None 则自动根据数据大小选
    - segment_label_fmt: 段内占比格式（百分比），默认"{p:.1f}%"
    """

    if pivot is None or len(pivot) == 0:
        raise ValueError("pivot 为空，无法绘图。")

    # 默认保存目录
    if save_dir is None:
        save_dir = os.path.join(os.getcwd(), "imgs")
    os.makedirs(save_dir, exist_ok=True)

    # 默认配色方案（6个）
    default_palette = [
        "#2F6EEB",  # 蓝
        "#F39C12",  # 橙
        "#9e9e9e",  # 灰
        "#F1C40F",  # 黄
        "#34B4F4",  # 浅蓝
        "#6A5ACD",  # 紫 
    ]

    # 统一列：缺的补0，多的丢弃
    pv = pivot.copy()
    for c in cols:
        if c not in pv.columns:
            pv[c] = 0.0
    pv = pv[cols].copy()

    # 强制数值化
    for c in cols:
        pv[c] = pd.to_numeric(pv[c], errors="coerce").fillna(0.0)
    pv = pv.sort_index()

    totals = pv.sum(axis=1).astype(float)
    pct = pv.div(totals.replace(0, np.nan), axis=0).fillna(0.0) * 100

    # 颜色映射：先用默认，再被用户 colors 覆盖
    color_map = {}
    for i, c in enumerate(cols):
        color_map[c] = default_palette[i % len(default_palette)]
    if colors:
        color_map.update(colors)

    # 顶部总和标签格式：默认智能一点
    if top_label_fmt is None:
        max_total = float(np.nanmax(totals.values)) if len(totals) else 0.0
        # 大多数你这里是“亿/万”这种已经缩放过的值，保留1位小数更常见；否则整数
        top_label_fmt = "{t:.1f}" if max_total < 1000 else "{t:.0f}"

    fig, ax = plt.subplots(figsize=(13.33, 7.5))
    x = np.arange(len(pv.index))
    bottom = np.zeros(len(pv.index), dtype=float)

    # 统一的白底黑字框
    bbox_kw = dict(boxstyle="round,pad=0.2", facecolor="white", edgecolor="white", alpha=0.75)

    # 堆积柱
    for c in cols:
        vals = pv[c].values.astype(float)
        ax.bar(x, vals, width, bottom=bottom, label=c, color=color_map.get(c))

        # 段内百分比标签（黑色 + 白底框）
        pct_vals = pct[c].values
        y_center = bottom + vals / 2
        for i, (xc, yc, percent, v) in enumerate(zip(x, y_center, pct_vals, vals)):
            if v > 0 and percent >= p:
                ax.text(
                    xc, yc,
                    segment_label_fmt.format(p=float(percent)),
                    ha="center", va="center",
                    fontsize=fontsize,
                    fontweight="bold",
                    color="black",
                    bbox=bbox_kw,
                )

        bottom += vals

    # 柱顶总数标签（黑色）
    for xc, t in zip(x, totals.values):
        if np.isfinite(t) and t != 0:
            # 让顶部标签稍微上移一点
            offset = max(float(np.nanmax(totals.values)) * 0.01, 0.01)
            ax.text(
                xc, t + offset,
                top_label_fmt.format(t=float(t)),
                ha="center", va="bottom",
                fontsize=fontsize + 1,
                fontweight="bold",
                color="black",
            )

    # 标题/坐标
    ax.set_title(title, fontsize=fontsize + 6, fontweight="bold", pad=20)
    if y_label:
        ax.set_ylabel(y_label)

    ax.set_xticks(x)
    ax.set_xticklabels(pv.index.astype(str), rotation=x_rotation, fontsize=fontsize)

    # y轴：不需要刻度线
    if hide_y_ticks:
        ax.tick_params(axis="y", length=0)
    if hide_y_tick_labels:
        ax.set_yticklabels([])

    # 去边框（保持一致风格）
    for sp in spine_off:
        ax.spines[sp].set_visible(False)

    # 图例放底部居中
    ax.legend(ncol = len(cols), loc="lower center", bbox_to_anchor=(0.5, -0.15), frameon=False, fontsize=fontsize)

    plt.tight_layout()

    # 保存：标题=文件名
    safe_name = title.replace("/", "_").replace("\", "_").strip()
    img_path = os.path.join(save_dir, f"{safe_name}{suffix}")
    fig.savefig(img_path, dpi=dpi, bbox_inches=bbox_inches)
    plt.show()
    plt.close(fig)
```
```python
def plot_lines_from_df(
    df: pd.DataFrame,
    x_col: str,
    y_cols: list,
    title: str,
    *,
    labels: dict = None,                 # {y_col: 显示名}
    colors: dict = None,                 # {y_col: color} 覆盖默认色
    linestyles: dict = None,             # {y_col: '-', ':' ...}
    markers: dict = None,                # {y_col: 'o', ...}
    offsets: dict = None,                # {y_col: (dx, dy)} 用于数据标签偏移（单位：points）
    aligns: dict = None,                 # {y_col: {"ha":"center","va":"bottom"}} 或 {y_col: (ha, va)}
    label_cfg: dict = None,              # 精细控制打标点位与样式 {y_col: {"idx":..., "offset":(dx,dy), "ha":..., "va":...}}
    fmt: str = "{y:.1f}%",               # 数据标签格式化（若percent=True则传入 y*100）
    percent: bool = True,                # True: y当作0~1比例；False: 原值
    fontsize: int = 14,                  # 基础字号
    dpi: int = 300,
    save_dir: str = None,
    suffix: str = ".png",
    bbox_inches: str = "tight",
    y_lim: tuple = None,                 # (ymin, ymax)；None则自动
    auto_y_pad: float = 1.15,            # 自动y上边距系数
    spine_off: tuple = ("top", "right", "left"),
    hide_y_ticks: bool = True,
    hide_y_tick_labels: bool = True,
):
    """
    - 画多折线图（统一为 16:9 大图，适配PPT），并为每个点添加数据标签，同时保存到 imgs 目录（或指定 save_dir），返回图片路径。

    Parameters
    ----------
    - df : pd.DataFrame 数据源表。至少包含 x_col + y_cols。
    - x_col : str x轴列名，函数内部会转成字符串并按字典序排序后绘制。
    - y_cols : list 需要绘制的多条折线列名列表。缺列会自动补 NaN（不报错）。
    - title : str 图标题，同时作为输出图片文件名（会做简单的 / 和 \ 替换）。

    Keyword Args
    ------------
    - labels : dict, default None
        显示名映射：{列名: 图例显示名}。不传则用列名本身。
    - colors : dict, default None
        颜色映射：{列名: '#RRGGBB'}。不传则使用函数内置默认调色盘。
    - linestyles : dict, default None
        线型映射：{列名: '-', '--', ':', '-.'}。不传则默认 '-' ,常用：{'总计': ':'} 让总计虚线。
    - markers : dict, default None
        点型映射：{列名: 'o', 's', '^', ...}。不传默认 'o'。
    - offsets : dict, default None
        数据标签偏移：{列名: (dx, dy)}，单位 points。
        dy>0 标签在点上方；dy<0 在下方；用于避免标签重叠。
        未指定的列默认 (0, 8)。
    - fmt : str, default "{y:.1f}%"
        数据标签字符串格式。会使用 fmt.format(y=shown)。
        - percent=True 时 shown = 原值*100
        - percent=False 时 shown = 原值
    - percent : bool, default True
        True：把 y 当作 0~1 的比例，并自动把标签显示为百分比（乘以100）。
        False：y 按原值绘制与标注。
    - fontsize : int, default 14
        基础字号：x刻度/图例用 fontsize；点标签用 fontsize-2；标题用 fontsize+6。
    - dpi : int, default 300
        保存图片DPI。
    - save_dir : str, default None
        保存目录。None 时默认 os.getcwd()/imgs。
    - suffix : str, default ".png"
        输出图片后缀。
    - bbox_inches : str, default "tight"
        fig.savefig 的 bbox_inches 参数（tight 可减少白边）。
    - legend_ncol : int, default None
        图例列数。None 时自动 min(len(y_cols), 6)。
    - y_lim : tuple, default None
        y轴范围 (ymin, ymax)。None 时自动计算并按 auto_y_pad 留上边距。
        percent=True 时自动把 ymin 固定为 0。
    - auto_y_pad : float, default 1.15
        自动y上边距系数（只在 y_lim=None 时生效）。
    - spine_off : tuple, default ("top","right","left")
        需要隐藏的边框集合。若要保留底框线，请不要包含 "bottom"。
    - hide_y_ticks : bool, default True
        True：隐藏 y 轴刻度线（length=0）。
    - hide_y_tick_labels : bool, default True
        True：隐藏 y 轴刻度文字（更“干净”的PPT风格）。
    - label_cfg = {
      "某列": {"idx": "min"/"max"/-1/0/[0,2,4]/"all"/[], "offset":(dx,dy), "ha":"center", "va":"top"}
    """

    if df is None or len(df) == 0:
        raise ValueError("df 为空，无法绘图。")

    # 默认保存目录
    if save_dir is None:
        save_dir = os.path.join(os.getcwd(), "imgs")
    os.makedirs(save_dir, exist_ok=True)

    # 默认配色（6个）
    default_palette = [
        "#2F6EEB",  # 蓝
        "#F39C12",  # 橙
        "#9e9e9e",  # 灰
        "#F1C40F",  # 黄
        "#34B4F4",  # 浅蓝
        "#6A5ACD",  # 紫 
    ]

    d = df.copy()
    if x_col not in d.columns:
        if d.index.name == x_col:
            d = d.reset_index()
        else:
            raise KeyError(f"x_col='{x_col}' 不在 df.columns，且 df.index.name != '{x_col}'。请先 reset_index() 或传入正确的 x_col。")

    d[x_col] = d[x_col].astype(str)
    d = d.sort_values(x_col)

    # x轴：严格按顺序
    x_labels = d[x_col].tolist()
    x = np.arange(len(d))

    # 颜色映射：默认 -> 用户覆盖
    color_map = {c: default_palette[i % len(default_palette)] for i, c in enumerate(y_cols)}
    if colors:
        color_map.update(colors)

    linestyles = linestyles or {}
    markers = markers or {}
    offsets = offsets or {}
    labels = labels or {}
    aligns = aligns or {}
    label_cfg = label_cfg or {}

    # offsets 也支持 list/tuple 简写（兼容你原来用法）
    if isinstance(offsets, (list, tuple)):
        if len(offsets) != len(y_cols):
            raise ValueError(f"offsets 长度({len(offsets)})必须等于 y_cols 长度({len(y_cols)})")
        offsets = {c: off for c, off in zip(y_cols, offsets)}

    def _resolve_label_indices(y: np.ndarray, idx_rule):
        finite_idx = [i for i, v in enumerate(y) if np.isfinite(v)]
        if not finite_idx:
            return []

        if idx_rule is None or idx_rule == "all":
            return finite_idx

        if idx_rule == "min":
            yy = np.where(np.isfinite(y), y, np.nan)
            return [int(np.nanargmin(yy))] if np.isfinite(yy).any() else []
        if idx_rule == "max":
            yy = np.where(np.isfinite(y), y, np.nan)
            return [int(np.nanargmax(yy))] if np.isfinite(yy).any() else []

        if isinstance(idx_rule, (int, np.integer)):
            idx_list = [int(idx_rule)]
        elif isinstance(idx_rule, (list, tuple, set)):
            idx_list = [int(i) for i in idx_rule]
        else:
            raise ValueError(f"idx 只支持 None/'all'/[]/int/list/'min'/'max'，当前为: {type(idx_rule)}")

        out = []
        for i in idx_list:
            if i == -1:
                out.append(finite_idx[-1])
            elif 0 <= i < len(y) and np.isfinite(y[i]):
                out.append(i)

        # 去重保持顺序
        seen = set()
        out2 = []
        for i in out:
            if i not in seen:
                out2.append(i)
                seen.add(i)
        return out2

    fig, ax = plt.subplots(figsize=(13.33, 7.5))

    plotted_any = False
    y_all = []

    for c in y_cols:
        if c not in d.columns:
            d[c] = np.nan

        y = pd.to_numeric(d[c], errors="coerce").astype(float).values
        y_all.append(y)

        # 画线（确保有 label）
        ax.plot(
            x, y,
            label=labels.get(c, c),
            color=color_map.get(c),
            marker=markers.get(c, "o"),
            linewidth=2.5,
            linestyle=linestyles.get(c, "-"),
        )
        plotted_any = True

        # ---------- 打标策略（默认全标；可用 label_cfg 指定） ----------
        cfg = label_cfg.get(c, {}) if isinstance(label_cfg, dict) else {}

        # offset：label_cfg 优先，其次 offsets
        dx, dy = cfg.get("offset", offsets.get(c, (0, 8)))

        # ha/va：label_cfg 优先，其次 aligns（再按 dy 推断 va）
        ha = "center"
        va = "bottom" if dy >= 0 else "top"

        # aligns 兼容 dict / tuple
        a = aligns.get(c, None)
        if isinstance(a, dict):
            if a.get("ha") is not None:
                ha = a["ha"]
            if a.get("va") is not None:
                va = a["va"]
        elif isinstance(a, (list, tuple)) and len(a) == 2:
            if a[0] is not None:
                ha = a[0]
            if a[1] is not None:
                va = a[1]

        if cfg.get("ha") is not None:
            ha = cfg["ha"]
        if cfg.get("va") is not None:
            va = cfg["va"]

        idx_rule = cfg.get("idx", "all")
        idxs_to_label = _resolve_label_indices(y, idx_rule)

        for i in idxs_to_label:
            yi = y[i]
            if not np.isfinite(yi):
                continue
            shown = (yi * 100.0) if percent else yi
            ax.annotate(
                fmt.format(y=shown),
                xy=(x[i], yi),
                xytext=(dx, dy),
                textcoords="offset points",
                ha=ha,
                va=va,
                fontsize=fontsize,
                color="black",
                fontweight="bold",
                annotation_clip=True,
            )

    if not plotted_any:
        raise ValueError("y_cols 没有可绘制列。")

    ax.set_title(title, fontsize=fontsize + 6, fontweight="bold")

    ax.set_xticks(x)
    ax.set_xticklabels(x_labels, rotation=0, fontsize=fontsize)

    ax.set_ylabel("")
    if hide_y_ticks:
        ax.tick_params(axis="y", length=0)
    if hide_y_tick_labels:
        ax.set_yticklabels([])

    # y范围：自动留白 or 指定
    if y_lim is not None:
        ax.set_ylim(*y_lim)
    else:
        y_concat = np.concatenate([v[np.isfinite(v)] for v in y_all if v is not None and np.any(np.isfinite(v))]) \
            if any(np.any(np.isfinite(v)) for v in y_all) else np.array([0.0])
        ymin = float(np.nanmin(y_concat)) if len(y_concat) else 0.0
        ymax = float(np.nanmax(y_concat)) if len(y_concat) else 1.0

        # 若是比例，默认从0开始更符合展示习惯
        if percent:
            ymin = 0.0

        # 上边距留白
        ymax = ymax * auto_y_pad if ymax > 0 else 1.0
        ax.set_ylim(ymin, ymax)
    
    # 去边框（保持一致风格）
    for sp in spine_off:
        ax.spines[sp].set_visible(False)

    # 图例放底部居中
    ax.legend(ncol=len(y_cols), loc="lower center", bbox_to_anchor=(0.5, -0.15), frameon=False, fontsize=14)
    plt.tight_layout()

    safe_name = title.replace("/", "_").replace("\", "_").strip()
    img_path = os.path.join(save_dir, f"{safe_name}{suffix}")
    fig.savefig(img_path, dpi=dpi, bbox_inches=bbox_inches)
    plt.show()
    plt.close(fig)

```

```python
### 老客月发起订单
plot_stacked_bar_from_pivot(
    pivot=pivot_order_cnt/1e4, #转化为万
    cols=["A", "B", "CD", "EF", "GH", "IJK"],
    title="老客月发起订单数（万）",
    p=0.5,               
)


### 老客月风险订单
plot_lines_from_df(
    df=agg_pct,
    x_col="apply_month",
    y_cols=["A", "B", "CD", "EF", "GH", "IJK", "总计"],
    title="老客月风险（订单）通过率",
    linestyles = {"总计": ":"},  # 总计设为虚线
    colors={"总计": "#111111"},
    fmt="{y:.1f}%",
    percent=True,
    y_lim = (0,1.0),
    label_cfg={
        "A": {"idx":'all', "offset": (0, 10),  "va": "bottom" , "ha" : "center"},
        "B": {"idx":'all', "offset": (0, 10),  "va": "bottom" , "ha" : "center"},
        "CD": {"idx": 'all', "offset": (0, 5),   "va": "bottom", "ha" : "center"},
        "EF": {"idx": 'all', "offset": (0, 5),  "va": "bottom", "ha" : "center"},
        "GH": {"idx": [0,1,2,3,4,5,6], "offset": (0, -5),  "va": "top", "ha" : "center"},
        "IJK":  {"idx": [0,1,2,3,4,5,6], "offset": (0, -5), "va": "top", "ha" : "center"},
        "总计": {"idx": 'all', "offset": (0, 5), "va": "bottom", "ha" : "center"},
    }
)
```

### 3.1.3 ppt图片替换流程

*   在PPT模版中找到需要替换的图片（如图中老客发起订单数），设置好目标图片的大小，呈现位置，后续替换图片会继承原有模版中的大小和位置
    
*   点击要替换的图片\-> 开始 -> 右侧选择 -> 选择窗格 -> 最右侧这一模块的名字设置为图片标题的名字，与3.1.2中生成的图片名字保持一致
    

```python
def replace_pictures_by_name_batch(
    ppt_path,
    img_dir,
    out_path=None,
    suffix=".png",
    dry_run=False,        # 是否试运行
    allowed_types=None,   # 例如 {MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.CHART}
    strict=False,         # True：图片缺失就报错；False：跳过
):
    """
    批量替换：图片文件名(不含后缀) = PPT 形状名称（shape.name）。

    - 会遍历 PPT 全部页的 shapes
    - 若 shape.name 在 img_dir 中存在同名图片，则删除该 shape 并按原 left/top/width/height 插入图片
    - dry_run=True：试运行，只打印“将要替换哪些 shape、用哪张图片替换”，但不删除 shape、不插入图片、不保存文件
    - allowed_types 为 None 时不限制 shape 类型；否则仅处理指定类型（如 CHART / PICTURE）
    - strict=True 时：遇到有匹配形状但找不到图片会抛错（方便做完整性校验）
    """
    assert os.path.isfile(ppt_path), f"PPT 不存在: {ppt_path}"
    assert os.path.isdir(img_dir), f"图片目录不存在: {img_dir}"

    # 建立：name -> img_path（同名覆盖取最后一个）
    files = {}
    for fn in os.listdir(img_dir):
        if not fn.lower().endswith(suffix.lower()):
            continue
        stem = os.path.splitext(fn)[0].strip()
        if stem:
            files[stem] = os.path.join(img_dir, fn)

    if not files:
        print(f"目录内无 {suffix} 图片：{img_dir}")
        return {"hit": 0, "miss": 0, "total_imgs": 0, "total_shapes": 0}

    prs = Presentation(ppt_path)

    hit, miss, skipped = 0, 0, 0
    total_shapes = 0

    # 收集PPT里“可参与匹配”的shape名
    ppt_names = set()

    # 遍历每页每个 shape
    for si, slide in enumerate(prs.slides):
        # list()：因为下面会 remove，避免迭代器被破坏
        for shp in list(slide.shapes):
            total_shapes += 1
            name = getattr(shp, "name", "").strip()
            if not name:
                continue

            # 类型过滤（可选）
            if allowed_types is not None and shp.shape_type not in allowed_types:
                skipped += 1
                continue
            ppt_names.add(name)
            
            # 仅当存在同名图片时才替换
            img_path = files.get(name)
            if not img_path:
                miss += 1
                continue

            left, top, width, height = shp.left, shp.top, shp.width, shp.height
            print(f"[替换] 第{si+1}页《{name}》 type={shp.shape_type} -> {img_path}")

            if dry_run:
                hit += 1
                continue

            # 先删除原 shape
            slide.shapes._spTree.remove(shp._element)

            # 再插入图片并保持尺寸 + 恢复 name
            pic = slide.shapes.add_picture(img_path, left, top)
            pic.width, pic.height = width, height

            # 关键：把新图片 shape 的 name 改回原来的 name（保持映射逻辑）
            try:
                pic.name = name
            except Exception as e:
                # 极少数情况下会失败，但一般都可用
                print(f"[警告] 无法设置图片名称为 {name}（第{si+1}页）：{e}")

            hit += 1

    # ✅ 直接打印：文件夹里没匹配上的图片名
    img_only_names = sorted(set(files.keys()) - ppt_names)
    print("\n========== 文件夹里没匹配上的图片（img_only_names）==========")
    if img_only_names:
        for n in img_only_names:
            print(n)
    else:
        print("无（全部图片都能在PPT里找到同名shape）")
    print("===========================================================\n")

    # strict 校验：如果要求“形状有名字的都必须有图”，可在这里加强
    if strict:
        # 找出 img_dir 里有图但PPT里找不到同名shape，也可以加校验；此处仅对“命名shape找不到图”提示
        if miss > 0:
            raise ValueError(f"strict=True：存在 {miss} 个命名shape未找到对应图片（suffix={suffix}）。")

    if not dry_run:
        save_path = out_path or ppt_path
        prs.save(save_path)
        print(f"已保存到：{save_path}")

    print(f"批量替换完成：命中 {hit}，未命中(无同名图片) {miss}，跳过(类型不符) {skipped}，候选图片 {len(files)}，遍历shape {total_shapes}。")
    return {"hit": hit, "miss": miss, "skipped": skipped, "total_imgs": len(files), "total_shapes": total_shapes}
```

# 参考资料

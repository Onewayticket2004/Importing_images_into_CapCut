# Importing_images_into_CapCut

## 功能说明

这个工具可以读取 Excel 文件中的时间戳，并将从 Excel 中提取的图片按指定时间插入到剪映（JianyingPro）草稿中，生成一个新的项目。

## 前提条件

1. 已安装剪映专业版
2. 已安装 Python 3.x
3. 已安装必要的 Python 包：
   ```bash
   pip install openpyxl
   pip install pycapcut
   ```

## 使用步骤

### 步骤 1：设置create_capcut_from_excel.py参数
修改代码块的“默认配置”部分即可
| 参数 | 必需 |
|------|------|
| `DEFAULT_EXCEL_FILE` | excel文件路径（绝对路径） | 
| `DEFAULT_DRAFT_DIR` | 项目所在路径（绝对路径，剪映可查看） ，一般为C:/Users/name/AppData/Local/JianyingPro/User Data/Projects/com.lveditor.draft|
| `DEFAULT_DRAFT_NAME` | 项目名称 | 
| `DEFAULT_RESOLUTION` | 默认画幅，如(1920,1080) |


### 步骤 2：创建剪映草稿
运行程序，输出总用时。


### 步骤 3：在剪映中查看草稿

1. 打开剪映专业版
2. 在草稿列表中找到你创建的草稿（例如："我的草稿"）
3. 双击打开草稿，查看视频轨道上的图片序列



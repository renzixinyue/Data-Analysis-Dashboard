# 学生成绩数据分析可视化大屏

本项目用于分析学生考试成绩（月考 vs 期中），并通过 Web 大屏进行可视化展示。

## 📁 项目结构

*   `diyiciyuekao.xlsx`: 第一次月考成绩单（源数据）
*   `qizhognchengji.xlsx`: 期中考试成绩单（源数据）
*   `analysis_result.xlsx`: 自动生成的分析结果（由脚本生成）
*   `analyze_data_full.py`: 数据清洗与分析脚本
*   `export_data_to_json.py`: 将分析结果导出为 Web 端可用的 JSON 数据
*   `dashboard/`: 大屏前端代码目录
    *   `index.html`: 大屏主页（Web 端入口，原 dashboard.html）
    *   `data.json`: 数据文件
    *   `assets/`: 静态资源（CSS, JS）

## 🚀 如何运行

### 方法一：使用一键脚本（推荐）

1.  **启动大屏**：双击运行 `start_dashboard.bat`。
    *   会自动启动一个本地服务器，并提示您在浏览器打开 `http://localhost:8000/`。
2.  **更新数据**：如果您替换了 Excel 源文件，请双击运行 `update_data.bat`。
    *   会自动重新计算分析结果并更新大屏数据。

### 方法二：手动运行

1.  **安装依赖**（首次运行前）：
    ```bash
    pip install pandas openpyxl matplotlib seaborn
    ```

2.  **更新数据分析**：
    ```bash
    python analyze_data_full.py
    python export_data_to_json.py
    ```

3.  **启动大屏**：
    进入 `dashboard` 目录并启动 HTTP 服务器：
    ```bash
    cd dashboard
    python -m http.server 8000
    ```
    然后在浏览器访问：[http://localhost:8000/](http://localhost:8000/)

## ⚠️ 注意事项

*   大屏使用了 `fetch` API 加载数据，**必须**通过 HTTP 服务器（如 `python -m http.server`）运行，不能直接双击 `html` 文件打开，否则会因为 CORS 跨域安全策略导致数据无法加载。

# 个股研究与报告生成工具

公开行情、收益表现、因子归因、基本面分析与 Word 报告生成脚本。

## 内容与运行

`code/` 保留以创元科技（000551）为案例的数据获取、收益表现、因子归因、财务提取、图表和 Word 报告生成器。

```bash
python -m pip install -r requirements.txt
python code/fetch_stock_health_data.py
```

随后按分析目标运行 `analysis_return_performance.py`、`analysis_factor_attribution.py`、`extract_fundamental_data.py` 和对应报告生成器。Python 3.10+。入口中的股票代码、日期和字段来自原研究案例，其他标的需要相应调整。

数据由公开行情接口或用户自行取得的财报提供，输出到 `data/` 和 `output/`。原始研报全文、实习内部报告与附件不随仓库发布。`extract_zhongbao_text.py` 需要将有权使用的财报保存为 `data/中报.pdf`。

当前是案例脚本集合，尚未封装成任意股票通用 API；本次只验证源码语法，未重新下载市场数据或核验报告结论。

## 项目说明

维护者：[lhowd301-lab](https://github.com/lhowd301-lab)。本仓库整理自既有课程、研究或工具项目，开发与文档整理使用了 AI 辅助。代码可运行性与研究结论有效性分别说明，不将历史模拟结果视为实盘业绩。

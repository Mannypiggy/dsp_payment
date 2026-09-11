# DSP 月度付款申请助手

把 HUION 美国 Amazon DSP 每月付款申请流程自动化：

> 上传文件 → 一键处理 → 自动核对 → 下载付款申请包

## 运行

```bash
cd dsp_payment
pip install -r requirements.txt
streamlit run app.py
```

## 测试

```bash
cd dsp_payment
python -m pytest tests/ -q
```

真实历史回归测试依赖 `C:\Users\41250\Desktop\2026年7月Huion-US对账单\` 下的 2026-07 文件，
若文件不存在，`tests/test_*july*.py` / `test_email.py` / `test_rebate.py` 会失败（属于预期）。

## 目录结构

```
dsp_payment/
├─ app.py                      # Streamlit 主页面
├─ pipeline.py                 # 一键处理编排（26 步）
├─ config.yaml                 # KPI/返点/邮件/收款方/店铺 等配置
├─ models/payment_summary.py   # 统一数据对象 PaymentSummary
├─ rules/                      # 业务规则（无 IO）
│  ├─ month_rules.py           #   月份编码/大促月
│  ├─ dsp_mapping.py           #   DSP 字段映射 + Strategy 识别
│  ├─ promoted_ads_mapping.py  #   广告类型转换 + SD 修正
│  └─ kpi_rules.py             #   KPI 计算 + 口径保护
├─ processors/                 # 业务处理（含 IO）
│  ├─ order_summary_processor.py   # Order Summary → DSPOrder
│  ├─ promoted_ads_processor.py    # 已推广 → 转换/聚合
│  ├─ reconciliation_processor.py  # 对账单（DSP数据源/已推广/DSP+SA）
│  ├─ spend_detail_processor.py    # DSP花费明细
│  ├─ spend_ratio_processor.py     # 店铺花费占比（.xls 兼容）
│  ├─ kpi_processor.py             # KPI 计算/核对/写回
│  ├─ meeting_processor.py         # 会议记录 → KPI 原因
│  ├─ email_processor.py           # 付款申请邮件
│  ├─ package_processor.py         # ZIP + 清单
│  └─ validation_processor.py      # 数据总校验
├─ utils/                      # 工具
│  ├─ excel_utils.py           #   Sheet 识别/清空/重算/格式签名
│  ├─ validation.py            #   格式保护对比
│  ├─ file_utils.py            #   文件类型识别
│  ├─ xls_utils.py             #   .xls → .xlsx 兼容
│  └─ config_loader.py         #   config.yaml 加载
└─ tests/                      # 113 个测试（含真实 2026-07 回归）
```

## 关键设计

- **只改数据单元格 value，不重建 Workbook**：禁止 pandas 重写。
- **统一数据对象**：所有 Excel/KPI/邮件/ZIP 从同一个 `PaymentSummary` 读取。
- **金额口径字段级绑定**（`rules/amount_fields.py`）：`dsp_cost` / `rebate_amount` / `payment_amount`
  三个金额语义明确区分，每个目标字段固定绑定，不再有全局口径开关。
- **店铺花费占比**：店铺名称与固定比例来自模板自身，程序不新增/删除/改名/改比例，
  不根据 Order Summary 推导店铺。
- **对账单 Pivot 保护**：DSP+SA 的策略表/SA 表是 Pivot 输出、综合行是 GETPIVOTDATA 公式，
  程序只更新两个数据源 Sheet。openpyxl 往返会破坏 Pivot 定义，故优先用 Windows Excel COM
  （`utils/excel_com.py`）写入数据源 + RefreshAll；COM 不可用时回退 openpyxl 并明确告警。
- **旧版 .xls**：优先 Excel COM 原格式编辑（保留 .xls），不可用时转 .xlsx 并明确告警，不静默转换。

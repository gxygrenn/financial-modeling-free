# 免费金融建模工具（替代 Claude for Financial Services）

> 来源仓库：<https://github.com/anthropics/financial-services>（Apache 2.0 协议，已保留版权）
> 改造目标：**去掉对 Claude 运行时 + 付费金融终端的依赖，用免费数据源在本地直接跑通。**

---

## 一、结论先行

Anthropic 那个仓库的价值分三层，本目录告诉你哪层能免费化：

| 组成部分 | 原依赖 | 能否免费化 | 本目录方案 |
|---|---|---|---|
| **金融方法论**（DCF/可比公司/LBO/三表） | 纯 Markdown 知识 | ✅ 100% 可免费 | 已迁移为 `dcf_model.py` |
| **运行时**（触发/执行） | Claude Code 订阅 / Managed Agents API | ✅ 可用本机 Python 替代 | 直接 `python dcf_model.py` |
| **数据源** | 付费 MCP（S&P/FactSet/PitchBook/Daloopa…） | ⚠️ 机构级数据无法免费，只能用免费渠道替代 | Yahoo Finance / SEC EDGAR |

**一句话**：方法论和"跑模型"的能力完全能免费化；唯一替代不了的是机构级数据（那种只有 Bloomberg/万得/彭博终端才有），个人用免费数据源顶替即可。

---

## 二、直接上手

### 1. 装依赖（一次性）

本机装了 Clash（127.0.0.1:7890）等系统代理，pip 直连会被拦（403/TLS 指纹被 CDN 拒）。**实测可用的安装命令**如下：

```bash
# 关键三点：① --proxy "" 让 pip 绕过系统代理 ② 用阿里源（清华源会拦 pip 指纹）③ 清空代理环境变量
VENV="C:/Users/49928/.workbuddy/binaries/python/envs/default"
unset http_proxy https_proxy HTTP_PROXY HTTPS_PROXY ALL_PROXY all_proxy
export NO_PROXY="*" no_proxy="*"
"$VENV/Scripts/python.exe" -m pip install --proxy "" -i https://mirrors.aliyun.com/pypi/simple openpyxl yfinance pandas
```

### 2. 生成一份 DCF 估值模型

```bash
python dcf_model.py AAPL          # 苹果
python dcf_model.py MSFT          # 微软
python dcf_model.py 0700.HK       # 腾讯（港股）
```

输出一个标准投行版式的 Excel：`AAPL_DCF_Model_2026-10-03.xlsx`，包含：
- 市场数据 + 关键输入（蓝色 = 输入）
- 5 年自由现金流投影
- 终端价值 + EV→股权价值→每股价值桥
- **WACC vs 终端增长率 5×5 敏感性表**
- 独立的 WACC 计算页

### 3. 调参数

```bash
python dcf_model.py AAPL --growth 12 10 9 8 7 --wacc 0.09 --terminal 0.03
```

---

## 三、保留了哪些方法论（免费）

这些是 Anthropic 仓库里最值钱、且与 Claude 无关的纯知识，已融入脚本：

### DCF（dcf-model 技能）
- EBIT → NOPAT → 加回 D&A → 减 CapEx → 减 ΔNWC = 无杠杆自由现金流
- WACC 用 CAPM：`Cost of Equity = 无风险利率 + β × 股权风险溢价`
- 终端价值（永续增长法）：`TV = FCF×(1+g)/(WACC−g)`，要求 g < WACC
- 年中惯例贴现、三情景（熊/基/牛）、敏感性表

### 可比公司分析（comps-analysis 技能）
核心指标与统计基准法（中位数 + 四分位），公式：EV/Revenue、EV/EBITDA、P/E、PEG 等。

### LBO / 三表模型 / 并购模型
同为标准投行方法论，逻辑与 DCF 同源（详见源仓库 `plugins/vertical-plugins/`）。

---

## 四、免费数据源对照表

| 付费 MCP（原仓库） | 免费替代 | 覆盖范围 |
|---|---|---|
| S&P Capital IQ / Kensho | SEC EDGAR（官方免费） | 美股财报 |
| FactSet | Yahoo Finance (yfinance) | 全球行情+基础财务 |
| Daloopa | yfinance / EDGAR | 历史财务 |
| LSEG / PitchBook | 无免费等价（PE/机构数据） | ❌ 只能手动输入 |
| Morningstar / Moody's | 无免费等价（评级） | ❌ 只能手动输入 |

**做不到的**：私募股权回报数据、债券评级、一级市场交易倍数——这些个人本来也拿不到。

---

## 五、A股 / 港股怎么办

本脚本用 yfinance，可覆盖美股、港股（`0700.HK`）。**A 股**则建议用 WorkBuddy 已内置的免费金融数据技能（行情、财报、资金流、估值等），把数据填进本脚本 `fetch_financials` 的返回值即可。

---

## 六、合规提醒

所有输出仅供**学习与研究**，不构成投资建议。数据来自免费公开渠道，使用前请自行核对。

---
name: "tableau-typhoon-twbx"
description: "Builds a Tableau-Public-uploadable typhoon-path .twbx + .zip from a raw Excel. Invoke when the user wants a Tableau 台风路径 case file / .twbx for rain-classroom or Public upload."
---

# Tableau 台风路径案例构建 (typhoon-twbx)

一键把原始 Excel（含 `台风信息` sheet 的台风轨迹点数据）转成：

- `basename.twbx` — 可被 Tableau Desktop / Public 直接打开的打包工作簿（内嵌 hyper 数据提取）
- `basename.zip` — 外层 zip（部分课程/雨课堂只接受 `.zip` 类附件，内含 `.twbx`）
- 可选：附带一张 `basename_预览.png` 路径预览图

全程 Python 自动完成，不依赖 Tableau Desktop 本机安装。

---

## 触发条件

用户说出或 imply 了以下任一意图时，调用本 skill：

| 用户原话 | 对应动作 |
|---------|---------|
| "用这份台风数据做一个 Tableau 台风路径案例" | 读取 Excel → 走 builder 出 .twbx / .zip |
| "把台风 1949-2023 做成 Tableau 文件交雨课堂作业" | 同上，并额外产出一张路径预览图 |
| "打包成 .twbx 给我上传 Tableau Public" | 同上，Hyper 用 database_version=1 兼容旧版 Public |
| 修复用户贴出的 Tableau 加载报错 (`Error(...): missing required attribute` / `Error: 不是数据提取`) | 先确认 `.twb` XML 结构 / `<extract enabled='true'>` 是否存在，再跑 builder 重生 |

---

## 工作流（必按顺序）

### 0. 依赖自检

```bash
pip install tableauhyperapi pandas openpyxl
```

### 1. 读取 Excel → 字段对齐

原始 Excel 列（顺序可乱，名称必须匹配）：

| 字段 | 类型 | 备注 |
|------|------|------|
| 移动方向 | string | WNW / W / N 等 |
| 台风级别 | string | 6级~17级以上 |
| 现在气压 | int | hPa |
| 移动速度 | real | 混合 int/str，pandas 用 `to_numeric(errors='coerce')` 清洗 |
| 预报时间 | int | 纯数字（非日期） |
| 国际名称 | string | |
| 现在纬度 | real | Geo.Latitude |
| 预报纬度 | real | |
| 现在时间 | int | |
| 中文名称 | string | |
| 预报气压 | real | |
| 预报经度 | real | |
| 现在风速 | int | m/s |
| 现在经度 | real | Geo.Longitude |
| 台风编号 | int | 主键 |
| 日期时间 | datetime | 原生 datetime64 |

### 2. 生成 Hyper 提取

```python
with HyperProcess(Telemetry.DO_NOT_SEND_USAGE_DATA_TO_TABLEAU,
                  parameters={'default_database_version': '1'}) as hyper:
```

**必须用 ASCII 文件名**（Windows 中文路径会导致 Tableau 加载失败）：`Data/Extracts/typhoon.hyper`。
**必须开启 `default_database_version=1`**（兼容 Tableau Public 旧版本）。

### 3. 生成 .twb XML（最容易踩坑的地方）

#### 合法结构 & 常见死坑

| 位置 | 正确 | 错误写法 |
|------|------|---------|
| `<worksheets>/<worksheet>/<table>` 内 | 先 `<view>` 再 `<style>` 再 `<panes>` 再 `<rows>` `<cols>` | `<rows>/<cols>` 放进 `<view>` 里 → `error: element 'rows' is not allowed` |
| `<window>` 子元素 | 必须同时有 `<cards>`（`edge → strip → card(type=mark/pages/filters/columns/rows)`）和 `<viewpoint>` | 只写 `<mark>` 或只写 `<zoom>` → `error: element 'window' is not allowed` |
| 数据源引用 hyper | `<datasource>` 末尾加 `<extract enabled='true'>`，内含 `<connection class='hyper' .../>` | 没有 `<extract>` → "保存到 Public 必须使用数据提取" |
| 字段语义角色 | 直接 `role='measure' type='quantitative'`，**不要写** `semantic-role='[Latitude]'` 等字符串 | Tableau 版本差异会解析失败 |
| map 底图 | 完全不要写 `<mapsources>` / `<mapprovider>` / `<map-base-style>` | 非声明元素 → `error: no declaration found for element 'mapsources'` |

#### worksheet 编码

```xml
<mark class='Line' />
<encodings>
  <color column='[{d_src}].[none:台风级别:nk]' />
  <size  column='[{d_src}].[sum:现在风速:qk]' />
  <path  column='[{d_src}].[none:日期时间:qk]' />
  <lod   column='[{d_src}].[none:台风编号:nk]' />
</encodings>
<rows>[{d_src}].[avg:现在纬度:qk]</rows>
<cols>[{d_src}].[avg:现在经度:qk]</cols>
```

### 4. 打包

```
<basename>.twbx  (zip)
├── <basename>.twb
├── Data/typhoon.xlsx
└── Data/Extracts/typhoon.hyper

<basename>.zip  (外层 zip，雨课堂/公选课作业用)
└── <basename>.twbx
```

### 5. 验证

```python
import xml.etree.ElementTree as ET
ET.parse("台风路径案例.twb")           # 1. XML 良构
import zipfile
zipfile.ZipFile("台风路径案例.twbx").testzip()   # 2. zip 完整
```

---

## 现成脚本

本 skill 目录内附了一个可直接跑的 builder：

```bash
python /workspace/.trae/skills/tableau-typhoon-twbx/typhoon_twbx_builder.py \
    --excel /workspace/typhoon.xlsx \
    --outdir /workspace/typhoon_out \
    --basename 台风路径案例
```

产出：
```
typhoon_out/
├── 台风路径案例.twb       # 合法 XML，ET.parse 通过
├── 台风路径案例.twbx      # Tableau Desktop/Public 可直接打开
├── 台风路径案例.zip       # 雨课堂可直接上传
└── Data/
    ├── typhoon.xlsx
    └── Extracts/typhoon.hyper
```

---

## 常见报错 → 对应修复

| 报错 | 含义 | 修复 |
|------|------|------|
| `Error(行,列): missing required attribute 'dim-percentage'` | `<view>` 里多写了非法 `<layout>`，它其实不是 `<view>` 的合法子元素 | 删掉整个 `<layout>` 节点，`rows/cols` 必须放在 `<table>` 顶层 |
| `Error: element 'mapsources' is not allowed` | `<view>` 里写了 `<mapsources>` | 删掉所有 `<mapsources>/<mapprovider>/map-base-style` 节点 |
| `Error: element 'window' is not allowed` | `<window>` 内容模型不完整 | 确保同时有 `<cards>` 和 `<viewpoint>` |
| `保存到 Public 的工作簿必须使用数据提取` | 数据源没有 `<extract enabled='true'>` | 在 `<datasource>` 末尾加 `<extract><connection class='hyper'/></extract>`，并把 `.hyper` 放在 `Data/Extracts/` |
| 只有 "加载无法成功完成" 无具体行号 | 几乎总是 `.twb` 里某个字符串属性值有非法字符，或 hyper 文件名含中文 | 统一用 ASCII 路径 + 重跑 builder |

---

## 限制

- 只支持 **台风路径（Line 标记 + 经纬度列）** 这一种 worksheet 布局。多 worksheet / 仪表板请另行扩展。
- 依赖 `tableauhyperapi`（已作为附属 pip 依赖，非 Tableau Desktop 必需）。
- Hyper 使用 `default_database_version=1`，兼容 Tableau 2023.1+。如果用户用的是 Tableau 2024 新版本，可去掉该参数以获得更快提取。

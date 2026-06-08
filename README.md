# SAN 下线影响分析报告生成说明

本仓库用于根据客户提供的 SAN 下线清单和两路 Fabric 的命令输出，生成下线影响分析结果，并对生成结果做自动化校验。

## 需求目标

输入数据包括：

- `wwn_with_vsan.txt`：待下线设备 WWN 清单，格式为 `wwn_vsan`。
- `Fabric A Command Result/show_fcalias.txt`：Fabric A 的 fcalias 输出。
- `Fabric A Command Result/show_zoneset_active.txt`：Fabric A 的 active zoneset 输出。
- `Fabric B Command Result/show_fcalias.txt`：Fabric B 的 fcalias 输出。
- `Fabric B Command Result/show_zoneset_active.txt`：Fabric B 的 active zoneset 输出。

需要输出：

1. 每个下线 WWN 对应的 `fcalias`、`zone` 和当前 `zone member`。
2. 如果某个 `zone member` 不在本次下线清单中，但与本次下线 WWN 处于同一个 Zone，则在 Excel 中标黄，并在说明列标识。
3. 保存实现思路、处理过程和测试结果，便于人工复核和后续复用。

## 文件说明

| 文件 | 说明 |
| --- | --- |
| `generate_san_offline_report.py` | 主生成脚本，解析输入文件并生成 `输出结果.xlsx` 和 `编码思路.txt`。 |
| `validate_san_offline_report.py` | 独立校验脚本，重新解析源文件并校验 `输出结果.xlsx` 是否与源数据一致。 |
| `输出结果.xlsx` | 运行脚本后在本地生成的 SAN 下线影响分析 Excel 结果；该文件是 XLSX 二进制文件，不提交到代码仓库。 |
| `编码思路.txt` | 脚本生成的编码思路与统计摘要。 |
| `测试结果.txt` | 校验脚本输出的测试结果。 |
| `README.md` | 本说明文件，集中记录实现思路、过程和测试方式。 |

## 实现思路

### 1. 读取下线清单

- 读取 `wwn_with_vsan.txt`。
- 每行按最后一个下划线拆分为 `WWN` 和 `VSAN`。
- 统一将 WWN 转为小写，减少大小写差异造成的匹配问题。
- VSAN `5` 对应 Fabric A，VSAN `6` 对应 Fabric B。
- 对下线清单进行去重，保持原始顺序。

### 2. 解析 fcalias

- 分别读取 Fabric A 和 Fabric B 的 `show_fcalias.txt`。
- 识别如下格式：

```text
fcalias name <alias名称> vsan <vsan编号>
  pwwn <wwn>
```

- 建立映射关系：

```text
(WWN, VSAN) -> [fcalias]
```

- 如果一个 WWN 对应多个 fcalias，则在输出中用逗号合并显示。

### 3. 解析 zoneset active

- 分别读取 Fabric A 和 Fabric B 的 `show_zoneset_active.txt`。
- 识别 Zone 头部：

```text
zone name <Zone名称> vsan <vsan编号>
```

- 收集 Zone 内成员，支持两类常见格式：

```text
pwwn <wwn> [alias]
fcid ... [pwwn <wwn>] [alias]
```

- 每个 Zone 会保存：
  - Fabric 名称
  - VSAN 编号
  - Zone 名称
  - 当前 Zone Member 原始行内容
  - Zone Member 对应的 WWN 和可选 alias

### 4. Zone 块去重

在测试过程中发现，部分命令输出中存在重复粘贴的 Zone 块。如果不去重，最终 Excel 会出现重复的 WWN/Zone 记录，导致统计结果偏大。

因此生成脚本会按以下字段对 Zone 块去重：

```text
Fabric + VSAN + Zone名称 + 完整Zone成员列表
```

只有完全一致的重复 Zone 块会被去除；如果 Zone 名称相同但成员不同，则不会被合并，以避免误删真实差异。

### 5. 生成输出行

- 遍历下线清单中的每个 `(WWN, VSAN)`。
- 查找该 WWN 对应的 fcalias。
- 查找该 WWN 所在的所有 active Zone。
- 对每个目标 WWN 按 Zone 展开，并输出该 Zone 的所有成员。
- 如果目标 WWN 未在 active zoneset 中找到，则在 Zone 列写入 `未在zoneset active中找到`，便于人工复核。

### 6. 标黄逻辑

对于每个 Zone Member：

- 如果该成员带有 WWN，且 `(成员WWN, 当前VSAN)` 不在本次下线清单中，则认为它不是本次下线对象，但与下线对象存在 Zone 业务关系。
- 该行在 Excel 中整行标黄。
- `说明` 列写入：

```text
本次不在提交的WWN列表中，但是和本次下线的WWN有Zone业务
```

### 7. 写入 Excel

- `generate_san_offline_report.py` 使用 Python 标准库 `zipfile` 和 XML 字符串生成最小 XLSX 文件，不依赖 `openpyxl`、`pandas` 等第三方库。
- Excel Sheet 名称为 `下线影响分析`。
- 表头字段为：

```text
Fabric, VSAN, WWN, Fcalias, Zone, Zone Member, 说明
```

## 处理过程

本次处理过程如下：

1. 运行生成脚本：

```bash
python3 generate_san_offline_report.py
```

2. 脚本读取以下源文件：
   - `wwn_with_vsan.txt`
   - `Fabric A Command Result/show_fcalias.txt`
   - `Fabric A Command Result/show_zoneset_active.txt`
   - `Fabric B Command Result/show_fcalias.txt`
   - `Fabric B Command Result/show_zoneset_active.txt`

3. 脚本解析并建立：
   - 下线 WWN/VSAN 清单
   - `(WWN, VSAN) -> fcalias` 映射
   - `(WWN, VSAN) -> active Zone` 映射
   - 每个 Zone 的当前成员列表

4. 脚本对重复 Zone 块进行去重。

5. 脚本生成：
   - `输出结果.xlsx`：本地生成的 Excel 二进制结果文件，仅用于交付或人工查看，不提交到代码仓库。
   - `编码思路.txt`：文本文件，可提交到代码仓库。

6. 运行校验脚本：

```bash
python3 validate_san_offline_report.py | tee 测试结果.txt
```

7. 校验脚本独立重新解析源文件，并对 `输出结果.xlsx` 做以下检查：
   - Excel ZIP 结构完整。
   - 表头符合预期。
   - Fabric 与 VSAN 映射正确。
   - 每个目标 WWN/VSAN/Zone 组合与源文件一致。
   - 每个 Zone Member 明细与源文件一致。
   - 外部成员说明列内容正确。
   - 需要标黄的行确实使用黄色样式。
   - 不需要标黄的行没有误标黄。

## 当前输出统计

本次生成和校验后的统计结果：

| 指标 | 数量 |
| --- | ---: |
| 下线清单 WWN 数量 | 148 |
| 匹配到 active Zone 的 WWN 数量 | 148 |
| 目标 WWN/VSAN/Zone 组合数量 | 260 |
| Excel 数据行数 | 1112 |
| 标黄非本次下线 Zone Member 行数 | 616 |

## 测试

已执行以下测试和检查：

```bash
python3 generate_san_offline_report.py
```

结果：成功生成本地 Excel 文件 `输出结果.xlsx` 和文本文件 `编码思路.txt`；其中 `输出结果.xlsx` 是二进制产物，不提交到代码仓库。

```bash
python3 validate_san_offline_report.py | tee 测试结果.txt
```

结果：校验通过，输出内容为：

```text
测试通过：输出结果.xlsx 与源文件解析结果一致
- 下线清单 WWN 数量：148
- 匹配到 active Zone 的 WWN 数量：148
- 目标 WWN/VSAN/Zone 组合数量：260
- Excel 数据行数：1112
- 标黄非本次下线 Zone Member 行数：616
- 表头、Fabric/VSAN 映射、Zone Member 明细、说明列和黄色样式均校验通过
```

```bash
python3 -m py_compile generate_san_offline_report.py validate_san_offline_report.py
```

结果：Python 语法检查通过。

```bash
python3 - <<'PY'
import zipfile, xml.etree.ElementTree as ET
with zipfile.ZipFile('输出结果.xlsx') as z:
    assert z.testzip() is None
    ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
    ET.fromstring(z.read('xl/styles.xml'))
print('xlsx zip and xml validation passed')
PY
```

结果：`输出结果.xlsx` 的 ZIP 完整性和核心 XML 解析检查通过。


## 二进制文件提交说明

当前不支持提交二进制文件，因此本次不应提交 `输出结果.xlsx` 这类 Excel 结果文件。需要查看结果时，请在本地运行生成脚本重新生成。

可提交的内容应以文本文件为主，例如：

- `generate_san_offline_report.py`
- `validate_san_offline_report.py`
- `README.md`
- `编码思路.txt`
- `测试结果.txt`

## 使用方式

如果源文件内容发生变化，重新执行以下命令即可更新结果并重新校验：

```bash
python3 generate_san_offline_report.py
python3 validate_san_offline_report.py | tee 测试结果.txt
```

如校验失败，应优先查看终端输出中的失败原因，再检查：

- `wwn_with_vsan.txt` 是否存在格式异常或重复项。
- Fabric A/B 的命令输出是否完整。
- `show_zoneset_active.txt` 中是否存在格式与当前解析规则不一致的新类型成员。
- 生成脚本和校验脚本的 Zone 解析规则是否需要同步调整。

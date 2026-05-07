# ETW (Excel To Word)

Excel转Word命令行工具，支持覆盖模式和智能匹配。

## 版本

当前版本：1.0

## 安装

1. 确保已安装 Python 3.10 或更高版本
2. 安装依赖：
```bash
pip install -r requirements.txt
```

## 使用方法

### 基本语法
```
etw.py -i <输入Excel> -o <输出Word> [-F] [--CK1:value1] [--CK2:value2] ... [--HM1:name1] [--HM2:name2] ...
```

### 参数说明

| 参数 | 说明 |
|------|------|
| `-i` | 输入Excel文件路径（必填） |
| `-o` | 输出Word文件路径（必填） |
| `-F` | 启用覆盖模式 |
| `--CK1` ~ `--CK4` | 根字段参数，用于匹配（最多4个） |
| `--HM1` ~ `--HM20` | 覆盖字段参数（最多20个） |

### 使用示例

#### 示例1：基础转换
```bash
etw.py -i source.xlsx -o dest.docx
```

#### 示例2：覆盖模式
```bash
etw.py -i source.xlsx -o dest.docx -F --CK1:试验标识
```

#### 示例3：指定覆盖字段
```bash
etw.py -i source.xlsx -o dest.docx -F --CK1:试验标识 --HM1:首轮结果 --HM2:首轮人员
```

## 版本历史

- **1.0** (2026-04-30) - 初始版本
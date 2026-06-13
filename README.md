# 老头表格助手

## 软件定位

老头表格助手是一个面向审计、财务、报表整理场景的 Excel 效率工具，主要运行环境为 Windows + Microsoft Excel。

本仓库用于维护工具源码、打包脚本和发布前检查资料。

## 主要功能

- 一簿按工作表拆分为多个文件
- 多表合并到一个新表
- 多簿到一簿
- 按指定列拆分工作表
- 生成带链接的工作表目录
- 批量删除工作表
- 批量更换多文件链接
- 按颜色汇总求和
- 按颜色清空内容
- 批量重命名文件
- 批量重命名工作表
- 数据穿透查询
- 按模板批量生成 Excel
- 选区 ROUND 保留两位

## 使用方式

普通用户优先使用发布包中的 `老头表格助手.exe`。

源码运行：

```powershell
python run_app.py
```

或：

```powershell
python -m src.excel_efficiency_toolkit.app
```

## 安全说明

- 批量删除、批量重命名、按颜色清空、按模板批量生成等功能会修改或生成较多文件，建议先备份或先用副本试跑。
- `.xls` 老格式在部分功能中不支持，建议先用 Excel 另存为 `.xlsx` 或 `.xlsm`。
- 带外部链接的 Excel 文件处理前建议先用副本试跑，并在处理后检查链接是否正常。
- 多数写入当前工作簿的功能不会自动保存，检查无误后再由用户自行保存。

## 文档

详细使用说明请查看：`使用说明.txt`

发布前检查请查看：`RELEASE_CHECKLIST.md`

## 开发与打包

安装依赖：

```powershell
pip install -r requirements.txt
```

发布前基础检查：

```powershell
python -m py_compile src\excel_efficiency_toolkit\app.py
python -m pytest tests/test_clear_by_color_ops.py --basetemp .pytest_clear_by_color_tmp
git diff --check
```

打包 Windows exe：

```powershell
powershell -ExecutionPolicy Bypass -File scripts\build_exe.ps1
```

打包产物：

```text
dist\老头表格助手.exe
```

## 当前状态

常用功能阶段性完成，当前版本号为 `0.1.0`。本阶段重点是发布收口、说明文档、帮助入口、版本信息和打包验证。

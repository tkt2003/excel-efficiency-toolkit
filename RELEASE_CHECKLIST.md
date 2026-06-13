# 发布检查清单

## 打包前检查

```powershell
git status --short
python -m py_compile src\excel_efficiency_toolkit\app.py
python -m pytest tests/test_clear_by_color_ops.py --basetemp .pytest_clear_by_color_tmp
git diff --check
Remove-Item -Recurse -Force .pytest_clear_by_color_tmp -ErrorAction SilentlyContinue
```

只改发布说明、帮助入口、版本号时，不需要跑全量测试；如改动 Excel 业务逻辑，应补充相关测试。

## 打包命令

```powershell
powershell -ExecutionPolicy Bypass -File scripts\build_exe.ps1
```

## 打包后检查

```powershell
Get-Item .\dist\老头表格助手.exe
```

- `dist\老头表格助手.exe` 是否存在。
- exe 文件名是否为 `老头表格助手.exe`，不能出现中文乱码。
- exe 能否正常启动。
- 主窗口标题是否显示 `老头表格助手 v0.1.0`。
- 主界面按钮是否正常显示。
- 【使用说明】入口是否能打开。
- exe 同目录没有 `使用说明.txt` 时，【使用说明】是否显示内置简版说明且不报错。
- 【关于】弹窗是否能打开。
- 【按颜色清空内容】按钮是否仍在。
- 【数据穿透查询】弹窗是否能打开。
- 【按模板批量生成 Excel】弹窗是否能打开。
- 日志区是否正常显示。

## 发布包建议

- 建议将 `老头表格助手.exe` 和 `使用说明.txt` 放在同一个文件夹发布。
- 如果只发 exe，也要保证软件内【使用说明】能显示内置简版说明。
- 发布前用副本文件完成最小人工冒烟。

## 不提交内容

- `dist/`
- `build/`
- `*.spec`
- `.pytest_clear_by_color_tmp`
- 临时规则表
- 临时 Excel 文件
- 测试输出文件

## 最小人工冒烟清单

1. 启动 exe。
2. 确认标题为 `老头表格助手 v0.1.0`。
3. 确认主按钮显示正常。
4. 点击【使用说明】，确认能打开 `使用说明.txt` 或显示内置简版说明。
5. 点击【关于】，确认版本号和阶段说明正确。
6. 确认【按颜色清空内容】按钮仍在。
7. 点击【数据穿透查询】，确认模式选择弹窗能打开，然后取消。
8. 点击【按模板批量生成 Excel】，确认模式选择弹窗能打开，然后取消。
9. 确认日志区正常显示操作记录。

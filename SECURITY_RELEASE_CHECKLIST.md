# 老头表格助手发布前安全与可靠性检查清单

本清单用于发布前快速自查，目标是用低成本方式减少低级风险。当前阶段保持简单、可执行，不做复杂安全工程化。

## 1. Git 与发布来源

- 确认当前分支为 master
- 确认 git status --short 为空，源码工作区干净
- 确认 git fetch origin 后，本地 master 与 origin/master 一致
- 确认 release zip 来自当前最新提交
- 不提交 dist/、build/、release/、*.spec、临时 Excel 文件

## 2. 敏感信息检查

- 不提交密钥、token、账号密码、真实客户文件
- 不提交客户底稿、财务数据、授权码清单
- 发布前用关键字搜索 api_key、secret、token、password、passwd、BEGIN PRIVATE KEY 等

示例命令：

```powershell
Select-String -Path .\* -Pattern "api_key|secret|token|password|passwd|BEGIN PRIVATE KEY" -Recurse
```

## 3. 用户文件安全

- 正式文件处理前建议用户自行备份
- 批量删除、批量重命名、批量替换链接等操作必须保留确认提示或备份提醒
- 不扩大文件读写范围
- 不默认修改未被用户选择的文件
- 操作日志不要写入完整敏感数据内容

## 4. 本地运行与数据上传

- 老头表格助手为本地 Excel 效率工具
- 不应新增上传用户 Excel 文件的逻辑
- 如未来新增联网、授权、云端或统计功能，必须单独评估并更新本清单

## 5. 依赖与打包

- 使用固定项目依赖环境打包
- 打包前确认源码界面可启动
- exe 冒烟通过后再生成 release zip
- release zip 解压后再次运行 exe
- 发布包只包含 exe 和必要说明文件

## 6. 发布前最小命令清单

```powershell
git status --short --branch
git fetch origin
git status --short --branch
git log --oneline --decorate -5
python -m py_compile src\excel_efficiency_toolkit\app.py
git diff --check
```

- 源码 run_app.py 冒烟
- exe 冒烟
- release zip 解压后冒烟

## 7. 不适合当前阶段的大工程

- 不引入账号系统
- 不引入复杂权限系统
- 不引入联网统计
- 不为安全名义重构全部文件处理逻辑

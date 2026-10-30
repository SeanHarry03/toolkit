# 一键更新 SVN 并同步到 Cocos

打开 `D:\Unity\toolkit\WorkSpace\publish\ExcelToJson.exe`，点击右上角 **一键更新并同步**。

每个工作表的 **其他配置 → 复制 JSON 到 Cocos** 开关默认关闭（已有旧规则也默认关闭）。仅为需要复制的工作表开启并保存规则。开关包含该工作表全部子表输出；关闭不影响生成 JSON。全部关闭时仍更新 SVN、解析并更新输出目录，但不会写入 Cocos，也无需有效的 Cocos 目标目录。

默认目录：
- SVN 和 Excel：`D:\SVNRoot\22.猛鬼`
- 普通 JSON 输出：`D:\SVNRoot\22.猛鬼\output`
- Cocos JSON：`D:\CocosProject\evil-spirit\assets\resources\GameConfig\V1`

展开“同步设置”可修改目录，点击“保存规则”保存。SVN 程序留空使用附带组件。复制工具到其他位置时，请保留整个 `publish` 文件夹，特别是 `rules.json`、`tools` 和运行库文件。

## 行为
1. 检查并更新 SVN；更新前后发现文本、属性或树冲突都会停止。不会自动解决冲突或提交 SVN。
2. 使用界面勾选的工作表、现有规则和自动行数设置解析，先将本次 JSON 写入独立目录。
3. 全部解析成功后，备份并更新输出目录全部 JSON，仅将开启复制的工作表 JSON 覆盖到 Cocos。不复制历史输出，不删除其他文件，不改 `.meta`。完成状态中的数量是实际复制到 Cocos 的文件数。

“开始转换”仍只导出 JSON，不更新 SVN、不复制到 Cocos。同步期间可点击“停止”；Excel 解析在当前阶段完成后停止，复制阶段停止会尝试回滚。

## 警告和失败
- 已不存在的工作表：沿用原工具的跳过行为，在日志和完成状态显示警告，保留目标的旧 JSON。当前规则中的“送礼-金宝箱”在真实 Excel 中不存在；未擅自改成“送礼-皮肤宝箱”，也未更改该表规则。
- Excel 文件缺失、解析失败、没有有效工作表或输出名冲突：停止同步。
- SVN 登录失败：先使用 TortoiseSVN 更新一次并保存凭据，然后重试。不会在配置文件保存 SVN 密码。
- SVN 更新存在冲突：先在 TortoiseSVN 中解决。停止更新后若提示工作副本锁定，使用 TortoiseSVN 的清理功能。
- 复制失败：尝试恢复已替换文件；若恢复失败，日志会列出文件及备份位置。

备份和本次生成文件保存在输出目录 `.sync-runs\时间-编号\` 中。`manifest.json` 记录原路径与备份文件；`backups` 保存覆盖前内容。备份不会自动清理。

## 验证记录
- 使用独立本地 SVN 仓库验证更新后读取新 Excel、中文子目录、冲突阻止同步、解析失败保护目标、被占用文件触发回滚、旧 JSON 排除、.meta 保留、配置持久化和取消。
- 原有 RowScanTests 和 WPF UiSmokeTests 回归通过；WPF 检查包含同步配置保存及展开设置后布局。
- 真实工作簿只读验证生成 13 个 JSON，输出在项目 `codex/real-data-check/`；没有执行真实仓库更新或覆盖真实 Cocos 文件。

测试命令（项目根目录）：
```powershell
& 'C:\Users\25330\.dotnet\dotnet.exe' run --project codex\SyncTests\SyncTests.csproj -p:NuGetAudit=false -p:UseSharedCompilation=false
& 'C:\Users\25330\.dotnet\dotnet.exe' run --project codex\RowScanTests\RowScanTests.csproj -p:NuGetAudit=false -p:UseSharedCompilation=false
& 'C:\Users\25330\.dotnet\dotnet.exe' run --project codex\UiSmokeTests\UiSmokeTests.csproj -p:NuGetAudit=false -p:UseSharedCompilation=false
```

## SVN 组件
来源：[VisualSVN 官方下载](https://www.visualsvn.com/downloads/)，Apache Subversion 1.15.0 x64 免安装包。
发行包地址：`https://www.visualsvn.com/files/Apache-Subversion-1.15.0-x64.zip`。
SHA-256：`98BB798DD6C70CC875B9BDACD600594351CF88F0A39A831A5055DF061C871BB5`。
`svn.exe` 的 VisualSVN 数字签名验证有效。许可证随包放在 `tools/svn/Licenses`。

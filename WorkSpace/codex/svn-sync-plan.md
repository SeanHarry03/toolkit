# SVN 更新与 Cocos JSON 同步

目标：在现有 WPF 工具中增加“一键更新并同步”，沿用当前勾选工作表及解析规则。

默认 SVN / Excel 目录：`D:\SVNRoot\22.猛鬼`。
默认目标：`D:\CocosProject\evil-spirit\assets\resources\GameConfig\V1`。

方案：在现有窗口增加 SVN、Cocos 目录配置与独立按钮，比额外脚本更容易复用规则、呈现日志和保存设置。保留普通转换入口。使用随工具提供的免安装 svn.exe，无需修改系统安装或 PATH。

运行：验证配置 → SVN 状态检查 → update --non-interactive --accept postpone → 再检查冲突 → 在独立暂存目录解析 → 验证 JSON → 备份并覆盖输出目录与 Cocos 同名 JSON。只复制本次生成的文件；保留其他文件及 .meta。失败停止，复制失败尝试回滚，并保留恢复备份。同步期间禁用重复操作；可取消 SVN 或在下一阶段停止。

不自动解决 SVN 冲突，不提交 SVN，不添加定时任务。本次真实项目只做读取检查；端到端验证在 codex 内独立 SVN 仓库和 Cocos 模拟目录完成。

## 实现步骤
- [x] 增加真实本地 SVN 仓库集成测试：更新后读取新 Excel、旧 JSON 排除、.meta 保留、更新失败/冲突/解析失败不覆盖。
- [x] 增加 SvnClient、AutomationRunner、配置字段；输出名限制为目录内 JSON，复制事务有备份和失败回滚。
- [x] 接入 WPF 目录选择、同步/停止按钮与日志，保存配置并保持普通转换行为。
- [x] 运行同步测试、原 RowScanTests、UiSmokeTests，构建并发布到 publish。

## 验证重点
缺失工作簿、缺失工作表（沿用原转换器跳过行为并显式警告，保留目标旧 JSON）、未启用任何表、输出名越界、同名输出冲突、目标写入失败、现有 SVN 冲突、中文空格路径、取消，以及配置持久化。

## 完成记录
2026-10-30 后续调整：增加工作表级 CopyToTarget，默认 false；新增“其他配置”页签，开关不放在基本设置。旧规则缺少属性时也不复制；开启后包含该表全部子表输出。输出目录仍生成全部已勾选工作表的 JSON，Cocos 仅接收开启复制的文件，所有复制关闭时无需 Cocos 目录。新增默认关闭/选择性复制/子表继承/保存与切换工作表测试，SyncTests、UiSmokeTests、RowScanTests 均通过。

2026-10-30：上述三组测试全部通过，Release win-x64 自包含发布成功，发布程序启动检查通过。真实 Excel 只读解析得到 13 个 JSON，发现旧规则“送礼-金宝箱”不存在，按兼容策略警告跳过；没有修改规则映射。发布目录的旧 FileRules 与备份逐项序列化比较保持一致。独立审查及兼容性变更复查均未发现 P1/P2 阻塞问题。未更新真实 SVN 或覆盖真实 Cocos 文件。

SVN 中文路径调查：直接通过 PowerShell native command 调用会在当前环境发生参数编码问题；生产代码使用 ProcessStartInfo.ArgumentList，可正确读取真实 SVN 中文目录，独立测试包含版本控制中的中文子目录。保留已验证签名的 1.15.0 组件。

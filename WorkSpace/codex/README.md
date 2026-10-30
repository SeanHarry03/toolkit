# 自动行数扫描验证

SVN 自动更新和 Cocos 同步功能说明见 [SVN同步使用说明.md](SVN同步使用说明.md)。打开 `publish\ExcelToJson.exe` 后点击“一键更新并同步”。

在项目根目录运行（本机 dotnet 位于 `C:\Users\25330\.dotnet\dotnet.exe`）：

```powershell
& 'C:\Users\25330\.dotnet\dotnet.exe' run --project codex\RowScanTests\RowScanTests.csproj -p:NuGetAudit=false
& 'C:\Users\25330\.dotnet\dotnet.exe' run --project codex\UiSmokeTests\UiSmokeTests.csproj -p:NuGetAudit=false
```

- `RowScanTests` 使用临时 Excel 验证自动扩展、遇空行停止、缩短/清空表、普通表头、手动模式、子表、规则持久化和旧规则默认开启。
- `UiSmokeTests` 在独立测试输出目录运行 WPF 窗口逻辑，验证扫描、转换、保存、编辑框刷新、开关绑定及重新打开；不修改发布目录中的规则。
- 测试文件和测试数据均位于 `codex/` 内。

运行行为：全局 `AutoDetectRows` 默认开启，旧配置没有该字段时也开启。扫描所选文件或转换所有勾选工作表后，将当前区域连续数据的最后一个物理行号写入 `rules.json` 的 `RowEnd`。空行按区域中未排除的列判断，空白文本和空字符串公式视为空，数字 0 和 false 视为数据。子表各自扫描。起始行即为空时，结束行记录为起始行减一，代表没有数据。

普通工作表仍用第一行作为表头，回写时保存 `UseFirstRowAsHeaders`，避免仅因新增结束行配置就改变表头模式。关闭自动读取后，可编辑结束行并按固定范围转换。

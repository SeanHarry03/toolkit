from pathlib import Path

root = Path(__file__).resolve().parent.parent
path = root / 'MainWindow.xaml'
old = path.read_text(encoding='utf-8-sig')
backup = root / 'codex' / 'before-layout'
backup.mkdir(exist_ok=True)
for name in ['MainWindow.xaml', 'MainWindow.xaml.cs', 'Theme.xaml']:
    dest = backup / name
    if not dest.exists():
        dest.write_bytes((root / name).read_bytes())

sidebar_start = old.index('            <Border Background="White" BorderBrush="{StaticResource LineBrush}" BorderThickness="1" CornerRadius="6" Padding="12">')
sidebar_end = old.index('            <GridSplitter Grid.Column="1"', sidebar_start)
sidebar = old[sidebar_start:sidebar_end]
sidebar = sidebar.replace('Background="White" BorderBrush="{StaticResource LineBrush}" BorderThickness="1" CornerRadius="6" Padding="12"', 'x:Name="WorkbooksPanel" Background="#F1F4F8" Padding="20,24" IsEnabled="{Binding IsEnabled, ElementName=ConvertButton}"', 1)
sidebar = sidebar.replace('Text="工作簿与工作表"', 'Text="工作簿"')
sidebar = sidebar.replace('AutomationProperties.Name="Excel 文件"/>', 'AutomationProperties.Name="Excel 文件" ToolTip="{Binding SelectedItem, RelativeSource={RelativeSource Self}}"/>')
sidebar = sidebar.replace('<TextBlock Text="勾选需要转换的 Sheet"', '<TextBlock Text="工作表" FontSize="15" FontWeight="SemiBold" Margin="0,16,0,8"/>\n                        <TextBlock Text="勾选参与生成，点击名称编辑规则"')
sidebar = sidebar.replace('Content="添加文件"', 'Content="添加"').replace('Content="移除文件"', 'Content="移除"')
sidebar = sidebar.replace('Content="添加 Sheet"', 'Content="添加工作表"').replace('Content="删除 Sheet"', 'Content="删除"')

rule_start = old.index('            <Border Grid.Column="2" Background="White"')
rule_end = old.index('        <GridSplitter Grid.Row="3"', rule_start)
rule_border = old[rule_start:rule_end]
rule = rule_border[rule_border.index('>') + 1:rule_border.rindex('</Border>')]
rule = rule.replace('<Grid>', '<Grid x:Name="RuleWorkspace" Margin="28,24" TextElement.FontWeight="Normal" IsEnabled="{Binding IsEnabled, ElementName=ConvertButton}">', 1)
rule = rule.replace('Text="工作表规则" FontSize="16" FontWeight="SemiBold" Margin="0,0,0,16"', 'Text="工作表规则" FontSize="20" FontWeight="SemiBold" Margin="0,0,0,24"')

header = '''<Window x:Class="ExcelToJson.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Excel → JSON 转换工具" Height="820" Width="1280" MinHeight="680" MinWidth="1000"
        WindowStartupLocation="CenterScreen" Style="{StaticResource {x:Type Window}}">
    <Grid x:Name="AppLayout" Background="{StaticResource CanvasBrush}">
        <Grid.RowDefinitions><RowDefinition Height="76"/><RowDefinition Height="*"/><RowDefinition Height="36"/></Grid.RowDefinitions>
        <Border Background="White" BorderBrush="{StaticResource LineBrush}" BorderThickness="0,0,0,1" Padding="24,0">
            <Grid>
                <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                <StackPanel VerticalAlignment="Center">
                    <TextBlock Text="Excel → JSON" FontSize="22" FontWeight="SemiBold"/>
                    <TextBlock Text="配置工作表 · 生成 JSON · 同步到项目" Style="{StaticResource HelpText}" Margin="0,2,0,0"/>
                </StackPanel>
                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                    <Button Content="保存规则" Click="SaveRules_Click" Margin="0,0,16,0" Height="38"
                            IsEnabled="{Binding IsEnabled, ElementName=ConvertButton}"/>
                    <Button x:Name="ConvertButton" Content="开始转换" Click="Convert_Click" Height="38" Padding="20,5"/>
                    <Button x:Name="SyncButton" Content="一键更新并同步" Height="38" Margin="8,0,0,0" Padding="20,5"
                            Style="{StaticResource PrimaryButton}" Click="Sync_Click"
                            IsEnabled="{Binding IsEnabled, ElementName=ConvertButton}"
                            ToolTip="更新 SVN，按规则生成 JSON，仅复制开启了复制开关的工作表"/>
                    <Button x:Name="CancelSyncButton" Content="停止" Margin="8,0,0,0" Height="38" Visibility="Collapsed"
                            Click="CancelSync_Click" ToolTip="停止 SVN；解析结束后停止；复制期间停止会回滚"/>
                </StackPanel>
            </Grid>
        </Border>
        <Grid Grid.Row="1">
            <Grid.ColumnDefinitions><ColumnDefinition Width="270" MinWidth="230" MaxWidth="360"/><ColumnDefinition Width="5"/><ColumnDefinition Width="*" MinWidth="620"/></Grid.ColumnDefinitions>
'''
middle = '''
            <GridSplitter Grid.Column="1" Width="5" HorizontalAlignment="Stretch" Background="{StaticResource LineBrush}" ShowsPreview="True"/>
            <TabControl x:Name="WorkspaceTabs" Grid.Column="2" Background="White" Padding="0" BorderThickness="0"
                        AutomationProperties.Name="工作区导航" FontWeight="Normal">
                <TabItem x:Name="EditorPage" Header="规则编辑" Style="{StaticResource WorkspaceTab}">
'''

def folder_field(label, name, handler, extra=''):
    return f'''<StackPanel Margin="0,0,0,24">
        <Label Content="{label}" Target="{{Binding ElementName={name}}}" Padding="0" Margin="0,0,0,8"/>
        <DockPanel>
            <Button DockPanel.Dock="Right" Content="浏览…" Click="{handler}" Margin="8,0,0,0" MinWidth="72" Height="36" AutomationProperties.Name="浏览{label}"/>
            {extra}
            <TextBox x:Name="{name}" Height="36" AutomationProperties.Name="{label}" ToolTip="{{Binding Text, RelativeSource={{RelativeSource Self}}}}"/>
        </DockPanel>
    </StackPanel>'''

directories = '''
                </TabItem>
                <TabItem x:Name="DirectoriesPage" Header="目录与同步" Style="{StaticResource WorkspaceTab}">
                    <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
                        <StackPanel x:Name="DirectorySettingsPanel" Margin="28,24" MaxWidth="920" HorizontalAlignment="Stretch"
                                    IsEnabled="{Binding IsEnabled, ElementName=ConvertButton}" TextElement.Foreground="{StaticResource TextBrush}" TextElement.FontWeight="Normal">
                            <TextBlock Text="目录与同步" FontSize="20" FontWeight="SemiBold"/>
                            <TextBlock Text="设置 Excel 来源、JSON 输出位置及 Cocos 同步目标。" Style="{StaticResource HelpText}" Margin="0,8,0,28"/>
                            <TextBlock Text="转换目录" FontSize="15" FontWeight="SemiBold" Margin="0,0,0,20"/>
'''
directories += folder_field('Excel 目录', 'ExcelDirBox', 'BrowseExcelDir_Click')
directories += folder_field('JSON 输出目录', 'OutputDirBox', 'BrowseOutputDir_Click', '<Button DockPanel.Dock="Right" Content="打开目录" Click="OpenOutput_Click" Margin="8,0,0,0" Height="36"/>')
directories += '''
                            <CheckBox x:Name="AutoDetectRowsBox" Content="自动读取行数，遇到空行停止" Margin="0,0,0,8"
                                      ToolTip="对所有工作表和子表生效。关闭后使用各表手动设置的结束行。"/>
                            <TextBlock Text="适用于所有工作表及子表；关闭后使用手动设置的结束行。" Style="{StaticResource HelpText}"/>
                            <Border BorderBrush="{StaticResource LineBrush}" BorderThickness="0,1,0,0" Margin="0,28,0,24"/>
                            <TextBlock Text="同步到 Cocos" FontSize="15" FontWeight="SemiBold" Margin="0,0,0,8"/>
                            <TextBlock Text="仅复制在工作表「其他配置」中开启了复制开关的 JSON，保留 .meta 和其他文件。" Style="{StaticResource HelpText}" Margin="0,0,0,20"/>
'''
directories += folder_field('SVN 更新目录', 'SvnDirBox', 'BrowseSvnDir_Click')
directories += folder_field('Cocos JSON 目录', 'TargetDirBox', 'BrowseTargetDir_Click')
directories += '''
                            <Expander x:Name="SyncSettingsExpander" Header="高级设置" Margin="0,4,0,24">
                                <StackPanel Margin="0,16,0,0">
'''
directories += folder_field('SVN 程序（可留空）', 'SvnExeBox', 'BrowseSvnExe_Click')
directories += '''
                                    <TextBlock Text="留空使用工具附带的 svn.exe。覆盖前的备份保存在 JSON 输出目录的 .sync-runs 中。" Style="{StaticResource HelpText}"/>
                                </StackPanel>
                            </Expander>
                            <Button Content="恢复默认规则" Click="ResetRules_Click" HorizontalAlignment="Left" ToolTip="恢复内置规则，将提示确认"/>
                        </StackPanel>
                    </ScrollViewer>
                </TabItem>
                <TabItem x:Name="LogPage" Header="运行日志" Style="{StaticResource WorkspaceTab}">
                    <Grid Margin="28,24">
                        <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
                        <Grid Margin="0,0,0,24">
                            <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                            <StackPanel><TextBlock Text="运行日志" FontSize="20" FontWeight="SemiBold"/>
                                <TextBlock Text="查看转换、SVN 更新和复制的详细结果。" Style="{StaticResource HelpText}" Margin="0,8,0,0"/></StackPanel>
                            <Button Grid.Column="1" Content="清空日志" Click="ClearLog_Click" VerticalAlignment="Top"/>
                        </Grid>
                        <TextBox x:Name="LogBox" Grid.Row="1" IsReadOnly="True" TextWrapping="Wrap" AcceptsReturn="True"
                                 VerticalScrollBarVisibility="Auto" FontFamily="Consolas, Microsoft YaHei UI" FontSize="13"
                                 BorderThickness="1" BorderBrush="{StaticResource LineBrush}" Background="#F8FAFC" Padding="16"
                                 FontWeight="Normal" Foreground="{StaticResource TextBrush}" VerticalContentAlignment="Top" AutomationProperties.Name="运行日志"/>
                    </Grid>
                </TabItem>
            </TabControl>
        </Grid>
        <Border Grid.Row="2" Background="White" BorderBrush="{StaticResource LineBrush}" BorderThickness="0,1,0,0" Padding="24,0">
            <Grid>
                <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                <TextBlock x:Name="StatusText" Text="就绪" VerticalAlignment="Center" Foreground="{StaticResource MutedBrush}"
                           TextTrimming="CharacterEllipsis" ToolTip="{Binding Text, RelativeSource={RelativeSource Self}}" AutomationProperties.LiveSetting="Polite"/>
                <Button Grid.Column="1" Content="查看日志" Click="ShowLog_Click" Padding="8,0" MinHeight="26" Margin="16,4,0,4"/>
            </Grid>
        </Border>
    </Grid>
</Window>
'''
path.write_text(header + sidebar + middle + rule + directories, encoding='utf-8')

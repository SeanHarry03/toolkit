# Excel 转换为json

# 运行

打开 ExcelToJSON.sln运行

需要转换的excel 文件和 output 同级

## 自定义转换规则

在路径的所有 .xlsx 需要跳过不转换的子表

```aiexclude
excludeSheets 字段
``` 

针对单独的.xlsx 转换规则

```plantuml
fileRules 字段
```

```plantuml
["策划案.xlsx"] = new FileRule
    {
    //策划案.xlsx 里面所有的子表的规则
        Sheets = new Dictionary<string, SheetRule>
        {
        //子表1
            ["道具配置表"] = new SheetRule
            {
                //输出的json文件名字
                OutputName = "xx.json",
                
                //转换数据 行范围[1,10]
                RowRange = new RowRangeDef(1, 10),
                
                // 列范围
                ColRange = new ColRangeDef("B", "AE"),
                
                // 跳过的列不转换
                ExcludeCols = ["G"],
                
                // 自定义表头映射（列字母 → 字段名）
                // 添加此规则，那么行1 就是实际的数据内容的第一行；
                // 不添加此规则则使用 行1 为json的键值
                CustomHeaders = new Dictionary<string, string>
                {
                    ["B"] = "id",
                    ["C"] = "name",
                    ["D"] = "des",
                    ["E"] = "describeValues",
                }
                
                /**
                把B列 对应上面的CustomHeaders 作为一列内容的json的key
                {
                    "id1":
                    {
                        "id":"id1,
                        "name":"张三"
                        "des":"描述文本1",
                    }                
                }
                */
                KeyColumn = "id",
                
                //选择在没有设置KeyColumn 的时候，输出数组还是使用自然序号输出 1、2、3
                NoKeyOutputMode = NoKeyOutputMode.RowIndexObject,
                NoKeyOutputMode = NoKeyOutputMode.Array,
                
                //指定哪些列为数组 (Excel 里面的内容格式 "1001,1002") 使用逗号分割
                ArrayFields=["describeValues"]
                
                //当某一列的数据为空，是否输出的json字段删除。
                RemoveEmpty = true,
                
                // RoundFields = new List<string> { "H", "upgradeGold", "buildGoldCost" }, 指定列取整
                // 指定全部列取整
                RoundFields = new List<string> { "*" },
                
                //计算得出，Excel原本不存在的列。写入到json
                ComputedColumns = new Dictionary<string, Func<IReadOnlyDictionary<string, object?>, object?>>
                {
                    ["spiritId"] = row =>
                    {
                        var spiritType = GetDouble(row, "spiritType");
                        var level = GetDouble(row, "level");
                        return spiritType * 100 + level;
                    }
                },
            }
        }
    }
```

在一个表Sheet 里面 还需要划分出来不同的表

```plantuml
    SubTables = new Dictionary<string, SubTableRule>
    {
        ["601"] = new SubTableRule
        {
            RowRange = new RowRangeDef(50, 80),
            ColRange = new ColRangeDef("C", "D"),
            CustomHeaders = new Dictionary<string, string>
            {
                ["C"] = "id",
                ["D"] = "probability",
            },
            WrapWithSubKey = true,
            OutputName = "默认道具和祈福签配置经典.json",
        },
    }
    
    //WrapWithSubKey
    // 使用外面的键作为子表的key。
    /**
    {
        601:[ 子表(50,80)行 [C,D]列 的所有的内容]
    }
    */
    //OutputName 设置相同的文件则把SubTable的内容输出到相同的文件中
    
 
    
```
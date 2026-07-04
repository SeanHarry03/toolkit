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
                
                //指定哪些列为数组 (Excel 里面的内容格式 "1001,1002") 使用逗号分割
                ArrayFields=["describeValues"]
                
                //当5行 B列的数据为空时是否在
                RemoveEmpty = true,
            }
        }
    }
```
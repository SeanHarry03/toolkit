namespace ExcelToJson;

/// <summary>内置默认规则种子（首次生成 rules.json）</summary>
public static class BuiltInRules
{
    public static ConversionConfig Create()
    {
        return new ConversionConfig
        {
            ExcludeSheets =
            [
                "任务列表",
                "文本配置",
                "动效需求",
                "人扮演鬼模式",
                "整体",
            ],
            ExcelFiles = ["策划案.xlsx"],
            FileRules = new Dictionary<string, FileRuleDto>(StringComparer.Ordinal)
            {
                ["策划案.xlsx"] = new FileRuleDto
                {
                    Sheets = new Dictionary<string, SheetRuleDto>(StringComparer.Ordinal)
                    {
                        ["道具配置表"] = Sheet(
                            enabled: true,
                            rowStart: 18, rowEnd: 155,
                            colStart: "B", colEnd: "AE",
                            excludeCols: ["G"],
                            headers: new Dictionary<string, string>
                            {
                                ["B"] = "id", ["C"] = "itemType", ["D"] = "itemLevel", ["E"] = "maxLevel",
                                ["F"] = "itemName", ["H"] = "itemHp", ["I"] = "upgradeGold",
                                ["J"] = "upgradeEnergy", ["K"] = "upgradeRequire", ["L"] = "removeGold",
                                ["M"] = "removeEnergy", ["N"] = "attack", ["O"] = "attackRange",
                                ["P"] = "buildGoldCost", ["Q"] = "buildEnergyCost", ["R"] = "goldPerSec",
                                ["S"] = "energyPerSec", ["T"] = "itemDescription", ["U"] = "describeValues",
                                ["V"] = "triggerCondition", ["W"] = "triggerProbability", ["X"] = "maxCount",
                                ["Y"] = "activeDuration", ["Z"] = "skillCd", ["AA"] = "skillEffect",
                                ["AB"] = "skillValue", ["AC"] = "skillDuration", ["AD"] = "useOnce",
                                ["AE"] = "spiritKillExp",
                            },
                            keyColumn: "id",
                            arrayFields: ["describeValues"],
                            outputName: "TowerConfig.json",
                            roundFields: ["*"]),

                        ["猛鬼配置表"] = Sheet(
                            enabled: true,
                            rowStart: 15, rowEnd: 270,
                            colStart: "D", colEnd: "N",
                            headers: new Dictionary<string, string>
                            {
                                ["D"] = "spiritType", ["E"] = "spiritName", ["F"] = "level",
                                ["G"] = "levelUpExp", ["H"] = "expPerSec", ["I"] = "hp",
                                ["J"] = "attack", ["K"] = "moveSpeed", ["L"] = "attackSpeed",
                                ["M"] = "damageToExpRate", ["N"] = "skillList",
                            },
                            keyColumn: "spiritId",
                            arrayFields: ["skillList"],
                            outputName: "EnemyConfig.json",
                            roundFields: ["*"],
                            computed: new Dictionary<string, string>
                            {
                                ["spiritId"] = "spiritType * 100 + level",
                            }),

                        ["猛鬼技能效果"] = Sheet(
                            enabled: true,
                            rowStart: 19, rowEnd: 109,
                            colStart: "E", colEnd: "N",
                            headers: new Dictionary<string, string>
                            {
                                ["E"] = "skillId", ["F"] = "skillType", ["G"] = "skillLevel",
                                ["H"] = "skillName", ["I"] = "skillProbability", ["J"] = "skillDuration",
                                ["K"] = "skillCd", ["L"] = "skillEffect", ["M"] = "skillValue",
                                ["N"] = "triggerCondition",
                            },
                            keyColumn: "skillId",
                            outputName: "EnemySkillConfig.json",
                            roundFields: ["*"]),

                        ["默认道具和祈福签配置"] = new SheetRuleDto
                        {
                            Enabled = true,
                            RoundFields = ["*"],
                            SubTables = new Dictionary<string, SubTableRuleDto>(StringComparer.Ordinal)
                            {
                                ["601"] = Sub(50, 80, "C", "D", headers: H(("C", "id"), ("D", "probability")),
                                    wrap: true, output: "默认道具和祈福签配置经典.json"),
                                ["701"] = Sub(50, 81, "G", "H", headers: H(("G", "id"), ("H", "probability")),
                                    wrap: true, output: "默认道具和祈福签配置经典.json"),
                                ["801"] = Sub(50, 90, "L", "M", headers: H(("L", "id"), ("M", "probability")),
                                    wrap: true, output: "默认道具和祈福签配置经典.json"),
                                ["6011"] = Sub(50, 73, "T", "U", headers: H(("T", "id"), ("U", "probability")),
                                    wrap: true, output: "默认道具和祈福签配置挑战.json"),
                                ["7011"] = Sub(50, 81, "X", "Y", headers: H(("X", "id"), ("Y", "probability")),
                                    wrap: true, output: "默认道具和祈福签配置挑战.json"),
                                ["8011"] = Sub(50, 83, "AC", "AD", headers: H(("AC", "id"), ("AD", "probability")),
                                    wrap: true, output: "默认道具和祈福签配置挑战.json"),
                                ["房间可以拾取的道具"] = Sub(25, 34, "T", "U",
                                    headers: H(("T", "id"), ("U", "probability")),
                                    wrap: false, output: "房间可以拾取的道具.json",
                                    mode: NoKeyOutputMode.RowIndexObject),
                                ["经典模式房间默认"] = Sub(25, 36, "J", "K",
                                    headers: H(("J", "id"), ("K", "probability")),
                                    wrap: false, output: "经典模式房间默认道具.json",
                                    mode: NoKeyOutputMode.RowIndexObject),
                                ["挑战模式房间默认道具"] = Sub(25, 35, "AA", "AB",
                                    headers: H(("AA", "id"), ("AB", "probability")),
                                    wrap: false, output: "挑战模式房间默认道具.json",
                                    mode: NoKeyOutputMode.RowIndexObject),
                            },
                        },

                        ["挑战模式"] = new SheetRuleDto
                        {
                            Enabled = true,
                            OutputName = "LevelConfig.json",
                            RoundFields = ["*"],
                            SubTables = new Dictionary<string, SubTableRuleDto>(StringComparer.Ordinal)
                            {
                                ["LevelConfig"] = Sub(14, 28, "H", "M",
                                    headers: H(
                                        ("H", "level"), ("I", "hpFactor"), ("J", "attackFactor"),
                                        ("K", "moveFactor"), ("L", "attackSpeed"), ("M", "waveList")),
                                    wrap: false, output: "LevelConfig.json",
                                    arrayFields: ["waveList"],
                                    mode: NoKeyOutputMode.RowIndexObject,
                                    removeEmpty: true),
                                ["WaveConfig"] = Sub(38, 89, "H", "I",
                                    headers: H(("H", "waveId"), ("I", "spiritIdList")),
                                    wrap: false, output: "WaveConfig.json",
                                    arrayFields: ["spiritIdList"],
                                    mode: NoKeyOutputMode.RowIndexObject,
                                    removeEmpty: true),
                            },
                        },

                        ["称号_商城_结算"] = new SheetRuleDto
                        {
                            Enabled = true,
                            OutputName = "LobbyConfig.json",
                            SubTables = new Dictionary<string, SubTableRuleDto>(StringComparer.Ordinal)
                            {
                                ["shop_config"] = Sub(43, 58, "E", "G",
                                    headers: H(("E", "item_id"), ("F", "type"), ("G", "price")),
                                    wrap: true,
                                    arrayFields: ["spiritIdList"],
                                    mode: NoKeyOutputMode.Array,
                                    roundFields: ["*"],
                                    removeEmpty: true),
                            },
                        },

                        ["送礼-银宝箱"] = Sheet(
                            enabled: true,
                            rowStart: 3, rowEnd: 21,
                            colStart: "C", colEnd: "H",
                            headers: new Dictionary<string, string>
                            {
                                ["E"] = "id", ["D"] = "item_type", ["H"] = "item_count",
                                ["F"] = "duration", ["C"] = "Clientprobability", ["G"] = "Serverprobability",
                            },
                            outputName: "送礼-银宝箱.json",
                            mode: NoKeyOutputMode.Array),

                        ["送礼-金宝箱"] = Sheet(
                            enabled: true,
                            rowStart: 3, rowEnd: 21,
                            colStart: "C", colEnd: "H",
                            headers: new Dictionary<string, string>
                            {
                                ["E"] = "id", ["D"] = "item_type", ["H"] = "item_count",
                                ["F"] = "duration", ["C"] = "Clientprobability", ["G"] = "Serverprobability",
                            },
                            outputName: "送礼-金宝箱.json",
                            mode: NoKeyOutputMode.Array),

                        ["人扮鬼-鬼增幅配置"] = Sheet(
                            enabled: true,
                            rowStart: 3, rowEnd: 14,
                            colStart: "B", colEnd: "G",
                            headers: new Dictionary<string, string>
                            {
                                ["B"] = "id", ["C"] = "type", ["D"] = "rarity",
                                ["E"] = "num", ["F"] = "name", ["G"] = "description",
                            },
                            keyColumn: "id",
                            outputName: "spirit_boost_config.json",
                            roundFields: ["*"],
                            mode: NoKeyOutputMode.RowIndexObject),
                    },
                },
            },
        };
    }

    private static Dictionary<string, string> H(params (string k, string v)[] pairs)
    {
        var d = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (k, v) in pairs)
            d[k] = v;
        return d;
    }

    private static SheetRuleDto Sheet(
        bool enabled,
        int? rowStart = null, int? rowEnd = null,
        string? colStart = null, string? colEnd = null,
        List<string>? excludeCols = null,
        Dictionary<string, string>? headers = null,
        string? keyColumn = null,
        List<string>? arrayFields = null,
        string? outputName = null,
        List<string>? roundFields = null,
        Dictionary<string, string>? computed = null,
        NoKeyOutputMode mode = NoKeyOutputMode.Array,
        bool removeEmpty = true)
    {
        return new SheetRuleDto
        {
            Enabled = enabled,
            RowStart = rowStart,
            RowEnd = rowEnd,
            ColStart = colStart,
            ColEnd = colEnd,
            ExcludeCols = excludeCols ?? [],
            CustomHeaders = headers ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase),
            KeyColumn = keyColumn,
            ArrayFields = arrayFields ?? [],
            RemoveEmpty = removeEmpty,
            OutputName = outputName,
            RoundFields = roundFields ?? [],
            ComputedColumns = computed ?? new Dictionary<string, string>(StringComparer.Ordinal),
            NoKeyOutputMode = mode,
        };
    }

    private static SubTableRuleDto Sub(
        int rowStart, int rowEnd, string colStart, string colEnd,
        Dictionary<string, string>? headers = null,
        bool wrap = true,
        string? output = null,
        List<string>? arrayFields = null,
        NoKeyOutputMode mode = NoKeyOutputMode.Array,
        List<string>? roundFields = null,
        bool removeEmpty = true)
    {
        return new SubTableRuleDto
        {
            RowStart = rowStart,
            RowEnd = rowEnd,
            ColStart = colStart,
            ColEnd = colEnd,
            CustomHeaders = headers ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase),
            WrapWithSubKey = wrap,
            OutputName = output,
            ArrayFields = arrayFields ?? [],
            NoKeyOutputMode = mode,
            RoundFields = roundFields ?? [],
            RemoveEmpty = removeEmpty,
        };
    }
}

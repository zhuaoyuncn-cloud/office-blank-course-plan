# 录入格式和运行方式

以下是虚构的完整小例子。班级可按任意顺序录入，程序会统一排序。只有用户明确指定排列顺序时，才按指定名单组织 `classes`，并在顶层设置 `"classOrder": "input"`。看全该班工作日的所有大节和备注后，才将 `complete` 置为 true。

```json
{
  "term": "示例学年第一学期",
  "grade": "示例年级",
  "endWeek": 18,
  "classes": [
    {
      "name": "示例一班",
      "complete": true,
      "entries": [
        {"days": [2], "periods": [2], "weeks": "7-10"},
        {"days": [2], "periods": [2], "weeks": "11-13"},
        {"days": [4], "periods": [1], "weeks": "1-15单,18"},
        {"days": "all", "periods": "all", "weeks": "14-16"}
      ]
    }
  ]
}
```

- `days` 的 1–5 为星期一至五，`periods` 从 1 起按本次表头的大节顺序编号。`"all"` 表示该维度全部。把周末和统计范围外周次从输入中剔除，完全在范围外的记录略去。
- `weeks` 按实际有课周填写。单周写 `7`，连续段写 `1-8`，列举写 `1,5,9-12`，奇数周写 `1-15单`，偶数周写 `2-16双`。明确整学期有课时填实际范围。班级整张表确认为空时可用空数组 `entries: []`。
- `endWeek` 可省略，默认 18。如果源课表的大节数量或时间不同，在顶层提供有序的 `periods` 字符串数组，如 `["第一大节\n08:00—09:30", "第二大节\n09:45—11:15"]`。只有节名时可仅写节名。`term` 和 `grade` 用本次资料确认的文字。
- 在整理过程中保留来源位置以复核。例中周二第二大节的有课周应为 7–16，无课周应为 `【1-6】【17-18】`。

在 Codex 桌面端使用带 `@oai/artifact-tool` 的工作区 Node 运行时。下例在技能目录中运行，两个路径改为本次临时 JSON 和新 Excel 文件的路径，输出目录应已存在。

```powershell
$node = "$env:USERPROFILE/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin/node.exe"
& $node scripts/build.mjs input.json "本学期 空课表.xlsx"
```

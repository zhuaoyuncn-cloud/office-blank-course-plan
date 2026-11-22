import fs from "node:fs/promises";
import { createRequire } from "node:module";
import os from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";

const runtime = createRequire(process.execPath);
const { Workbook, SpreadsheetFile } = await import(pathToFileURL(runtime.resolve("@oai/artifact-tool")));
const days = ["星期一", "星期二", "星期三", "星期四", "星期五"];
const defaultPeriods = ["第一大节\n08:30—10:00", "第二大节\n10:15—11:45", "第三大节\n13:30—15:00",
  "第四大节\n15:15—16:45", "第五大节\n19:00—20:35"];

function requireValue(ok, message) {
  if (!ok) throw new Error(message);
}

function hasText(value) {
  return typeof value === "string" && value.trim().length > 0;
}

const sampleMajors = ["机械", "机工", "智造", "成型", "机职贯", "型职贯"];
const byClassName = new Intl.Collator("zh-Hans-CN-u-co-pinyin", { numeric: true });

function classSortKey(name) {
  const digits = "零一二三四五六七八九";
  const label = name.normalize("NFKC").replace(/\s+/gu, "").replace(
    /([一二三四五六七八九]?十[一二三四五六七八九]?|[一二三四五六七八九])(?=班$)/u,
    (_, number) => {
      const [tens, ones] = number.split("十");
      if (ones === undefined) return String(digits.indexOf(tens));
      return String((tens ? digits.indexOf(tens) : 1) * 10 + (ones ? digits.indexOf(ones) : 0));
    },
  );
  const major = label.split(/\d/u, 1)[0];
  const index = sampleMajors.indexOf(major);
  return { rank: index < 0 ? sampleMajors.length : index, label };
}

function indices(selection, count, field) {
  if (selection === "all") return Array.from({ length: count }, (_, i) => i);
  requireValue(Array.isArray(selection) && selection.length &&
    selection.every(n => Number.isInteger(n) && n >= 1 && n <= count), field + " 超出范围或为空");
  return selection.map(n => n - 1);
}

function parseWeeks(spec, endWeek) {
  requireValue(hasText(spec), "缺少周次");
  const result = new Set();
  for (const token of spec.split(",")) {
    const match = /^(\d+)(?:-(\d+))?([单双]?)$/u.exec(token.trim());
    requireValue(match, "无法识别周次 " + token);
    const first = Number(match[1]);
    const last = Number(match[2] ?? match[1]);
    requireValue(first >= 1 && first <= last && last <= endWeek, "周次超出统计范围 " + token);
    let selected = false;
    for (let n = first; n <= last; n++) {
      if (match[3] && n % 2 !== Number(match[3] === "单")) continue;
      result.add(n);
      selected = true;
    }
    requireValue(selected, "单双周标记未选中任何周 " + token);
  }
  return result;
}

function ranges(numbers) {
  const result = [];
  for (let i = 0; i < numbers.length; i++) {
    const start = numbers[i];
    while (numbers[i + 1] === numbers[i] + 1) i++;
    result.push(start === numbers[i] ? String(start) : start + "-" + numbers[i]);
  }
  return result;
}

function build(data) {
  const { term, grade, classes, periods = defaultPeriods, endWeek = 18, classOrder = "auto" } = data;
  requireValue(hasText(term) && hasText(grade), "缺少年级或学期");
  requireValue(Number.isInteger(endWeek) && endWeek >= 1 && endWeek <= 18, "末周应在 1–18 之间");
  requireValue(Array.isArray(periods) && periods.length && periods.every(hasText), "缺少有效大节名称");
  requireValue(Array.isArray(classes) && classes.length, "缺少班级");
  requireValue(classOrder === "auto" || classOrder === "input", "班级顺序应为 auto 或 input");
  const weeks = Array.from({ length: endWeek }, (_, i) => i + 1);
  const lines = periods.map(() => days.map(() => []));
  const audit = [["班级", "节次", "星期", "有课周次", "无课周次（空课表所填）"]];
  const names = new Set();
  const ordered = classes.map(item => {
    requireValue(hasText(item?.name) && !names.has(item.name.trim()), "班名缺失或重复 " + item?.name);
    names.add(item.name.trim());
    return { item, ...classSortKey(item.name) };
  });
  if (classOrder === "auto") {
    ordered.sort((a, b) => a.rank - b.rank || byClassName.compare(a.label, b.label) ||
      a.item.name.localeCompare(b.item.name, "zh-CN"));
  }
  let dayWidth = 155;
  for (const { item: { name, complete, entries } } of ordered) {
    requireValue(complete === true && Array.isArray(entries), name + " 的课表或备注尚未核全");
    const busy = periods.map(() => days.map(() => new Set()));
    for (const [i, entry] of entries.entries()) {
      try {
        const selectedDays = indices(entry.days, days.length, "星期");
        const selectedPeriods = indices(entry.periods, periods.length, "大节");
        const selectedWeeks = parseWeeks(entry.weeks, endWeek);
        for (const p of selectedPeriods) for (const d of selectedDays) {
          for (const w of selectedWeeks) busy[p][d].add(w);
        }
      } catch (cause) {
        throw new Error(name + " 第 " + (i + 1) + " 条记录：" + cause.message);
      }
    }
    periods.forEach((period, p) => days.forEach((day, d) => {
      const occupied = busy[p][d];
      const free = ranges(weeks.filter(w => !occupied.has(w))).map(r => "【" + r + "】").join("") || "【无】";
      const line = name.trim() + free;
      lines[p][d].push(line);
      const units = Array.from(line).reduce((sum, ch) => sum + (ch.codePointAt(0) > 127 ? 2 : 1), 0);
      dayWidth = Math.max(dayWidth, units * 7 + 20);
      audit.push([name.trim(), period.split("\n")[0], day, ranges(weeks.filter(w => occupied.has(w))).join(",") || "—", free]);
    }));
  }

  const book = Workbook.create();
  const main = book.worksheets.add("空课表");
  const basis = book.worksheets.add("课表依据（核对用）");
  const borders = { preset: "all", style: "thin", color: "#A6A6A6" };
  function header(range) {
    range.format.fill = "#EDEDED";
    range.format.font.bold = true;
    range.format.horizontalAlignment = "center";
  }
  function banner(row, text, height, size) {
    const r = main.getRange("A" + row + ":F" + row);
    r.merge();
    main.getRange("A" + row).values = [[text]];
    r.format.rowHeight = height;
    r.format.font.size = size;
    r.format.horizontalAlignment = "center";
  }
  main.showGridLines = false;
  main.getRange("A1:A3").format.columnWidth = 17;
  main.getRange("B3:F3").format.columnWidthPx = dayWidth;
  banner(1, term + "  空课表", 36, 20);
  main.getRange("A1").format.font.bold = true;
  banner(2, "年级：" + grade + "　　　统计范围：星期一至星期五，第1—" + endWeek + "周", 24, 11);
  const headings = main.getRange("A3:F3");
  headings.values = [["节次", ...days]];
  header(headings);
  headings.format.rowHeight = 25;
  // Excel 单行高度有限，班级较多时同一大节续行。
  let row = 4;
  periods.forEach((period, p) => {
    for (let start = 0; start < classes.length; start += 24) {
      const chunk = lines[p].map(list => list.slice(start, start + 24));
      const range = main.getRange("A" + row + ":F" + row);
      range.values = [[period, ...chunk.map(list => list.join("\n"))]];
      range.format.font.size = 10;
      range.format.wrapText = true;
      range.format.rowHeight = Math.max(38, chunk[0].length * 15 + 9);
      const label = main.getRange("A" + row);
      header(label);
      row++;
    }
  });
  main.getRange("A3:F" + (row - 1)).format.borders = borders;
  banner(row, "备注　【】内为无课周次，间断周分列；【无】表示该时段每周均有课。", 28, 10);
  main.getRange("A1:F" + row).format.verticalAlignment = "center";

  const table = basis.getRange("A1:E" + audit.length);
  table.values = audit;
  table.format.borders = borders;
  table.format.verticalAlignment = "center";
  table.format.rowHeight = 20;
  header(basis.getRange("A1:E1"));
  basis.getRange("A1:C1").format.columnWidth = 18;
  basis.getRange("D1").format.columnWidth = 42;
  basis.getRange("E1").format.columnWidthPx = Math.max(220, dayWidth);
  basis.freezePanes.freezeRows(1);
  return book;
}

const [input, output] = process.argv.slice(2);
requireValue(input && output && /\.xlsx$/i.test(output), "用法：build.mjs 输入.json 新空课表.xlsx");
const book = build(JSON.parse((await fs.readFile(input, "utf8")).replace(/^\uFEFF/, "")));
const temp = await fs.mkdtemp(path.join(os.tmpdir(), "office-blank-"));
try {
  const staged = path.join(temp, "workbook.xlsx");
  await (await SpreadsheetFile.exportXlsx(book)).save(staged);
  await fs.writeFile(output, await fs.readFile(staged), { flag: "wx" });
} finally {
  await fs.rm(temp, { recursive: true, force: true });
}

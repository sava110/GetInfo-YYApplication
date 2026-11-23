const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const { test } = require("node:test");
const vm = require("node:vm");

const context = vm.createContext({
  Date,
  Utilities: {
    formatDate(date, zone) {
      const parts = new Intl.DateTimeFormat("sv-SE", {
        timeZone: zone, dateStyle: "short", timeStyle: "short",
      }).format(date);
      return parts.replaceAll("-", "/");
    },
  },
});
vm.runInContext(fs.readFileSync(path.join(__dirname, "../apps-script/Code.gs"), "utf8"), context);

function mockSheet(dataRows = []) {
  const rows = [[], [], ...dataRows];
  let capacity = rows.length;
  return {
    rows,
    getLastRow: () => rows.length,
    getMaxRows: () => capacity,
    insertRowsAfter(after, count) {
      assert.equal(after, capacity);
      capacity += count;
    },
    getRange(row, column, height, width) {
      assert.ok(row + height - 1 <= capacity);
      const read = (field) => Array.from({ length: height }, (_, r) =>
        Array.from({ length: width }, (_, c) => rows[row + r - 1]?.[column + c - 1]?.[field] ?? "")
      );
      const write = (values) => values.forEach((line, r) => {
        const target = rows[row + r - 1] ||= [];
        line.forEach((cell, c) => { target[column + c - 1] = cell; });
      });
      return {
        getValues: () => read("value"),
        getFormulas: () => read("formula"),
        getNotes: () => read("note"),
        setNumberFormat() {},
        setWrap() {},
        setValues: (values) => write(Array.from(values, (line) =>
          Array.from(line, (value) => ({ value }))
        )),
        write,
        moveTo(target) {
          const source = rows[row - 1];
          target.write([source.slice(column - 1, column - 1 + width)]);
          source.fill(undefined, column - 1, column - 1 + width);
        },
      };
    },
  };
}

function entry(day, version, recentChanges) {
  return { fetchedAt: `2026-11-${String(day).padStart(2, "0")}T00:00:00Z`, version, recentChanges };
}

test("packs each app independently and appends all new entries in order", () => {
  const manual = [
    { value: new Date("2026-11-02T00:00:00Z") },
    { value: "2.0" },
    { value: "old" },
    { value: "反映済み" },
    { formula: '=HYPERLINK("https://example.com", "記事")', note: "確認済み" },
  ];
  const sheet = mockSheet([[], [], manual.slice()]);
  const [ios, android] = context.APPS;
  const document = { apps: {
    [ios.id]: { history: [
      entry(1, "1.0", "earlier"),
      entry(2, "2.0", "old"),
      entry(3, "2.0", "iOS notes only"),
      entry(4, "3.0", "next"),
      entry(5, "2.0", "rollback"),
      entry(6, "4.0", "latest"),
    ] },
    [android.id]: { history: [
      entry(1, "Varies with device", "first"),
      entry(2, "Varies with device", "second"),
      entry(3, "Varies with device", "second"),
    ] },
  } };

  context.syncAppHistories_(sheet, document);
  assert.deepEqual(sheet.rows[2].slice(0, 5), manual);
  assert.deepEqual(sheet.rows.slice(2).map((row) => row[1]?.value), ["2.0", "3.0", "2.0", "4.0"]);
  assert.deepEqual(sheet.rows.slice(2).map((row) => row[7]?.value), ["first", "second", undefined, undefined]);
  assert.equal(sheet.rows[3][3], undefined);
  assert.equal(sheet.rows[3][4], undefined);
  const after = JSON.stringify(sheet.rows);
  context.syncAppHistories_(sheet, document);
  assert.equal(JSON.stringify(sheet.rows), after);
});

test("uses the recorded date when an existing row has no exact match", () => {
  const sheet = mockSheet([["2026/11/2", "manual version", "old"].map((value) => ({ value }))]);
  const [ios] = context.APPS;
  const record = { history: [entry(1, "1.0", "old"), entry(3, "3.0", "latest")] };
  context.syncAppHistories_(sheet, { apps: { [ios.id]: record } });
  assert.deepEqual(sheet.rows.slice(2).map((row) => row[1].value), ["manual version", "3.0"]);
});

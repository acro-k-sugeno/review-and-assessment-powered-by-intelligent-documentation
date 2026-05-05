import {
  CheckRepository,
  makePrismaCheckRepository,
} from "../domain/repository";
import { CheckListItemDetail } from "../domain/model/checklist";
import {
  RequestUser,
  assertHasOwnerAccessOrThrow,
} from "../../../core/middleware/authorization";
import { ChecklistError } from "../../../core/errors/application-errors";

const EXCEL_MIME_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

type ExportRow = {
  numbers: number[];
  name: string;
  description: string;
};

export const getChecklistExcelMimeType = (): string => EXCEL_MIME_TYPE;

const xmlEscape = (value: string): string =>
  value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");

const getColumnName = (index: number): string => {
  let n = index;
  let name = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    name = String.fromCharCode(65 + rem) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
};

const createStringCell = (
  columnIndex: number,
  rowIndex: number,
  value: string,
  styleId?: number
): string => {
  const style = styleId !== undefined ? ` s="${styleId}"` : "";
  return `<c r="${getColumnName(columnIndex)}${rowIndex}" t="inlineStr"${style}><is><t>${xmlEscape(
    value
  )}</t></is></c>`;
};

const createNumberCell = (
  columnIndex: number,
  rowIndex: number,
  value: number
): string => `<c r="${getColumnName(columnIndex)}${rowIndex}"><v>${value}</v></c>`;

export const buildChecklistExportRows = (
  items: CheckListItemDetail[]
): ExportRow[] => {
  if (items.length === 0) {
    return [];
  }

  const byId = new Map(items.map((item) => [item.id, item]));
  const childrenByParentId = new Map<string, CheckListItemDetail[]>();
  const roots: CheckListItemDetail[] = [];

  for (const item of items) {
    if (!item.parentId) {
      roots.push(item);
      continue;
    }

    if (!byId.has(item.parentId)) {
      throw new ChecklistError(
        `Checklist item parent not found: ${item.parentId}`,
        "CHECKLIST_EXPORT_INVALID_HIERARCHY"
      );
    }

    const siblings = childrenByParentId.get(item.parentId) ?? [];
    siblings.push(item);
    childrenByParentId.set(item.parentId, siblings);
  }

  if (roots.length === 0) {
    throw new ChecklistError(
      "Checklist hierarchy has no root item",
      "CHECKLIST_EXPORT_INVALID_HIERARCHY"
    );
  }

  const rows: ExportRow[] = [];
  const visited = new Set<string>();

  const visit = (item: CheckListItemDetail, numbers: number[]): void => {
    if (visited.has(item.id)) {
      throw new ChecklistError(
        "Checklist hierarchy contains a cycle",
        "CHECKLIST_EXPORT_INVALID_HIERARCHY"
      );
    }

    visited.add(item.id);
    rows.push({
      numbers,
      name: item.name,
      description: item.description ?? "",
    });

    const children = childrenByParentId.get(item.id) ?? [];
    children.forEach((child, index) => {
      visit(child, [...numbers, index + 1]);
    });
  };

  roots.forEach((root, index) => {
    visit(root, [index + 1]);
  });

  if (visited.size !== items.length) {
    throw new ChecklistError(
      "Checklist hierarchy contains unreachable items",
      "CHECKLIST_EXPORT_INVALID_HIERARCHY"
    );
  }

  return rows;
};

const createWorksheetXml = (rows: ExportRow[]): string => {
  const maxDepth = Math.max(1, ...rows.map((row) => row.numbers.length));
  const headers = [
    ...Array.from({ length: maxDepth }, (_, index) => `項番L${index + 1}`),
    "名前",
    "説明",
  ];

  const headerXml = `<row r="1">${headers
    .map((header, index) => createStringCell(index + 1, 1, header, 1))
    .join("")}</row>`;

  const dataXml = rows
    .map((row, index) => {
      const rowIndex = index + 2;
      const cells: string[] = [];
      for (let i = 0; i < maxDepth; i++) {
        const number = row.numbers[i];
        if (number !== undefined) {
          cells.push(createNumberCell(i + 1, rowIndex, number));
        }
      }
      cells.push(createStringCell(maxDepth + 1, rowIndex, row.name, 2));
      cells.push(createStringCell(maxDepth + 2, rowIndex, row.description, 2));
      return `<row r="${rowIndex}">${cells.join("")}</row>`;
    })
    .join("");

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
  <cols>
    <col min="1" max="${maxDepth}" width="10" customWidth="1"/>
    <col min="${maxDepth + 1}" max="${maxDepth + 1}" width="50" customWidth="1"/>
    <col min="${maxDepth + 2}" max="${maxDepth + 2}" width="80" customWidth="1"/>
  </cols>
  <sheetData>${headerXml}${dataXml}</sheetData>
</worksheet>`;
};

const createStylesXml = (): string => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

const crcTable = (() => {
  const table: number[] = [];
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let j = 0; j < 8; j++) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    table[i] = c >>> 0;
  }
  return table;
})();

const crc32 = (buffer: Buffer): number => {
  let crc = 0xffffffff;
  for (const byte of buffer) {
    crc = crcTable[(crc ^ byte) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
};

const dosDateTime = (date: Date): { date: number; time: number } => {
  const year = Math.max(1980, date.getFullYear());
  return {
    time:
      (date.getHours() << 11) |
      (date.getMinutes() << 5) |
      Math.floor(date.getSeconds() / 2),
    date:
      ((year - 1980) << 9) | ((date.getMonth() + 1) << 5) | date.getDate(),
  };
};

const createZip = (files: Array<{ name: string; content: string }>): Buffer => {
  const localParts: Buffer[] = [];
  const centralParts: Buffer[] = [];
  let offset = 0;
  const now = dosDateTime(new Date());

  for (const file of files) {
    const nameBuffer = Buffer.from(file.name, "utf8");
    const contentBuffer = Buffer.from(file.content, "utf8");
    const crc = crc32(contentBuffer);

    const localHeader = Buffer.alloc(30);
    localHeader.writeUInt32LE(0x04034b50, 0);
    localHeader.writeUInt16LE(20, 4);
    localHeader.writeUInt16LE(0x0800, 6);
    localHeader.writeUInt16LE(0, 8);
    localHeader.writeUInt16LE(now.time, 10);
    localHeader.writeUInt16LE(now.date, 12);
    localHeader.writeUInt32LE(crc, 14);
    localHeader.writeUInt32LE(contentBuffer.length, 18);
    localHeader.writeUInt32LE(contentBuffer.length, 22);
    localHeader.writeUInt16LE(nameBuffer.length, 26);
    localHeader.writeUInt16LE(0, 28);
    localParts.push(localHeader, nameBuffer, contentBuffer);

    const centralHeader = Buffer.alloc(46);
    centralHeader.writeUInt32LE(0x02014b50, 0);
    centralHeader.writeUInt16LE(20, 4);
    centralHeader.writeUInt16LE(20, 6);
    centralHeader.writeUInt16LE(0x0800, 8);
    centralHeader.writeUInt16LE(0, 10);
    centralHeader.writeUInt16LE(now.time, 12);
    centralHeader.writeUInt16LE(now.date, 14);
    centralHeader.writeUInt32LE(crc, 16);
    centralHeader.writeUInt32LE(contentBuffer.length, 20);
    centralHeader.writeUInt32LE(contentBuffer.length, 24);
    centralHeader.writeUInt16LE(nameBuffer.length, 28);
    centralHeader.writeUInt16LE(0, 30);
    centralHeader.writeUInt16LE(0, 32);
    centralHeader.writeUInt16LE(0, 34);
    centralHeader.writeUInt16LE(0, 36);
    centralHeader.writeUInt32LE(0, 38);
    centralHeader.writeUInt32LE(offset, 42);
    centralParts.push(centralHeader, nameBuffer);

    offset += localHeader.length + nameBuffer.length + contentBuffer.length;
  }

  const centralOffset = offset;
  const centralDirectory = Buffer.concat(centralParts);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(0, 4);
  end.writeUInt16LE(0, 6);
  end.writeUInt16LE(files.length, 8);
  end.writeUInt16LE(files.length, 10);
  end.writeUInt32LE(centralDirectory.length, 12);
  end.writeUInt32LE(centralOffset, 16);
  end.writeUInt16LE(0, 20);

  return Buffer.concat([...localParts, centralDirectory, end]);
};

export const createChecklistWorkbookBuffer = (rows: ExportRow[]): Buffer =>
  createZip([
    {
      name: "[Content_Types].xml",
      content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`,
    },
    {
      name: "_rels/.rels",
      content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
    },
    {
      name: "xl/workbook.xml",
      content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
    },
    {
      name: "xl/_rels/workbook.xml.rels",
      content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`,
    },
    {
      name: "xl/worksheets/sheet1.xml",
      content: createWorksheetXml(rows),
    },
    {
      name: "xl/styles.xml",
      content: createStylesXml(),
    },
  ]);

const sanitizeFileName = (value: string): string =>
  value.replace(/[\\/:*?"<>|]/g, "_").replace(/\s+/g, " ").trim() ||
  "checklist";

export const createChecklistExportFileName = (params: {
  checkListSetName: string;
  checkListSetId: string;
}): string =>
  `チェックリストセット_${sanitizeFileName(
    params.checkListSetName
  )}_${sanitizeFileName(params.checkListSetId)}_エクスポート.xlsx`;

export const exportChecklistSetToExcel = async (params: {
  checkListSetId: string;
  user: RequestUser;
  deps?: {
    repo?: CheckRepository;
  };
}): Promise<{ buffer: Buffer; fileName: string }> => {
  const repo = params.deps?.repo || (await makePrismaCheckRepository());
  const checkListSet = await repo.findCheckListSetDetailById(
    params.checkListSetId
  );

  assertHasOwnerAccessOrThrow(params.user, checkListSet.userId, {
    api: "exportChecklistSetToExcel",
    resourceId: params.checkListSetId,
    logger: console,
  });

  const items = await repo.findCheckListItems(
    params.checkListSetId,
    undefined,
    true
  );
  const rows = buildChecklistExportRows(items);

  return {
    buffer: createChecklistWorkbookBuffer(rows),
    fileName: createChecklistExportFileName({
      checkListSetName: checkListSet.name,
      checkListSetId: checkListSet.id,
    }),
  };
};

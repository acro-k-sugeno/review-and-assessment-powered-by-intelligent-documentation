import { inflateRawSync } from "zlib";
import { TextDecoder } from "util";
import { ulid } from "ulid";
import {
  CheckRepository,
  makePrismaCheckRepository,
} from "../domain/repository";
import { CheckListItemEntity } from "../domain/model/checklist";
import {
  RequestUser,
  assertHasOwnerAccessOrThrow,
} from "../../../core/middleware/authorization";
import { ValidationError } from "../../../core/errors";

type WorksheetRow = Map<number, string>;

type ImportRow = {
  rowNumber: number;
  depth: number;
  name: string;
  description: string;
};

const textDecoder = new TextDecoder("utf-8");

const xmlDecode = (value: string): string =>
  value
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");

const normalizeCellValue = (value: string | undefined): string =>
  (value ?? "").trim();

const getColumnIndex = (cellRef: string): number => {
  const letters = cellRef.match(/^[A-Z]+/)?.[0] ?? "";
  let index = 0;
  for (const letter of letters) {
    index = index * 26 + (letter.charCodeAt(0) - 64);
  }
  return index;
};

const findEndOfCentralDirectory = (buffer: Buffer): number => {
  const minOffset = Math.max(0, buffer.length - 0xffff - 22);
  for (let i = buffer.length - 22; i >= minOffset; i--) {
    if (buffer.readUInt32LE(i) === 0x06054b50) {
      const commentLength = buffer.readUInt16LE(i + 20);
      if (i + 22 + commentLength !== buffer.length) {
        continue;
      }
      return i;
    }
  }
  throw new ValidationError("Invalid xlsx file");
};

const unzipXlsx = (buffer: Buffer): Map<string, Buffer> => {
  if (buffer.length < 4 || buffer.readUInt32LE(0) !== 0x04034b50) {
    throw new ValidationError("Invalid xlsx file");
  }

  const files = new Map<string, Buffer>();
  const endOffset = findEndOfCentralDirectory(buffer);
  const entryCount = buffer.readUInt16LE(endOffset + 10);
  let offset = buffer.readUInt32LE(endOffset + 16);

  if (offset < 0 || offset >= buffer.length) {
    throw new ValidationError("Invalid xlsx central directory offset");
  }

  for (let i = 0; i < entryCount; i++) {
    if (offset + 46 > buffer.length) {
      throw new ValidationError("Invalid xlsx central directory");
    }

    if (buffer.readUInt32LE(offset) !== 0x02014b50) {
      throw new ValidationError("Invalid xlsx central directory");
    }

    const compressionMethod = buffer.readUInt16LE(offset + 10);
    const compressedSize = buffer.readUInt32LE(offset + 20);
    const fileNameLength = buffer.readUInt16LE(offset + 28);
    const extraLength = buffer.readUInt16LE(offset + 30);
    const commentLength = buffer.readUInt16LE(offset + 32);
    const localHeaderOffset = buffer.readUInt32LE(offset + 42);
    const fileName = buffer
      .subarray(offset + 46, offset + 46 + fileNameLength)
      .toString("utf8");

    if (
      offset + 46 + fileNameLength + extraLength + commentLength >
      buffer.length
    ) {
      throw new ValidationError("Invalid xlsx central directory entry");
    }

    if (localHeaderOffset < 0 || localHeaderOffset + 30 > buffer.length) {
      throw new ValidationError("Invalid xlsx local header offset");
    }

    if (buffer.readUInt32LE(localHeaderOffset) !== 0x04034b50) {
      throw new ValidationError("Invalid xlsx local header");
    }

    const localNameLength = buffer.readUInt16LE(localHeaderOffset + 26);
    const localExtraLength = buffer.readUInt16LE(localHeaderOffset + 28);
    const dataStart =
      localHeaderOffset + 30 + localNameLength + localExtraLength;
    if (dataStart + compressedSize > buffer.length) {
      throw new ValidationError("Invalid xlsx file data");
    }
    const compressed = buffer.subarray(dataStart, dataStart + compressedSize);

    if (compressionMethod === 0) {
      files.set(fileName, compressed);
    } else if (compressionMethod === 8) {
      files.set(fileName, inflateRawSync(compressed));
    }

    offset += 46 + fileNameLength + extraLength + commentLength;
  }

  return files;
};

const parseSharedStrings = (xml: string | undefined): string[] => {
  if (!xml) return [];

  return Array.from(xml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)).map(
    ([, si]) =>
      Array.from(si.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g))
        .map(([, text]) => xmlDecode(text))
        .join("")
  );
};

const parseWorksheetRows = (
  xml: string,
  sharedStrings: string[]
): Map<number, WorksheetRow> => {
  const rows = new Map<number, WorksheetRow>();

  for (const [, rowAttrs, rowXml] of xml.matchAll(
    /<row\b([^>]*)>([\s\S]*?)<\/row>/g
  )) {
    const rowNumber = Number(rowAttrs.match(/\br="(\d+)"/)?.[1]);
    if (!Number.isFinite(rowNumber)) continue;

    const row = new Map<number, string>();
    for (const [, cellAttrs, cellXml] of rowXml.matchAll(
      /<c\b([^>]*)>([\s\S]*?)<\/c>/g
    )) {
      const cellRef = cellAttrs.match(/\br="([^"]+)"/)?.[1];
      if (!cellRef) continue;

      const columnIndex = getColumnIndex(cellRef);
      const type = cellAttrs.match(/\bt="([^"]+)"/)?.[1];
      let value = "";

      if (type === "inlineStr") {
        value = Array.from(cellXml.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g))
          .map(([, text]) => xmlDecode(text))
          .join("");
      } else {
        const rawValue = cellXml.match(/<v>([\s\S]*?)<\/v>/)?.[1] ?? "";
        if (type === "s") {
          value = sharedStrings[Number(rawValue)] ?? "";
        } else if (type === "b") {
          value = rawValue === "1" ? "TRUE" : "FALSE";
        } else {
          value = xmlDecode(rawValue);
        }
      }

      row.set(columnIndex, normalizeCellValue(value));
    }

    rows.set(rowNumber, row);
  }

  return rows;
};

export const parseChecklistImportRows = (buffer: Buffer): ImportRow[] => {
  const files = unzipXlsx(buffer);
  const worksheetBuffer = files.get("xl/worksheets/sheet1.xml");
  if (!worksheetBuffer) {
    throw new ValidationError("Worksheet not found");
  }

  const sharedStringsBuffer = files.get("xl/sharedStrings.xml");
  const sharedStrings = parseSharedStrings(
    sharedStringsBuffer ? textDecoder.decode(sharedStringsBuffer) : undefined
  );
  const rows = parseWorksheetRows(textDecoder.decode(worksheetBuffer), sharedStrings);
  const header = rows.get(1);
  if (!header) {
    throw new ValidationError("Header row not found");
  }

  const headersByName = new Map<string, number>();
  for (const [columnIndex, value] of header.entries()) {
    if (value) {
      headersByName.set(value, columnIndex);
    }
  }

  const nameColumn = headersByName.get("名前");
  const descriptionColumn = headersByName.get("説明");
  if (!nameColumn || !descriptionColumn) {
    throw new ValidationError("Required columns are missing");
  }

  const numberColumns = Array.from(headersByName.entries())
    .map(([headerName, columnIndex]) => {
      const match = headerName.match(/^項番L(\d+)$/);
      return match ? { level: Number(match[1]), columnIndex } : null;
    })
    .filter((value): value is { level: number; columnIndex: number } => !!value)
    .sort((a, b) => a.level - b.level);

  numberColumns.forEach((column, index) => {
    if (column.level !== index + 1) {
      throw new ValidationError("Numbering columns must be sequential");
    }
  });

  const importRows: ImportRow[] = [];
  for (const [rowNumber, row] of Array.from(rows.entries()).sort(
    ([a], [b]) => a - b
  )) {
    if (rowNumber === 1) continue;

    const name = normalizeCellValue(row.get(nameColumn));
    if (!name) continue;

    const description = normalizeCellValue(row.get(descriptionColumn));
    let lastNumberColumn: { level: number; columnIndex: number } | undefined;
    for (const column of numberColumns) {
      if (normalizeCellValue(row.get(column.columnIndex))) {
        lastNumberColumn = column;
      }
    }
    const depth = lastNumberColumn?.level ?? 1;

    importRows.push({
      rowNumber,
      depth,
      name,
      description,
    });
  }

  if (importRows.length === 0) {
    throw new ValidationError("No checklist items found");
  }

  return importRows;
};

export const buildChecklistImportItems = (params: {
  setId: string;
  rows: ImportRow[];
}): CheckListItemEntity[] => {
  const stack: CheckListItemEntity[] = [];
  const items: CheckListItemEntity[] = [];

  for (const row of params.rows) {
    if (row.depth > stack.length + 1) {
      throw new ValidationError(
        `Invalid hierarchy at row ${row.rowNumber}: parent row not found`
      );
    }

    const parent = row.depth === 1 ? undefined : stack[row.depth - 2];
    const item: CheckListItemEntity = {
      id: ulid(),
      setId: params.setId,
      parentId: parent?.id,
      name: row.name,
      description: row.description,
    };

    stack[row.depth - 1] = item;
    stack.length = row.depth;
    items.push(item);
  }

  return items;
};

const extractMultipartFile = (params: {
  body: Buffer;
  contentType?: string;
}): Buffer => {
  const boundary = params.contentType?.match(/boundary=([^;]+)/i)?.[1];
  if (!boundary) {
    throw new ValidationError("Multipart boundary not found");
  }

  const delimiter = Buffer.from(`--${boundary.replace(/^"|"$/g, "")}`);
  let offset = params.body.indexOf(delimiter);

  while (offset >= 0) {
    const partStart = offset + delimiter.length;
    if (params.body.subarray(partStart, partStart + 2).toString() === "--") {
      break;
    }

    const headerStart =
      params.body.subarray(partStart, partStart + 2).toString() === "\r\n"
        ? partStart + 2
        : partStart;
    const headerEnd = params.body.indexOf(Buffer.from("\r\n\r\n"), headerStart);
    if (headerEnd < 0) break;

    const headers = params.body.subarray(headerStart, headerEnd).toString("utf8");
    const dataStart = headerEnd + 4;
    const nextBoundary = params.body.indexOf(
      Buffer.from(`\r\n--${boundary.replace(/^"|"$/g, "")}`),
      dataStart
    );
    if (nextBoundary < 0) break;

    if (
      /name="file"/i.test(headers) &&
      /filename="[^"]+"/i.test(headers)
    ) {
      return params.body.subarray(dataStart, nextBoundary);
    }

    offset = params.body.indexOf(delimiter, nextBoundary + 2);
  }

  throw new ValidationError("Excel file is required");
};

export const extractChecklistImportFile = extractMultipartFile;

export const importChecklistSetFromExcel = async (params: {
  checkListSetId: string;
  user: RequestUser;
  fileBuffer: Buffer;
  deps?: {
    repo?: CheckRepository;
  };
}): Promise<{ createdCount: number; deletedCount: number }> => {
  const repo = params.deps?.repo || (await makePrismaCheckRepository());
  const checkListSet = await repo.findCheckListSetDetailById(
    params.checkListSetId
  );

  assertHasOwnerAccessOrThrow(params.user, checkListSet.userId, {
    api: "importChecklistSetFromExcel",
    resourceId: params.checkListSetId,
    logger: console,
  });

  const isEditable = await repo.checkSetEditable({
    setId: params.checkListSetId,
  });
  if (!isEditable) {
    throw new ValidationError("Set is not editable");
  }

  const rows = parseChecklistImportRows(params.fileBuffer);
  const items = buildChecklistImportItems({
    setId: params.checkListSetId,
    rows,
  });

  return repo.replaceCheckListItems({
    setId: params.checkListSetId,
    items,
  });
};

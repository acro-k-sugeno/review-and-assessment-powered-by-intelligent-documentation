import { describe, expect, it, vi } from "vitest";
import { CHECK_LIST_STATUS } from "../domain/model/checklist";
import { createChecklistWorkbookBuffer } from "./checklist-export";
import {
  buildChecklistImportItems,
  importChecklistSetFromExcel,
  parseChecklistImportRows,
} from "./checklist-import";
import {
  ForbiddenError,
  ValidationError,
} from "../../../core/errors/application-errors";

describe("parseChecklistImportRows", () => {
  it("parses exported workbook rows and ignores numbering values for stored data", () => {
    const buffer = createChecklistWorkbookBuffer([
      { numbers: [10], name: "Root", description: "" },
      { numbers: [10, 99], name: "Child", description: "Description" },
    ]);

    expect(parseChecklistImportRows(buffer)).toEqual([
      { rowNumber: 2, depth: 1, name: "Root", description: "" },
      { rowNumber: 3, depth: 2, name: "Child", description: "Description" },
    ]);
  });

  it("accepts base64 encoded xlsx content", () => {
    const buffer = createChecklistWorkbookBuffer([
      { numbers: [1], name: "Root", description: "" },
    ]);
    const base64Buffer = Buffer.from(buffer.toString("base64"), "utf8");

    expect(parseChecklistImportRows(base64Buffer)).toEqual([
      { rowNumber: 2, depth: 1, name: "Root", description: "" },
    ]);
  });

  it("rejects hierarchy jumps", () => {
    const rows = [
      { rowNumber: 2, depth: 1, name: "Root", description: "" },
      { rowNumber: 3, depth: 3, name: "Grandchild", description: "" },
    ];

    expect(() =>
      buildChecklistImportItems({ setId: "set-1", rows })
    ).toThrow(ValidationError);
  });
});

describe("importChecklistSetFromExcel", () => {
  const checkListSet = {
    id: "set-1",
    name: "Set 1",
    description: "",
    userId: "owner-1",
    documents: [],
    processingStatus: CHECK_LIST_STATUS.COMPLETED,
    isEditable: true,
    hasError: false,
  };

  it("replaces all existing checklist items with new imported items", async () => {
    const buffer = createChecklistWorkbookBuffer([
      { numbers: [1], name: "Root", description: "" },
      { numbers: [1, 1], name: "Child", description: "Description" },
    ]);
    const repo = {
      findCheckListSetDetailById: vi.fn().mockResolvedValue(checkListSet),
      checkSetEditable: vi.fn().mockResolvedValue(true),
      replaceCheckListItems: vi
        .fn()
        .mockResolvedValue({ deletedCount: 3, createdCount: 2 }),
    };

    const result = await importChecklistSetFromExcel({
      checkListSetId: "set-1",
      user: { userId: "owner-1", isAdmin: false },
      fileBuffer: buffer,
      deps: { repo: repo as any },
    });

    expect(result).toEqual({ deletedCount: 3, createdCount: 2 });
    expect(repo.replaceCheckListItems).toHaveBeenCalledWith({
      setId: "set-1",
      items: expect.arrayContaining([
        expect.objectContaining({ name: "Root", parentId: undefined }),
        expect.objectContaining({ name: "Child", description: "Description" }),
      ]),
    });
    const items = repo.replaceCheckListItems.mock.calls[0][0].items;
    expect(items[1].parentId).toBe(items[0].id);
  });

  it("throws ForbiddenError when user is not owner", async () => {
    const repo = {
      findCheckListSetDetailById: vi.fn().mockResolvedValue(checkListSet),
      checkSetEditable: vi.fn().mockResolvedValue(true),
      replaceCheckListItems: vi.fn(),
    };

    await expect(
      importChecklistSetFromExcel({
        checkListSetId: "set-1",
        user: { userId: "other-1", isAdmin: false },
        fileBuffer: createChecklistWorkbookBuffer([
          { numbers: [1], name: "Root", description: "" },
        ]),
        deps: { repo: repo as any },
      })
    ).rejects.toBeInstanceOf(ForbiddenError);
  });

  it("rejects non-editable checklist sets before replacing items", async () => {
    const repo = {
      findCheckListSetDetailById: vi.fn().mockResolvedValue(checkListSet),
      checkSetEditable: vi.fn().mockResolvedValue(false),
      replaceCheckListItems: vi.fn(),
    };

    await expect(
      importChecklistSetFromExcel({
        checkListSetId: "set-1",
        user: { userId: "owner-1", isAdmin: false },
        fileBuffer: createChecklistWorkbookBuffer([
          { numbers: [1], name: "Root", description: "" },
        ]),
        deps: { repo: repo as any },
      })
    ).rejects.toBeInstanceOf(ValidationError);
    expect(repo.replaceCheckListItems).not.toHaveBeenCalled();
  });
});

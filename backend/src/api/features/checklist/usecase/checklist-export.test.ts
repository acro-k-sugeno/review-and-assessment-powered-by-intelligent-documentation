import { describe, expect, it, vi } from "vitest";
import {
  buildChecklistExportRows,
  createChecklistExportFileName,
  createChecklistWorkbookBuffer,
  exportChecklistSetToExcel,
} from "./checklist-export";
import {
  ChecklistError,
  ForbiddenError,
} from "../../../core/errors/application-errors";
import { CHECK_LIST_STATUS, CheckListItemDetail } from "../domain/model/checklist";

const makeItem = (params: {
  id: string;
  parentId?: string;
  name: string;
  description?: string;
}): CheckListItemDetail => ({
  id: params.id,
  parentId: params.parentId,
  setId: "set-1",
  name: params.name,
  description: params.description,
  hasChildren: false,
});

describe("buildChecklistExportRows", () => {
  it("numbers root and nested checklist items by hierarchy", () => {
    const rows = buildChecklistExportRows([
      makeItem({ id: "root-1", name: "Root 1" }),
      makeItem({ id: "child-1", parentId: "root-1", name: "Child 1" }),
      makeItem({ id: "child-2", parentId: "root-1", name: "Child 2" }),
      makeItem({
        id: "grandchild-1",
        parentId: "child-2",
        name: "Grandchild 1",
      }),
      makeItem({ id: "root-2", name: "Root 2" }),
    ]);

    expect(rows).toEqual([
      { numbers: [1], name: "Root 1", description: "" },
      { numbers: [1, 1], name: "Child 1", description: "" },
      { numbers: [1, 2], name: "Child 2", description: "" },
      { numbers: [1, 2, 1], name: "Grandchild 1", description: "" },
      { numbers: [2], name: "Root 2", description: "" },
    ]);
  });

  it("allows checklist items deeper than six levels", () => {
    const items = [
      makeItem({ id: "l1", name: "L1" }),
      makeItem({ id: "l2", parentId: "l1", name: "L2" }),
      makeItem({ id: "l3", parentId: "l2", name: "L3" }),
      makeItem({ id: "l4", parentId: "l3", name: "L4" }),
      makeItem({ id: "l5", parentId: "l4", name: "L5" }),
      makeItem({ id: "l6", parentId: "l5", name: "L6" }),
      makeItem({ id: "l7", parentId: "l6", name: "L7" }),
    ];

    expect(buildChecklistExportRows(items).at(-1)).toEqual({
      numbers: [1, 1, 1, 1, 1, 1, 1],
      name: "L7",
      description: "",
    });
  });

  it("rejects orphan checklist items", () => {
    expect(() =>
      buildChecklistExportRows([
        makeItem({ id: "orphan", parentId: "missing", name: "Orphan" }),
      ])
    ).toThrow(ChecklistError);
  });
});

describe("createChecklistWorkbookBuffer", () => {
  it("creates an xlsx zip buffer", () => {
    const buffer = createChecklistWorkbookBuffer([
      {
        numbers: [1],
        name: "構造関連事項",
        description: "",
      },
    ]);

    expect(buffer.subarray(0, 4).toString("binary")).toBe("PK\u0003\u0004");
    expect(buffer.includes(Buffer.from("xl/worksheets/sheet1.xml"))).toBe(
      true
    );
  });

  it("places name and description after the deepest numbering column", () => {
    const buffer = createChecklistWorkbookBuffer([
      {
        numbers: [1, 1, 1],
        name: "Level 3",
        description: "Description",
      },
    ]);
    const content = buffer.toString("utf8");

    expect(content).toContain("項番L3");
    expect(content).not.toContain("項番L4");
    expect(content).toContain('r="D1"');
    expect(content).toContain('r="E1"');
  });
});

describe("createChecklistExportFileName", () => {
  it("sanitizes invalid filename characters", () => {
    expect(
      createChecklistExportFileName({
        checkListSetName: 'A/B:C*D?E"F<G>H|',
        checkListSetId: "set-1",
      })
    ).toBe("チェックリストセット_A_B_C_D_E_F_G_H__set-1_エクスポート.xlsx");
  });
});

describe("exportChecklistSetToExcel", () => {
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

  it("throws ForbiddenError when user is not owner", async () => {
    const repo = {
      findCheckListSetDetailById: vi.fn().mockResolvedValue(checkListSet),
      findCheckListItems: vi.fn().mockResolvedValue([]),
    };

    await expect(
      exportChecklistSetToExcel({
        checkListSetId: "set-1",
        user: { userId: "other-1", isAdmin: false },
        deps: { repo: repo as any },
      })
    ).rejects.toBeInstanceOf(ForbiddenError);
  });

  it("allows an admin to export another user's checklist set", async () => {
    const repo = {
      findCheckListSetDetailById: vi.fn().mockResolvedValue(checkListSet),
      findCheckListItems: vi
        .fn()
        .mockResolvedValue([makeItem({ id: "root-1", name: "Root" })]),
    };

    const result = await exportChecklistSetToExcel({
      checkListSetId: "set-1",
      user: { userId: "admin-1", isAdmin: true },
      deps: { repo: repo as any },
    });

    expect(result.fileName).toBe(
      "チェックリストセット_Set 1_set-1_エクスポート.xlsx"
    );
    expect(result.buffer.length).toBeGreaterThan(0);
  });
});

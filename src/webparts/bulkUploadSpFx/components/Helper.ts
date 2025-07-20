// Helper: Combine a date and time into one local Date object.
export const combineDateAndTimeLocal = (date: Date, time: Date): Date => {
  const year = date.getFullYear();
  const month = date.getMonth();
  const day = date.getDate();
  const hours = time.getHours();
  const minutes = time.getMinutes();
  const seconds = time.getSeconds();
  // Returns a local time Date
  return new Date(year, month, day, hours, minutes, seconds);
};

export const buttonStyle = {
  backgroundColor: "#0059a8",
};

export const stateHelpers = {
  openDialog: () => ({
    isDialogOpen: true,
    missingRows: [],
  }),
  closeDialog: () => ({
    isDialogOpen: false,
    missingRows: [],
  }),
  closePanel: () => ({
    isPanelOpen: false,
  }),
  openPanel: (formMode: number) => ({
    isPanelOpen: true,
    formMode: formMode,
  }),
};

// helper.ts

export const filterItems = (items: any[], searchTerm: string): any[] => {
  const term = searchTerm.toLowerCase();
  return items.filter(
    (item) =>
      item.EmployeeName.toLowerCase().includes(term) ||
      item.EmployeeWorkEmail.toLowerCase().includes(term) ||
      item.EmployeePersonalEmail.toLowerCase().includes(term)
  );
};
// helper.ts

import { IColumn } from "@fluentui/react";

// Updates the columns: toggles the sort order on the clicked column and resets others.
// Returns the updated columns array and the clicked column object (if found).
export function updateColumnsOnClick(
  columns: IColumn[],
  clickedColumnKey: string
): { updatedColumns: IColumn[]; sortedColumn: IColumn | undefined } {
  // Create a shallow copy of the columns.
  const newColumns = columns.slice();
  // Find the column that was clicked.
  const currColumn = newColumns.find((col) => col.key === clickedColumnKey);
  if (!currColumn) {
    return { updatedColumns: newColumns, sortedColumn: undefined };
  }

  // Update each column: toggle sort for the clicked one, and reset others.
  newColumns.forEach((col: IColumn) => {
    if (col.key === currColumn.key) {
      currColumn.isSortedDescending = !currColumn.isSortedDescending;
      currColumn.isSorted = true;
    } else {
      col.isSorted = false;
      col.isSortedDescending = true;
    }
  });
  return { updatedColumns: newColumns, sortedColumn: currColumn };
}

// Sorts the items based on a column field and sort order.
// The function assumes the field exists on each item.
export function copyAndSort<T>(
  items: T[],
  columnKey: string,
  isSortedDescending: boolean
): T[] {
  const key = columnKey as keyof T;
  return items.slice(0).sort((a: T, b: T) => {
    // Note: Adjust the sort condition if you need to handle numbers or dates differently.
    return (isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1;
  });
}
export const convertDateOnlyToUTC = (dateValue: any): string => {
  const date = new Date(dateValue);
  // Extract year, month, and day
  const year = date.getFullYear();
  const month = date.getMonth(); // zero-based
  const day = date.getDate();
  // Create a date at midnight UTC
  const utcDate = new Date(Date.UTC(year, month, day));
  return utcDate.toISOString();
};

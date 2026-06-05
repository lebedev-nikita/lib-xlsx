import { Column } from "./DataSheet.js";

export function column<const TId extends string>(
  id: TId,
  label: string,
  options?: {
    skipped?: boolean;
    rowSpan?: number;
    width?: number;
  },
) {
  return new Column({ id, label, ...options });
}

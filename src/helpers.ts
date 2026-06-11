import { Column, Group, HeaderItem } from "./DataSheet.js";

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

export function colgroup<const TId extends string>(label: string, children: HeaderItem<TId>[]) {
  return new Group({ label, children });
}

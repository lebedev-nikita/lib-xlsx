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

export function colgroup<const TId extends string>(
  label: string,
  options: { rowSpan?: number },
  children: HeaderItem<TId>[],
): Group<TId>;
export function colgroup<const TId extends string>(
  label: string,
  children: HeaderItem<TId>[],
): Group<TId>;

export function colgroup(label: string, optionsOrChildren: any, children?: any) {
  if (!children) {
    children = optionsOrChildren;
    optionsOrChildren = {};
  }
  return new Group({ label, children, rowSpan: optionsOrChildren.rowSpan });
}

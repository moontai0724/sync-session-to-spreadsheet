export function toMapById<T extends { id: string }>(
  items: T[],
): Record<string, T> {
  return items.reduce((all, item) => {
    all[item.id] = item;
    return all;
  }, {} as Record<string, T>);
}

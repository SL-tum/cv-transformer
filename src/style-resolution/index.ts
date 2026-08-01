import type { NodeStyle } from "../model/core/node.js";
import type { StyleRegistry } from "../model/core/registries.js";

export interface ResolveStyleOptions<T extends Record<string, unknown>> {
  registry: StyleRegistry;
  styleId?: string;
  themeDefaults?: Partial<T>;
  direct?: Partial<T>;
}

/** Resolves defaults → theme → basedOn chain → referenced style → direct formatting. */
export function resolveStyle<T extends Record<string, unknown> = Record<string, unknown>>(
  options: ResolveStyleOptions<T>,
): T {
  const output: Record<string, unknown> = {
    ...(options.registry.documentDefaults ?? {}),
    ...(options.themeDefaults ?? {}),
  };
  const visiting = new Set<string>();
  const resolved = new Set<string>();
  const apply = (id: string) => {
    if (resolved.has(id)) return;
    if (visiting.has(id)) throw new Error(`Style inheritance cycle detected at ${id}`);
    const style = options.registry.styles[id];
    if (!style) return;
    visiting.add(id);
    if (style.basedOn) apply(style.basedOn);
    Object.assign(output, style.properties);
    visiting.delete(id);
    resolved.add(id);
  };
  if (options.styleId) apply(options.styleId);
  Object.assign(output, options.direct ?? {});
  return output as T;
}

export function computeNodeStyle<T extends Record<string, unknown>>(
  style: NodeStyle<T>,
  registry: StyleRegistry,
  themeDefaults?: Partial<T>,
): T {
  const styleId = style.references.at(-1)?.registryId;
  const computed = resolveStyle<T>({
    registry,
    ...(styleId ? { styleId } : {}),
    ...(themeDefaults ? { themeDefaults } : {}),
    ...(style.direct ? { direct: style.direct } : {}),
  });
  style.computed = computed;
  return computed;
}

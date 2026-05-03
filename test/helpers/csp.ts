export type ParsedCspDirective = { directive: string; values: string[] };

export function parseCsp(csp: string): ParsedCspDirective[] {
  return csp
    .split(";")
    .map((seg) => {
      const [directive, ...values] = seg.trim().split(/\s+/);
      return { directive: directive ?? "", values };
    })
    .filter((d) => d.directive);
}

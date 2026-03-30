/**
 * AST-based eval code validator (Layer 2 defense-in-depth).
 *
 * Primary defense is runtime global shadowing in the eval wrapper (Layer 1).
 * This module provides static validation that catches patterns shadows miss:
 * - Prototype chain access to Function constructor (.constructor, .__proto__)
 * - Dynamic import() expressions
 * - Defense-in-depth identifier and member expression blocks
 */

import * as acorn from "acorn";

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

export const MAX_EVAL_CODE_LENGTH = 10_000;

/** Blocked property names on ANY object (prototype chain / meta-programming) */
const BLOCKED_PROPERTY_NAMES = new Set(["constructor", "__proto__"]);

/**
 * Defense-in-depth: blocked identifiers as standalone references.
 * The runtime shadow layer is the primary defense for these — this catches
 * patterns that might somehow bypass shadowing.
 */
const BLOCKED_IDENTIFIERS = new Set([
  "fetch",
  "XMLHttpRequest",
  "WebSocket",
  "EventSource",
  "Worker",
  "SharedWorker",
  "importScripts",
  "eval",
  "Function",
  "Image",
]);

/**
 * Defense-in-depth: blocked object.property member expression pairs.
 * These are also blocked by the document proxy at runtime.
 */
const BLOCKED_MEMBER_ACCESS = new Map<string, Set<string>>([
  ["document", new Set(["cookie", "domain", "write", "writeln"])],
  ["navigator", new Set(["sendBeacon"])],
]);

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export type EvalValidationResult = { safe: true } | { safe: false; reason: string };

// ---------------------------------------------------------------------------
// AST node type predicates (acorn uses a generic Node type)
// ---------------------------------------------------------------------------

interface AcornNode {
  type: string;
  [key: string]: unknown;
}

interface IdentifierNode extends AcornNode {
  type: "Identifier";
  name: string;
}

interface MemberExpressionNode extends AcornNode {
  type: "MemberExpression";
  object: AcornNode;
  property: AcornNode;
  computed: boolean;
}

interface PropertyNode extends AcornNode {
  type: "Property";
  key: AcornNode;
  value: AcornNode;
  shorthand: boolean;
  computed: boolean;
}

// ---------------------------------------------------------------------------
// AST walker
// ---------------------------------------------------------------------------

/**
 * Recursively walk the AST. Returns a violation reason string or null if safe.
 *
 * The walker handles the critical subtlety of MemberExpression properties:
 * - Non-computed `.property` identifiers are only checked against BLOCKED_PROPERTY_NAMES
 *   (not the full BLOCKED_IDENTIFIERS set) to avoid false positives on `el.dataset.fetch`
 * - Standalone identifiers (call targets, variable refs, etc.) ARE checked against
 *   BLOCKED_IDENTIFIERS
 * - MemberExpression object+property pairs are checked against BLOCKED_MEMBER_ACCESS
 */
function checkNode(
  node: AcornNode,
  parentNode: AcornNode | null,
  parentKey: string | null,
): string | null {
  if (!node || typeof node !== "object" || !node.type) {
    return null;
  }

  // --- ImportExpression ---
  if (node.type === "ImportExpression") {
    return "dynamic import() is blocked";
  }

  // --- MemberExpression: check property names and object.property pairs ---
  if (node.type === "MemberExpression") {
    const mem = node as MemberExpressionNode;

    // Check property name against BLOCKED_PROPERTY_NAMES (constructor, __proto__)
    if (!mem.computed && mem.property.type === "Identifier") {
      const propName = (mem.property as IdentifierNode).name;
      if (BLOCKED_PROPERTY_NAMES.has(propName)) {
        return `access to .${propName} is blocked`;
      }
    }

    // Check object.property against BLOCKED_MEMBER_ACCESS (document.cookie, etc.)
    if (!mem.computed && mem.object.type === "Identifier" && mem.property.type === "Identifier") {
      const objName = (mem.object as IdentifierNode).name;
      const propName = (mem.property as IdentifierNode).name;
      const blocked = BLOCKED_MEMBER_ACCESS.get(objName);
      if (blocked?.has(propName)) {
        return `${objName}.${propName} is blocked`;
      }
    }

    // Recurse into object (but NOT into non-computed property — handled above)
    const objViolation = checkNode(mem.object, node, "object");
    if (objViolation) {
      return objViolation;
    }

    // Only recurse into computed properties (dynamic expressions that could contain identifiers)
    if (mem.computed) {
      const propViolation = checkNode(mem.property, node, "property");
      if (propViolation) {
        return propViolation;
      }
    }

    return null;
  }

  // --- Identifier: defense-in-depth check ---
  if (node.type === "Identifier") {
    const id = node as IdentifierNode;

    // Skip if this identifier is a non-computed MemberExpression property
    // (already handled above — el.dataset.fetch should NOT trigger)
    if (
      parentNode?.type === "MemberExpression" &&
      parentKey === "property" &&
      !(parentNode as MemberExpressionNode).computed
    ) {
      return null;
    }

    // Skip if this is a Property key in an object literal or destructuring pattern
    // (e.g., { fetch: "value" } or { fetch: alias } in destructuring)
    if (
      parentNode?.type === "Property" &&
      parentKey === "key" &&
      !(parentNode as PropertyNode).computed
    ) {
      return null;
    }

    if (BLOCKED_IDENTIFIERS.has(id.name)) {
      return `reference to '${id.name}' is blocked`;
    }
    return null;
  }

  // --- Property node: handle key/value traversal carefully ---
  if (node.type === "Property") {
    const prop = node as PropertyNode;
    // Only check key if it's computed (dynamic expression)
    if (prop.computed) {
      const keyViolation = checkNode(prop.key, node, "key");
      if (keyViolation) {
        return keyViolation;
      }
    } else {
      // For non-computed keys, still check for BLOCKED_PROPERTY_NAMES
      // but pass context so Identifier handler skips BLOCKED_IDENTIFIERS
      const keyViolation = checkNode(prop.key, node, "key");
      if (keyViolation) {
        return keyViolation;
      }
    }
    // Always check value
    const valViolation = checkNode(prop.value, node, "value");
    if (valViolation) {
      return valViolation;
    }
    return null;
  }

  // --- Generic recursive walk for all other node types ---
  for (const key of Object.keys(node)) {
    if (key === "type" || key === "start" || key === "end" || key === "loc" || key === "range") {
      continue;
    }
    const child = node[key];
    if (Array.isArray(child)) {
      for (const item of child) {
        if (item && typeof item === "object" && (item as AcornNode).type) {
          const violation = checkNode(item as AcornNode, node, key);
          if (violation) {
            return violation;
          }
        }
      }
    } else if (child && typeof child === "object" && (child as AcornNode).type) {
      const violation = checkNode(child as AcornNode, node, key);
      if (violation) {
        return violation;
      }
    }
  }

  return null;
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

export function validateBrowserEvalCode(code: string): EvalValidationResult {
  const trimmed = code.trim();
  if (!trimmed) {
    return { safe: false, reason: "empty code" };
  }
  if (trimmed.length > MAX_EVAL_CODE_LENGTH) {
    return { safe: false, reason: "code exceeds max length" };
  }

  let ast: acorn.Node;
  try {
    // Wrap in parens to parse as expression (same as eval("(" + code + ")"))
    ast = acorn.parse("(" + trimmed + ")", {
      ecmaVersion: 2022,
      sourceType: "script",
    });
  } catch {
    return { safe: false, reason: "invalid JavaScript syntax" };
  }

  const violation = checkNode(ast as unknown as AcornNode, null, null);
  if (violation) {
    return { safe: false, reason: violation };
  }

  return { safe: true };
}

export function assertSafeEvalCode(code: string): void {
  const result = validateBrowserEvalCode(code);
  if (!result.safe) {
    throw new Error("Blocked unsafe browser eval code: " + result.reason);
  }
}

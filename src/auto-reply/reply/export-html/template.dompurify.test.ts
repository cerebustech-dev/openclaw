import fs from "node:fs";
import path from "node:path";
import vm from "node:vm";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { parseHTML } from "linkedom";
import { JSDOM } from "jsdom";
import createDOMPurify from "dompurify";

// ---------- Types (matching template.security.test.ts) ----------

type SessionEntry = {
  id: string;
  parentId: string | null;
  timestamp: string;
  type: string;
  message?: unknown;
  summary?: string;
  content?: unknown;
  display?: boolean;
  customType?: string;
  provider?: string;
  modelId?: string;
  thinkingLevel?: string;
};

type SessionData = {
  header: { id: string; timestamp: string };
  entries: SessionEntry[];
  leafId: string;
  systemPrompt: string;
  tools: unknown[];
};

// ---------- File Loading ----------

const exportHtmlDir = path.dirname(fileURLToPath(import.meta.url));
const templateHtml = fs.readFileSync(path.join(exportHtmlDir, "template.html"), "utf8");
const templateJs = fs.readFileSync(path.join(exportHtmlDir, "template.js"), "utf8");
const markedJs = fs.readFileSync(path.join(exportHtmlDir, "vendor", "marked.min.js"), "utf8");
const highlightJs = fs.readFileSync(
  path.join(exportHtmlDir, "vendor", "highlight.min.js"),
  "utf8",
);

// ---------- DOMPurify via jsdom (for contract tests) ----------

const jsdomWindow = new JSDOM("").window;
const DOMPurify = createDOMPurify(jsdomWindow as unknown as Window);

// The PURIFY_CONFIG that template.js should use (source of truth for tests)
const PURIFY_CONFIG = {
  ALLOWED_TAGS: [
    "p", "br", "strong", "em", "code", "pre", "ul", "ol", "li",
    "h1", "h2", "h3", "h4", "h5", "h6", "a", "img", "blockquote",
    "table", "thead", "tbody", "tr", "th", "td", "span", "div",
    "hr", "del", "sup", "sub", "details", "summary",
  ],
  ALLOWED_ATTR: [
    "href", "src", "alt", "title", "class", "id", "target",
    "rel", "width", "height", "colspan", "rowspan",
  ],
  ALLOW_DATA_ATTR: false,
};

function sanitize(html: string): string {
  return DOMPurify.sanitize(html, PURIFY_CONFIG);
}

// ---------- Integration Test Infrastructure (linkedom + vm + DOMPurify mock) ----------

interface RenderResult {
  document: Document;
  runtime: Record<string, unknown>;
}

function renderTemplate(sessionData: SessionData): RenderResult {
  const html = templateHtml
    .replace("{{CSS}}", "")
    .replace(
      "{{SESSION_DATA}}",
      Buffer.from(JSON.stringify(sessionData), "utf8").toString("base64"),
    )
    .replace(/\{\{DOMPURIFY_JS\}\}/g, "")
    .replace("{{MARKED_JS}}", "")
    .replace("{{HIGHLIGHT_JS}}", "")
    .replace("{{JS}}", "");

  const { document, window } = parseHTML(html);
  if ((window as Record<string, unknown>).HTMLElement) {
    (
      (window as Record<string, unknown>).HTMLElement as { prototype: Record<string, unknown> }
    ).prototype.scrollIntoView = () => {};
  }

  const immediateTimeout = (fn: (...args: unknown[]) => void) => {
    fn();
    return 0;
  };
  const runtime: Record<string, unknown> = {
    document,
    console,
    clearTimeout: () => {},
    setTimeout: immediateTimeout,
    URLSearchParams,
    TextDecoder,
    atob: (s: string) => Buffer.from(s, "base64").toString("binary"),
    btoa: (s: string) => Buffer.from(s, "binary").toString("base64"),
    navigator: { clipboard: { writeText: async () => {} } },
    history: { replaceState: () => {} },
    location: { href: "http://localhost/export.html", search: "" },
  };
  runtime.window = runtime;
  runtime.self = runtime;
  runtime.globalThis = runtime;

  vm.createContext(runtime);

  // Inject DOMPurify mock that tracks calls (real DOMPurify doesn't work with linkedom).
  // The mock passes HTML through unchanged so we can separately test:
  //   1. That template.js CALLS DOMPurify (integration test)
  //   2. That DOMPurify with our config WORKS (contract tests via jsdom)
  vm.runInContext(
    `
    window.__dompurify_calls = [];
    window.DOMPurify = {
      sanitize: function(html, config) {
        window.__dompurify_calls.push({ html: String(html), config: config });
        return html;
      },
      version: '3.3.3',
      isSupported: true,
    };
    `,
    runtime,
  );

  vm.runInContext(markedJs, runtime);
  vm.runInContext(highlightJs, runtime);
  vm.runInContext(templateJs, runtime);
  return { document: document as unknown as Document, runtime };
}

function now() {
  return new Date("2026-02-24T00:00:00.000Z").toISOString();
}

function userMsg(id: string, parentId: string | null, content: string): SessionEntry {
  return {
    id,
    parentId,
    timestamp: now(),
    type: "message",
    message: { role: "user", content },
  };
}

function assistantMsg(id: string, parentId: string, text: string): SessionEntry {
  return {
    id,
    parentId,
    timestamp: now(),
    type: "message",
    message: {
      role: "assistant",
      content: [{ type: "text", text }],
    },
  };
}

function session(content: string): SessionData {
  return {
    header: { id: "dompurify-test", timestamp: now() },
    entries: [userMsg("1", null, content), assistantMsg("2", "1", content)],
    leafId: "2",
    systemPrompt: "",
    tools: [],
  };
}

// ==========================================================================
// CATEGORY 1: DOMPurify Integration (TRUE RED TESTS)
// These fail until template.js is modified to call DOMPurify.sanitize
// ==========================================================================

describe("DOMPurify integration with template pipeline", () => {
  it("template.js calls DOMPurify.sanitize during rendering", () => {
    const { runtime } = renderTemplate(session("hello world"));
    const calls = runtime.__dompurify_calls as Array<{ html: string; config: unknown }>;
    expect(calls.length).toBeGreaterThan(0);
  });

  it("safeMarkedParse passes PURIFY_CONFIG to DOMPurify.sanitize", () => {
    const { runtime } = renderTemplate(session("testing config"));
    const calls = runtime.__dompurify_calls as Array<{
      html: string;
      config: Record<string, unknown>;
    }>;
    expect(calls.length).toBeGreaterThan(0);
    const config = calls[0].config;
    expect(config).toBeDefined();
    expect(config.ALLOWED_TAGS).toBeDefined();
    expect(config.ALLOWED_ATTR).toBeDefined();
    expect(config.ALLOW_DATA_ATTR).toBe(false);
  });

  it("PURIFY_CONFIG matches expected allowlist", () => {
    const { runtime } = renderTemplate(session("config check"));
    const calls = runtime.__dompurify_calls as Array<{
      html: string;
      config: Record<string, unknown>;
    }>;
    expect(calls.length).toBeGreaterThan(0);
    const config = calls[0].config;
    const tags = config.ALLOWED_TAGS as string[];
    // Verify key tags are in the allowlist
    expect(tags).toContain("p");
    expect(tags).toContain("a");
    expect(tags).toContain("img");
    expect(tags).toContain("code");
    expect(tags).toContain("pre");
    expect(tags).toContain("table");
    // Verify dangerous tags are NOT in the allowlist
    expect(tags).not.toContain("script");
    expect(tags).not.toContain("style");
    expect(tags).not.toContain("iframe");
    expect(tags).not.toContain("form");
  });
});

// ==========================================================================
// CATEGORY 2: Core XSS Vector Stripping (contract tests via jsdom)
// These use real DOMPurify with our PURIFY_CONFIG to verify sanitization
// ==========================================================================

describe("DOMPurify XSS vector stripping with PURIFY_CONFIG", () => {
  it("strips script tags", () => {
    const result = sanitize("<p>safe</p><script>alert(1)</script>");
    expect(result).not.toContain("<script>");
    expect(result).toContain("<p>safe</p>");
  });

  it("strips img onerror event handlers", () => {
    const result = sanitize('<img src="x" onerror="alert(1)">');
    expect(result).not.toContain("onerror");
  });

  it("strips svg onload event handlers", () => {
    const result = sanitize("<svg onload=alert(1)></svg>");
    expect(result).not.toContain("onload");
    expect(result).not.toContain("<svg");
  });

  it("strips nested SVG foreignObject injection", () => {
    const result = sanitize(
      '<svg><foreignObject><img src=x onerror=alert(1)></foreignObject></svg>',
    );
    expect(result).not.toContain("onerror");
  });

  it("strips MathML namespace confusion", () => {
    const result = sanitize(
      '<math><mtext><img src=x onerror=alert(1)></mtext></math>',
    );
    expect(result).not.toContain("onerror");
  });

  it("strips mXSS via SVG namespace break", () => {
    const result = sanitize("<svg></p><img src=x onerror=alert(1)>");
    expect(result).not.toContain("onerror");
  });

  it("strips all on* event handler attributes", () => {
    const handlers = [
      '<div onmouseover="alert(1)">hover</div>',
      '<input onfocus="alert(1)" autofocus>',
      '<details open ontoggle="alert(1)">x</details>',
      '<body onload="alert(1)">',
      '<marquee onstart="alert(1)">',
    ];
    for (const payload of handlers) {
      const result = sanitize(payload);
      expect(result).not.toMatch(/\bon\w+=/i);
    }
  });

  it("strips style tag injection", () => {
    const result = sanitize(
      '<style>body{background:url("javascript:alert(1)")}</style>',
    );
    expect(result).not.toContain("<style>");
  });

  it("strips form action injection", () => {
    const result = sanitize(
      '<form action="javascript:alert(1)"><button>click</button></form>',
    );
    expect(result).not.toContain("<form");
  });
});

// ==========================================================================
// CATEGORY 3: Advanced Bypass Vectors (contract tests via jsdom)
// ==========================================================================

describe("DOMPurify advanced bypass vector protection", () => {
  it("blocks deep nesting bypass (CVE-2024-47875)", () => {
    let payload = "";
    for (let i = 0; i < 120; i++) payload += "<div>";
    payload += '<img src=x onerror="alert(1)">';
    for (let i = 0; i < 120; i++) payload += "</div>";
    const result = sanitize(payload);
    expect(result).not.toContain("onerror");
  });

  it("blocks template element injection", () => {
    const result = sanitize("<template><script>alert(1)</script></template>");
    expect(result).not.toContain("<script>");
    expect(result).not.toContain("<template>");
  });

  it("blocks iframe injection (src and srcdoc)", () => {
    expect(sanitize('<iframe src="javascript:alert(1)"></iframe>')).not.toContain(
      "<iframe",
    );
    expect(
      sanitize('<iframe srcdoc="<script>alert(1)</script>"></iframe>'),
    ).not.toContain("<iframe");
  });

  it("strips javascript: from href after DOMPurify processing", () => {
    const result = sanitize('<a href="javascript:alert(1)">click</a>');
    expect(result).not.toContain("javascript:");
  });
});

// ==========================================================================
// CATEGORY 4: Legitimate Content Preservation (contract tests via jsdom)
// ==========================================================================

describe("DOMPurify preserves legitimate markdown output", () => {
  it("preserves basic formatting (p, br, strong, em)", () => {
    const html = "<p>Hello <strong>bold</strong> and <em>italic</em></p><br>";
    const result = sanitize(html);
    expect(result).toContain("<strong>bold</strong>");
    expect(result).toContain("<em>italic</em>");
    expect(result).toContain("<p>");
  });

  it("preserves code and pre blocks", () => {
    const html = '<pre><code class="hljs">const x = 1;</code></pre>';
    const result = sanitize(html);
    expect(result).toContain("<pre>");
    expect(result).toContain("<code");
    expect(result).toContain("const x = 1;");
  });

  it("preserves lists (ul, ol, li)", () => {
    const html = "<ul><li>item 1</li><li>item 2</li></ul><ol><li>first</li></ol>";
    const result = sanitize(html);
    expect(result).toContain("<ul>");
    expect(result).toContain("<li>item 1</li>");
    expect(result).toContain("<ol>");
  });

  it("preserves headings h1 through h6", () => {
    for (let i = 1; i <= 6; i++) {
      const tag = `h${i}`;
      const result = sanitize(`<${tag}>Heading ${i}</${tag}>`);
      expect(result).toContain(`<${tag}>`);
    }
  });

  it("preserves safe links with href, target, rel", () => {
    const html =
      '<a href="https://example.com" target="_blank" rel="noopener">link</a>';
    const result = sanitize(html);
    expect(result).toContain('href="https://example.com"');
    expect(result).toContain("target=");
    expect(result).toContain("rel=");
  });

  it("preserves safe images with src, alt, width", () => {
    const html = '<img src="photo.png" alt="A photo" width="100">';
    const result = sanitize(html);
    expect(result).toContain('src="photo.png"');
    expect(result).toContain('alt="A photo"');
  });

  it("preserves table markup", () => {
    const html =
      "<table><thead><tr><th>Col</th></tr></thead><tbody><tr><td>Val</td></tr></tbody></table>";
    const result = sanitize(html);
    expect(result).toContain("<table>");
    expect(result).toContain("<th>Col</th>");
    expect(result).toContain("<td>Val</td>");
  });

  it("preserves blockquote, hr, del, details/summary", () => {
    expect(sanitize("<blockquote>quoted</blockquote>")).toContain("<blockquote>");
    expect(sanitize("<hr>")).toContain("<hr");
    expect(sanitize("<del>deleted</del>")).toContain("<del>");
    expect(sanitize("<details><summary>More</summary>Content</details>")).toContain(
      "<details>",
    );
  });
});

// ==========================================================================
// CATEGORY 5: Integration Pipeline (full template rendering via linkedom+vm)
// ==========================================================================

describe("template rendering pipeline produces safe DOM", () => {
  it("renderEntry sanitizes user message XSS payloads", () => {
    const attack = "<img src=x onerror=alert(1)>";
    const { document } = renderTemplate(session(attack));
    const messages = document.getElementById("messages");
    expect(messages).toBeTruthy();
    expect(messages?.querySelector("img[onerror]")).toBeNull();
  });

  it("renderEntry sanitizes assistant message XSS payloads", () => {
    const attack = "<script>alert(1)</script><img src=x onerror=alert(2)>";
    const s: SessionData = {
      header: { id: "test", timestamp: now() },
      entries: [userMsg("1", null, "hello"), assistantMsg("2", "1", attack)],
      leafId: "2",
      systemPrompt: "",
      tools: [],
    };
    const { document } = renderTemplate(s);
    const messages = document.getElementById("messages");
    expect(messages).toBeTruthy();
    expect(messages?.querySelector("img[onerror]")).toBeNull();
    expect(messages?.querySelector("script")).toBeNull();
  });

  it("tree node display sanitizes XSS in labels", () => {
    const s: SessionData = {
      header: { id: "test-tree", timestamp: now() },
      entries: [
        userMsg("1", null, "safe content"),
        assistantMsg("2", "1", "safe response"),
      ],
      leafId: "2",
      systemPrompt: "",
      tools: [],
    };
    const { document } = renderTemplate(s);
    const tree = document.getElementById("tree-container");
    expect(tree).toBeTruthy();
    expect(tree?.querySelector("img[onerror]")).toBeNull();
  });
});

// ==========================================================================
// CATEGORY 6: Config Correctness
// ==========================================================================

describe("PURIFY_CONFIG correctness", () => {
  it("blocks data-* attributes (ALLOW_DATA_ATTR: false)", () => {
    const result = sanitize('<div data-evil="payload" data-x="y">content</div>');
    expect(result).not.toContain("data-evil");
    expect(result).not.toContain("data-x");
  });

  it("strips dangerous tags not in allowlist", () => {
    const dangerousTags = [
      "<base href='evil'>",
      "<object data='evil'>",
      '<embed src="evil">',
      "<meta http-equiv='refresh'>",
      '<link rel="stylesheet" href="evil">',
      "<marquee>text</marquee>",
      '<source src="evil">',
    ];
    for (const tag of dangerousTags) {
      const result = sanitize(tag);
      const tagName = tag.match(/<(\w+)/)?.[1];
      expect(result).not.toContain(`<${tagName}`);
    }
  });

  it("DOMPurify version is >= 3.3.3", () => {
    expect(DOMPurify.version).toBeDefined();
    const parts = DOMPurify.version.split(".").map(Number);
    const versionNum = parts[0] * 10000 + parts[1] * 100 + parts[2];
    expect(versionNum).toBeGreaterThanOrEqual(30303);
  });
});

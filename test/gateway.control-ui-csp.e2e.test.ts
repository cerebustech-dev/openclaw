import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { parseCsp } from "./helpers/csp.js";
import {
  type GatewayInstance,
  spawnGatewayInstance,
  stopGatewayInstance,
} from "./helpers/gateway-e2e-harness.js";

const E2E_TIMEOUT_MS = 240_000;

describe("Control UI CSP — served header", () => {
  let gateway: GatewayInstance | null = null;

  beforeAll(async () => {
    gateway = await spawnGatewayInstance("control-ui-csp", { controlUi: true });
  }, E2E_TIMEOUT_MS);

  afterAll(async () => {
    if (gateway) {
      await stopGatewayInstance(gateway);
      gateway = null;
    }
  });

  it(
    "serves Content-Security-Policy with worker-src and connect-src locked to 'self'",
    { timeout: E2E_TIMEOUT_MS },
    async () => {
      if (!gateway) {
        throw new Error("gateway not initialized");
      }
      const res = await fetch(`http://127.0.0.1:${gateway.port}/`, {
        method: "GET",
        redirect: "manual",
      });

      const csp = res.headers.get("content-security-policy");
      expect(csp, "Control UI response must include a CSP header").not.toBeNull();
      const directives = parseCsp(csp ?? "");

      const workerSrc = directives.filter((d) => d.directive === "worker-src");
      expect(workerSrc, "exactly one worker-src directive").toHaveLength(1);
      expect(workerSrc[0]?.values).toEqual(["'self'"]);

      const connectSrc = directives.filter((d) => d.directive === "connect-src");
      expect(connectSrc, "exactly one connect-src directive").toHaveLength(1);
      expect(connectSrc[0]?.values).toEqual(["'self'"]);

      // Forbid scheme-wildcard relaxations regardless of where they appear.
      for (const banned of ["ws:", "wss:", "blob:", "https:", "*"]) {
        expect(
          connectSrc[0]?.values.includes(banned),
          `connect-src must not include ${banned}`,
        ).toBe(false);
        expect(
          workerSrc[0]?.values.includes(banned),
          `worker-src must not include ${banned}`,
        ).toBe(false);
      }

      const frameAncestors = directives.filter((d) => d.directive === "frame-ancestors");
      expect(frameAncestors[0]?.values).toEqual(["'none'"]);
      const objectSrc = directives.filter((d) => d.directive === "object-src");
      expect(objectSrc[0]?.values).toEqual(["'none'"]);
    },
  );
});

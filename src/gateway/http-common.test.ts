import { describe, expect, it } from "vitest";
import { setDefaultSecurityHeaders } from "./http-common.js";
import { makeMockHttpResponse } from "./test-http-response.js";

describe("setDefaultSecurityHeaders", () => {
  it("sets X-Frame-Options to DENY", () => {
    const { res, setHeader } = makeMockHttpResponse();
    setDefaultSecurityHeaders(res);
    expect(setHeader).toHaveBeenCalledWith("X-Frame-Options", "DENY");
  });

  it("does NOT set X-Frame-Options to SAMEORIGIN", () => {
    const { res, setHeader } = makeMockHttpResponse();
    setDefaultSecurityHeaders(res);
    const frameOptionsCall = setHeader.mock.calls.find(
      ([name]: string[]) => name === "X-Frame-Options",
    );
    expect(frameOptionsCall).toBeDefined();
    expect(frameOptionsCall![1]).not.toBe("SAMEORIGIN");
  });

  it("sets X-Content-Type-Options to nosniff", () => {
    const { res, setHeader } = makeMockHttpResponse();
    setDefaultSecurityHeaders(res);
    expect(setHeader).toHaveBeenCalledWith("X-Content-Type-Options", "nosniff");
  });

  it("sets Referrer-Policy", () => {
    const { res, setHeader } = makeMockHttpResponse();
    setDefaultSecurityHeaders(res);
    expect(setHeader).toHaveBeenCalledWith("Referrer-Policy", "no-referrer");
  });

  it("sets Permissions-Policy", () => {
    const { res, setHeader } = makeMockHttpResponse();
    setDefaultSecurityHeaders(res);
    expect(setHeader).toHaveBeenCalledWith(
      "Permissions-Policy",
      "camera=(), microphone=(), geolocation=()",
    );
  });

  it("sets Strict-Transport-Security when provided", () => {
    const { res, setHeader } = makeMockHttpResponse();
    setDefaultSecurityHeaders(res, {
      strictTransportSecurity: "max-age=63072000; includeSubDomains; preload",
    });
    expect(setHeader).toHaveBeenCalledWith(
      "Strict-Transport-Security",
      "max-age=63072000; includeSubDomains; preload",
    );
  });

  it("does not set Strict-Transport-Security when not provided", () => {
    const { res, setHeader } = makeMockHttpResponse();
    setDefaultSecurityHeaders(res);
    expect(setHeader).not.toHaveBeenCalledWith("Strict-Transport-Security", expect.anything());
  });

  it("does not set Strict-Transport-Security for empty string", () => {
    const { res, setHeader } = makeMockHttpResponse();
    setDefaultSecurityHeaders(res, { strictTransportSecurity: "" });
    expect(setHeader).not.toHaveBeenCalledWith("Strict-Transport-Security", expect.anything());
  });});

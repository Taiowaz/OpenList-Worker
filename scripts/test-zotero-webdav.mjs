// Run from the repository root: node --test scripts/test-zotero-webdav.mjs
// No live OneDrive account is required. Authentication and storage selection are
// fixtures; the OneDrive request code, WebDAV router and HTTP responses are real.
import assert from "node:assert/strict"
import { test } from "node:test"
import { build } from "esbuild"
import path from "node:path"
import { fileURLToPath } from "node:url"

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..")
const result = await build({
  stdin: {
    contents: `
      export { webdavRouter } from "./src/backend/server/webdav.ts";
      export { requestApi as regularApi } from "./src/backend/drivers/onedrive/util.ts";
      export { requestApi as appApi } from "./src/backend/drivers/onedrive_app/util.ts";
      export { setRead, setWrite } from "test-storage";
    `,
    resolveDir: root,
  },
  bundle: true,
  platform: "node",
  format: "esm",
  write: false,
  plugins: [{
    name: "isolated-storage-fixtures",
    setup(builder) {
      builder.onResolve({ filter: /^(test-storage|\.\.\/internal\/op\/storage)$/ },
        () => ({ path: "storage", namespace: "fixture" }))
      builder.onResolve({ filter: /^\.\/auth$/ }, (args) =>
        args.importer.endsWith("/server/webdav.ts")
          ? { path: "auth", namespace: "fixture" } : undefined)
      builder.onLoad({ filter: /.*/, namespace: "fixture" }, (args) => ({
        contents: args.path === "auth" ? `
          export async function authUserFromReq() {
            return { user: { role: 2, permission: 0 } };
          }
          export async function getOrInitUsers() { return { users: [] }; }
          export async function verifyUserPassword() { return false; }
        ` : `
          let read, write;
          export const setRead = fn => { read = fn; };
          export const setWrite = fn => { write = fn; };
          export const getItem = (...args) => read(...args);
          export const putItem = (...args) => write(...args);
          export const listItems = async () => ({ content: [] });
          export const makeDirectory = async () => {};
          export const removeItems = async () => {};
          export const moveItems = async () => {};
          export const copyItems = async () => {};
        `,
        loader: "js",
      }))
    },
  }],
})
const { webdavRouter, regularApi, appApi, setRead, setWrite } =
  await import(`data:text/javascript;base64,${Buffer.from(result.outputFiles[0].text).toString("base64")}`)

const dav = "/dav/onedrive/app/zotero/nonexistent.prop"
const graph = "https://graph.microsoft.com/v1.0/me/drive/root:/zotero/nonexistent.prop:"
const headers = { Authorization: "Bearer fixture" }

async function withGraph(response, fn) {
  const original = globalThis.fetch
  globalThis.fetch = async (url) => {
    assert.equal(url, graph, "the test must never contact a live server")
    return response.clone()
  }
  try { return await fn() } finally { globalThis.fetch = original }
}

for (const [driver, api] of [["OneDrive", regularApi], ["OneDrive APP", appApi]]) {
  for (const [label, status, body, expected] of [
    ["Zotero missing-file probe", 404, { error: { code: "itemNotFound", message: "The resource could not be found." } }, 404],
    ["localized missing-file response", 404, { error: { code: "itemNotFound", message: "找不到请求的资源。" } }, 404],
    ["non-JSON missing-file response", 404, "Missing resource", 404],
    ["actual internal failure with identical wording", 500, { error: { code: "generalException", message: "The resource could not be found." } }, 500],
    ["forbidden is not missing", 403, { error: { code: "accessDenied", message: "Access denied" } }, 500],
    ["rate limit is not missing", 429, { error: { code: "tooManyRequests", message: "Too many requests" } }, 500],
    ["authentication failure is not missing", 401, { error: { code: "InvalidAuthenticationToken", message: "Invalid token" } }, 500],
  ]) {
    test(`${driver}: ${label}`, async () => {
      const upstream = new Response(typeof body === "string" ? body : JSON.stringify(body), { status })
      setRead(async (requestPath) => {
        assert.equal(requestPath, "/onedrive/app/zotero/nonexistent.prop")
        await api({ accessToken: "fixture" }, graph, "GET", undefined, true)
        throw new Error("expected upstream failure")
      })
      await withGraph(upstream, async () => {
        const response = await webdavRouter.request(dav, { headers })
        assert.equal(response.status, expected)
        if (expected === 404) assert.equal(await response.text(), "Not Found")
      })
    })
  }

  test(`${driver}: existing file keeps its download redirect`, async () => {
    setRead(async () => {
      const item = await api({ accessToken: "fixture" }, graph, "GET")
      return { item: { name: item.name, is_dir: false }, rawUrl: "/api/d/existing.prop" }
    })
    await withGraph(new Response(JSON.stringify({ name: "existing.prop" }), { status: 200 }), async () => {
      const response = await webdavRouter.request(dav, { headers })
      assert.equal(response.status, 302)
      assert.equal(response.headers.get("Location"), "/api/d/existing.prop")
    })
  })
}

test("Zotero verification still advertises WebDAV", async () => {
  const response = await webdavRouter.request("/dav/onedrive/app/zotero/", { method: "OPTIONS", headers })
  assert.equal(response.status, 200)
  assert.equal(response.headers.get("DAV"), "1, 2")
})

test("Zotero test-file upload still returns 201 with the original content", async () => {
  let written
  setWrite(async (requestPath, data) => { written = { requestPath, data } })
  const response = await webdavRouter.request("/dav/onedrive/app/zotero/zotero-test-file.prop", {
    method: "PUT", headers, body: " ",
  })
  assert.equal(response.status, 201)
  assert.equal(written.requestPath, "/onedrive/app/zotero/zotero-test-file.prop")
  assert.equal(written.data.toString(), " ")
})

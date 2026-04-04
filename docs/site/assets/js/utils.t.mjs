import assert from "node:assert";
import "./utils.mjs";
import {
  isLinkToMarkdownWithoutSuffix,
  isSiteRelativeLink,
  isNonSiteRelativeLink,
} from "./utils.mjs";

describe("util functions", function () {
  describe("#isSiteRelativeLink", function () {
    for (const [p, expect] of Object.entries({
      index: false,
      "translation/ddl": false,
      "translation/ddl.md": false,
      "./translation/ddl": false,
      "assets/css/page.css": false,
      "assets/css/page": false,
      "./translation/example-hidden.png": false,
      "translation/example-hidden.png": false,
      "/translation/example-hidden.png": true,
      "https://example.com": false,
    })) {
      it(`Should return ${expect} for ${p}`, () => {
        assert.equal(isSiteRelativeLink(p), expect);
      });
    }
  });
  describe("#isLinkToMarkdownWithoutSuffix", function () {
    for (const [p, expect] of Object.entries({
      index: true,
      "translation/ddl": true,
      "translation/ddl.md": false,
      "./translation/ddl": true,
      "assets/css/page.css": false,
      "assets/css/page": false,
      "./translation/example-hidden.png": false,
      "translation/example-hidden.png": false,
      "/translation/example-hidden.png": false,
    })) {
      it(`Should return ${expect} for ${p}`, async () => {
        assert.equal(await isLinkToMarkdownWithoutSuffix(p), expect);
      });
    }
  });
  describe("#isNonSiteRelativeLink", function () {
    for (const [p, expect] of Object.entries({
      index: true,
      "translation/ddl": true,
      "translation/ddl.md": true,
      "./translation/ddl": true,
      "assets/css/page.css": true,
      "assets/css/page": true,
      "./translation/example-hidden.png": true,
      "translation/example-hidden.png": true,
      "/translation/example-hidden.png": false,
      "https://example.com": false,
    })) {
      it(`Should return ${expect} for ${p}`, async () => {
        assert.equal(await isNonSiteRelativeLink(p), expect);
      });
    }
  });
});

// link fixer-upper

import * as fs from "node:fs/promises";
import * as path from "node:path";
import {
  traverse,
  mkdirNoisily,
  dstIsNewer,
  isLinkToMarkdownWithoutSuffix,
  isNonSiteRelativeLink,
  isExternalLink,
  isMarkdownLink,
  SKIP_THESE,
  isPageLink,
  isRelativeLink,
} from "./utils.mjs";
import { marked } from "marked";
// import { link } from "node:fs";

/** rerender a token in markdown, flattening its tokens */
function renderAndFlatten(t)
{
  if (t.tokens) {
    t.text = flattenTokens(t.tokens);
    t.tokens = [];
  }
  if (t.type === "link") {
    return `[${t.text}](${t.href})`;
  } else if (t.type === "image") {
    return `![${t.text}](${t.href})`;
  } else if (t.type === "list_item") {
    if (!t.item_prefix) {
        throw Error(`Need item_prefix for ${t.raw}`);
    }
    return t.item_prefix + t.text + t.item_suffix;
  } else if (t.type === 'list') {
    return flattenTokens(t.items);
  } else {
    throw Error(`Can't rerender type ${t.type}`);
  }
}

/**
 * Recursively fix an array of tokens
 * @param {Object[]} srcTokens array of source tokens
 * @returns true if any fix was needed
 */
async function fixTokens(srcTokens, o) {
  const { srcPath } = o;
  const parentPath = path.dirname(srcPath); // foo/bar/baz.md => foo/bar
  // were any children fixed?
  let didFix = false;
  for (let i in srcTokens) {
    const t = srcTokens[i];
    const { type, raw, text, href, tokens, items } = t;

    // store these off so that renderAndFlatten knows what the prefix/suffix were
      if (t.type === 'list_item') {
          const prefixLen = t.raw.indexOf(t.text);
          if (prefixLen === -1) {
            // throw Error(`Could not find prefix for ${t.raw}`);
            t.item_prefix = undefined;
            t.item_suffix = undefined;
          } else {
            t.item_prefix = t.raw.substring(0, prefixLen);
            if (t.raw.endsWith('\n')) {
                t.item_suffix = '\n';
            } else {
                t.item_suffix = '';
            }
        }
    }

    if(t.raw.indexOf('Understand the basics') !== -1) {
        console.dir(t);
    }

    // if (raw && raw.indexOf('Survey Tool Guide') !== -1) {
    //     console.dir(t);
    // }
        if (href === 'translation/getting-started/guide.md') {
            console.dir({t});
        }

    if (type === "link" || type === "image") {
      // TODO: lint here CLDR-18011
      // if (isSiteRelativeLink(href)) {
      //     throw Error(`Err! Site Relative Link ${href}, use relative link to .md file instead`);
      // }
      // TODO: lint here CLDR-18011
      // if (isRelativeLink(href) && !isMarkdownLink(href)) {
      //     throw Error(`Err! Relative Link ${href}, use relative link to ${href}.md instead`);
      // }
      if (!isExternalLink(href) && isMarkdownLink(href)) {
        //DUMP //console.dir({type, raw, href});

        t.href = href.slice(0, href.length - 3); // remove .md

        // confirm the sub-tokens situation and rerender
        didFix = true; // so that the parent re-renders
        t.raw = renderAndFlatten(t);
      } else if (
        isNonSiteRelativeLink(href) &&
        !isPageLink(href) &&
        !(await isLinkToMarkdownWithoutSuffix(href))
      ) {
        // links to png etc
        const oldHref = t.href;
        t.href = path.join("/", parentPath, href); // convert to site relative.
        // confirm the sub-tokens situation and rerender
        didFix = true; // so that the parent re-renders
        t.raw = renderAndFlatten(t);
      } else {
        // no fixup needed for this link.
        // we aren't processing subtokens here, however, they should not be needed,
        // since t.raw remains valid
      }
    } else if (type === 'list') {
        if(await fixTokens(items, o)) {
            didFix = true;
            t.raw = renderAndFlatten(t);
        }
    } else if (await fixTokens(tokens, o)) {
      didFix = true; // so parents get re-rendered
      if (type === 'list_item') {
        t.raw = renderAndFlatten(t);
      } else if (type === "paragraph" || type === "em" || type === "strong" || type === 'text') {
        delete t.raw; // recombine this paragraph
      } else {
        throw Error(
          `Can't fixup parent type ${type}: in ${srcPath}::${
            t.raw
          } : ${JSON.stringify(t)}`
        );
      }
    }
  }
  return didFix; // propagate up
}

/**
 * Convert array of tokens to string
 * The convention is that if 'raw' is unset, then 'tokens' is used.
 * @param {object[]} tokens token array
 * @returns markdown source
 */
function flattenTokens(srcTokens) {
  const out = [];
  for (const { type, raw, tokens } of srcTokens) {
    if (tokens && !raw) {
      // if the raw was deleted…
      if (type === "em") out.push("_");
      if (type === "strong") out.push("**");
      out.push(flattenTokens(tokens));
      if (type === "em") out.push("_");
      if (type === "strong") out.push("**");
    } else {
      out.push(raw);
    }
  }
  return out.join("");
}

/**
 * Fix links in one .md file
 * @param {string} srcPath
 * @param {string} dstPath
 */
async function linkFix(srcPath, dstPath) {
  const str = (await fs.readFile(srcPath, "utf-8")).replaceAll(/\r\n/g, "\n");

    const tokens = marked.lexer(str);

    if (srcPath === 'translation.md') {
        console.dir(tokens, {depth: Infinity});
    }

    const didFix = await fixTokens(tokens, {
    srcPath,
  });

  // DUMP console.dir({ tokens }, { depth: Infinity });

  let outStr = flattenTokens(tokens);
  // TODO: READ the prev file, don't rewrite if unchanged. CLDR-18011
  let existingText;
  try {
    existingText = await fs.readFile(dstPath, "utf-8");
  } catch (e) {
    existingText = null;
  }

  // Wasn't broken, don't fix: Don't rewrite files if fixTokens returned false.
  if (!didFix) outStr = str;

  if (outStr === existingText) {
    // File same content, no rewrite needed
  } else if (str === outStr) {
    // console.log(`# ${dstPath} [no change needed] ${didFix}`)
  } else {
    console.log(`# ${dstPath} [updated] ${didFix}`);
  }

  await fs.writeFile(dstPath, outStr, "utf-8");
}

/**
 * Sync from srcDir to dstDir, and fixup links
 * @param {string} srcDir
 * @param {string} dstDir
 */
export async function syncAndFixLinks(srcDir, dstDir) {
  const out = {};
  await traverse(
    srcDir,
    out,
    async (dirPath, srcPath, out, e) => {
      await mkdirNoisily(path.join(dstDir, dirPath)); // mkdir each time
      const dstPath = path.join(dstDir, srcPath);
      if (await dstIsNewer(srcPath, dstPath)) {
        return; // skip if unchanged
      }
      if (!SKIP_THESE.test(srcPath) && e.name.endsWith(".md")) {
        await linkFix(srcPath, dstPath);
      } else {
        fs.copyFile(srcPath, dstPath);
      }
    },
    /^(.jekyll-cache)/
  );
}

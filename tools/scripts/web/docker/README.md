# Docker for local builds of the CLDR Site

See <https://cldr.unicode.org/development/updating-site> for more details about updating the site.

## Using Docker to build and preview the site

This is the recommended mechanism

1. install https://docker.io
2. `cd docs/site`
3. `docker compose up`
4. visit <http://127.0.0.1:4000>
5. hit control-C to cancel the docker run.
6. (on Windows, you may need to restart the container to pickup changes)

## Manually running the site build

### Installing assets

This is a prerequisite, and `build` needs to re-run if the sitemap changes.

1. run `npm i` and `npm run build` in `docs/site`
   (You can use `npm run watch` if you want to automatically rebuild if any files change)

### Building the static site

1. In `docs/site` run `jekyll build` (will need prereqs. see build-site.sh)
2. output is in `../../_site` here in this dir.

## Production Build

Cloudflare runs `sh tools/scripts/web/build-site.sh` from the repo root.  The `wrangler.jsonc` file at the root controls the deploy.

## Link Check

You can locally run lychee to link check \- [https://github.com/lycheeverse](https://github.com/lycheeverse).

   1. `lychee --cache http://127.0.0.1:4000/`

   2. 1-liner for link checking the entire site ( from docs/site dir ):
      `for p in $(jq  -r '.usermap | keys | flatten[]' < assets/json/tree.json); do echo; echo $p; lychee --cache http://127.0.0.1:4000/${p}; done`

FROM node:lts
RUN npm i -g wrangler
VOLUME /_site
EXPOSE 4000
CMD [ "wrangler", "pages", "dev",  "/_site", "--port", "4000", "--ip", "0.0.0.0" , "--compatibility-date=2026-04-04"]

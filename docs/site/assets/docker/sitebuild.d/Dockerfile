FROM node:lts
WORKDIR /app
ADD package.json /app/package.json
ADD package-lock.json /app/package-lock.json
RUN ls -l /app
RUN npm ci
ADD assets/ /app/assets/
VOLUME /site
WORKDIR /site
RUN pwd
CMD [ "node", "/app/assets/js/build.mjs", "--watch" ]


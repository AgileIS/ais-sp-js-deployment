# Deployment / Provisioning mit nodejs

# parameters
-f : <configfile>

## dev
npm install
typings install
tsc || tsc -w (mit watch)

## dev-run
node dist/deploy -f config/<config>.json

## build
gulp && npm pack || npm publish

## install
npm install ais-sp-js-deployment [--save]

## run
node deploy -f config/<config>.json

## run with child process debug
node deploy -f config/<config>.json -d

# Publish
npm publish
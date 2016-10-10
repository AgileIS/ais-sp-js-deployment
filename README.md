# Deployment / Provisioning mit nodejs

#parameters
-f : <configfile>

##dev
npm install
typings install
tsc || tsc -w (mit watch)

##dev-run
node dist/deploy -f config/<config>.json

##build
gulp && npm pack || npm publish

##install
npm install ais-sp-js-deployment [--save]

##run
node deploy -f config/<config>.json

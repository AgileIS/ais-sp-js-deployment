# SharePoint client side deployment and content provisioning

The package provides a powerful client side deployment engine for SharePoint 2013 / 2016 and SharePoint online (office 365). All structural elements can be defined within versioned configuration files. Those will be combined with the environmental information’s of each stage. A deployment run will then create/update only the difference between the defined structure and the actual instantiated artefact of the target SharePoint site. Hence the script is working in an application life cycle mode.

## Install package

### Prerequisites

* You definitely need nodejs on your system. :)
* Create a new provisioning project
  * mkdir \<your project dir name\>
  * npm init 

### Install npm package

Add the npm package ais-sp-js-deployment to your new deployment project
```
$ npm install ais-sp-js-deployment --save 
```

## Usage

hier fehlen noch die Demofiles und ich habe noch einen deployment Fehler



parameters
-f : <configfile>

## dev
npm install
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
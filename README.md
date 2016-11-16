# SharePoint client side deployment and content provisioning engine

The package provides a powerful client side deployment engine for SharePoint 2013 / 2016 and SharePoint online
(office 365). All structural elements can be defined within versioned configuration files. Those will be combined with the
environmental information’s of each stage. A deployment run will then create/update only the difference between the defined
structure and the actual instantiated artefact of the target SharePoint site. Hence the script is working in an application
life cycle mode.

## Install package

### Prerequisites

* You definitely need nodejs on your system. :)
* Create a new provisioning project
  * mkdir \<your project dir name\>
  * npm init

### Install npm package

Add the npm package ais-sp-js-deployment to your new deployment project

```cmd
npm install ais-sp-js-deployment --save
```

## Usage

To use the client side provisioning engine, you need to prepare configuration files containing your instance description
and in addition at least one environment related configuration. Those files will be compiled to one configuration file per
staging environment.

### Prepare configuration files

The ais-sp-js-deployment package will provide some example configuration files during installation. Those can be found in
the `./config`. The naming convention of the configuration files is as follows:

* `partial_*.json` >> all environmental independent elements like content types, list columns, etc. All files starting
  with `partial_` will be combined in the order of their name. For example partial_v1.json, partial_v1_1.json, partial_v2.json
* `stage_*.json` >> Configuration files starting with `stage_` can be used to define stage dependent parts of the configuration
  like site url, deployment user, etc. 
* `config_<stagename>.json` Will contain the combined configuration after you have run the config build task `npm run sp:buildconfig`

### Compile configuration files

```cmd
npm run sp:buildconfig
```

### Deploy your solution artefacts

```cmd
npm run sp:deploy
```

Be aware that the script above is configured to use a sample configuration file. please change settings in your .\package.json. 
In addition you can run deployment with different file by using the following.

```cmd
node .\deploy\deploy.js -f .\config\<filename>.json
```

## Development

### Clone and Install

```
git clone https://github.com/agileis/  
npm install  
tsc || tsc -w (watch)
```

### dev-run
```
node dist/deploy.js -f config/<config>.json
```

### Build
```
gulp && npm pack || npm publish
```

### run with child process debug
```
node deploy -f config/<config>.json -d
```

## Feel free to contribute, report bugs and share your thoughts. 

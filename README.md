# SharePoint client side deployment and content provisioning engine

[![Build Status](https://travis-ci.org/AgileIS/ais-sp-js-deployment.svg?branch=master)](https://travis-ci.org/AgileIS/ais-sp-js-deployment)
[![npm version](https://badge.fury.io/js/ais-sp-js-deployment.svg)](https://badge.fury.io/js/ais-sp-js-deployment)

The package provides a powerful client side deployment engine for SharePoint 2013 / 2016 and SharePoint online
(office 365). All structural elements can be defined within versioned configuration files. Those will be combined with the
environmental information of each stage. A deployment run will then create/update only the difference between the defined
structure and the actual instantiated artefact of the target SharePoint site. Hence the script is working in an application
life cycle mode.

## Install package

### Prerequisites

* You definitely need [Node.js](https://nodejs.org/en/) on your system :) . We did our tests with Node.js 6.x
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
the `./config` folder. The naming convention of the configuration files is as follows:

* `partial_*.json` >> all environmental independent elements like content types, list columns, etc. All files starting
  with `partial_` will be combined in the order of their name. For example partial_v1.json, partial_v1_1.json, partial_v2.json
* `stage_*.json` >> Configuration files starting with `stage_` can be used to define stage dependent parts of the configuration
  like site url, deployment user, etc.
* `config_<stagename>.json` >> Will contain the combined configuration after you have run the config build task `npm run sp:buildconfig`

### Compile configuration files

```cmd
npm run sp:buildconfig
```

### Deploy your solution artefacts

```cmd
npm run sp:deploy
```

Be aware that the script above is configured to use a sample configuration file. You may need to change settings of your .\package.json.
In addition, you can run deployment with a different file by using the following.

```cmd
node .\deploy\deploy.js -f .\config\<filename>.json
```

**Note:** You can run the deployment script multiple times. Any existing items won’t be changed unless you have defined `ControlOption="Delete||Update"`

---

## Development

### Clone and Install

```cmd
git clone https://github.com/AgileIS/ais-sp-js-deployment.git
npm install
tsc || tsc -w (watch)
```

### dev-run

```cmd
node dist/deploy.js -f config/<config>.json
```

### Build

```cmd
gulp && npm pack || npm publish
```

### run with child process debug

```cmd
node deploy -f config/<config>.json -d
```

## Applies to

* SharePoint 2013 - windows authentication (ntlm, basic)
* SharePoint 2016 - windows authentication (ntlm, basic)

### comming but not implemented yet

* SharePoint Online (ACS)

## See also

* We tested a lot of [Vesa Juvonen's](https://github.com/VesaJuvonen) work regarding SharePoint provisioning.
* We also tested [OfficeDev/PnP-PowerShell](https://github.com/OfficeDev/PnP-powershell)
* Finally we started our provisioning engine project based on the [SharePoint/PnP-JS-Core](https://github.com/SharePoint/PnP-JS-Core)
  package. During development we noticed that the underlying SharePoint RESTful API wont fit to all our requirements. Hence
  we fell back to integrating JSOM and faking some environmental variables.

> Thanks to all projects and developers influencing our work.

## Next

* Some more authentication providers (ACS, AD FS), hence SharePoint Online support
* Subsite support
* Config generator
* …

---

Feel free to contribute, report bugs and share your thoughts.

---

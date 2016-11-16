const fs = require('fs');
const fse = require('fs.extra');
var cmd = require('node-cmd');

const destPath = '../../deploy';
const gulpSrc = 'gulpfile.merge.js';
const gulpDest = '../../deploy/gulpfile.js';
const deploySrc = './dist/deploy.js';
const deployDest = '../../deploy/deploy.js';
const demoSrc = './demofiles';
const demoDest = '../../config/';
const packageName = 'ais-sp-js-deployment';
const confDestReg = /configDest\s=\s'.*'/;
let confDest = 'configDest = \'../config/\'';
const confPrefixReg = /partialConfigPrefix\s=\s'.*'/;
let confDemoPrefix = 'partialConfigPrefix = \'democonfig_*.json\'';
const deployScript = 'cd deploy && gulp && cd .. && node ./deploy/deploy -f config/config_demo.json';

function processGulpfile() {
    fs.exists(gulpDest, exists => {
        if (exists) {
            let file = fs.readFileSync(gulpDest, 'utf8');
            let fileConfDest = file.match(confDestReg);
            if (fileConfDest) confDest = fileConfDest[0];
            let fileConfDemoPrefix = file.match(confPrefixReg);
            if (fileConfDemoPrefix) confDemoPrefix = fileConfDemoPrefix[0];
        }
        fs.rename(gulpSrc, gulpDest, error => {
            console.log("- copy gulpfile for config merge");
            if (error) {
                console.error(error);
            } else {
                fs.readFile(gulpDest, 'utf8', (error, data) => {
                    console.log('- gulpfile replacements');
                    if (error) {
                        console.error(error);
                    } else {
                        let result = data.replace(confPrefixReg, confDemoPrefix);
                        result = result.replace(confDestReg, confDest);
                        fs.writeFile(gulpDest, result, 'utf8', error => {
                            if (error) {
                                return console.log(error);
                            };
                        });
                    }
                });
            }
        });
    });
}

function processDeployJs() {
    fs.rename(deploySrc, deployDest, error => {
        console.log('- copy deploy.js');
        if (error) {
            console.error(error);
        } else {
            fs.readFile(deployDest, 'utf8', (error, data) => {
                console.log('- fix deploy.js requires');
                if (error) {
                    console.error(error);
                } else {
                    let result = data.replace('./index', packageName);
                    result = result.replace('//# sourceMappingURL=deploy.js.map', '');
                    fs.writeFile(deployDest, result, 'utf8', error => {
                        if (error) {
                            return console.error(error);
                        };
                    });
                }
            });
        }
    });
}

function createPackageScript() {
    fs.readFile('../../package.json', 'utf8', (error, data) => {
        console.log('- add "deploy" script to package.json');
        if (error) {
            console.error(error);
        } else {
            let packageObj = JSON.parse(data);
            packageObj.scripts['sp:deploy'] = deployScript;
            packageObj.scripts['sp:buildconfig'] = deployScript;
            fs.writeFile('../../package.json', JSON.stringify(packageObj, null, 2), 'utf8', error => {
                if (error) {
                    return console.error(error);
                };
            });
        }
    });
}

if(__dirname.includes("node_modules")) {
    console.log('Initialize ais-sp-js-deployment package: ');
    if(!fs.existsSync(destPath)) {
        fs.mkdirSync(destPath);
    }

    processGulpfile();
    processDeployJs();

    if(!fs.existsSync(demoDest)) {
        console.log('- copy demo configs');
        fs.mkdirSync(demoDest);
        fse.copyRecursive(demoSrc, demoDest, error => { 
            if (error) console.error(error);
        });
    }

    createPackageScript();
} else {
    cmd.run('typings install');
}

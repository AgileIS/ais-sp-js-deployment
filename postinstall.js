var fs = require('fs');


var gulpSrc = 'gulpfile.js.template';
var gulpDest = '../../gulpfile.js';
var deploySrc = './dist/deploy.js';
var deployDest = '../../deploy.js';


console.log("Initialize ais-sp-js-deployment package: ");

fs.rename(gulpSrc, gulpDest, (error) => {
    console.log("- copied gulpfile for config merge");
    if (error) console.error(error);
});

fs.rename(deploySrc, deployDest, (error) => {
    console.log("- copied deploy.js");
    if (error) console.error(error);

    fs.readFile(deployDest, 'utf8', (error, data) => {
        console.log("- fix deploy.js requires");
        if (error) {
            console.error(error);
        } else {
            var result = data.replace('./index', 'ais-sp-js-deployment');
            var result = result.replace('//# sourceMappingURL=deploy.js.map', '');
            fs.writeFile(deployDest, result, 'utf8', (error) => {
                if (error) {
                    return console.log(error);
                };
            });
        }
    });
});
const gulp = require('gulp');
const jsoncombine = require('gulp-jsoncombine');
const foreach = require('gulp-foreach');
const mergeStreams = require('merge-stream');
const jsonFormat = require('gulp-json-format');
const fs = require('fs');
const rename = require('gulp-rename');
const concat = require('gulp-concat');
const replace = require('gulp-replace');


const userConfigPrefix = 'config/userconfig_*.json';
const partialConfigPrefix = 'config/insiderverzeichnis_*.json';
//const partialConfigPrefix = 'config/partialconfig_*.json';
const configPrefix = 'config_';
/*
 * value: string => merge by field
 * value: string Array => merge by each field, op: operator ["&&", "||"
 * value: tuple Array => cross merge (item[tuple[0]] === searchItem[tuple[1]] + vise versa), op: operator ["&&", "||"]
 */
let _parent = "";
const assignments = {
  Lists: {
    value: 'InternalName'
  },
  Fields: {
    value: 'InternalName',
  },
  Views: {
    value: 'InternalName',
  },
  Sites: {
    value: 'Identifier'
  },
  Features: {
    value: 'Id'
  },
  ContentTypes: {
    value: 'Id'
  },
  QuickLaunch: {
    value: 'Title'
  },
  Files: {
    value: 'Name'
  },
  Items: {
    value: null
  },
  Solutions: {
    value: 'Title'
  }
};

function assignmentsFilter(prop, item, searchItem) {
  let found = false;
  if (assignments[prop] && !assignments[prop].value) {
    found = true;
  } else if (assignments[prop] && assignments[prop].value instanceof Array) {
    let condition = "";
    for (let assignment of assignments[prop].value) {
      condition += (condition) ? ` ${assignments[prop].op} ` : "";
      if (assignment instanceof Array) {
        condition += (item[assignment[0]] === searchItem[assignment[1]]);
      } else {
        condition += (item[assignment] === searchItem[assignment]);
      }
    }
    found =  eval(condition);
  } else if (assignments[prop]) {
    found = searchItem[assignments[prop].value] === item[assignments[prop].value];
  }
  return found;
}

function xMerge(target, source, parent, prop) {
  let isXMerge = false
  if (assignments[parent] && assignments[parent].value instanceof Array && assignments[parent].value[0] instanceof Array) {
    if (prop === assignments[parent].value[0][0]) {
      target[prop] = source[assignments[parent].value[0][1]];
      isXMerge =  true;
    } else if (prop === assignments[parent].value[0][1]) {
      target[prop] = source[assignments[parent].value[0][0]];
      isXMerge =  true;
    }
  }
  return isXMerge;
}

function merge(target, source) {
  for (let prop in source) {
    _parent = (assignments[prop]) ? prop : _parent;
    if (source[prop].constructor === Object) {
      target[prop] = merge(target[prop] ? target[prop] : {}, source[prop]);
    } else if (source[prop].constructor === Array) {
      if (!target.hasOwnProperty(prop)) target[prop] = [];
      for (let item of source[prop]) {
        let result = target[prop].filter(assignmentsFilter.bind(this, prop, item));
        if (result.length === 1) {
          result[0] = merge(result[0], item);
          break;
        } else {
          target[prop].push(item);
        }
      };
    } else {
      target[prop] = source[prop];
    }
  }
  return target;
}

function combine(userConfigName) {

  return gulp.src([userConfigName, partialConfigPrefix])
    .pipe(jsoncombine(configPrefix + userConfigName.split('_')[1], (configs) => {
      let configNames = Object.getOwnPropertyNames(configs);
      let resultConfig = {};

      configNames.forEach((configName) => {
        merge(resultConfig, configs[configName]);
      }, this);

      return new Buffer(JSON.stringify(resultConfig), 'utf8');
    }))
    .pipe(jsonFormat(4))
    .pipe(gulp.dest('./config'));
}

gulp.task('combineAll', () => {
  let streams = [];

  streams.push(gulp.src(userConfigPrefix, {
      read: false
    })
    .pipe(foreach((stream, file) => {
      streams.push(combine(file.path.replace(file.cwd + '\\', '')));
      return stream;
    })));

  return mergeStreams(streams);
});

gulp.task('spjsom-clean', function () {
  return fs.stat('./src/node-spjsom/index.js', function (err, stat) {
    if (err == null) {
      fs.unlinkSync('./src/node-spjsom/index.js');
    }
  });
});

gulp.task('spjcom-merge-scripts', function () {
  return gulp.src(['./src/node-spjsom/scripts/INIT.debug.js',
      './src/node-spjsom/scripts/MicrosoftAjax-4.0.0.0.debug.js',
      './src/node-spjsom/scripts/SP.Core.debug.js',
      './src/node-spjsom/scripts/SP.Runtime.debug.js',
      './src/node-spjsom/scripts/SP.debug.js'
    ])
    .pipe(concat({
      path: 'spjsom.js'
    }))
    .pipe(gulp.dest('./src/node-spjsom/scripts'))

});

gulp.task('spjsom-insert-scripts', ['spjcom-merge-scripts'], function () {
  var fileContent = fs.readFileSync('./src/node-spjsom/scripts/spjsom.js');
  return gulp.src('./src/node-spjsom/indexDev.js')
    .pipe(replace('//spjcsom', fileContent))
    .pipe(rename('index.js'))
    .pipe(gulp.dest('./src/node-spjsom'));
});

gulp.task('spjsom-init', ['spjsom-clean', 'spjsom-insert-scripts', 'spjcom-merge-scripts'], function () {
  return fs.stat('./src/node-spjsom/scripts/spjsom.js', function (err, stat) {
    if (err == null) {
      fs.unlinkSync('./src/node-spjsom/scripts/spjsom.js');
    }
  });
});


gulp.task('default', ['combineAll']);
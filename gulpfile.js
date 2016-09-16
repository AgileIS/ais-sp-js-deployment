const gulp = require('gulp');
const jsoncombine = require('gulp-jsoncombine');
const foreach = require('gulp-foreach');
const mergeStreams = require('merge-stream');
const jsonFormat = require('gulp-json-format');

const userConfigPrefix = 'config/userconfig_*.json';
const partialConfigPrefix = 'config/partialconfig_*.json';
const configPrefix = 'config_';
const assignments = {
  List: 'InternalName',
  Field: 'InternalName',
  View: 'Title',
  Site: 'Url'
};

function merge(target, source) {
  for (let prop in source) {
    if (source[prop].constructor === Object) {
      target[prop] = merge(target[prop] ? target[prop] : {}, source[prop]);
    } else if (source[prop].constructor === Array) {
      if (!target.hasOwnProperty(prop)) target[prop] = [];
      source[prop].forEach((item) => {
        let result = target[prop].filter((searchItem) => { return searchItem[assignments[prop]] === item[assignments[prop]] });
        if (result.length === 1) result[0] = merge(result[0], item);
        else target[prop].push(item);
      });
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

  streams.push(gulp.src(userConfigPrefix, { read: false })
    .pipe(foreach((stream, file) => {
      streams.push(combine(file.path.replace(file.cwd + '\\', '')));
      return stream;
    })));

  return mergeStreams(streams);
});

gulp.task('default', ['combineAll']);
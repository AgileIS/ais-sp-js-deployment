var gulp = require('gulp'),
    tsc = require('gulp-typescript'),
    gulpTslint = require("gulp-tslint");
    tslint = require('tslint'),
    sourcemaps = require('gulp-sourcemaps'),
    del = require('del'),
    path = require('path'),
    merge = require('merge2'),
    chug = require('gulp-chug');

var tsProject = tsc.createProject('tsconfig.json');
var tslintProgram = tslint.createProgram("tsconfig.json");

gulp.task('ts-lint', ['clean-ts'], function () {
    return gulp.src(['./**/*.ts', '!typings/**/*.ts', '!node_modules/**/*.ts']).pipe(gulpTslint({
            formatter: "prose",
            program: tslintProgram
        }))
        .pipe(gulpTslint.report());
});

gulp.task('compile-ts', ['ts-lint'], function () {
    var tsResult = tsProject.src()
        .pipe(sourcemaps.init())
        .pipe(tsProject());
    return merge([
        tsResult.dts.pipe(gulp.dest('./dist')),
        tsResult.js.pipe(sourcemaps.write('.', {
            // Return relative source map root directories per file.
            includeContent: false,
            sourceRoot: function (file) {
                var sourceFile = path.join(file.cwd, file.sourceMap.file);
                return "../" + path.relative(path.dirname(sourceFile), __dirname);
            }
        })).pipe(gulp.dest('./dist'))
    ]);
});

gulp.task('clean-ts', function (cb) {
    var typeScriptGenFiles = [
        './dist/**'
    ];

    // delete the files
    return del(typeScriptGenFiles, cb);
});

gulp.task('merge', function (cb) {
    return gulp.src('./gulpfile.merge.js')
        .pipe(chug())
});

gulp.task('default', ['clean-ts', 'ts-lint', 'compile-ts']);
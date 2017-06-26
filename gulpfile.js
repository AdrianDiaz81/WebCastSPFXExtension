'use strict';

const gulp = require('gulp');
const build = require('@microsoft/sp-build-web');

require('./gulp-tasks/gulp-serve-info.js');

build.initialize(gulp);

gulp.tasks['serve-info'].dep.push('serve');
'use strict';

const gulp = require('gulp');
const build = require('@microsoft/sp-build-web');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};

build.initialize(gulp);

gulp.task('version-sync', async function () {
    // import file system utilities form nodeJS
    const fs = require('fs');
    // read package.json
    const pkgConfig = require('./package.json');
    // read configuration of web part solution file
    const pkgSolution = require('./config/package-solution.json');
    // Generate new MS compliant version number
    const newVersionNumber = pkgConfig.version.split('-')[0] + '.0';
    // assign newly generated version number to web part version
    pkgSolution.solution.version = newVersionNumber;
    // write changed package-solution file
    fs.writeFile('./config/package-solution.json', JSON.stringify(pkgSolution, null, 4), () => { });
});


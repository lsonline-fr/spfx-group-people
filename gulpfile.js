'use strict';

const gulp = require('gulp');
const build = require('@microsoft/sp-build-web');
build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);
build.addSuppression(`Warning - [sass] The local CSS class 'ms-Persona-secondaryText' is not camelCase and will not be type-safe.`);
build.addSuppression(`Warning - tslint - src/webparts/groupPeople/test/GroupPeople.test.ts(38,7): error no-unused-expression: unused expression, expected an assignment or function call`);

build.initialize(gulp);

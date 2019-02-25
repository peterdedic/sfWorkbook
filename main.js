#!/usr/bin/env node

const _wb = require('./workbook');

let args = process.argv.slice(2);

_wb.get(args[0]);
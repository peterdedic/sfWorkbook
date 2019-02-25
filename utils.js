const _fs = require('fs');
const _path = require('path');
const _rl = require('readline');

Array.prototype.sortByProp = function(prop) {
    // console.log('sort',this)
    return this.sort(function(a,b) {
        //console.log(a[prop], '-' , b[prop])
        if (a[prop] < b[prop])
            return -1;
        if (a[prop] > b[prop])
            return 1;
        return 0;
    });
}

function pad(n) {return n<10?'0'+n:n}
function tzo() {let t=this.getTimezoneOffset(); return t>0?'-':'+' + pad(Math.abs(t/60))+':'+pad(Math.abs(t%60))}

Date.prototype.toLocaleISOString = function() {
    return this.getFullYear() +
    '-' + pad(this.getMonth() + 1) +
    '-' + pad(this.getDate()) +
    'T' + pad(this.getHours()) +
    ':' + pad(this.getMinutes()) +
    ':' + pad(this.getSeconds()) + tzo();
}

Date.prototype.toLocaleDateTimeString = function() {
    return this.getFullYear() +
    '-' + pad(this.getMonth() + 1) +
    '-' + pad(this.getDate()) +
    ' ' + pad(this.getHours()) +
    ':' + pad(this.getMinutes()) +
    ':' + pad(this.getSeconds());
}

function toLocaleDateTimeString(dt) {
    let td = new Date(dt)
    return td.getFullYear() +
    '-' + pad(td.getMonth() + 1) +
    '-' + pad(td.getDate()) +
    ' ' + pad(td.getHours()) +
    ':' + pad(td.getMinutes()) +
    ':' + pad(td.getSeconds());
}


function mkDir(p) {
    if (!_fs.existsSync(p)) {
        let tempPath = '';
        p.split(_path.sep).forEach(dSeg => {
            tempPath = _path.join(tempPath, dSeg);
            if(!_fs.existsSync(tempPath))
                _fs.mkdirSync(tempPath);
        });
    }
}

function copy(from, to) {
    if (_fs.statSync(from).isDirectory()) {
        let files = _fs.readdirSync(from);
        files.forEach(f => {
            copy(_path.join(from, f), _path.join(to, f))
        })
    } else {
        // console.log(`[>] ${to.split(_path.sep).slice(0, -1).join(_path.sep)}`)
        // console.log(`[f] ${from} --> ${to}`);
        mkDir(to.split(_path.sep).slice(0, -1).join(_path.sep));
        _fs.copyFileSync(from, to);
    }
}

function lsDir(source) {
    let list = [];
    _fs.readdirSync(source).map(i => _path.join(source, i)).forEach(i => {
        if (_fs.statSync(i).isDirectory()) {
            list = list.concat(lsDir(i));
        } else {
            list.push(i);
        }
    });
    return list;
}

function terminalWrite(msg) {
    _rl.clearLine(process.stdout, 0)
    _rl.cursorTo(process.stdout, 0)
    process.stdout.write(msg)
}

function readVal(val, alt = '') {
    return val ? val : alt;
}
function readArr(val) {
    if (!val) return [];
    if (Array.isArray(val)) return val;
    return [val];
}

function isDir(path) {
    try {
        return _fs.lstatSync(path).isDirectory();
    } catch (e) {
        // lstatSync throws an error if path doesn't exist
        return false;
    }
}

function Serial() {
    this.startTime = 0;
    this.fnArray = [];
}
Serial.prototype.next = function(fn) {
    this.fnArray.push(fn);
    return this;
}
Serial.prototype.done = function(input) {
    var fn = this.fnArray.shift();
    if (fn) 
        try {
            fn.apply(this, [this.done.bind(this), this.fail.bind(this), input]);
        } catch(e) {
            this.fail(e);
        } 
    return this;
}
Serial.prototype.fail = function(err) {this.onError.apply(this, [err]); return this;}
Serial.prototype.catch = function(fn) {this.onError = fn; return this;}

module.exports = {
    mkdir: mkDir,
    copy: copy,
    Serial: Serial,
    readVal: readVal,
    readArr: readArr,
    lsDir: lsDir,
    isDir: isDir,
    terminalWrite: terminalWrite,
    toLocaleDateTimeString: toLocaleDateTimeString
}
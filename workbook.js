const _xml2js = require('xml2js')
        _utils = require('./utils.js')
        _fs = require('fs')
        _path = require('path');

function objectPermissionTable(ws, profiles, objList, profList) {
    let row = ws.addRow();
    row.getCell(1).value = 'Profiles';
    profList.forEach((p,i) => {
        let idx = (i * 7) + 2;
        row.getCell(idx).value = p.split('.')[0]
        ws.mergeCells(row.number,idx,row.number,idx+5)
    })

    row = ws.addRow();
    row.getCell(1).value = 'Object';
    profList.forEach((p, i) => {
        ['R','C','E','D','VA','MA'].forEach((a, j) => {
            row.getCell((i * 6) + 2 + j).value = a
        })
    })
    
    objList.forEach(o => {
        let apiName = o.split('.')[0];
        let r = [apiName];
        
        profList.forEach(p => {
            let ops = profiles[p].Profile.objectPermissions;
            if (ops) {
                let op = ops.find(i => i.object === apiName);
                if (op) {
                    ['allowRead','allowCreate','allowEdit','allowDelete','viewAllRecords','modifyAllRecords'].forEach(k => r.push(op[k] === 'true' ? 'x' : ''))
                }
            }
        });
        ws.addRow(r).eachCell((c, coln) => c.alignment = {horizontal: 'center'})
    });
    
    // additional formatting
    for (let ii=1; ii<ws.columnCount; ii++) {
        ws.getColumn(ii+1).width = 5;
        ws.getColumn(ii+1).alignment = {horizontal: 'center'}
    }
    ws.getColumn(1).width = 25;
    ws.getColumn(1).alignment = {horizontal: 'left', vertical: 'top'};
    ws.addRow();
}
function wb_getObjectPermissions(ws, dir, callback) {
    var parser = _xml2js.Parser({explicitArray: false});
        
    let serial = new _utils.Serial();

    let profiles = {};
    _fs.readdirSync(_path.join(dir, 'profiles')).forEach(f => {
        serial.next((done,fail,input) => {
            parser.parseString(_fs.readFileSync(_path.join(dir, 'profiles', f), 'utf8'), (err, res) => {profiles[f] = res; done();});
        })
    })
    let objects = {};
    _fs.readdirSync(_path.join(dir, 'objects')).forEach(f => {
        serial.next((done,fail,input) => {
            parser.parseString(_fs.readFileSync(_path.join(dir, 'objects', f), 'utf8'), (err, res) => {objects[f] = res; done();});
        })
    })
    serial.next((done,fail,input) => {
        let standardObj = [];
        let customObj = [];
        
        Object.keys(objects).forEach(o => {
            if (o.match(/__[cex]\./g)) customObj.push(o)
            else standardObj.push(o)
        })

        let row = ws.addRow();
        row.getCell(1).value = 'OBJECT PERMISSIONS';

        let profileList = ['Admin.profile','CN Standard.profile','Customer Community.profile'];
        objectPermissionTable(ws, profiles, standardObj, profileList)
        objectPermissionTable(ws, profiles, customObj, profileList)
        
        if (callback) callback();
        done();
    });
    serial.catch(err => console.log('err:', err))
    serial.done();
}

function parseTranslations(transFileNames) {
    return transFileNames.map(n => n.slice(n.lastIndexOf('-') + 1, n.lastIndexOf('.'))).filter((v, i, s) => s.indexOf(v) === i);
}
function parseTranslationsForObject(transFileNames, objName) {
    return transFileNames.filter(n => n.indexOf(objName) === 0).map(n => n.slice(n.lastIndexOf('-') + 1, n.lastIndexOf('.')));
}

function isJSdocComment(txt) {
    return txt.indexOf('/**') === 0;
}
function indexOfClosingBracket(str, pos) {
    var stack = [''], tp = pos+1;
    while (stack.length > 0 && tp < str.length) {
        if (str[tp] == '}') stack.pop();
        else if (str[tp] == '{') stack.push('');
        tp++;
    }
    return tp-1;
}
function parseJSDoc_Tags(txt) {
    let res = [];
    // split jsDoc into individual section based on new line beginning with @
    txt.substr(txt.indexOf('\n@')).split('\n@').forEach(sec => {
        let m = /([\w]+)\s+([\s\S]+)/.exec(sec);
        if (m)
            res.push({
                type: m[1],
                value: m[2].trim()
            })
    })
    return res
}
function parseJSDoc_Block(jsdoc) {
    var txt = jsdoc.replace(/(^\s*\/\*\*)|(^\s*\*\/*\s?)/gm, '').trim();
    // var tagsFlat = jsdoc.match(/(@\w+)\s+(\{[\w<>]+\})\s+(\w*)\s*-\s*([\w\s.*]+)(?!@)/gm);
    var tagsFlat = jsdoc.match(/(@\w+)\s+([\w\s.]+)(?!@)/gm);

    var firstLineBreakIdx = txt.indexOf('\n'),
        openBracketIdx = txt.indexOf('['),
        closeBracketIdx = txt.indexOf(']'),
        refs = [],
        desc = '';

    if (openBracketIdx > -1 && openBracketIdx < firstLineBreakIdx) {
        refs = txt.substr(openBracketIdx, closeBracketIdx - openBracketIdx).match(/[\w]{2}-[0-9]{4}/g);
        txt = txt.substr(closeBracketIdx+1).trim();
    }

    var tags = [];
    var firstTagSignIdx = txt.indexOf('\n@');
    if (firstTagSignIdx > -1) {
        desc = txt.substr(0, firstTagSignIdx);
        tags = parseJSDoc_Tags(txt);
    } else {
        desc = txt;
    }
    
    return {
        ref: refs,
        desc: desc.trim(),
        tags: tags
    }
}
function parseFiles(dir, callback) {
    let parser = _xml2js.Parser({explicitArray: false});
    let serial = new _utils.Serial();

    let objects = {};
    _fs.readdirSync(dir).forEach(f => {
        serial.next((done,fail,input) => {
            parser.parseString(_fs.readFileSync(_path.join(dir, f), 'utf8'), (err, res) => {objects[f] = res; done();});
        })
    })
    serial.next((done,fail,input) => callback(objects))
    serial.catch(err => console.log('err:', err))
    serial.done();
}
function parseClassParams(txt) {
    let paramRe = /(\w+(?:<.*>)*)\s+([\w]+)/g,
        res = [],
        temp;
    while (temp = paramRe.exec(txt)) {
        res.push({
            name: temp[2],
            type: temp[1]
        })
    }
    return res;
}
function parseClasses(txt) {
    let classRe = /class\s+([\w]+)/g;
    let methodRe = /([\w.]+(?:<.*>)*)\s+([\w]+)(\(.*\))\s*{/g;
    let blockCommentRe = /\/\*([^\*]|(\*(?!\/)))*\*\//g;
    let inlineCommentRe = /\/\/[^\n]*/g;
    
    // process.stdout.write(o)
    let comments = [];
    let temp;
    while (temp = blockCommentRe.exec(txt)) {
        comments.push({
            s: temp.index,
            e: blockCommentRe.lastIndex,
            txt: temp[0]
        })
    }
    while (temp = inlineCommentRe.exec(txt)) {
        comments.push({
            s: temp.index,
            e: inlineCommentRe.lastIndex,
            txt: temp[0]
        })
    }
    let classes = [];
    while (temp = classRe.exec(txt)) {
        // exclude those found inside comments
        if (comments.findIndex(c => temp.index > c.s && temp.index < c.e) === -1) {
            classes.push({
                name: temp[1],
                pos: {
                    s: temp.index,
                    e: classRe.lastIndex
                },
                body: {
                    s: classRe.lastIndex,
                    e: indexOfClosingBracket(txt, classRe.lastIndex + 3)
                }
            })
        }
    }

    let methods = [];
    while (temp = methodRe.exec(txt)) {
        // syntax of some keywords share the same pattern => exclude
        if (['if','do','while','for'].indexOf(temp[2]) === -1) {
            // exclude those found inside comments
            if (comments.findIndex(c => temp.index > c.s && temp.index < c.e) === -1) {
                methods.push({
                    type: temp[1],
                    name: temp[2],
                    params: temp[3],
                    pos: {
                        s: temp.index,
                        e: methodRe.lastIndex
                    },
                    doc: {}
                })
            }
        }
    }

    // find jsDoc block belonging to a method and parse it
    methods.forEach(mtd => {
        let min = Number.MAX_VALUE,
            comment = {};
        comments.filter(c => isJSdocComment(c.txt)).forEach(c => {
            if (Math.abs(c.e - mtd.pos.s) < min) {
                min = Math.abs(c.e - mtd.pos.s);
                comment = c;
            }
        })
        mtd.desc = parseJSDoc_Block(_utils.readVal(comment.txt));
    })
    // methods.forEach(m => console.log(m.desc));
    
    // find class that a method belongs to based on position in the text and body-scope of the class
    methods.forEach(m => {
        let min = {
            name: '-'
        }
        classes.forEach(c => {
            if (m.pos.s > c.body.s && m.pos.s < c.body.e) {
                min.name = c.name;
            }
        })
        m.class = min.name;
    })

    let res = [];
    classes.forEach(cls => {
        res.push({
            name: cls.name,
            methods: methods.filter(m => m.class == cls.name).map(m => {
                return {
                    type: m.type,
                    name: m.name,
                    param: parseClassParams(m.params),
                    doc: m.desc
                }
            })
        })
    })
    return res;
}
function setupColums(ws, cols) {
    row = ws.addRow();
    cols.forEach((col, i) => row.getCell(i+1).value = col.header);
    for (let ii=0; ii<cols.length; ii++) {
        let col = ws.getColumn(ii+1);
        col.width = cols[ii].width;
        if (cols[ii].key) col.key = cols[ii].key;
    }
}
function fieldTypeToString(f) {
    let res = '';
    switch(f.type) {
        case 'LongTextArea':
        case 'Text': res = `${f.type} (${f.length})`; break;
        case 'Lookup': res = `${f.type} (${f.referenceTo})`; break;
        case 'Number': res = `${f.type} (${f.precision - f.scale},${f.scale})`; break;
        default: res = `${f.type}`; break;
    }
    return res;
}
function valuesFormulaToString(f) {
    let res = '';
    try {
        switch(f.type) {
            case 'Picklist': res = `${f.valueSet.valueSetDefinition.value.map(v => v.fullName).join(';')}`; break;
            case 'Formula': res = `${f.formula}`; break;
            default: res = `${f.type}`; break;
        }
    } catch (e) {
        res = 'error reading'
    }
    return res;
}
function dependentOnToString(f) {
    if (f.type === 'Picklist') {
        return '-';
    } else return '';
}
function wb_getFields(ws, dir, callback) {
    new _utils.Serial()
    .next((done,fail,input) => parseFiles(_path.join(dir, 'objects'), res => done(res)))
    .next((done,fail,input) => {
        ws.addRow().getCell(1).value = 'Fields';
        ws.addRow().getCell(1).value = 'Note:';
        let cols = [
            { header: 'Reference',                  key: 'refc', width: 30 },
            { header: 'Object',                     key: 'objn', width: 30 },
            { header: 'Field Label',                key: 'flab', width: 30 },
            { header: 'Fieldname SAP',              key: 'fsap', width: 30 },
            { header: 'API Name',                   key: 'apin', width: 30 },
            { header: 'Type',                       key: 'type', width: 30 },
            { header: 'Description',                key: 'desc', width: 30 },
            { header: 'Dependent On',               key: 'depo', width: 30 },
            { header: 'Value, Details, Formula',    key: 'vdfo', width: 30 },
            { header: 'HelpText',                   key: 'htxt', width: 30 },
            { header: 'History Tracking',           key: 'hist', width: 30 },
            { header: 'Chatter Feed Tracking',      key: 'chft', width: 30 }
        ];
        setupColums(ws, cols);
        
        row = ws.addRow()
        Object.keys(input).forEach((o, oi) => {
            let root = input[o].CustomObject;
            if (!root.fields) {
                console.log('no fields for ' + o);
                return;
            }
            root.fields.forEach((f, fi) => {
                row = ws.addRow();
                let i = 1;
                row.getCell(i++).value = '';
                row.getCell(i++).value = root.label ? root.label : o.split('.')[0];
                row.getCell(i++).value = f.label;
                row.getCell(i++).value = '';
                row.getCell(i++).value = f.fullName;
                row.getCell(i++).value = fieldTypeToString(f);
                row.getCell(i++).value = f.description ? f.description : '';
                row.getCell(i++).value = dependentOnToString(f);
                row.getCell(i++).value = valuesFormulaToString(f);
                row.getCell(i++).value = f.inlineHelpText ? f.inlineHelpText : '';
                row.getCell(i++).value = f.trackHistory ? f.trackHistory : '';
                row.getCell(i++).value = f.trackFeedHistor ? f.trackFeedHistor : '';
            })
            //console.log(res)
        });
        if (callback) callback();
        done();
    })
    .catch(err => console.log('err:', err))
    .done();
}
function wb_getOWD(ws, dir, callback) {
    new _utils.Serial()
    .next((done,fail,input) => parseFiles(_path.join(dir, 'objects'), res => done(res)))
    .next((done,fail,input) => {
        ws.addRow().getCell(1).value = 'Sharing Rules';
        let cols = [
            { header: 'Object',                   width: 30 },
            { header: 'Default Internal Access',  width: 30 },
            { header: 'Default External Access',  width: 30 }
        ];
        setupColums(ws, cols);

        Object.keys(input).forEach(o => {
            let root = input[o].CustomObject;
            row = ws.addRow();
            row.getCell(1).value = _utils.readVal(root.label, o.split('.')[0]);
            row.getCell(2).value = _utils.readVal(root.externalSharingModel);
            row.getCell(3).value = _utils.readVal(root.sharingModel);
        })
        if (callback) callback();
        done();
    })
    .catch(err => console.log('err:', err))
    .done();
}
function wb_getAPP(ws, dir, callback) {
    new _utils.Serial()
    .next((done,fail,input) => parseFiles(_path.join(dir, 'applications'), res => done(res)))
    .next((done,fail,input) => {
        ws.addRow().getCell(1).value = 'Application';
        ws.addRow().getCell(1).value = 'Note:';
        let cols = [
            { header: 'App Label',                  width: 30 },
            { header: 'App Name',                   width: 30 },
            { header: 'Included Tab',               width: 30 },
            { header: 'Default Landing Tab',        width: 30 }
        ];
        setupColums(ws, cols);

        Object.keys(input).forEach(o => {
            let root = input[o].CustomApplication;
            let appName = o.split('.')[0];
            row = ws.addRow();
            row.getCell(1).value = _utils.readVal(root.label, appName);
            row.getCell(2).value = appName;
            row.getCell(3).value = _utils.readArr(root.tabs).join('; ');
            row.getCell(4).value = _utils.readVal(root.defaultLandingTab, 'Home');
        })
        if (callback) callback();
        done();
    })
    .catch(err => console.log('err:', err))
    .done();
}
function wb_getObjects(ws, dir, callback) {
    new _utils.Serial()
    .next((done,fail,input) => parseFiles(_path.join(dir, 'objects'), res => done(res)))
    .next((done,fail,input) => {
        ws.addRow().getCell(1).value = 'Objects';
        ws.addRow().getCell(1).value = 'Note:';
        let cols = [
            { header: 'Label',        width: 30 },
            { header: 'Object Name',  width: 30 },
            { header: 'Description',  width: 30 }
        ];
        setupColums(ws, cols);

        Object.keys(input).forEach(o => {
            let root = input[o].CustomObject;
            let name = o.split('.')[0];
            row = ws.addRow();
            row.getCell(1).value = _utils.readVal(root.label, name);
            row.getCell(2).value = _utils.readVal(name);
            row.getCell(3).value = _utils.readVal(root.description);
        })
        if (callback) callback();
        done();
    })
    .catch(err => console.log('err:', err))
    .done();
}
function wb_getPicklistValues(ws, dir, callback) {
    new _utils.Serial()
    .next((done,fail,input) => parseFiles(_path.join(dir, 'objects'), res => done(res)))
    .next((done,fail,input) => parseFiles(_path.join(dir, 'objectTranslations'), res => done({objects: input, translations: res})))
    .next((done,fail,input) => {
        ws.addRow().getCell(1).value = 'Picklist Values';
        ws.addRow().getCell(1).value = 'Note:';
        let cols = [
            { header: 'Reference',        width: 30 },
            { header: 'Object',  width: 30 },
            { header: 'Field API Name',  width: 30 },
            { header: 'Field Label',  width: 30 },
            { header: 'Picklist Value',  width: 30 }
        ];
        parseTranslations(Object.keys(input.translations)).forEach(t => cols.push(
            {header: `Translation (${t})`, width: 30, key: `t_${t}`}
        ))
        setupColums(ws, cols);
        let headerRow = ws.lastRow;

        Object.keys(input.objects).forEach(o => {
            // console.log(o)
            let root = input.objects[o].CustomObject;
            let objName = o.split('.')[0];
            
            _utils.readArr(root.fields).filter(f => f.type === 'Picklist').forEach(pf => {
                let values = [];
                if (pf.valueSet.valueName) values = [{fullName: pf.valueSet.valueName}];
                if (pf.valueSet.valueSetDefinition) values = _utils.readArr(pf.valueSet.valueSetDefinition.value);
                // console.log('   ' + pf.fullName)
                values.forEach(v => {
                    row = ws.addRow();
                    // console.log('      ' + v.fullName)
                    row.getCell(1).value = '';
                    row.getCell(2).value = _utils.readVal(root.label, objName);
                    row.getCell(3).value = _utils.readVal(pf.fullName);
                    row.getCell(4).value = _utils.readVal(pf.label);
                    row.getCell(5).value = _utils.readVal(v.fullName);
                    
                    let langList = parseTranslationsForObject(Object.keys(input.translations), objName);
                    langList.forEach((l, li) => {
                        let transField = _utils.readArr(input.translations[objName + '-' + l + '.objectTranslation'].CustomObjectTranslation.fields).find(f => f.name === pf.fullName);
                        
                        if (transField) {
                            let trans = _utils.readArr(transField.picklistValues).find(tpv => tpv.masterLabel === v.fullName);
                            if (trans) {
                                // row.getCell(6 + li).value = _utils.readVal(trans.translation);
                                row.getCell(ws.getColumn(`t_${l}`).number).value = _utils.readVal(trans.translation);
                            }
                        }
                    })
                })
            })
        })
        if (callback) callback();
        done();
    })
    .catch(err => console.log('err:', err))
    .done();
}
function wb_getRecordTypes(ws, dir, callback) {
    new _utils.Serial()
    .next((done,fail,input) => parseFiles(_path.join(dir, 'objects'), res => done(res)))
    .next((done,fail,input) => {
        ws.addRow().getCell(1).value = 'Record Types';
        ws.addRow().getCell(1).value = 'Note:';
        let cols = [
            { header: 'Entity',      key: 'enti', width: 30 },
            { header: 'Record Type', key: 'rtyp', width: 30 },
            { header: 'API Name',    key: 'apin', width: 30 },
            { header: 'Description', key: 'desc', width: 30 }
        ];
        setupColums(ws, cols);
        
        Object.keys(input).forEach((o, oi) => {
            let root = input[o].CustomObject;
            let name = o.split('.')[0];
            
            _utils.readArr(root.recordTypes).forEach(rt => {
                row = ws.addRow();
                let i = 1;
                row.getCell(i++).value = _utils.readVal(root.label, name);
                row.getCell(i++).value = _utils.readVal(rt.label);
                row.getCell(i++).value = _utils.readVal(rt.fullName);
                row.getCell(i++).value = _utils.readVal(rt.description);
            })
        });
        if (callback) callback();
        done();
    })
    .catch(err => console.log('err:', err))
    .done();
}
function wb_getCustomLabels(ws, dir, callback) {
    new _utils.Serial()
    .next((done,fail,input) => parseFiles(_path.join(dir, 'labels'), res => done(res)))
    .next((done,fail,input) => {
        ws.addRow().getCell(1).value = 'Custom Labels';
        ws.addRow().getCell(1).value = 'Note:';
        let cols = [
            { header: 'Name',                   key: 'name', width: 30 },
            { header: 'Categories',             key: 'cate', width: 30 },
            { header: 'Short Description',      key: 'desc', width: 30 },
            { header: 'Value',                  key: 'valu', width: 30 },
            { header: 'Language',               key: 'lang', width: 30 }
        ];
        setupColums(ws, cols);
        
        Object.keys(input).forEach((o, oi) => {
            let root = input[o].CustomLabels;
            // let name = o.split('.')[0];
            
            _utils.readArr(root.labels).forEach(cl => {
                row = ws.addRow();
                let i = 1;
                row.getCell(i++).value = _utils.readVal(cl.fullName);
                row.getCell(i++).value = _utils.readVal(cl.categories);
                row.getCell(i++).value = _utils.readVal(cl.shortDescription);
                row.getCell(i++).value = _utils.readVal(cl.value);
                row.getCell(i++).value = _utils.readVal(cl.language);
            })
        });
        if (callback) callback();
        done();
    })
    .catch(err => console.log('err:', err))
    .done();
}
function wb_getTriggers(ws, dir, callback) {
    new _utils.Serial()
    // .next((done,fail,input) => parseFiles(_path.join(dir, 'triggers'), res => done(res)))
    .next((done,fail,input) => _fs.readdir(_path.join(dir, 'triggers'), (err, files) => done(files)))
    .next((done,fail,input) => {
        ws.addRow().getCell(1).value = 'Triggers';
        ws.addRow().getCell(1).value = 'Note:';
        let cols = [
            { header: 'Type',                   key: 'type', width: 30 },
            { header: 'Name',             key: 'name', width: 30 },
            { header: 'Code Coverage',      key: 'code', width: 30 },
            { header: 'Object',                  key: 'object', width: 30 },
            { header: 'Description',               key: 'desc', width: 30 }
        ];
        setupColums(ws, cols);
        
        input.forEach(o => {
            let name = o.split('.')[0];
            
            row = ws.addRow();
            let c = 1;
            row.getCell(c++).value = _utils.readVal('');
            row.getCell(c++).value = _utils.readVal(name);
            row.getCell(c++).value = _utils.readVal('');
            row.getCell(c++).value = _utils.readVal('');
            row.getCell(c++).value = _utils.readVal('');
        });
        if (callback) callback();
        done();
    })
    .catch(err => console.log('err:', err))
    .done();
}
function wb_getClasses(ws, dir, callback) {
    new _utils.Serial()
    // .next((done,fail,input) => parseFiles(_path.join(dir, 'triggers'), res => done(res)))
    .next((done,fail,input) => _fs.readdir(_path.join(dir, 'classes'), (err, files) => done(files)))
    .next((done,fail,input) => {
        ws.addRow().getCell(1).value = 'Classes';
        ws.addRow().getCell(1).value = 'Note:';
        ws.views = [
            {state: 'frozen', xSplit: 0, ySplit: 3, showGridLines: false}
        ]
        // ws.autoFilter = 'A3:G3'
        ws.autoFilter = {
            from: {row: 3, column: 1},
            to: {row: 3, column: 7}
        }
        let cols = [
            { header: 'Class Name',             key: 'cname', width: 30 },
            { header: 'Code Coverage',          key: 'code', width: 30 },
            { header: 'Method Type',            key: 'mtype', width: 30 },
            { header: 'Method Name',            key: 'mname', width: 30 },
            { header: 'Method Params',          key: 'mpars', width: 30 },
            { header: 'Method Description',     key: 'mdesc', width: 60 },
            { header: 'Comments',               key: 'comm', width: 30 }
        ];
        setupColums(ws, cols);
        
        input.forEach(o => {
            let name = o.split('.')[0];
            if (o.split('.')[1] !== 'cls') return;

            let txt = _fs.readFileSync(_path.join(dir, 'classes', o), 'utf8');
            
            let classes = parseClasses(txt);

            // console.log(classes)
            // txt.match(classRe).forEach()
            
            classes.forEach(cls => {
                let methods = cls.methods.length === 0 ? [{param:[]}] : cls.methods;
                _utils.readArr(methods).forEach(mtd => {
                    row = ws.addRow();
                    let c = 1;
                    row.getCell(c++).value = _utils.readVal(cls.name);
                    row.getCell(c++).value = _utils.readVal('');
                    row.getCell(c++).value = _utils.readVal(mtd.type);
                    row.getCell(c++).value = _utils.readVal(mtd.name);
                    row.getCell(c++).value = _utils.readVal(mtd.param.map(p => p.type + ' ' + p.name).join('; '));
                    if (mtd.doc) {
                        let val = {richText: []};
                        val.richText.push({font: {italic: true}, text: _utils.readVal(mtd.doc.desc)})
                        _utils.readArr(mtd.doc.tags).forEach(t => {
                            // val.richText.push({text: '\n'});
                            val.richText.push({font: {bold: true}, text: '\n' + _utils.readVal(t.type) + ': ' });
                            val.richText.push({text: _utils.readVal(t.value)});
                        })

                        let cell = row.getCell(c++);
                        cell.value = val;
                        cell.alignment = {wrapText: true};
                    }
                    //
                    // row.getCell(c++).value = _utils.readVal('');
                    // row.getCell(c++).value = _utils.readVal('');
                })
            })
        });
        if (callback) callback();
        done();
    })
    .catch(err => console.log('err:', err))
    .done();
}

module.exports = {
    get : (dir, outFile = 'output.xlsx') => {
        if (!_utils.isDir(dir)) {
            console.log(`error: "${dir}" is not a directory`);
            return;
        }

        Excel = require('exceljs');
        let wb = new Excel.Workbook();

        new _utils.Serial()
        .next((done,fail,input) => {
            let ws = wb.addWorksheet('Object Permissions');
            wb_getObjectPermissions(ws, dir, res => done(input));
        })
        .next((done,fail,input) => {
            let ws = wb.addWorksheet('Fields');
            wb_getFields(ws, dir, res => done(input));
        })
        .next((done,fail,input) => {
            let ws = wb.addWorksheet('OWD');
            wb_getOWD(ws, dir, res => done(input));
        })
        .next((done,fail,input) => {
            let ws = wb.addWorksheet('APP');
            wb_getAPP(ws, dir, res => done(input));
        })
        .next((done,fail,input) => {
            let ws = wb.addWorksheet('Objects');
            wb_getObjects(ws, dir, res => done(input));
        })
        .next((done,fail,input) => {
            let ws = wb.addWorksheet('Picklist Values');
            wb_getPicklistValues(ws, dir, res => done(input));
        })
        .next((done,fail,input) => {
            let ws = wb.addWorksheet('Record Types');
            wb_getRecordTypes(ws, dir, res => done(input));
        })
        .next((done,fail,input) => {
            let ws = wb.addWorksheet('Custom Labels');
            wb_getCustomLabels(ws, dir, res => done(input));
        })
        .next((done,fail,input) => {
            let ws = wb.addWorksheet('Triggers');
            wb_getTriggers(ws, dir, res => done(input));
        })
        .next((done,fail,input) => {
            let ws = wb.addWorksheet('Classes');
            wb_getClasses(ws, dir, res => done(input));
        })
        .next((done,fail,input) => {done(input)})
        .next((done,fail,input) => {done(input)})
        .next((done,fail,input) => {
            wb.xlsx.writeFile(outFile);
        })
        .catch(err => console.log('err', err))
        .done({})
    }
}
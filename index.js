flat = require('flatjson');
xlsx = require('node-xlsx');

function createSheet(obj, sheetname, delimiter, filter) {
    var sheet = sheetname || "Sheet 1";
    
    var input = [];
    if (obj instanceof Array) {
        input = obj;
    }
    else {
        input.push(obj);
    }

    //build headers
    var headers = []
    for ( var i = 0; i < input.length; i++ ) {
        var fo = flat(input[i], delimiter, filter);
        var keys = Object.keys(fo);
        for (var j = 0; j < keys.length; j++ ) {
            if ( headers.indexOf(keys[j]) < 0) {
                headers.push(keys[j]);
            }
        }
    }

    var data = [];
    data.push(headers);
    for ( var i = 0; i < input.length; i++ ) {
        var actual_data = []
        var fo = flat(input[i], delimiter, filter);
        for (key in headers) {
            actual_data.push(fo[headers[key]]);
        }
        data.push(actual_data);
    }
    return {
        name: sheet,
        data: data
    }
}

module.exports = function (options) {
    
    if(!(options instanceof Array)) {
        options = [options];    
    }
    var sheets = options.map(function(value, index, array) {
        return createSheet(
            value.obj,
            value.sheetname,
            value.delimiter,
            value.filter
        );
    });
    return xlsx.build(sheets);
}

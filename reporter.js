var fs = require('fs'),
    path = require('canonical-path'),
    _ = require('lodash'),
    excel4node = require('excel4node');

// Custom reporter
var Reporter = function(options) {

    var workBook = new excel4node.Workbook(),
        workSheet = workBook.addWorksheet('Test');

    var _defaultOutputFile = path.resolve(process.cwd(), './_test-output', 'protractor-results.xlsx');
    options.outputFile = options.outputFile || _defaultOutputFile;

    initOutputFile(options.outputFile);
    options.appDir = options.appDir ||  './';
    var _root = { appDir: options.appDir, suites: [] };
    log('AppDir: ' + options.appDir, +1);
    var _currentSuite;

    this.suiteStarted = function(suite) {
        _currentSuite = { description: suite.description, status: null, specs: [] };
        _root.suites.push(_currentSuite);
        log('Suite: ' + suite.description, +1);
    };

    this.suiteDone = function(suite) {
        var statuses = _currentSuite.specs.map(function(spec) {
            return spec.status;
        });
        statuses = _.uniq(statuses);
        var status = statuses.indexOf('failed') >= 0 ? 'failed' : statuses.join(', ');
        _currentSuite.status = status;
        log('Suite ' + _currentSuite.status + ': ' + suite.description, -1);
    };

    this.specStarted = function(spec) {

    };

    this.specDone = function(spec) {
        var currentSpec = {
            description: spec.description,
            status: spec.status
        };
        if (spec.failedExpectations.length > 0) {
            currentSpec.failedExpectations = spec.failedExpectations;
        }

        _currentSuite.specs.push(currentSpec);
        log(spec.status + ' - ' + spec.description);

    };

    this.jasmineDone = function() {
        outputFile = options.outputFile;
        var output = formatOutput(_root);
        fs.appendFileSync(outputFile, output);
    };

    function ensureDirectoryExistence(filePath) {
        var dirname = path.dirname(filePath);
        if (directoryExists(dirname)) {
            return true;
        }
        ensureDirectoryExistence(dirname);
        fs.mkdirSync(dirname);
    }

    function directoryExists(path) {
        try {
            return fs.statSync(path).isDirectory();
        }
        catch (err) {
            return false;
        }
    }

    function initOutputFile(outputFile) {
        ensureDirectoryExistence(outputFile);
    }

    // for output file output
    function formatOutput(output) {
        var header = "Protractor results for: " + (new Date()).toLocaleString() + "\n\n";
        workSheet.cell(1,1,1,4,true).string(header);
        workSheet.cell(2,1,2,2,true).string('Directory: ' + output.appDir);
        var i = 3;
        output.suites.forEach(function(suite) {
            workSheet.cell(i,1).string('Suite:');
            workSheet.cell(i,2,i,4,true).string(suite.description);
            workSheet.cell(i,5).string(suite.status);
            i++;
            suite.specs.forEach(function(spec) {
                workSheet.cell(i,5).string(spec.status);
                workSheet.cell(i,2,i,4,true).string(spec.description).style({ alignment: {wrapText: true} });
                if (spec.failedExpectations) {
                    spec.failedExpectations.forEach(function (fe) {
                        workSheet.cell(i,6).string('message: ' + fe.message).style({alignment: {wrapText: true}});
                    });
                }
                i++;
            });
        });
        
        return workBook.write(options.outputFile);

    }   

    // for console output
    var _pad;
    function log(str, indent) {
        _pad = _pad || '';
        if (indent == -1) {
            _pad = _pad.substr(2);
        }
        console.log(_pad + str);
        if (indent == 1) {
            _pad = _pad + '  ';
        }
    }
};

module.exports = Reporter;

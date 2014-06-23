var global,
    console,
    require;

global = (function () {
    'use strict';

    return {
        NaN:                NaN,
        Infinity:           Infinity,
        undefined:          undefined,
        eval:               eval,
        parseInt:           parseInt,
        parseFloat:         parseFloat,
        isNaN:              isNaN,
        isFinite:           isFinite,
        decodeURI:          decodeURI,
        decodeURIComponent: decodeURIComponent,
        encodeURI:          encodeURI,
        encodeURIComponent: encodeURIComponent,
        Object:             Object,
        Function:           Function,
        Array:              Array,
        String:             String,
        Boolean:            Boolean,
        Number:             Number,
        Date:               Date,
        RegExp:             RegExp,
        Error:              Error,
        EvalError:          EvalError,
        RangeError:         RangeError,
        SyntaxError:        SyntaxError,
        TypeError:          TypeError,
        URIError:           URIError,
        Math:               Math
    };
}());

console = (function () {
    'use strict';

    function Console() {
        this._logLevel = 1;
    }

    Console.prototype._log = function (type, args) {
        var i;
        if (type < this._logLevel) {
            return;
        }

        for (i = 0; i < args.length; i += 1) {
            if ("object" === typeof args[i]) {
                this._log(type, args[i]);
            } else {
                WScript.Stdout.WriteLine(String(args[i]));
            }
        }
    };

    Console.prototype.log = function () {
        this._log(Console.TYPE_LOG, arguments);
    };

    Console.prototype.trace = function () {
        this._log(Console.TYPE_TRACE, arguments);
    };

    Console.prototype.info = function () {
        this._log(Console.TYPE_INFO, arguments);
    };

    Console.prototype.warn = function () {
        this._log(Console.TYPE_WARN, arguments);
    };

    Console.prototype.error = function () {
        this._log(Console.TYPE_ERROR, arguments);
    };

    Console.TYPE_TRACE = 1;
    Console.TYPE_INFO = 2;
    Console.TYPE_WARN = 3;
    Console.TYPE_ERROR = 4;
    Console.TYPE_LOG = 5;

    return new Console(1);
}());

require = (function () {
    'use strict';

    var _shell = new ActiveXObject("WScript.Shell"),
        _fso = new ActiveXObject("Scripting.FileSystemObject"),
        _dirname = _shell.CurrentDirectory,
        _require;

    function read(filePath) {
        var fd = _fso.OpenTextFile(filePath, 1),
            content = fd.AtEndOfStream ? "" : fd.ReadAll();

        return content;
    }

    function bind(fn, context) {
        return function () {
            return fn.apply(context, arguments);
        };
    }

    function ModuleScope() {
        this._varNames = [];
        this._varValues = [];
    }

    ModuleScope.prototype.add = function (name, value) {
        this._varNames.push(name);
        this._varValues.push(value);

        return this;
    };

    ModuleScope.prototype.getVarNames = function () {
        return this._varNames;
    };

    ModuleScope.prototype.getVarValues = function () {
        return this._varValues;
    };

    function Require(dirname) {
        this._dirname = dirname;
    }

    Require.prototype.require = function (moduleName) {
        var modulePath = this.resolve(moduleName),
            module,
            moduleContent,
            moduleScope,
            moduleInitializer,
            moduleRequire,
            __dirname,
            __filename;

        console.trace("$> requiring " + modulePath);

        module = this.cache[modulePath];

        if (module !== undefined) {
            console.trace("$> module cache found");

            return module.exports;
        }

        moduleContent = read(modulePath);

        module = {
            loaded:  false,
            exports: {}
        };

        moduleRequire = new Require(_fso.GetParentFolderName(modulePath));

        __dirname = _fso.GetParentFolderName(modulePath);
        __filename = modulePath;

        this.cache[modulePath] = module;

        moduleScope = new ModuleScope();

        moduleScope
            .add("global", global)
            .add("console", console)
            .add("require", bind(moduleRequire.require, moduleRequire))
            .add("module", module)
            .add("exports", module.exports)
            .add("__dirname", __dirname)
            .add("__filename", __filename);

        moduleInitializer = Function(
            moduleScope.getVarNames().toString(),
                "{" + moduleContent + "}"
        );

        try {
            moduleInitializer.apply(global, moduleScope.getVarValues());

            module.loaded = true;
        } catch (e) {
            delete this.cache[modulePath];

            throw e;
        }

        return module.exports;
    };

    Require.prototype.resolve = function (moduleName) {
        return _fso.GetAbsolutePathName(this._dirname + moduleName);
    };

    Require.prototype.cache = {};

    _require = new Require(_dirname);

    return bind(_require.require, _require);
}());
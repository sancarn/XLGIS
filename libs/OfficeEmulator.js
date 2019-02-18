"use strict";
exports.__esModule = true;
exports["default"] = OfficeEmulator = {
    EventType: {
        SettingsChanged: "settings-changed"
    },
    context: {
        document: {
            settings: {
                get: function (name) {
                    return this.data[name];
                },
                set: function (name, value) {
                    this.data[name] = value;
                },
                addHandlerAsync: function (type, handler, options, callback) {
                    if (!options)
                        options = {};
                    try {
                        if (type == "settings-changed") {
                            this._private_data.handlers.push(handler);
                        }
                        ;
                        callback({
                            result: "succeeded",
                            asyncContext: options["asyncContext"],
                            value: undefined,
                            error: undefined
                        });
                    }
                    catch (e) {
                        callback({
                            result: "failed",
                            error: e,
                            value: undefined,
                            asyncContext: options["asyncContext"]
                        });
                    }
                },
                saveAsync: function (options, callback) {
                    if (!options)
                        options = {};
                    try {
                        var This = this;
                        this._private_data.handlers.forEach(function (handler) {
                            handler({
                                settings: This,
                                type: "settings-changed"
                            });
                        });
                        callback({
                            status: "succeeded",
                            value: This,
                            error: undefined,
                            asyncContext: options["asyncContext"]
                        });
                        return This;
                    }
                    catch (e) {
                        callback({
                            status: "failed",
                            value: This,
                            error: e,
                            asyncContext: options["asyncContext"]
                        });
                        return undefined;
                    }
                },
                refreshAsync: function (callback) {
                    return callback(this);
                },
                remove: function (name) {
                    delete this.data[name];
                },
                data: {},
                _private_data: {
                    handlers: []
                }
            }
        }
    }
};
//# sourceMappingURL=OfficeEmulator.js.map
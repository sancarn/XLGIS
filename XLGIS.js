"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
exports.__esModule = true;
var OfficeEmulator_1 = require("./libs/OfficeEmulator");
//Leaflet extension for settings cog
var LeafletHelpers = {
    SettingsCog: L.Control.extend({
        onAdd: function ( /* map */) {
            var container = L.DomUtil.create('div');
            container.classList.add("leaflet-settingsCog");
            //Basic div properties
            //container.style.backgroundColor = "ffffff";
            container.style.width = "27px";
            container.style.height = "27px";
            container.style.margin = "10px";
            container.style.padding = "3px";
            //Border shadow style
            container.style.borderColor = "rgba(0,0,0,0.2)";
            container.style.borderWidth = "2px";
            container.style.borderRadius = "4px";
            container.style.borderStyle = "solid";
            //Settings image icon
            container.style.backgroundImage = 'url("https://cdn2.iconfinder.com/data/icons/web/512/Cog-512.png")';
            container.style.backgroundRepeat = "no-repeat";
            container.style.backgroundSize = "23px";
            container.style.backgroundPosition = "center";
            //element to append events to
            this.domElement = container;
            //Make callback
            var This = this;
            this.callback = function (ev) {
                L.DomEvent.stopPropagation(ev);
                This.options.handler(ev);
            };
            //Register click listener
            L.DomEvent.on(this.domElement, 'click', this.callback);
            return this.domElement;
        },
        onRemove: function ( /* map */) {
            //Unregister click listener 
            L.DomEvent.off(this.domElement, 'click', this.callback);
        }
    }),
    settingsCog: function (opts) {
        return new this.SettingsCog(opts);
    }
};
var XLGIS = /** @class */ (function () {
    function XLGIS() {
    }
    XLGIS.initialise = function () {
        if (!Office)
            var Office = OfficeEmulator_1["default"];
    };
    return XLGIS;
}());
XLGIS.initialise = function (Office) {
    var Office;
    return __awaiter(this, void 0, void 0, function () {
        var settings, defaults;
        return __generator(this, function (_a) {
            if (!Office) {
                Office = XLGIS.test_Office();
            }
            XLGIS._Office = Office;
            settings = Office.context.document.settings;
            XLGIS._Settings = settings;
            //If settings don't exist, add them.
            if (settings.get("center") == null)
                settings.set("center", [51.505, -0.09]);
            if (settings.get("zoom") == null)
                settings.set("zoom", 13);
            if (settings.get("tileLayers") == null) { //tileLayers open street and topo maps as default
                defaults = [];
                defaults.push({
                    displayName: "Open street map",
                    tilePattern: 'https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',
                    attributionHTML: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors'
                });
                defaults.push({
                    displayName: "Open topo map",
                    tilePattern: 'https://{s}.tile.opentopomap.org/{z}/{x}/{y}.png',
                    attributionHTML: 'Map data: &copy; <a href="http://www.openstreetmap.org/copyright">OpenStreetMap</a>, <a href="http://viewfinderpanoramas.org">SRTM</a> | Map style: &copy; <a href="https://opentopomap.org">OpenTopoMap</a> (<a href="https://creativecommons.org/licenses/by-sa/3.0/">CC-BY-SA</a>)'
                });
                settings.set("tileLayers", defaults);
            }
            ;
            if (settings.get("frontLayers") == null)
                settings.set("frontLayers", []);
            if (settings.get("projections") == null) { //earth (aka latlong) and british national grid.
                settings.set("projections", {
                    "Earth": '+proj=longlat +datum=WGS84 +no_defs ',
                    "British National Grid": '+proj=tmerc +lat_0=49 +lon_0=-2 +k=0.9996012717 +x_0=400000 +y_0=-100000 +ellps=airy +towgs84=446.448,-125.157,542.06,0.15,0.247,0.842,-20.489 +units=m +no_defs '
                });
            }
            ;
            //Save settings and then initialise map
            settings.saveAsync(undefined, function () {
                //After save
                //--------------------------------------------
                //Create leaflet map
                XLGIS.mapElement = document.getElementById("mainMap");
                XLGIS.map = L.map('mainMap');
                //setView to settings or default center.
                XLGIS.map.setView(settings.get("center"), settings.get("zoom"));
                //Back layers
                var tileLayers = {};
                settings.get("tileLayers").forEach(function (layer) {
                    tileLayers[layer.displayName] = L.tileLayer(layer.tilePattern, { attribution: layer.attributionHTML });
                });
                //Front layers
                var frontLayers = {};
                settings.get("frontLayers").forEach(function (layer) {
                    frontLayers[layer.displayName] = L.layerGroup(XLGIS.layers.getLayer(layer));
                });
                //Add to map:
                L.control.layers(tileLayers, frontLayers).addTo(XLGIS.map);
                //Settings control:
                LeafletHelpers.settingsCog({
                    position: "bottomleft",
                    handler: function () {
                        document.querySelector("#settings-main").classList.remove("hidden");
                    }
                }).addTo(XLGIS.map);
                console.warn("Mapper initialisation finished");
                //...
            });
            //Add handlers:
            settings.addHandlerAsync(Office.EventType.SettingsChanged, function () {
                XLGIS.map.setView(settings.get("center"));
                //...
            });
            callTestCase();
            return [2 /*return*/, true];
        });
    });
};
//@ts-ignore
window.XLGIS = XLGIS;
//Error handling and listeners.
XLGIS.errprs;
GenericObject;
XLGIS.errors.value = [];
XLGIS.errors.groups = {};
XLGIS.errors.listeners = {};
XLGIS.errors.on = function (event, func) { this.listeners[event] = func; };
XLGIS.errors.raise = function (error, groups) {
    error.groups = groups;
    this.value.push(error);
    console.error(error);
    if (groups) {
        groups.forEach(function (group) {
            XLGIS.errors.groups[group] = error;
        });
    }
    ;
    this.listeners["raise"].forEach(function (listener) { return listener(error, groups); });
};
//Used by combobox
XLGIS.data = {};
XLGIS.data.getNamedDatabodies = function () {
    return __awaiter(this, void 0, void 0, function () {
        return __generator(this, function (_a) {
            return [2 /*return*/, Excel.run(function (context) {
                    return __awaiter(this, void 0, void 0, function () {
                        var names, tables;
                        return __generator(this, function (_a) {
                            switch (_a.label) {
                                case 0:
                                    names = context.workbook.names;
                                    tables = context.workbook.tables;
                                    names.load("items");
                                    tables.load("items");
                                    return [4 /*yield*/, context.sync()];
                                case 1:
                                    _a.sent();
                                    return [2 /*return*/, names.items.map(function (namedRange) {
                                            return { name: namedRange.name, type: "namedRange" };
                                        }).concat(tables.items.map(function (table) {
                                            return { name: table.name, type: "table" };
                                        }))];
                            }
                        });
                    });
                })];
        });
    });
};
//XL Wrappers around Leaflet
XLGIS.layers = {};
XLGIS.layers.getLayerPart = function (geotype, geometry, options) {
    try {
        if (geotype.toUpperCase() == "POINT") {
            return L.CircleMarker(L.latlng(geometry), options);
        }
        else if (geotype.toUpperCase() == "LINE") {
            return L.polyline(geometry, options);
        }
        else if (geotype.toUpperCase() == "POLYGON") {
            return L.polygon(geometry, options);
        }
        else if (geotype.toUpperCase() == "CIRCLE") {
            return L.Circle(geometry, options);
        }
        else if (geotype.toUpperCase() == "RECT") {
            return L.rectangle(geometry, options);
        }
        else if (geotype.toUpperCase() == "MARKER") {
            return L.marker(geometry, options);
        }
        else if (geotype.toUpperCase() == "IMAGE") {
            return L.Marker(geometry, options);
        }
    }
    catch (e) {
        return e;
    }
};
XLGIS.layers.getLayer = function (layer) {
    return __awaiter(this, void 0, void 0, function () {
        var data, geometries;
        return __generator(this, function (_a) {
            /*
            {
              type: "table/range",
              name: "name of table, range or range address",
              displayName:"Layer name as in leaflet control",
              projection: "projectionName",
          
            }
            */
            //By Default have a click handler on points which writes the ID of the point into the range named "<<layerName>>_click"
            if (layer.type == "table") {
                return [2 /*return*/, Excel.run(function (context) {
                        return __awaiter(this, void 0, void 0, function () {
                            var table, geotype, geometry, style, geotypes, geodatas, styles, geometries, i, geodata, style_1, geotype_1, geopart;
                            return __generator(this, function (_a) {
                                switch (_a.label) {
                                    case 0:
                                        table = context.workbook.tables.getItem(layer.name);
                                        geotype = table.columns.getItem("Geometry Type");
                                        geometry = table.columns.getItem("Geometry");
                                        style = table.columns.getItem("Style");
                                        geotype.load("values");
                                        geometry.load("values");
                                        style.load("values");
                                        return [4 /*yield*/, context.sync()];
                                    case 1:
                                        _a.sent();
                                        geotypes = geotype.values.slice(1);
                                        geodatas = geometry.values.slice(1);
                                        styles = style.values.slice(1);
                                        geometries = [];
                                        for (i = 0; i < geotype.length; i++) {
                                            geodata = XLGIS.projections.project(layer.projection, JSON.parse(geodatas[i]));
                                            style_1 = JSON.parse(styles[i]);
                                            geotype_1 = geotypes[i];
                                            geopart = XLGIS.layers.getLayerPart(geotype_1, geodata, style_1);
                                            if (geopart.__proto__.name != "Error") {
                                                geometries.push(geopart);
                                            }
                                            else {
                                                window.XLGIS.errors.raise(new Error("Cannot create geopart with args: " + JSON.stringify(geodata)), [
                                                    "XLGIS.layers",
                                                    "geotype:" + geotype_1,
                                                    "layer.name:" + layer.name,
                                                    "layer.type:" + layer.type,
                                                    "layer.projection:" + layer.projection,
                                                    "layer.displayName:" + layer.displayName
                                                ]);
                                            }
                                            ;
                                        }
                                        ;
                                        return [2 /*return*/, geometries];
                                }
                            });
                        });
                    })];
            }
            else if (layer.type == "range") {
                return [2 /*return*/, Excel.run(function (context) {
                        return __awaiter(this, void 0, void 0, function () {
                            var name, e_1, sheet, matches, sheetName, rng, headers, i, header, geometries, i, geodata, style, geotype, geopart;
                            return __generator(this, function (_a) {
                                switch (_a.label) {
                                    case 0:
                                        _a.trys.push([0, 2, , 3]);
                                        name = context.workbook.names.getItem(layer.name);
                                        name.load("formula");
                                        return [4 /*yield*/, context.sync()];
                                    case 1:
                                        _a.sent();
                                        address = name.formula;
                                        return [3 /*break*/, 3];
                                    case 2:
                                        e_1 = _a.sent();
                                        address = "=" + name;
                                        return [3 /*break*/, 3];
                                    case 3:
                                        ;
                                        matches = /=(.+)\!/.exec(address);
                                        if (matches) {
                                            sheetName = matches[1];
                                            sheet = context.workbook.worksheets.getItem(sheetName);
                                        }
                                        else {
                                            sheet = context.workbook.worksheets.getActiveWorksheet();
                                        }
                                        ;
                                        rng = sheet.getRange(address);
                                        rng.load("values");
                                        return [4 /*yield*/, context.sync()];
                                    case 4:
                                        _a.sent();
                                        headers = {};
                                        for (i = 0; i < rng.values[0].length; i++) {
                                            header = rng.values[0][i];
                                            if (header == "Geometry Type") {
                                                headers["Geometry Type"] = i;
                                            }
                                            else if (header == "Geometry") {
                                                headers["Geometry"] = i;
                                            }
                                            else if (header == "Style") {
                                                headers["Style"] = i;
                                            }
                                            ;
                                        }
                                        ;
                                        if (!headers["Geometry Type"] && !headers["Geometry"] && !headers["Style"]) {
                                            window.XLGIS.errors.raise(new Error("Cannot find geometry headers."), [
                                                "XLGIS.layers"
                                            ]);
                                        }
                                        geometries = [];
                                        for (i = 1; i < rng.values.length; i++) {
                                            geodata = XLGIS.projections.project(layer.projection, JSON.parse(rng.values[i][headers["Geometry"]]));
                                            style = JSON.parse(rng.values[i][headers["Style"]]);
                                            geotype = rng.values[i][headers["Geometry Type"]];
                                            geopart = XLGIS.layers.getLayerPart(geotype, geodata, style);
                                            if (geopart.__proto__.name != "Error") {
                                                geometries.push(geopart);
                                            }
                                            else {
                                                window.XLGIS.errors.raise(new Error("Cannot create geopart with args: " + JSON.stringify(geodata)), [
                                                    "XLGIS.layers",
                                                    "geotype:" + geotype,
                                                    "layer.name:" + layer.name,
                                                    "layer.type:" + layer.type,
                                                    "layer.projection:" + layer.projection,
                                                    "layer.displayName:" + layer.displayName
                                                ]);
                                            }
                                            ;
                                        }
                                        ;
                                        return [2 /*return*/, geometries];
                                }
                            });
                        });
                    })];
            }
            else if (layer.type == "json") {
                data = JSON.parse(layer.name);
                geometries = [];
                data.forEach(function (geometry) {
                    var geodata = XLGIS.projections.project(layer.projection, geometry.d || geometry.data);
                    var geotype = geometry.t || geometry.type;
                    var style = geometry.s || geometry.style;
                    var geopart = XLGIS.layers.getLayerPart(geotype, geodata, style);
                    if (geopart.__proto__.name != "Error") {
                        geometries.push(geopart);
                    }
                    else {
                        window.XLGIS.errors.raise(new Error("Cannot create geopart with args: " + JSON.stringify(geodata)), [
                            "XLGIS.layers",
                            "geotype:" + geotype,
                            "layer.name:" + layer.name,
                            "layer.type:" + layer.type,
                            "layer.projection:" + layer.projection,
                            "layer.displayName:" + layer.displayName
                        ]);
                    }
                    ;
                });
                return [2 /*return*/, geometries];
            }
            else {
                window.XLGIS.errors.raise(new Error("Unknown layer type '" + layer.type + "'."), [
                    "XLGIS.layers",
                    "layer.name:" + layer.name,
                    "layer.type:" + layer.type,
                    "layer.projection:" + layer.projection,
                    "layer.displayName:" + layer.displayName
                ]);
            }
            ;
            return [2 /*return*/];
        });
    });
};
//Proj4 Wrapper for dealing with projections:
XLGIS.projections = {};
var projections = XLGIS.projections;
projections.listeners = {
    "add": []
};
projections.data = {
    Earth: '+proj=longlat +datum=WGS84 +no_defs '
};
projections.add = function (name, projection) {
    //Save data
    projections.data[name] = projection;
    //Call listener
    projections.listeners["add"].forEach(function (listener) {
        listener(name, projection);
    });
};
projections.on = function (eventID, func) {
    if (!projections.listeners[eventID])
        projections.listeners[eventID] = [];
    projections.listeners[eventID].push(func);
};
projections.project = function (srcProjection, point) {
    var EarthProjection = new Proj4.Proj(projection.data["Earth"]);
    return proj4(new Proj4.Proj(srcProjection), EarthProjection, point);
};
XLGIS.forms = {
    openForm: function (form) {
        if (form.id)
            document.getElementById(form.id).classList.remove("hidden");
        if (form.parent)
            document.getElementById(form.parent).classList.add("hidden");
    },
    closeForm: function (form) {
        if (form.parent)
            document.getElementById(form.parent).classList.remove("hidden");
        if (form.id)
            document.getElementById(form.id).classList.add("hidden");
    },
    Settings: {
        id: "settings-main",
        Open: function () { XLGIS.forms.openForm(this); },
        Close: function () { XLGIS.forms.closeForm(this); },
        General: {
            parent: "settings-main",
            id: "settings-general",
            Open: function () { XLGIS.forms.openForm(this); },
            Close: function () { XLGIS.forms.closeForm(this); }
        },
        Layers: {
            parent: "settings-main",
            id: "settings-layers",
            Open: function () {
                //Instantiate grids
                $("#grid-tileLayers").jsGrid({
                    autoload: true,
                    editing: true,
                    inserting: true,
                    width: "100%",
                    controller: {
                        loadData: function () {
                            var data = XLGIS._Settings.get("tileLayers");
                            data = data.map(function (el) {
                                var newEl = {};
                                newEl["Display Name"] = el.displayName;
                                newEl["Tile URL"] = el.tilePattern;
                                newEl["Attribution"] = el.attributionHTML;
                                return newEl;
                            });
                            return data;
                        },
                        insertItem: function () {
                        },
                        updateItem: function () {
                        }
                    },
                    fields: [
                        { type: "text", name: "Display Name" },
                        { type: "text", name: "Tile URL" },
                        { type: "text", name: "Attribution" },
                        { type: "control" }
                    ]
                });
                $("#grid-frontLayers").jsGrid({
                    autoload: true,
                    editing: true,
                    inserting: true,
                    width: "100%",
                    controller: {
                        loadData: function () {
                            var data = XLGIS._Settings.get("frontLayers");
                            data = data.map(function (el) {
                                var newEl = {};
                                newEl.id = el.id;
                                newEl.Data = el.name;
                                newEl.Type = el.type;
                                newEl.Projection = el.projection;
                                newEl["Display Name"] = el.displayName;
                                return newEl;
                            });
                            return data;
                        },
                        insertItem: function (item, otherArg) {
                            return XLGIS._Settings.refreshAsync(function (settings) {
                                return __awaiter(this, void 0, void 0, function () {
                                    var frontLayers, countOfDupes, newSetting;
                                    return __generator(this, function (_a) {
                                        frontLayers = settings.get("frontLayers");
                                        countOfDupes = frontLayers.filter(function (e) { return e.displayName == item["Display Name"]; }).length;
                                        if (countOfDupes > 0) {
                                            setTimeout(function () {
                                                setTimeout(function () {
                                                    $("#grid-frontLayers>.jsgrid-grid-header").notify("Cannot add 2 rows with the same display name.");
                                                });
                                                $("#grid-frontLayers").jsGrid();
                                            });
                                        }
                                        else {
                                            newSetting = {};
                                            newSetting.displayName = item["Display Name"];
                                            newSetting.name = item["Data"];
                                            newSetting.type = item["Type"];
                                            newSetting.projection = item["Projection"];
                                            frontLayers.push(newSetting);
                                            settings.set("frontLayers", frontLayers);
                                            return [2 /*return*/, settings.saveAsync(undefined, function () {
                                                    return __awaiter(this, void 0, void 0, function () {
                                                        return __generator(this, function (_a) {
                                                            return [2 /*return*/, item];
                                                        });
                                                    });
                                                })];
                                        }
                                        ;
                                        return [2 /*return*/];
                                    });
                                });
                            });
                        },
                        updateItem: function (item) {
                            debugger;
                            console.log(item);
                        }
                    },
                    fields: [
                        { type: "text", name: "Display Name" },
                        { type: "select", name: "Type", items: [
                                { Name: "TABLE", Type: "TABLE" },
                                { Name: "RANGE", Type: "RANGE" },
                                { Name: "JSON", Type: "JSON" }
                            ], valueField: "Type", textField: "Name" },
                        { type: "text", name: "Data" },
                        { type: "select", name: "Projection", items: Object.keys(XLGIS._Settings.get("projections")).map(function (key) {
                                return { Name: key, Type: key };
                            }),
                            valueField: "Type", textField: "Name" },
                        { type: "control" }
                    ]
                });
                //Show form
                XLGIS.forms.openForm(this);
            },
            Close: function () {
                //Hide form
                XLGIS.forms.closeForm(this);
                //Destroy grids
            }
        },
        Projections: {
            parent: "settings-main",
            id: "settings-projections",
            Open: function () { XLGIS.forms.openForm(this); },
            Close: function () { XLGIS.forms.closeForm(this); }
        },
        About: {
            parent: "settings-main",
            id: "settings-about",
            Open: function () { XLGIS.forms.openForm(this); },
            Close: function () { XLGIS.forms.closeForm(this); }
        }
    }
};
function callTestCase() {
    XLGIS._Settings.data.frontLayers.push({
        name: "coolLayer",
        type: "TABLE",
        projection: "Earth",
        displayName: "Cool layer"
    });
    window.setTimeout(function () {
        XLGIS.forms.Settings.Open();
        XLGIS.forms.Settings.Layers.Open();
    }, 100);
}
/*
//LAYERS GRID VIEW:

Example:
------------
XLGIS._Settings.data.frontLayers.push({
  name:"coolLayer",
  type:"table",
  projection:"Earth",
  displayName:"Cool layer"
})


*/
//# sourceMappingURL=XLGIS.js.map
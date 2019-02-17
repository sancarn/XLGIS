"use strict";

function asyncGeneratorStep(gen, resolve, reject, _next, _throw, key, arg) { try { var info = gen[key](arg); var value = info.value; } catch (error) { reject(error); return; } if (info.done) { resolve(value); } else { Promise.resolve(value).then(_next, _throw); } }

function _asyncToGenerator(fn) { return function () { var self = this, args = arguments; return new Promise(function (resolve, reject) { var gen = fn.apply(self, args); function _next(value) { asyncGeneratorStep(gen, resolve, reject, _next, _throw, "next", value); } function _throw(err) { asyncGeneratorStep(gen, resolve, reject, _next, _throw, "throw", err); } _next(undefined); }); }; }

//Leaflet extension for settings cog
L.Control.SettingsCog = L.Control.extend({
  onAdd: function onAdd(map) {
    var container = L.DomUtil.create('div');
    container.classList.add("leaflet-settingsCog"); //Basic div properties
    //container.style.backgroundColor = "ffffff";

    container.style.width = "27px";
    container.style.height = "27px";
    container.style.margin = "10px";
    container.style.padding = "3px"; //Border shadow style

    container.style.borderColor = "rgba(0,0,0,0.2)";
    container.style.borderWidth = "2px";
    container.style.borderRadius = "4px";
    container.style.borderStyle = "solid"; //Settings image icon

    container.style.backgroundImage = 'url("https://cdn2.iconfinder.com/data/icons/web/512/Cog-512.png")';
    container.style.backgroundRepeat = "no-repeat";
    container.style.backgroundSize = "23px";
    container.style.backgroundPosition = "center"; //element to append events to

    this.domElement = container; //Make callback

    var This = this;

    this.callback = function (ev) {
      L.DomEvent.stopPropagation(ev);
      This.options.handler(ev);
    }; //Register click listener


    L.DomEvent.on(this.domElement, 'click', this.callback);
    return this.domElement;
  },
  onRemove: function onRemove(map) {
    //Unregister click listener 
    L.DomEvent.off(this.domElement, 'click', this.callback);
  }
});

L.control.settingsCog = function (opts) {
  return new L.Control.SettingsCog(opts);
};

var XLGIS = {};
window.XLGIS = XLGIS;

XLGIS.test_Office = function () {
  var Office = {
    EventType: {
      SettingsChanged: "settings-changed"
    },
    context: {
      document: {
        settings: {
          get: function get(name) {
            return this.data[name];
          },
          set: function set(name, value) {
            this.data[name] = value;
          },
          addHandlerAsync: function addHandlerAsync(type, handler, options, callback) {
            if (!options) options = {};

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
            } catch (e) {
              callback({
                result: "failed",
                error: e,
                value: undefined,
                asyncContext: options["asyncContext"]
              });
            }
          },
          saveAsync: function saveAsync(options, callback) {
            if (!options) options = {};

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
            } catch (e) {
              callback({
                status: "failed",
                value: This,
                error: e,
                asyncContext: options["asyncContext"]
              });
              return undefined;
            }
          },
          refreshAsync: function refreshAsync(callback) {
            return callback(this);
          },
          remove: function remove(name) {
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
  return Office;
};

XLGIS.initialise =
/*#__PURE__*/
function () {
  var _ref = _asyncToGenerator(
  /*#__PURE__*/
  regeneratorRuntime.mark(function _callee(Office) {
    var settings, defaults;
    return regeneratorRuntime.wrap(function _callee$(_context) {
      while (1) {
        switch (_context.prev = _context.next) {
          case 0:
            if (!Office) {
              Office = XLGIS.test_Office();
            }

            XLGIS._Office = Office; //Get office settings

            settings = Office.context.document.settings;
            XLGIS._Settings = settings; //If settings don't exist, add them.

            if (settings.get("center") == null) settings.set("center", [51.505, -0.09]);
            if (settings.get("zoom") == null) settings.set("zoom", 13);

            if (settings.get("tileLayers") == null) {
              //tileLayers open street and topo maps as default
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
            if (settings.get("frontLayers") == null) settings.set("frontLayers", []);

            if (settings.get("projections") == null) {
              //earth (aka latlong) and british national grid.
              settings.set("projections", {
                "Earth": '+proj=longlat +datum=WGS84 +no_defs ',
                "British National Grid": '+proj=tmerc +lat_0=49 +lon_0=-2 +k=0.9996012717 +x_0=400000 +y_0=-100000 +ellps=airy +towgs84=446.448,-125.157,542.06,0.15,0.247,0.842,-20.489 +units=m +no_defs '
              });
            }

            ; //Save settings and then initialise map

            settings.saveAsync(undefined, function () {
              //After save
              //--------------------------------------------
              //Create leaflet map
              XLGIS.mapElement = document.getElementById("mainMap");
              XLGIS.map = L.map('mainMap'); //setView to settings or default center.

              XLGIS.map.setView(settings.get("center"), settings.get("zoom")); //Back layers

              var tileLayers = {};
              settings.get("tileLayers").forEach(function (layer) {
                tileLayers[layer.displayName] = L.tileLayer(layer.tilePattern, {
                  attribution: layer.attributionHTML
                });
              }); //Front layers

              var frontLayers = {};
              settings.get("frontLayers").forEach(function (layer) {
                frontLayers[layer.displayName] = L.layerGroup(XLGIS.layers.getLayer(layer));
              }); //Add to map:

              L.control.layers(tileLayers, frontLayers).addTo(XLGIS.map); //Settings control:

              L.control.settingsCog({
                position: "bottomleft",
                handler: function handler() {
                  document.querySelector("#settings-main").classList.remove("hidden");
                }
              }).addTo(XLGIS.map);
              console.warn("Mapper initialisation finished"); //...
            }); //Add handlers:

            settings.addHandlerAsync(Office.EventType.SettingsChanged, function () {
              XLGIS.map.setView(settings.get("center")); //...
            });
            callTestCase();
            return _context.abrupt("return", true);

          case 15:
          case "end":
            return _context.stop();
        }
      }
    }, _callee, this);
  }));

  return function (_x) {
    return _ref.apply(this, arguments);
  };
}(); //Error handling and listeners.


XLGIS.errors = [];
XLGIS.errors.groups = {};
XLGIS.errors.listeners = {};

XLGIS.errors.on = function (event, func) {
  this.listeners[event] = func;
};

XLGIS.errors.raise = function (error, groups) {
  error.groups = groups;
  this.push(error);
  console.error(error);

  if (groups) {
    groups.forEach(function (group) {
      XLGIS.errors.groups[group] = error;
    });
  }

  ;
  this.listeners["raise"].forEach(function (listener) {
    return listener(error, groups);
  });
}; //Used by combobox


XLGIS.data = {};
XLGIS.data.getNamedDatabodies =
/*#__PURE__*/
_asyncToGenerator(
/*#__PURE__*/
regeneratorRuntime.mark(function _callee3() {
  return regeneratorRuntime.wrap(function _callee3$(_context3) {
    while (1) {
      switch (_context3.prev = _context3.next) {
        case 0:
          return _context3.abrupt("return", Excel.run(
          /*#__PURE__*/
          function () {
            var _ref3 = _asyncToGenerator(
            /*#__PURE__*/
            regeneratorRuntime.mark(function _callee2(context) {
              var names, tables;
              return regeneratorRuntime.wrap(function _callee2$(_context2) {
                while (1) {
                  switch (_context2.prev = _context2.next) {
                    case 0:
                      names = context.workbook.names;
                      tables = context.workbook.tables;
                      names.load("items");
                      tables.load("items");
                      _context2.next = 6;
                      return context.sync();

                    case 6:
                      return _context2.abrupt("return", names.items.map(function (namedRange) {
                        return {
                          name: namedRange.name,
                          type: "namedRange"
                        };
                      }).concat(tables.items.map(function (table) {
                        return {
                          name: table.name,
                          type: "table"
                        };
                      })));

                    case 7:
                    case "end":
                      return _context2.stop();
                  }
                }
              }, _callee2, this);
            }));

            return function (_x2) {
              return _ref3.apply(this, arguments);
            };
          }()));

        case 1:
        case "end":
          return _context3.stop();
      }
    }
  }, _callee3, this);
})); //XL Wrappers around Leaflet

XLGIS.layers = {};

XLGIS.layers.getLayerPart = function (geotype, geometry, options) {
  try {
    if (geotype.toUpperCase() == "POINT") {
      return L.CircleMarker(L.latlng(geometry), options);
    } else if (geotype.toUpperCase() == "LINE") {
      return L.polyline(geometry, options);
    } else if (geotype.toUpperCase() == "POLYGON") {
      return L.polygon(geometry, options);
    } else if (geotype.toUpperCase() == "CIRCLE") {
      return L.Circle(geometry, options);
    } else if (geotype.toUpperCase() == "RECT") {
      return L.rectangle(geometry, options);
    } else if (geotype.toUpperCase() == "MARKER") {
      return L.marker(geometry, options);
    } else if (geotype.toUpperCase() == "IMAGE") {
      return L.Marker(geometry, options);
    }
  } catch (e) {
    return e;
  }
};

XLGIS.layers.getLayer =
/*#__PURE__*/
function () {
  var _ref4 = _asyncToGenerator(
  /*#__PURE__*/
  regeneratorRuntime.mark(function _callee6(layer) {
    var data, geometries;
    return regeneratorRuntime.wrap(function _callee6$(_context6) {
      while (1) {
        switch (_context6.prev = _context6.next) {
          case 0:
            if (!(layer.type == "table")) {
              _context6.next = 4;
              break;
            }

            return _context6.abrupt("return", Excel.run(
            /*#__PURE__*/
            function () {
              var _ref5 = _asyncToGenerator(
              /*#__PURE__*/
              regeneratorRuntime.mark(function _callee4(context) {
                var table, geotype, geometry, style, geotypes, geodatas, styles, geometries, i, geodata, _style, _geotype, geopart;

                return regeneratorRuntime.wrap(function _callee4$(_context4) {
                  while (1) {
                    switch (_context4.prev = _context4.next) {
                      case 0:
                        table = context.workbook.tables.getItem(layer.name);
                        geotype = table.columns.getItem("Geometry Type");
                        geometry = table.columns.getItem("Geometry");
                        style = table.columns.getItem("Style");
                        geotype.load("values");
                        geometry.load("values");
                        style.load("values");
                        _context4.next = 9;
                        return context.sync();

                      case 9:
                        geotypes = geotype.values.slice(1);
                        geodatas = geometry.values.slice(1);
                        styles = style.values.slice(1);
                        geometries = [];

                        for (i = 0; i < geotype.length; i++) {
                          geodata = XLGIS.projections.project(layer.projection, JSON.parse(geodatas[i]));
                          _style = JSON.parse(styles[i]);
                          _geotype = geotypes[i];
                          geopart = XLGIS.layers.getLayerPart(_geotype, geodata, _style);

                          if (geopart.__proto__.name != "Error") {
                            geometries.push(geopart);
                          } else {
                            window.XLGIS.errors.raise(new Error("Cannot create geopart with args: " + JSON.stringify(geodata)), ["XLGIS.layers", "geotype:" + _geotype, "layer.name:" + layer.name, "layer.type:" + layer.type, "layer.projection:" + layer.projection, "layer.displayName:" + layer.displayName]);
                          }

                          ;
                        }

                        ;
                        return _context4.abrupt("return", geometries);

                      case 16:
                      case "end":
                        return _context4.stop();
                    }
                  }
                }, _callee4, this);
              }));

              return function (_x4) {
                return _ref5.apply(this, arguments);
              };
            }()));

          case 4:
            if (!(layer.type == "range")) {
              _context6.next = 8;
              break;
            }

            return _context6.abrupt("return", Excel.run(
            /*#__PURE__*/
            function () {
              var _ref6 = _asyncToGenerator(
              /*#__PURE__*/
              regeneratorRuntime.mark(function _callee5(context) {
                var name, sheet, matches, sheetName, rng, headers, i, header, geometries, geodata, style, geotype, geopart;
                return regeneratorRuntime.wrap(function _callee5$(_context5) {
                  while (1) {
                    switch (_context5.prev = _context5.next) {
                      case 0:
                        _context5.prev = 0;
                        name = context.workbook.names.getItem(layer.name);
                        name.load("formula");
                        _context5.next = 5;
                        return context.sync();

                      case 5:
                        address = name.formula;
                        _context5.next = 11;
                        break;

                      case 8:
                        _context5.prev = 8;
                        _context5.t0 = _context5["catch"](0);
                        address = "=" + name;

                      case 11:
                        ; //Get sheet:

                        matches = /=(.+)\!/.exec(address);

                        if (matches) {
                          sheetName = matches[1];
                          sheet = context.workbook.worksheets.getItem(sheetName);
                        } else {
                          sheet = context.workbook.worksheets.getActiveWorksheet();
                        }

                        ; //Get the range and load it's values

                        rng = sheet.getRange(address);
                        rng.load("values");
                        _context5.next = 19;
                        return context.sync();

                      case 19:
                        //Get headers
                        headers = {};

                        for (i = 0; i < rng.values[0].length; i++) {
                          header = rng.values[0][i];

                          if (header == "Geometry Type") {
                            headers["Geometry Type"] = i;
                          } else if (header == "Geometry") {
                            headers["Geometry"] = i;
                          } else if (header == "Style") {
                            headers["Style"] = i;
                          }

                          ;
                        }

                        ;

                        if (!headers["Geometry Type"] && !headers["Geometry"] && !headers["Style"]) {
                          window.XLGIS.errors.raise(new Error("Cannot find geometry headers."), ["XLGIS.layers"]);
                        }

                        geometries = [];

                        for (i = 1; i < rng.values.length; i++) {
                          geodata = XLGIS.projections.project(layer.projection, JSON.parse(rng.values[i][headers["Geometry"]]));
                          style = JSON.parse(rng.values[i][headers["Style"]]);
                          geotype = rng.values[i][headers["Geometry Type"]];
                          geopart = XLGIS.layers.getLayerPart(geotype, geodata, style);

                          if (geopart.__proto__.name != "Error") {
                            geometries.push(geopart);
                          } else {
                            window.XLGIS.errors.raise(new Error("Cannot create geopart with args: " + JSON.stringify(geodata)), ["XLGIS.layers", "geotype:" + geotype, "layer.name:" + layer.name, "layer.type:" + layer.type, "layer.projection:" + layer.projection, "layer.displayName:" + layer.displayName]);
                          }

                          ;
                        }

                        ;
                        return _context5.abrupt("return", geometries);

                      case 27:
                      case "end":
                        return _context5.stop();
                    }
                  }
                }, _callee5, this, [[0, 8]]);
              }));

              return function (_x5) {
                return _ref6.apply(this, arguments);
              };
            }()));

          case 8:
            if (!(layer.type == "json")) {
              _context6.next = 15;
              break;
            }

            //layer.type        --> "json"
            //layer.name        --> [{t:"Point",d:[50,3],s:{}},{t:"Point",d:[49,4],s:{}},...]
            //layer.displayName --> "CoolLayer"
            //layer.projection  --> "Earth"
            data = JSON.parse(layer.name);
            geometries = [];
            data.forEach(function (geometry) {
              var geodata = XLGIS.projections.project(layer.projection, geometry.d || geometry.data);
              var geotype = geometry.t || geometry.type;
              var style = geometry.s || geometry.style;
              var geopart = XLGIS.layers.getLayerPart(geotype, geodata, style);

              if (geopart.__proto__.name != "Error") {
                geometries.push(geopart);
              } else {
                window.XLGIS.errors.raise(new Error("Cannot create geopart with args: " + JSON.stringify(geodata)), ["XLGIS.layers", "geotype:" + geotype, "layer.name:" + layer.name, "layer.type:" + layer.type, "layer.projection:" + layer.projection, "layer.displayName:" + layer.displayName]);
              }

              ;
            });
            return _context6.abrupt("return", geometries);

          case 15:
            window.XLGIS.errors.raise(new Error("Unknown layer type '" + layer.type + "'."), ["XLGIS.layers", "layer.name:" + layer.name, "layer.type:" + layer.type, "layer.projection:" + layer.projection, "layer.displayName:" + layer.displayName]);

          case 16:
            ;

          case 17:
          case "end":
            return _context6.stop();
        }
      }
    }, _callee6, this);
  }));

  return function (_x3) {
    return _ref4.apply(this, arguments);
  };
}(); //Proj4 Wrapper for dealing with projections:


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
  projections.data[name] = projection; //Call listener

  projections.listeners["add"].forEach(function (listener) {
    listener(name, projection);
  });
};

projections.on = function (eventID, func) {
  if (!projections.listeners[eventID]) projections.listeners[eventID] = [];
  projections.listeners[eventID].push(func);
};

projections.project = function (srcProjection, point) {
  var EarthProjection = new Proj4.Proj(projection.data["Earth"]);
  return proj4(new Proj4.Proj(srcProjection), EarthProjection, point);
};

XLGIS.forms = {
  openForm: function openForm(form) {
    if (form.id) document.getElementById(form.id).classList.remove("hidden");
    if (form.parent) document.getElementById(form.parent).classList.add("hidden");
  },
  closeForm: function closeForm(form) {
    if (form.parent) document.getElementById(form.parent).classList.remove("hidden");
    if (form.id) document.getElementById(form.id).classList.add("hidden");
  },
  Settings: {
    id: "settings-main",
    Open: function Open() {
      XLGIS.forms.openForm(this);
    },
    Close: function Close() {
      XLGIS.forms.closeForm(this);
    },
    General: {
      parent: "settings-main",
      id: "settings-general",
      Open: function Open() {
        XLGIS.forms.openForm(this);
      },
      Close: function Close() {
        XLGIS.forms.closeForm(this);
      }
    },
    Layers: {
      parent: "settings-main",
      id: "settings-layers",
      Open: function Open() {
        //Instantiate grids
        $("#grid-tileLayers").jsGrid({
          autoload: true,
          editing: true,
          inserting: true,
          width: "100%",
          controller: {
            loadData: function loadData() {
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
            insertItem: function insertItem() {},
            updateItem: function updateItem() {}
          },
          fields: [{
            type: "text",
            name: "Display Name"
          }, {
            type: "text",
            name: "Tile URL"
          }, {
            type: "text",
            name: "Attribution"
          }, {
            type: "control"
          }]
        });
        $("#grid-frontLayers").jsGrid({
          autoload: true,
          editing: true,
          inserting: true,
          width: "100%",
          controller: {
            loadData: function loadData() {
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
            insertItem: function insertItem(item, otherArg) {
              return XLGIS._Settings.refreshAsync(
              /*#__PURE__*/
              function () {
                var _ref7 = _asyncToGenerator(
                /*#__PURE__*/
                regeneratorRuntime.mark(function _callee8(settings) {
                  var frontLayers, countOfDupes, newSetting;
                  return regeneratorRuntime.wrap(function _callee8$(_context8) {
                    while (1) {
                      switch (_context8.prev = _context8.next) {
                        case 0:
                          frontLayers = settings.get("frontLayers");
                          countOfDupes = frontLayers.filter(function (e) {
                            return e.displayName == item["Display Name"];
                          }).length;

                          if (!(countOfDupes > 0)) {
                            _context8.next = 6;
                            break;
                          }

                          setTimeout(function () {
                            setTimeout(function () {
                              $("#grid-frontLayers>.jsgrid-grid-header").notify("Cannot add 2 rows with the same display name.");
                            });
                            $("#grid-frontLayers").jsGrid();
                          });
                          _context8.next = 14;
                          break;

                        case 6:
                          newSetting = {};
                          newSetting.displayName = item["Display Name"];
                          newSetting.name = item["Data"];
                          newSetting.type = item["Type"];
                          newSetting.projection = item["Projection"];
                          frontLayers.push(newSetting);
                          settings.set("frontLayers", frontLayers);
                          return _context8.abrupt("return", settings.saveAsync(undefined,
                          /*#__PURE__*/
                          _asyncToGenerator(
                          /*#__PURE__*/
                          regeneratorRuntime.mark(function _callee7() {
                            return regeneratorRuntime.wrap(function _callee7$(_context7) {
                              while (1) {
                                switch (_context7.prev = _context7.next) {
                                  case 0:
                                    return _context7.abrupt("return", item);

                                  case 1:
                                  case "end":
                                    return _context7.stop();
                                }
                              }
                            }, _callee7, this);
                          }))));

                        case 14:
                          ;

                        case 15:
                        case "end":
                          return _context8.stop();
                      }
                    }
                  }, _callee8, this);
                }));

                return function (_x6) {
                  return _ref7.apply(this, arguments);
                };
              }());
            },
            updateItem: function updateItem(item) {
              debugger;
              console.log(item);
            }
          },
          fields: [{
            type: "text",
            name: "Display Name"
          }, {
            type: "select",
            name: "Type",
            items: [{
              Name: "TABLE",
              Type: "TABLE"
            }, {
              Name: "RANGE",
              Type: "RANGE"
            }, {
              Name: "JSON",
              Type: "JSON"
            }],
            valueField: "Type",
            textField: "Name"
          }, {
            type: "text",
            name: "Data"
          }, {
            type: "select",
            name: "Projection",
            items: Object.keys(XLGIS._Settings.get("projections")).map(function (key) {
              return {
                Name: key,
                Type: key
              };
            }),
            valueField: "Type",
            textField: "Name"
          }, {
            type: "control"
          }]
        }); //Show form

        XLGIS.forms.openForm(this);
      },
      Close: function Close() {
        //Hide form
        XLGIS.forms.closeForm(this); //Destroy grids
      }
    },
    Projections: {
      parent: "settings-main",
      id: "settings-projections",
      Open: function Open() {
        XLGIS.forms.openForm(this);
      },
      Close: function Close() {
        XLGIS.forms.closeForm(this);
      }
    },
    About: {
      parent: "settings-main",
      id: "settings-about",
      Open: function Open() {
        XLGIS.forms.openForm(this);
      },
      Close: function Close() {
        XLGIS.forms.closeForm(this);
      }
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


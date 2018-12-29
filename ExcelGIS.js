//Leaflet extension for settings cog
L.Control.SettingsCog = L.Control.extend({
  onAdd: function(map) {
    var container = L.DomUtil.create('div');
    container.classList.add("leaflet-settingsCog")
    //Basic div properties
    //container.style.backgroundColor = "ffffff";
    container.style.width = "27px";
    container.style.height = "27px";
    container.style.margin = "10px";
    container.style.padding = "3px";
    
    //Border shadow style
    container.style.borderColor="rgba(0,0,0,0.2)";
    container.style.borderWidth="2px";
    container.style.borderRadius="4px";
    container.style.borderStyle="solid";

    //Settings image icon
    container.style.backgroundImage='url("https://cdn2.iconfinder.com/data/icons/web/512/Cog-512.png")';
    container.style.backgroundRepeat="no-repeat";
    container.style.backgroundSize = "23px";
    container.style.backgroundPosition="center";
    console.log(container.style)
    
    

    //element to append events to
    this.domElement = container;

    //Make callback
    var This = this;
    this.callback = function(ev){
      L.DomEvent.stopPropagation(ev);
      This.options.handler(ev);
    };
    
    //Register click listener
    L.DomEvent.on(this.domElement,'click',this.callback);

    return this.domElement;
  },
  onRemove: function(map) {
      //Unregister click listener 
      L.DomEvent.off(this.domElement,'click',this.callback);
  }
});

L.control.settingsCog = function(opts) {
  return new L.Control.SettingsCog(opts);
}

var XLGIS = {}
window.XLGIS = XLGIS;
XLGIS.test_Office = function(){
  var Office =  {
    EventType:{
      SettingsChanged:"settings-changed"
    },
    context:{
      document:{
        settings:{
          get:function(name){
            return this.data[name];
          },
          set: function(name,value){
            this.data[name]=value;
          },
          addHandlerAsync:function(type,handler){
            if(type=="settings-changed"){
              this._private_data.handlers.push(handler);
            }
          },
          saveAsync:function(optons,callback){
            var This=this;
            this._private_data.handlers.forEach(function(handler){
              handler({
                settings:This,
                type:"settings-changed"
              });
            });
            callback()
          },
          refreshAsync:function(){},
          data: {},
          _private_data: {
            handlers:[]
          }
        }
      }
    }
  }
  return Office;
}

XLGIS.initialise = async function(Office){
  if(!Office){
    var Office = XLGIS.test_Office()
  }
  XLGIS._Office = Office;
  
  //Get office settings
  let settings = Office.context.document.settings;

  //If settings don't exist, add them.
  if(settings.get("center") == null) settings.set("center",[51.505, -0.09]);
  if(settings.get("zoom")   == null) settings.set("zoom",13);
  if(settings.get("tileLayers")==null){ //tileLayers open street and topo maps as default
    var defaults = [];
    defaults.push({
      displayName:"Open street map",
      tilePattern:'https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',
      attributionHTML:'&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors'
    });
    defaults.push({
      displayName:"Open topo map",
      tilePattern:'https://{s}.tile.opentopomap.org/{z}/{x}/{y}.png',
      attributionHTML: 'Map data: &copy; <a href="http://www.openstreetmap.org/copyright">OpenStreetMap</a>, <a href="http://viewfinderpanoramas.org">SRTM</a> | Map style: &copy; <a href="https://opentopomap.org">OpenTopoMap</a> (<a href="https://creativecommons.org/licenses/by-sa/3.0/">CC-BY-SA</a>)'
    });
    settings.set("tileLayers",defaults);
  };
  if(settings.get("frontLayers")==null) settings.set("frontLayers",[]);
  if(settings.get("projections") == null){ //earth (aka latlong) and british national grid.
    settings.set("projections",{
      "Earth"                 : '+proj=longlat +datum=WGS84 +no_defs ',
      "British National Grid" : '+proj=tmerc +lat_0=49 +lon_0=-2 +k=0.9996012717 +x_0=400000 +y_0=-100000 +ellps=airy +towgs84=446.448,-125.157,542.06,0.15,0.247,0.842,-20.489 +units=m +no_defs '
    });
  };

  //Save settings and then initialise map
  settings.saveAsync(undefined,function(){
    //After save
    //--------------------------------------------
    //Create leaflet map
    XLGIS.mapElement = document.getElementById("mainMap");
    XLGIS.map = L.map('mainMap');

    //setView to settings or default center.
    XLGIS.map.setView(settings.get("center"), settings.get("zoom"));

    //Back layers
    var tileLayers = {};
    settings.get("tileLayers").forEach(function(layer){
      tileLayers[layer.displayName] = L.tileLayer(layer.tilePattern, {attribution: layer.attributionHTML});
    });

    //Front layers
    var frontLayers = {};
    settings.get("frontLayers").forEach(function(layer){
      frontLayers[layer.displayName] = L.layerGroup(XLGIS.layers.getLayer(layer));
    });

    //Add to map:
    L.control.layers(tileLayers, frontLayers).addTo(XLGIS.map);

    //Settings control:
    L.control.settingsCog({
      position:"bottomleft",
      handler:function(){
        document.querySelector("#settings-main").classList.remove("hidden")
      }
    }).addTo(XLGIS.map)

    console.warn("Mapper initialisation finished")
    //...
  });
  
  //Add handlers:
  settings.addHandlerAsync(Office.EventType.SettingsChanged,function(){
    XLGIS.map.setView(settings.get("center"));
    //...
  });
  
  
  return true;
}

//Error handling and listeners.
XLGIS.errors = [];
XLGIS.errors.groups = {};
XLGIS.errors.listeners = {};
XLGIS.errors.on = function(event,func){this.listeners[event]=func;}
XLGIS.errors.raise = function(error,groups){
  error.groups = groups;
  this.push(error);
  console.error(error);
  if(groups){
    groups.forEach(function(group){
      XLGIS.errors.groups[group]=error;
    });
  };
  this.listeners["raise"].forEach(listener=>listener(error,groups));
};

//Used by combobox
XLGIS.data = {}
XLGIS.data.getNamedDatabodies = async function () {
  return Excel.run(async function (context) {
      var names = context.workbook.names;
      var tables = context.workbook.tables;
      names.load("items");
      tables.load("items");
      await context.sync();
      return names.items.map(function (namedRange) {
          return { name: namedRange.name, type: "namedRange" }
      }).concat(tables.items.map(function (table) {
          return { name: table.name, type: "table" }
      }));
  });
};

//XL Wrappers around Leaflet
XLGIS.layers = {}
XLGIS.layers.getLayerPart = function(geotype,geometry,options){
  try {
    if(geotype.toUpperCase()=="POINT"){
      return L.CircleMarker(L.latlng(geometry),options);
    }else if(geotype.toUpperCase()=="LINE"){
      return L.polyline(geometry, options);
    }else if(geotype.toUpperCase()=="POLYGON"){
      return L.polygon(geometry, options);
    }else if(geotype.toUpperCase()=="CIRCLE"){
      return L.Circle(geometry, options);
    }else if(geotype.toUpperCase()=="RECT"){
      return L.rectangle(geometry, options);
    }else if(geotype.toUpperCase()=="MARKER"){
      return L.marker(geometry, options);
    }else if(geotype.toUpperCase()=="IMAGE"){
      return L.Marker(geometry, options);
    }
  } catch(e){
    return e;
  }
};
XLGIS.layers.getLayer = async function(layer){
  /*
  {
    displayName:"",
    type: "table/range",
    name: "Asdf",
    projection: "projectionName"
  }
  */
 //By Default have a click handler on points which writes the ID of the point into the range named "<<layerName>>_click"
  if(layer.type=="table"){
    return Excel.run(async function(context){
      var table = context.workbook.tables.getItem(layer.name);
      var geotype = table.columns.getItem("Geometry Type");
      var geometry = table.columns.getItem("Geometry");
      var style = table.columns.getItem("Style");

      geotype.load("values");
      geometry.load("values");
      style.load("values");

      await context.sync();

      let geotypes = geotype.values.slice(1);
      let geodatas = geometry.values.slice(1);
      let styles   = style.values.slice(1);
      
      let geometries = [];
      for(var i=0;i<geotype.length;i++){
        let geodata = XLGIS.projections.project(layer.projection,JSON.parse(geodatas[i]));
        let style   = JSON.parse(styles[i]);
        let geotype = geotypes[i];
        let geopart = XLGIS.layers.getLayerPart(geotype, geodata,style);
        if(geopart.__proto__.name!="Error"){
          geometries.push(geopart);
        } else {
          window.XLGIS.errors.raise(new Error("Cannot create geopart with args: " + JSON.stringify(geodata)),[
            "XLGIS.layers",
            "geotype:"+geotype,
            "layer.name:"+layer.name,
            "layer.type:" + layer.type,
            "layer.projection:" + layer.projection,
            "layer.displayName:" + layer.displayName
          ]);
        };
      };
      return geometries;
    });
  } else if(layer.type=="range"){
    return Excel.run(async function(context){
      //Try to get name
      try {
        var name = context.workbook.names.getItem(layer.name);
        name.load("formula");
        await context.sync();
        address=name.formula;
      } catch(e){
        address="=" + name;
      };
      
      //Get sheet:
      var sheet;
      var matches = /=(.+)\!/.exec(address);
      if(matches){
        var sheetName = matches[1]
        sheet = context.workbook.worksheets.getItem(sheetName);
      }else{
        sheet = context.workbook.worksheets.getActiveWorksheet();
      };
      
      //Get the range and load it's values
      var rng = sheet.getRange(address);
      rng.load("values");
      await context.sync();
      
      //Get headers
      var headers = {};
      for(var i=0;i<rng.values[0].length;i++){
        var header=rng.values[0][i];
        if(header=="Geometry Type"){
          headers["Geometry Type"]=i;
        }else if(header=="Geometry"){
          headers["Geometry"]=i;
        }else if(header=="Style"){
          headers["Style"]=i;
        };
      };
      if(!headers["Geometry Type"] && !headers["Geometry"] && !headers["Style"]){
        window.XLGIS.errors.raise(new Error("Cannot find geometry headers."),[
          "XLGIS.layers"
        ]);
      }
      let geometries = [];
      for(var i=1;i<rng.values.length;i++){
        let geodata = XLGIS.projections.project(layer.projection,JSON.parse(rng.values[i][headers["Geometry"]]));
        let style   = JSON.parse(rng.values[i][headers["Style"]]);
        let geotype = rng.values[i][headers["Geometry Type"]];
        let geopart = XLGIS.layers.getLayerPart(geotype, geodata,style)
        if(geopart.__proto__.name!="Error"){
          geometries.push(geopart);
        } else {
          window.XLGIS.errors.raise(new Error("Cannot create geopart with args: " + JSON.stringify(geodata)),[
            "XLGIS.layers",
            "geotype:"+geotype,
            "layer.name:"+layer.name,
            "layer.type:" + layer.type,
            "layer.projection:" + layer.projection,
            "layer.displayName:" + layer.displayName
          ]);
        };
      };
      return geometries;
    })
  } else if(layer.type=="json"){
    //layer.type    --> "json"
    //layer.name    --> [{t:"Point",d:[50,3],s:{}},{t:"Point",d:[49,4],s:{}},...]
    //layer.displayName --> "CoolLayer"
    //layer.projection  --> "Earth"
    //layer.name ,,,? Where does this come from?
    let data = JSON.parse(layer.name);
    var geometries=[]
    data.forEach(function(geometry){
      let geodata = XLGIS.projections.project(layer.projection,geometry.d||geometry.data);
      let geotype = geometry.t || geometry.type;
      let style = geometry.s || geometry.style;
      let geopart = XLGIS.layers.getLayerPart(geotype, geodata,style)
      if(geopart.__proto__.name!="Error"){
        geometries.push(geopart);
      } else {
        window.XLGIS.errors.raise(new Error("Cannot create geopart with args: " + JSON.stringify(geodata)),[
          "XLGIS.layers",
          "geotype:"+geotype,
          "layer.name:"+layer.name,
          "layer.type:" + layer.type,
          "layer.projection:" + layer.projection,
          "layer.displayName:" + layer.displayName
        ]);
      };
    });
    return geometries;
  } else {
    window.XLGIS.errors.raise(new Error("Unknown layer type '" + layer.type +  "'."),[
      "XLGIS.layers",
      "layer.name:"+layer.name,
      "layer.type:" + layer.type,
      "layer.projection:" + layer.projection,
      "layer.displayName:" + layer.displayName
    ]);
  };
};

//Proj4 Wrapper for dealing with projections:
XLGIS.projections = {};
var projections = XLGIS.projections
projections.listeners = {
  "add":[]
};
projections.data = {
  Earth: '+proj=longlat +datum=WGS84 +no_defs '
};
projections.add = function(name,projection){
  //Save data
  projections.data[name]=projection;
  
  //Call listener
  projections.listeners["add"].forEach(function(listener){
    listener(name,projection);
  });
};
projections.on = function(eventID,func){
  if(!projections.listeners[eventID]) projections.listeners[eventID] = [];
  projections.listeners[eventID].push(func);
};
projections.project = function(srcProjection,point){
  let EarthProjection = projection.data["Earth"];
  return proj4(srcProjection,EarthProjection,point);
};
import OfficeEmulator from "./libs/OfficeEmulator"
import * as L from 'leaflet';
import * as Proj4 from 'proj4'


interface ILayer {
  displayName:string,
  tilePattern:string,
  attributionHTML: string
}
interface IGenericObject {
  [key:string]: any
}
interface IGenericListeners {
  [key:string]:Function[]
}
interface XLGIS_ctorOpts {
  tests:Function[]
}



//Leaflet extension for settings cog
var LeafletHelpers = {
  SettingsCog: L.Control.extend({
    onAdd: function(/* map */) {
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
      

      //element to append events to
      this.domElement = container;

      //Make callback
      var This = this;
      this.callback = function(ev : any){
        L.DomEvent.stopPropagation(ev);
        This.options.handler(ev);
      };
      
      //Register click listener
      L.DomEvent.on(this.domElement,'click',this.callback);

      return this.domElement;
    },
    onRemove: function(/* map */) {
        //Unregister click listener 
        L.DomEvent.off(this.domElement,'click',this.callback);
    }
  }),

  settingsCog: function(opts:any) {
    return new this.SettingsCog(opts);
  }
}


class XLGIS {
  public _Office : any;
  public _Settings : any;
  public map : L.Map;
  public mapElement : HTMLElement;
  public errors : XLGIS_Errors = new XLGIS_Errors(this);
  public data : XLGIS_Data = new XLGIS_Data(this);
  public projections : XLGIS_Projections = new XLGIS_Projections(this);
  public layers : XLGIS_Layers = new XLGIS_Layers(this);

  public static Errors : XLGIS_Errors;
  public static Data : XLGIS_Data;
  public static Projections : XLGIS_Projections;
  public static Layers : XLGIS_Layers;

  constructor(opts : XLGIS_ctorOpts){
    if(!Office) var Office=OfficeEmulator;
    
    //Save for debugging:
    this._Office = Office;

    //Get office settings:
    let settings = Office.context.document.settings;
    this._Settings = settings;
    
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
      this.mapElement = document.getElementById("mainMap");
      this.map = L.map('mainMap');

      //setView to settings or default center.
      this.map.setView(settings.get("center"), settings.get("zoom"));

      //Back layers
      var tileLayers : IGenericObject = {};
      settings.get("tileLayers").forEach(function(layer : ILayer){
        tileLayers[layer.displayName] = L.tileLayer(layer.tilePattern, {attribution: layer.attributionHTML});
      });

      //Front layers
      var frontLayers : IGenericObject = {};
      settings.get("frontLayers").forEach(function(layer:ILayer){
        frontLayers[layer.displayName] = L.layerGroup(this.layers.getLayer(layer));
      });

      //Add to map:
      L.control.layers(tileLayers, frontLayers).addTo(this.map);

      //Settings control:
      LeafletHelpers.settingsCog({
        position:"bottomleft",
        handler:function(){
          document.querySelector("#settings-main").classList.remove("hidden")
        }
      }).addTo(this.map)

      console.info("Mapper initialisation finished")
    });

    //Add handlers:
    settings.addHandlerAsync(Office.EventType.SettingsChanged,function(){
      this.map.setView(settings.get("center"));
    },null,null);

    

    //Call optional tests:
    opts.tests.forEach(function(func){
      func();
    })
  }
}



type IXLGIS_Error = Error & {groups:string[]}
interface IXLGISError_Groups {
  [key:string]:IXLGIS_Error[]
}

class XLGIS_Errors {
  public value : IXLGIS_Error[] = []
  public parent : XLGIS
  private groups : IXLGISError_Groups = {}
  private listeners : IGenericListeners = {}
  
  //Set parent
  constructor(parent:XLGIS){
    this.parent = parent;
  }

  /**
   * Listen to an event from this class. This will attach to all errors raised by XLGIS object.
   * @param {string} event   - Error to listen to.
   * @param {Function} func  - Function to execute.
   */
  on(event:string,func:Function){
    if(!this.listeners[event]) this.listeners[event] = [];
    this.listeners[event].push(func);
  }
  
  /**
   * Raise an error. Use this function to raise XLGIS errors.
   * @param {IXLGIS_Error} error  - The error to raise. E.G. `new Error("hello world")`.
   * @param {string[]}     groups - A set of groups to help find the error raised.
   */
  raise(error:IXLGIS_Error,groups:string[]){
    //Set groups parameter, useful for seeing which groups the error is part of
    error.groups = groups;

    //Push error to error list
    this.value.push(error);

    //Log erro in console
    console.error(error);

    //If groups exist, push error to all groups this error is part of.
    if(groups){
      let This = this
      groups.forEach(function(group){
        if(!This.groups[group]) This.groups[group]=[];
        This.groups[group].push(error);
      });
    };
    
    //Call on-raise events
    this.listeners['raise'].forEach(listener=>listener(error,groups));
  }
}

class XLGIS_Data {
    public parent : XLGIS
    public value = {}
    constructor(parent : XLGIS){
      this.parent = parent;
    }
    async getNamedDatabodies(){
      return Excel.run(async function(context){
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
}

//Proj4 Wrapper for dealing with projections:
type XLGIS_ProjectionPoint = {x:number, y:number}
class XLGIS_Projections {
  public parent : XLGIS
  public listeners : IGenericListeners = {}
  public data : IGenericObject = {}

  constructor(parent : XLGIS){
    this.parent = parent;
  }
  add(name:string, projection:string){
    //Save data
    this.data[name]=projection;

    //Call listener
    this.listeners["add"].forEach(function(listener:Function){
      listener(name,projection);
    });
  }
  on(eventID:string,func:Function){
    if(!this.listeners[eventID]) this.listeners[eventID]=[];
    this.listeners[eventID].push(func);
  }
  project(srcProjection:string,point:XLGIS_ProjectionPoint){
    //@ts-ignore
    let EarthProjection = new Proj4.Proj(this.data["Earth"]);
    
    //@ts-ignore
    return proj4(new Proj4.Proj(srcProjection),EarthProjection,point);
  }
  revProject(finProjection:string,point:XLGIS_ProjectionPoint){
    //@ts-ignore
    let EarthProjection = new Proj4.Proj(this.data["Earth"]);
    
    //@ts-ignore
    return proj4(EarthProjection,new Proj4.Proj(finProjection),point);
  }
}







interface IXLGISLayer {
  type : "table" | "range"
  name: string //Name of table, name of range or range address
  displayName: string
  projection: string

}

enum XLGISGeoType {
  Point="POINT",
  Line="LINE",
  Polygon="POLYGON",
  Circle="CIRCLE",
  Rect="RECT",
  Marker="MARKER",
  Image="IMAGE"
} 

class XLGIS_Layers {
  public parent : XLGIS
  public value : ILayer[] = []
  constructor(parent : XLGIS){
    this.parent = parent
  }
  getLayerPart(geotype : string,geometry:any, options:any) : L.Layer{
    try {
      geotype = geotype.toUpperCase();
      switch(geotype){
        case XLGISGeoType.Point:
          return L.circleMarker(geometry,options);
        case XLGISGeoType.Line:
          return L.polyline(geometry,options);
        case XLGISGeoType.Polygon:
          return L.polygon(geometry,options);
        case XLGISGeoType.Circle:
          return L.circle(geometry,options);
        case XLGISGeoType.Rect:
          return L.rectangle(geometry,options);
        case XLGISGeoType.Marker:
          return L.marker(geometry,options);
        case XLGISGeoType.Image:
          return L.marker(geometry,options);
        default:
          return null
      }
    } catch(e){
      return e;
    }
  }
  /**
   * Creates a new layer and initialises all settings
   */
  public newLayer(layerName){
    //initialise settings
  }
  public async getLayer(layer : IXLGISLayer) {
    let This = this;
    //By Default have a click handler on points which writes the ID of the point into the range named "<<layerName>>_click"
    if(layer.type=="table"){
      let geometries = await Excel.run(async function(context){
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
        for(var i=0;i<geotypes.length;i++){
          let geodata = This.parent.projections.project(layer.projection,JSON.parse(geodatas[i]));
          let style   = JSON.parse(styles[i]);
          let geotype = geotypes[i];
          let geopart = This.parent.layers.getLayerPart(geotype, geodata,style);
          if(geopart.__proto__.name!="Error"){
            geometries.push(geopart);
          } else {
            This.parent.errors.raise(new Error("Cannot create geopart with args: " + JSON.stringify(geodata)),[
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
      return L.layerGroup(geometries)
    } else if(layer.type=="range"){
      let geometries = await Excel.run(async function(context){
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
          This.parent.errors.raise(new Error("Cannot find geometry headers."),[
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
            This.parent.errors.raise(new Error("Cannot create geopart with args: " + JSON.stringify(geodata)),[
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
      return L.layerGroup(geometries)
    } else if(layer.type=="json"){
      //layer.type        --> "json"
      //layer.name        --> [{t:"Point",d:[50,3],s:{}},{t:"Point",d:[49,4],s:{}},...]
      //layer.displayName --> "CoolLayer"
      //layer.projection  --> "Earth"
      let data = JSON.parse(layer.name);
      var geometries=[]
      data.forEach(function(geometry){
        let geodata = This.parent.projections.project(layer.projection,geometry.d||geometry.data);
        let geotype = geometry.t || geometry.type;
        let style = geometry.s || geometry.style;
        let geopart = This.parent.layers.getLayerPart(geotype, geodata,style)
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
      return L.layerGroup(geometries);
    } else {
      This.parent.errors.raise(new Error("Unknown layer type '" + layer.type +  "'."),[
        "XLGIS.layers",
        "layer.name:"+layer.name,
        "layer.type:" + layer.type,
        "layer.projection:" + layer.projection,
        "layer.displayName:" + layer.displayName
      ]);
    };
  }


}

//XL Wrappers around Leaflet
XLGIS.layers = {}
XLGIS.layers.getLayerPart = function(geotype,geometry,options){
  
};
XLGIS.layers.getLayer = async function(layer){

};



XLGIS.forms = {
  openForm:function(form){
    if(form.id    ) document.getElementById(form.id    ).classList.remove("hidden");
    if(form.parent) document.getElementById(form.parent).classList.add("hidden");
  },
  closeForm:function(form){
    if(form.parent) document.getElementById(form.parent).classList.remove("hidden");
    if(form.id    ) document.getElementById(form.id    ).classList.add("hidden");
  },
  Settings: {
    id:"settings-main",
    Open: function(){XLGIS.forms.openForm(this) },
    Close:function(){XLGIS.forms.closeForm(this)},
    General: {
      parent:"settings-main",
      id:"settings-general",
      Open: function(){XLGIS.forms.openForm(this) },
      Close:function(){XLGIS.forms.closeForm(this)},
    },
    Layers:{
      parent:"settings-main",
      id:"settings-layers",
      Open: function(){
        //Instantiate grids
        $("#grid-tileLayers").jsGrid({
          autoload:true,
          editing:true,
          inserting:true,
          width:"100%",
          controller:{
            loadData:function(){
              var data = XLGIS._Settings.get("tileLayers");
              data = data.map(function(el){
                var newEl = {};
                newEl["Display Name"] = el.displayName;
                newEl["Tile URL"]     = el.tilePattern;
                newEl["Attribution"]  = el.attributionHTML;
                return newEl;
              });
              return data;
            },
            insertItem:function(){

            },
            updateItem:function(){

            }
          },
          fields:[
            {type:"text", name:"Display Name"},
            {type:"text", name:"Tile URL"},
            {type:"text", name:"Attribution"},
            {type:"control"}
          ]

        })
        $("#grid-frontLayers").jsGrid({
          autoload:true,
          editing:true,
          inserting:true,
          width:"100%",
          controller:{
            loadData:function(){
              var data = XLGIS._Settings.get("frontLayers");
              data = data.map(function(el){
                var newEl = {};
                newEl.id = el.id;
                newEl.Data=el.name;
                newEl.Type=el.type;
                newEl.Projection=el.projection;
                newEl["Display Name"]=el.displayName;
                return newEl;
              }); 
              return data;
            },
            insertItem:function(item,otherArg){
              return XLGIS._Settings.refreshAsync(async function(settings){
                var frontLayers = settings.get("frontLayers")
                var countOfDupes = frontLayers.filter(e=>e.displayName==item["Display Name"]).length
                if(countOfDupes>0){
                  setTimeout(function(){
                    setTimeout(function(){
                      $("#grid-frontLayers>.jsgrid-grid-header").notify("Cannot add 2 rows with the same display name.");
                    });
                    $("#grid-frontLayers").jsGrid();
                  });
                } else {
                  var newSetting = {};
                  newSetting.displayName = item["Display Name"];
                  newSetting.name        = item["Data"];
                  newSetting.type        = item["Type"];
                  newSetting.projection  = item["Projection"];
                  frontLayers.push(newSetting)
                  settings.set("frontLayers",frontLayers);
                  return settings.saveAsync(undefined,async function(){
                    return item;
                  });
                };
              });
            },
            updateItem:function(item){
              debugger;
              console.log(item);
            }
          },
          fields:[
            {type:"text", name:"Display Name"},
            {type:"select", name:"Type", items:[
              {Name:"TABLE",Type:"TABLE"},
              {Name:"RANGE",Type:"RANGE"},
              {Name:"JSON",Type:"JSON"}
            ], valueField:"Type", textField:"Name"},
            {type:"text", name:"Data"},
            {type:"select", name:"Projection", items:
              Object.keys(XLGIS._Settings.get("projections")).map(function(key){
                return {Name:key,Type:key};
              }),
            valueField:"Type", textField:"Name"},
            {type:"control"}
          ]
        });

        //Show form
        XLGIS.forms.openForm(this)
      },
      Close:function(){
        //Hide form
        XLGIS.forms.closeForm(this)
        
        //Destroy grids
      }
    },
    Projections:{
      parent:"settings-main",
      id:"settings-projections",
      Open: function(){XLGIS.forms.openForm(this) },
      Close:function(){XLGIS.forms.closeForm(this)}
    },
    About:{
      parent:"settings-main",
      id:"settings-about",
      Open: function(){XLGIS.forms.openForm(this) },
      Close:function(){XLGIS.forms.closeForm(this)}
    }
  }
}

function callTestCase(){
  XLGIS._Settings.data.frontLayers.push({
    name:"coolLayer",
    type:"TABLE",
    projection:"Earth",
    displayName:"Cool layer"
  });
  window.setTimeout(function(){
    XLGIS.forms.Settings.Open();
    XLGIS.forms.Settings.Layers.Open();
  },100)
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

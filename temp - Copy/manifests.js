(function webpackUniversalModuleDefinition(root, factory) {
	if(typeof exports === 'object' && typeof module === 'object')
		module.exports = factory();
	else if(typeof define === 'function' && define.amd)
		define([], factory);
	else if(typeof exports === 'object')
		exports["debugManifests"] = factory();
	else
		root["debugManifests"] = factory();
})(window, function() {
return /******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};
/******/
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/
/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId]) {
/******/ 			return installedModules[moduleId].exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			i: moduleId,
/******/ 			l: false,
/******/ 			exports: {}
/******/ 		};
/******/
/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/ 		// Flag the module as loaded
/******/ 		module.l = true;
/******/
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/
/******/
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;
/******/
/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;
/******/
/******/ 	// define getter function for harmony exports
/******/ 	__webpack_require__.d = function(exports, name, getter) {
/******/ 		if(!__webpack_require__.o(exports, name)) {
/******/ 			Object.defineProperty(exports, name, { enumerable: true, get: getter });
/******/ 		}
/******/ 	};
/******/
/******/ 	// define __esModule on exports
/******/ 	__webpack_require__.r = function(exports) {
/******/ 		if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 			Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 		}
/******/ 		Object.defineProperty(exports, '__esModule', { value: true });
/******/ 	};
/******/
/******/ 	// create a fake namespace object
/******/ 	// mode & 1: value is a module id, require it
/******/ 	// mode & 2: merge all properties of value into the ns
/******/ 	// mode & 4: return value when already ns object
/******/ 	// mode & 8|1: behave like require
/******/ 	__webpack_require__.t = function(value, mode) {
/******/ 		if(mode & 1) value = __webpack_require__(value);
/******/ 		if(mode & 8) return value;
/******/ 		if((mode & 4) && typeof value === 'object' && value && value.__esModule) return value;
/******/ 		var ns = Object.create(null);
/******/ 		__webpack_require__.r(ns);
/******/ 		Object.defineProperty(ns, 'default', { enumerable: true, value: value });
/******/ 		if(mode & 2 && typeof value != 'string') for(var key in value) __webpack_require__.d(ns, key, function(key) { return value[key]; }.bind(null, key));
/******/ 		return ns;
/******/ 	};
/******/
/******/ 	// getDefaultExport function for compatibility with non-harmony modules
/******/ 	__webpack_require__.n = function(module) {
/******/ 		var getter = module && module.__esModule ?
/******/ 			function getDefault() { return module['default']; } :
/******/ 			function getModuleExports() { return module; };
/******/ 		__webpack_require__.d(getter, 'a', getter);
/******/ 		return getter;
/******/ 	};
/******/
/******/ 	// Object.prototype.hasOwnProperty.call
/******/ 	__webpack_require__.o = function(object, property) { return Object.prototype.hasOwnProperty.call(object, property); };
/******/
/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "";
/******/
/******/
/******/ 	// Set the webpack public path
/******/ 	(function () {
/******/ 	  var scripts = document.getElementsByTagName('script');
/******/ 	  var regex = new RegExp('manifests\\.js', 'i');
/******/ 	  var publicPath;
/******/ 	
/******/ 	  if (scripts && scripts.length) {
/******/ 	    for (var i = 0; i < scripts.length; i++) {
/******/ 	      if (!scripts[i]) continue;
/******/ 	      var path = scripts[i].getAttribute('src');
/******/ 	      if (path && path.match(regex)) {
/******/ 	        publicPath = path.substring(0, path.lastIndexOf('/') + 1);
/******/ 	        break;
/******/ 	      }
/******/ 	    }
/******/ 	  }
/******/ 	
/******/ 	  if (!publicPath) {
/******/ 	    for (var global in window.__setWebpackPublicPathLoaderSrcRegistry__) {
/******/ 	      if (global && global.match(regex)) {
/******/ 	        publicPath = global.substring(0, global.lastIndexOf('/') + 1);
/******/ 	        break;
/******/ 	      }
/******/ 	    }
/******/ 	  }
/******/ 	  __webpack_require__.p = publicPath;
/******/ 	})();
/******/ 	
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(__webpack_require__.s = 0);
/******/ })
/************************************************************************/
/******/ ([
/* 0 */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getManifests", function() { return getManifests; });
var MANIFESTS_ARRAY = [
  {
    "id": "73e1dc6c-8441-42cc-ad47-4bd3659f8a3a",
    "alias": "SPLodashSubset",
    "componentType": "Library",
    "version": "1.10.0",
    "manifestVersion": 2,
    "loaderConfig": {
      "entryModuleId": "sp-lodash-subset",
      "internalModuleBaseUrls": [
        "https://localhost:4321/node_modules/@microsoft/sp-lodash-subset/"
      ],
      "scriptResources": {
        "sp-lodash-subset": {
          "type": "path",
          "path": "dist/sp-lodash-subset.js"
        }
      }
    },
    "isInternal": true
  },
  {
    "id": "7263c7d0-1d6a-45ec-8d85-d4d1d234171b",
    "alias": "SPCoreLibrary",
    "componentType": "Library",
    "version": "1.10.0",
    "manifestVersion": 2,
    "loaderConfig": {
      "entryModuleId": "sp-core-library",
      "internalModuleBaseUrls": [
        "https://localhost:4321/node_modules/@microsoft/sp-core-library/"
      ],
      "scriptResources": {
        "sp-core-library": {
          "type": "localizedPath",
          "defaultPath": "dist/sp-core-library_en-us.js"
        },
        "@microsoft/sp-lodash-subset": {
          "type": "component",
          "version": "1.10.0",
          "id": "73e1dc6c-8441-42cc-ad47-4bd3659f8a3a"
        }
      }
    },
    "isInternal": true
  },
  {
    "id": "29bd516f-4ece-40b7-8028-597cbc65a223",
    "alias": "SpOfficeUIFabricCore",
    "componentType": "Library",
    "version": "1.10.0",
    "manifestVersion": 2,
    "loaderConfig": {
      "entryModuleId": "office-ui-fabric-core",
      "internalModuleBaseUrls": [
        "https://localhost:4321/node_modules/@microsoft/sp-office-ui-fabric-core/"
      ],
      "scriptResources": {
        "office-ui-fabric-core": {
          "type": "path",
          "path": "dist/office-ui-fabric-core.js"
        },
        "@microsoft/load-themed-styles": {
          "type": "component",
          "version": "0.1.2",
          "id": "229b8d08-79f3-438b-8c21-4613fc877abd"
        }
      }
    },
    "isInternal": true
  },
  {
    "id": "f97266fb-ccb7-430e-9384-4124d05295d3",
    "alias": "Decorators",
    "componentType": "Library",
    "version": "1.10.0",
    "manifestVersion": 2,
    "loaderConfig": {
      "entryModuleId": "decorators",
      "internalModuleBaseUrls": [
        "https://localhost:4321/node_modules/@microsoft/decorators/"
      ],
      "scriptResources": {
        "decorators": {
          "type": "path",
          "path": "dist/decorators.js"
        }
      }
    },
    "isInternal": true
  },
  {
    "id": "02a01e42-69ab-403d-8a16-acd128661f8e",
    "alias": "OfficeUIFabricReact",
    "componentType": "Library",
    "version": "1.10.0",
    "manifestVersion": 2,
    "loaderConfig": {
      "entryModuleId": "office-ui-fabric-react-bundle",
      "internalModuleBaseUrls": [
        "https://localhost:4321/node_modules/@microsoft/office-ui-fabric-react-bundle/"
      ],
      "scriptResources": {
        "office-ui-fabric-react-bundle": {
          "type": "path",
          "path": "dist/office-ui-fabric-react-bundle.js"
        },
        "react": {
          "type": "component",
          "version": "16.8.5",
          "id": "0d910c1c-13b9-4e1c-9aa4-b008c5e42d7d",
          "failoverPath": "node_modules/react/dist/react.js"
        },
        "react-dom": {
          "type": "component",
          "version": "16.8.5",
          "id": "aa0a46ec-1505-43cd-a44a-93f3a5aa460a",
          "failoverPath": "node_modules/react-dom/dist/react-dom.js"
        },
        "@ms/uifabric-styling-bundle": {
          "type": "component",
          "version": "0.1.0",
          "id": "17ce0976-e69a-4355-be84-89b69a74717d"
        },
        "@microsoft/load-themed-styles": {
          "type": "component",
          "version": "0.1.2",
          "id": "229b8d08-79f3-438b-8c21-4613fc877abd"
        }
      }
    },
    "isInternal": true
  },
  {
    "manifestVersion": 2,
    "id": "78359e4b-07c2-43c6-8d0b-d060b4d577e8",
    "alias": "SPDiagnostics",
    "componentType": "Library",
    "version": "1.10.0",
    "loaderConfig": {
      "entryModuleId": "sp-diagnostics",
      "internalModuleBaseUrls": [
        "https://localhost:4321/node_modules/@microsoft/sp-diagnostics/"
      ],
      "scriptResources": {
        "sp-diagnostics": {
          "type": "path",
          "path": "dist/sp-diagnostics.js"
        },
        "@microsoft/sp-lodash-subset": {
          "type": "component",
          "version": "1.10.0",
          "id": "73e1dc6c-8441-42cc-ad47-4bd3659f8a3a"
        },
        "@microsoft/sp-core-library": {
          "type": "component",
          "version": "1.10.0",
          "id": "7263c7d0-1d6a-45ec-8d85-d4d1d234171b"
        }
      }
    },
    "isInternal": true
  },
  {
    "manifestVersion": 2,
    "id": "e40f8203-b39d-425a-a957-714852e33b79",
    "alias": "SPDynamicData",
    "componentType": "Library",
    "version": "1.10.0",
    "loaderConfig": {
      "entryModuleId": "sp-dynamic-data",
      "internalModuleBaseUrls": [
        "https://localhost:4321/node_modules/@microsoft/sp-dynamic-data/"
      ],
      "scriptResources": {
        "sp-dynamic-data": {
          "type": "localizedPath",
          "defaultPath": "dist/sp-dynamic-data_en-us.js"
        },
        "@microsoft/sp-lodash-subset": {
          "type": "component",
          "version": "1.10.0",
          "id": "73e1dc6c-8441-42cc-ad47-4bd3659f8a3a"
        },
        "@microsoft/sp-core-library": {
          "type": "component",
          "version": "1.10.0",
          "id": "7263c7d0-1d6a-45ec-8d85-d4d1d234171b"
        },
        "@microsoft/sp-diagnostics": {
          "type": "component",
          "version": "1.10.0",
          "id": "78359e4b-07c2-43c6-8d0b-d060b4d577e8"
        }
      }
    },
    "isInternal": true
  },
  {
    "id": "1e384972-6346-49b4-93c7-b2e6763938e6",
    "alias": "sp-polyfills",
    "componentType": "Library",
    "version": "1.10.0",
    "manifestVersion": 2,
    "loaderConfig": {
      "entryModuleId": "sp-polyfills",
      "internalModuleBaseUrls": [
        "https://localhost:4321/node_modules/@microsoft/sp-polyfills/"
      ],
      "scriptResources": {
        "sp-polyfills": {
          "type": "path",
          "path": "dist/sp-polyfills.js"
        }
      }
    }
  },
  {
    "id": "c07208f0-ea3b-4c1a-9965-ac1b825211a6",
    "alias": "SPHttp",
    "componentType": "Library",
    "version": "1.10.0",
    "manifestVersion": 2,
    "preloadComponents": [],
    "loaderConfig": {
      "entryModuleId": "sp-http",
      "internalModuleBaseUrls": [
        "https://localhost:4321/node_modules/@microsoft/sp-http/"
      ],
      "scriptResources": {
        "sp-http": {
          "type": "localizedPath",
          "defaultPath": "dist/sp-http_en-us.js"
        },
        "@microsoft/sp-core-library": {
          "type": "component",
          "version": "1.10.0",
          "id": "7263c7d0-1d6a-45ec-8d85-d4d1d234171b"
        },
        "@microsoft/sp-diagnostics": {
          "type": "component",
          "version": "1.10.0",
          "id": "78359e4b-07c2-43c6-8d0b-d060b4d577e8"
        }
      }
    },
    "isInternal": true
  },
  {
    "id": "1c4541f7-5c31-41aa-9fa8-fbc9dc14c0a8",
    "alias": "SPPageContext",
    "componentType": "Library",
    "version": "1.10.0",
    "manifestVersion": 2,
    "loaderConfig": {
      "entryModuleId": "sp-page-context",
      "internalModuleBaseUrls": [
        "https://localhost:4321/node_modules/@microsoft/sp-page-context/"
      ],
      "scriptResources": {
        "sp-page-context": {
          "type": "localizedPath",
          "defaultPath": "dist/sp-page-context_en-us.js"
        },
        "@microsoft/sp-dynamic-data": {
          "type": "component",
          "version": "1.10.0",
          "id": "e40f8203-b39d-425a-a957-714852e33b79"
        },
        "@microsoft/sp-core-library": {
          "type": "component",
          "version": "1.10.0",
          "id": "7263c7d0-1d6a-45ec-8d85-d4d1d234171b"
        },
        "@microsoft/sp-diagnostics": {
          "type": "component",
          "version": "1.10.0",
          "id": "78359e4b-07c2-43c6-8d0b-d060b4d577e8"
        }
      }
    },
    "isInternal": true
  },
  {
    "id": "1c6c9123-7aac-41f3-a376-3caea41ed83f",
    "alias": "SPLoader",
    "componentType": "Library",
    "version": "1.10.0",
    "manifestVersion": 2,
    "loaderConfig": {
      "entryModuleId": "sp-loader",
      "internalModuleBaseUrls": [
        "https://localhost:4321/node_modules/@microsoft/sp-loader/"
      ],
      "scriptResources": {
        "sp-loader": {
          "type": "localizedPath",
          "defaultPath": "dist/sp-loader_en-us.js"
        },
        "@ms/sp-telemetry": {
          "type": "component",
          "version": "0.8.16",
          "id": "8217e442-8ed3-41fd-957d-b112e841286a"
        },
        "@microsoft/sp-dynamic-data": {
          "type": "component",
          "version": "1.10.0",
          "id": "e40f8203-b39d-425a-a957-714852e33b79"
        },
        "@microsoft/sp-lodash-subset": {
          "type": "component",
          "version": "1.10.0",
          "id": "73e1dc6c-8441-42cc-ad47-4bd3659f8a3a"
        },
        "@microsoft/sp-core-library": {
          "type": "component",
          "version": "1.10.0",
          "id": "7263c7d0-1d6a-45ec-8d85-d4d1d234171b"
        },
        "@microsoft/sp-page-context": {
          "type": "component",
          "version": "1.10.0",
          "id": "1c4541f7-5c31-41aa-9fa8-fbc9dc14c0a8"
        },
        "@microsoft/load-themed-styles": {
          "type": "component",
          "version": "0.1.2",
          "id": "229b8d08-79f3-438b-8c21-4613fc877abd"
        },
        "@microsoft/sp-diagnostics": {
          "type": "component",
          "version": "1.10.0",
          "id": "78359e4b-07c2-43c6-8d0b-d060b4d577e8"
        },
        "@microsoft/sp-http": {
          "type": "component",
          "version": "1.10.0",
          "id": "c07208f0-ea3b-4c1a-9965-ac1b825211a6"
        }
      }
    },
    "isInternal": true
  },
  {
    "id": "467dc675-7cc5-4709-8aac-78e3b71bd2f6",
    "alias": "SPComponentBase",
    "componentType": "Library",
    "version": "1.10.0",
    "manifestVersion": 2,
    "loaderConfig": {
      "entryModuleId": "sp-component-base",
      "internalModuleBaseUrls": [
        "https://localhost:4321/node_modules/@microsoft/sp-component-base/"
      ],
      "scriptResources": {
        "sp-component-base": {
          "type": "localizedPath",
          "defaultPath": "dist/sp-component-base_en-us.js"
        },
        "@microsoft/sp-dynamic-data": {
          "type": "component",
          "version": "1.10.0",
          "id": "e40f8203-b39d-425a-a957-714852e33b79"
        },
        "@microsoft/sp-lodash-subset": {
          "type": "component",
          "version": "1.10.0",
          "id": "73e1dc6c-8441-42cc-ad47-4bd3659f8a3a"
        },
        "@microsoft/sp-core-library": {
          "type": "component",
          "version": "1.10.0",
          "id": "7263c7d0-1d6a-45ec-8d85-d4d1d234171b"
        },
        "@microsoft/sp-page-context": {
          "type": "component",
          "version": "1.10.0",
          "id": "1c4541f7-5c31-41aa-9fa8-fbc9dc14c0a8"
        },
        "@microsoft/load-themed-styles": {
          "type": "component",
          "version": "0.1.2",
          "id": "229b8d08-79f3-438b-8c21-4613fc877abd"
        },
        "@microsoft/sp-diagnostics": {
          "type": "component",
          "version": "1.10.0",
          "id": "78359e4b-07c2-43c6-8d0b-d060b4d577e8"
        },
        "@microsoft/sp-http": {
          "type": "component",
          "version": "1.10.0",
          "id": "c07208f0-ea3b-4c1a-9965-ac1b825211a6"
        },
        "@microsoft/decorators": {
          "type": "component",
          "version": "1.10.0",
          "id": "f97266fb-ccb7-430e-9384-4124d05295d3"
        }
      }
    },
    "isInternal": true
  },
  {
    "id": "f9e737b7-f0df-4597-ba8c-3060f82380db",
    "alias": "SPPropertyPane",
    "componentType": "Library",
    "version": "1.10.0",
    "manifestVersion": 2,
    "loaderConfig": {
      "entryModuleId": "sp-property-pane",
      "internalModuleBaseUrls": [
        "https://localhost:4321/node_modules/@microsoft/sp-property-pane/"
      ],
      "scriptResources": {
        "sp-property-pane": {
          "type": "localizedPath",
          "defaultPath": "dist/sp-property-pane_en-us.js"
        },
        "@microsoft/sp-component-base": {
          "type": "component",
          "version": "1.10.0",
          "id": "467dc675-7cc5-4709-8aac-78e3b71bd2f6"
        },
        "@microsoft/office-ui-fabric-react-bundle": {
          "type": "component",
          "version": "1.10.0",
          "id": "02a01e42-69ab-403d-8a16-acd128661f8e"
        },
        "@microsoft/sp-lodash-subset": {
          "type": "component",
          "version": "1.10.0",
          "id": "73e1dc6c-8441-42cc-ad47-4bd3659f8a3a"
        },
        "@microsoft/sp-core-library": {
          "type": "component",
          "version": "1.10.0",
          "id": "7263c7d0-1d6a-45ec-8d85-d4d1d234171b"
        },
        "react": {
          "type": "component",
          "version": "16.8.5",
          "id": "0d910c1c-13b9-4e1c-9aa4-b008c5e42d7d",
          "failoverPath": "node_modules/react/dist/react.js"
        },
        "react-dom": {
          "type": "component",
          "version": "16.8.5",
          "id": "aa0a46ec-1505-43cd-a44a-93f3a5aa460a",
          "failoverPath": "node_modules/react-dom/dist/react-dom.js"
        },
        "@ms/uifabric-styling-bundle": {
          "type": "component",
          "version": "0.1.0",
          "id": "17ce0976-e69a-4355-be84-89b69a74717d"
        },
        "@microsoft/load-themed-styles": {
          "type": "component",
          "version": "0.1.2",
          "id": "229b8d08-79f3-438b-8c21-4613fc877abd"
        },
        "@microsoft/sp-diagnostics": {
          "type": "component",
          "version": "1.10.0",
          "id": "78359e4b-07c2-43c6-8d0b-d060b4d577e8"
        },
        "@microsoft/decorators": {
          "type": "component",
          "version": "1.10.0",
          "id": "f97266fb-ccb7-430e-9384-4124d05295d3"
        }
      }
    },
    "isInternal": true
  },
  {
    "id": "974a7777-0990-4136-8fa6-95d80114c2e0",
    "alias": "SPWebPartBase",
    "componentType": "Library",
    "version": "1.10.0",
    "manifestVersion": 2,
    "preloadComponents": [
      "f9e737b7-f0df-4597-ba8c-3060f82380db"
    ],
    "loaderConfig": {
      "entryModuleId": "sp-webpart-base",
      "internalModuleBaseUrls": [
        "https://localhost:4321/node_modules/@microsoft/sp-webpart-base/"
      ],
      "scriptResources": {
        "sp-webpart-base": {
          "type": "localizedPath",
          "defaultPath": "dist/sp-webpart-base_en-us.js"
        },
        "@ms/sp-telemetry": {
          "type": "component",
          "version": "0.8.16",
          "id": "8217e442-8ed3-41fd-957d-b112e841286a"
        },
        "@ms/sp-load-themed-styles": {
          "type": "component",
          "version": "0.1.2",
          "id": "229b8d08-79f3-438b-8c21-4613fc877abd"
        },
        "@microsoft/sp-component-base": {
          "type": "component",
          "version": "1.10.0",
          "id": "467dc675-7cc5-4709-8aac-78e3b71bd2f6"
        },
        "@microsoft/sp-loader": {
          "type": "component",
          "version": "1.10.0",
          "id": "1c6c9123-7aac-41f3-a376-3caea41ed83f"
        },
        "@microsoft/sp-lodash-subset": {
          "type": "component",
          "version": "1.10.0",
          "id": "73e1dc6c-8441-42cc-ad47-4bd3659f8a3a"
        },
        "@microsoft/sp-core-library": {
          "type": "component",
          "version": "1.10.0",
          "id": "7263c7d0-1d6a-45ec-8d85-d4d1d234171b"
        },
        "@microsoft/sp-page-context": {
          "type": "component",
          "version": "1.10.0",
          "id": "1c4541f7-5c31-41aa-9fa8-fbc9dc14c0a8"
        },
        "@microsoft/load-themed-styles": {
          "type": "component",
          "version": "0.1.2",
          "id": "229b8d08-79f3-438b-8c21-4613fc877abd"
        },
        "@microsoft/sp-diagnostics": {
          "type": "component",
          "version": "1.10.0",
          "id": "78359e4b-07c2-43c6-8d0b-d060b4d577e8"
        },
        "@microsoft/sp-http": {
          "type": "component",
          "version": "1.10.0",
          "id": "c07208f0-ea3b-4c1a-9965-ac1b825211a6"
        },
        "@microsoft/decorators": {
          "type": "component",
          "version": "1.10.0",
          "id": "f97266fb-ccb7-430e-9384-4124d05295d3"
        }
      }
    },
    "isInternal": true
  },
  {
    "id": "2ad06bf8-18c2-44ef-bf46-c9effb647f12",
    "alias": "iConnectFlipCardNew",
    "componentType": "WebPart",
    "version": "0.0.1",
    "manifestVersion": 2,
    "requiresCustomScript": false,
    "supportedHosts": [
      "SharePointWebPart"
    ],
    "preconfiguredEntries": [
      {
        "groupId": "16956a8e-5f15-46ce-b9ca-cbfc61584992",
        "group": {
          "default": "Other"
        },
        "title": {
          "default": "iConnect-Flip Card-New"
        },
        "description": {
          "default": "Flip Card description  2.0"
        },
        "iconImageUrl": "data:image/png;base64,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",
        "officeFabricIconFontName": "Page",
        "properties": {
          "description": "Gallery"
        }
      }
    ],
    "loaderConfig": {
      "entryModuleId": "modern-site-search-web-part",
      "internalModuleBaseUrls": [
        "https://localhost:4321/"
      ],
      "scriptResources": {
        "modern-site-search-web-part": {
          "type": "path",
          "path": "dist/modern-site-search-web-part.js"
        },
        "ControlStrings": {
          "defaultPath": "node_modules/@pnp/spfx-controls-react/lib/loc/en-us.js",
          "type": "localizedPath"
        },
        "IESModernSiteListFlipFlopWebPartStrings": {
          "defaultPath": "lib/webparts/IESModernSiteListFlipFlop/loc/en-us.js",
          "type": "localizedPath"
        },
        "PropertyControlStrings": {
          "defaultPath": "node_modules/@pnp/spfx-property-controls/lib/loc/en-us.js",
          "type": "localizedPath"
        },
        "@microsoft/sp-property-pane": {
          "type": "component",
          "version": "1.10.0",
          "id": "f9e737b7-f0df-4597-ba8c-3060f82380db"
        },
        "@microsoft/sp-lodash-subset": {
          "type": "component",
          "version": "1.10.0",
          "id": "73e1dc6c-8441-42cc-ad47-4bd3659f8a3a"
        },
        "@microsoft/sp-core-library": {
          "type": "component",
          "version": "1.10.0",
          "id": "7263c7d0-1d6a-45ec-8d85-d4d1d234171b"
        },
        "@microsoft/sp-page-context": {
          "type": "component",
          "version": "1.10.0",
          "id": "1c4541f7-5c31-41aa-9fa8-fbc9dc14c0a8"
        },
        "@microsoft/sp-webpart-base": {
          "type": "component",
          "version": "1.10.0",
          "id": "974a7777-0990-4136-8fa6-95d80114c2e0"
        },
        "react": {
          "type": "component",
          "version": "16.8.5",
          "id": "0d910c1c-13b9-4e1c-9aa4-b008c5e42d7d",
          "failoverPath": "node_modules/react/dist/react.js"
        },
        "react-dom": {
          "type": "component",
          "version": "16.8.5",
          "id": "aa0a46ec-1505-43cd-a44a-93f3a5aa460a",
          "failoverPath": "node_modules/react-dom/dist/react-dom.js"
        },
        "@microsoft/sp-http": {
          "type": "component",
          "version": "1.10.0",
          "id": "c07208f0-ea3b-4c1a-9965-ac1b825211a6"
        },
        "@microsoft/decorators": {
          "type": "component",
          "version": "1.10.0",
          "id": "f97266fb-ccb7-430e-9384-4124d05295d3"
        }
      }
    }
  }
];
/**
 * Get the manifest array.
 */
function getManifests() {
    // Clone manifestsArray
    var manifests = JSON.parse(JSON.stringify(MANIFESTS_ARRAY));
    var manifestsFileUrl = __webpack_require__.p;
    if (manifestsFileUrl && manifestsFileUrl !== '') {
        manifests.forEach(function (manifest) {
            if (!manifest.loaderConfig.internalModuleBaseUrls || manifest.loaderConfig.internalModuleBaseUrls.length === 0) {
                manifest.loaderConfig.internalModuleBaseUrls = [manifestsFileUrl];
            }
        });
    }
    else {
        console.error("Unable to determine " + "manifests.js" + " file URL. Using default base URL. " +
            'This is expected if you are running "gulp serve."');
    }
    return manifests;
}
//# sourceMappingURL=manifestsFile.js.map

/***/ })
/******/ ]);
});
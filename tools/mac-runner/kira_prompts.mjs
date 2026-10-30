var __create = Object.create;
var __defProp = Object.defineProperty;
var __getOwnPropDesc = Object.getOwnPropertyDescriptor;
var __getOwnPropNames = Object.getOwnPropertyNames;
var __getProtoOf = Object.getPrototypeOf;
var __hasOwnProp = Object.prototype.hasOwnProperty;
var __commonJS = (cb, mod) => function __require() {
  return mod || (0, cb[__getOwnPropNames(cb)[0]])((mod = { exports: {} }).exports, mod), mod.exports;
};
var __copyProps = (to, from, except, desc) => {
  if (from && typeof from === "object" || typeof from === "function") {
    for (let key of __getOwnPropNames(from))
      if (!__hasOwnProp.call(to, key) && key !== except)
        __defProp(to, key, { get: () => from[key], enumerable: !(desc = __getOwnPropDesc(from, key)) || desc.enumerable });
  }
  return to;
};
var __toESM = (mod, isNodeMode, target) => (target = mod != null ? __create(__getProtoOf(mod)) : {}, __copyProps(
  // If the importer is in node compatibility mode or this is not an ESM
  // file that has been converted to a CommonJS file using a Babel-
  // compatible transform (i.e. "__esModule" has not been set), then set
  // "default" to the CommonJS "module.exports" for node compatibility.
  isNodeMode || !mod || !mod.__esModule ? __defProp(target, "default", { value: mod, enumerable: true }) : target,
  mod
));

// ../../node_modules/react/cjs/react.production.js
var require_react_production = __commonJS({
  "../../node_modules/react/cjs/react.production.js"(exports) {
    "use strict";
    var REACT_ELEMENT_TYPE = Symbol.for("react.transitional.element");
    var REACT_PORTAL_TYPE = Symbol.for("react.portal");
    var REACT_FRAGMENT_TYPE = Symbol.for("react.fragment");
    var REACT_STRICT_MODE_TYPE = Symbol.for("react.strict_mode");
    var REACT_PROFILER_TYPE = Symbol.for("react.profiler");
    var REACT_CONSUMER_TYPE = Symbol.for("react.consumer");
    var REACT_CONTEXT_TYPE = Symbol.for("react.context");
    var REACT_FORWARD_REF_TYPE = Symbol.for("react.forward_ref");
    var REACT_SUSPENSE_TYPE = Symbol.for("react.suspense");
    var REACT_MEMO_TYPE = Symbol.for("react.memo");
    var REACT_LAZY_TYPE = Symbol.for("react.lazy");
    var REACT_ACTIVITY_TYPE = Symbol.for("react.activity");
    var MAYBE_ITERATOR_SYMBOL = Symbol.iterator;
    function getIteratorFn(maybeIterable) {
      if (null === maybeIterable || "object" !== typeof maybeIterable) return null;
      maybeIterable = MAYBE_ITERATOR_SYMBOL && maybeIterable[MAYBE_ITERATOR_SYMBOL] || maybeIterable["@@iterator"];
      return "function" === typeof maybeIterable ? maybeIterable : null;
    }
    var ReactNoopUpdateQueue = {
      isMounted: function() {
        return false;
      },
      enqueueForceUpdate: function() {
      },
      enqueueReplaceState: function() {
      },
      enqueueSetState: function() {
      }
    };
    var assign = Object.assign;
    var emptyObject = {};
    function Component(props, context, updater) {
      this.props = props;
      this.context = context;
      this.refs = emptyObject;
      this.updater = updater || ReactNoopUpdateQueue;
    }
    Component.prototype.isReactComponent = {};
    Component.prototype.setState = function(partialState, callback) {
      if ("object" !== typeof partialState && "function" !== typeof partialState && null != partialState)
        throw Error(
          "takes an object of state variables to update or a function which returns an object of state variables."
        );
      this.updater.enqueueSetState(this, partialState, callback, "setState");
    };
    Component.prototype.forceUpdate = function(callback) {
      this.updater.enqueueForceUpdate(this, callback, "forceUpdate");
    };
    function ComponentDummy() {
    }
    ComponentDummy.prototype = Component.prototype;
    function PureComponent(props, context, updater) {
      this.props = props;
      this.context = context;
      this.refs = emptyObject;
      this.updater = updater || ReactNoopUpdateQueue;
    }
    var pureComponentPrototype = PureComponent.prototype = new ComponentDummy();
    pureComponentPrototype.constructor = PureComponent;
    assign(pureComponentPrototype, Component.prototype);
    pureComponentPrototype.isPureReactComponent = true;
    var isArrayImpl = Array.isArray;
    function noop() {
    }
    var ReactSharedInternals = { H: null, A: null, T: null, S: null };
    var hasOwnProperty = Object.prototype.hasOwnProperty;
    function ReactElement(type, key, props) {
      var refProp = props.ref;
      return {
        $$typeof: REACT_ELEMENT_TYPE,
        type,
        key,
        ref: void 0 !== refProp ? refProp : null,
        props
      };
    }
    function cloneAndReplaceKey(oldElement, newKey) {
      return ReactElement(oldElement.type, newKey, oldElement.props);
    }
    function isValidElement(object) {
      return "object" === typeof object && null !== object && object.$$typeof === REACT_ELEMENT_TYPE;
    }
    function escape(key) {
      var escaperLookup = { "=": "=0", ":": "=2" };
      return "$" + key.replace(/[=:]/g, function(match) {
        return escaperLookup[match];
      });
    }
    var userProvidedKeyEscapeRegex = /\/+/g;
    function getElementKey(element, index) {
      return "object" === typeof element && null !== element && null != element.key ? escape("" + element.key) : index.toString(36);
    }
    function resolveThenable(thenable) {
      switch (thenable.status) {
        case "fulfilled":
          return thenable.value;
        case "rejected":
          throw thenable.reason;
        default:
          switch ("string" === typeof thenable.status ? thenable.then(noop, noop) : (thenable.status = "pending", thenable.then(
            function(fulfilledValue) {
              "pending" === thenable.status && (thenable.status = "fulfilled", thenable.value = fulfilledValue);
            },
            function(error) {
              "pending" === thenable.status && (thenable.status = "rejected", thenable.reason = error);
            }
          )), thenable.status) {
            case "fulfilled":
              return thenable.value;
            case "rejected":
              throw thenable.reason;
          }
      }
      throw thenable;
    }
    function mapIntoArray(children, array, escapedPrefix, nameSoFar, callback) {
      var type = typeof children;
      if ("undefined" === type || "boolean" === type) children = null;
      var invokeCallback = false;
      if (null === children) invokeCallback = true;
      else
        switch (type) {
          case "bigint":
          case "string":
          case "number":
            invokeCallback = true;
            break;
          case "object":
            switch (children.$$typeof) {
              case REACT_ELEMENT_TYPE:
              case REACT_PORTAL_TYPE:
                invokeCallback = true;
                break;
              case REACT_LAZY_TYPE:
                return invokeCallback = children._init, mapIntoArray(
                  invokeCallback(children._payload),
                  array,
                  escapedPrefix,
                  nameSoFar,
                  callback
                );
            }
        }
      if (invokeCallback)
        return callback = callback(children), invokeCallback = "" === nameSoFar ? "." + getElementKey(children, 0) : nameSoFar, isArrayImpl(callback) ? (escapedPrefix = "", null != invokeCallback && (escapedPrefix = invokeCallback.replace(userProvidedKeyEscapeRegex, "$&/") + "/"), mapIntoArray(callback, array, escapedPrefix, "", function(c) {
          return c;
        })) : null != callback && (isValidElement(callback) && (callback = cloneAndReplaceKey(
          callback,
          escapedPrefix + (null == callback.key || children && children.key === callback.key ? "" : ("" + callback.key).replace(
            userProvidedKeyEscapeRegex,
            "$&/"
          ) + "/") + invokeCallback
        )), array.push(callback)), 1;
      invokeCallback = 0;
      var nextNamePrefix = "" === nameSoFar ? "." : nameSoFar + ":";
      if (isArrayImpl(children))
        for (var i = 0; i < children.length; i++)
          nameSoFar = children[i], type = nextNamePrefix + getElementKey(nameSoFar, i), invokeCallback += mapIntoArray(
            nameSoFar,
            array,
            escapedPrefix,
            type,
            callback
          );
      else if (i = getIteratorFn(children), "function" === typeof i)
        for (children = i.call(children), i = 0; !(nameSoFar = children.next()).done; )
          nameSoFar = nameSoFar.value, type = nextNamePrefix + getElementKey(nameSoFar, i++), invokeCallback += mapIntoArray(
            nameSoFar,
            array,
            escapedPrefix,
            type,
            callback
          );
      else if ("object" === type) {
        if ("function" === typeof children.then)
          return mapIntoArray(
            resolveThenable(children),
            array,
            escapedPrefix,
            nameSoFar,
            callback
          );
        array = String(children);
        throw Error(
          "Objects are not valid as a React child (found: " + ("[object Object]" === array ? "object with keys {" + Object.keys(children).join(", ") + "}" : array) + "). If you meant to render a collection of children, use an array instead."
        );
      }
      return invokeCallback;
    }
    function mapChildren(children, func, context) {
      if (null == children) return children;
      var result = [], count = 0;
      mapIntoArray(children, result, "", "", function(child) {
        return func.call(context, child, count++);
      });
      return result;
    }
    function lazyInitializer(payload) {
      if (-1 === payload._status) {
        var ctor = payload._result;
        ctor = ctor();
        ctor.then(
          function(moduleObject) {
            if (0 === payload._status || -1 === payload._status)
              payload._status = 1, payload._result = moduleObject;
          },
          function(error) {
            if (0 === payload._status || -1 === payload._status)
              payload._status = 2, payload._result = error;
          }
        );
        -1 === payload._status && (payload._status = 0, payload._result = ctor);
      }
      if (1 === payload._status) return payload._result.default;
      throw payload._result;
    }
    var reportGlobalError = "function" === typeof reportError ? reportError : function(error) {
      if ("object" === typeof window && "function" === typeof window.ErrorEvent) {
        var event = new window.ErrorEvent("error", {
          bubbles: true,
          cancelable: true,
          message: "object" === typeof error && null !== error && "string" === typeof error.message ? String(error.message) : String(error),
          error
        });
        if (!window.dispatchEvent(event)) return;
      } else if ("object" === typeof process && "function" === typeof process.emit) {
        process.emit("uncaughtException", error);
        return;
      }
      console.error(error);
    };
    var Children = {
      map: mapChildren,
      forEach: function(children, forEachFunc, forEachContext) {
        mapChildren(
          children,
          function() {
            forEachFunc.apply(this, arguments);
          },
          forEachContext
        );
      },
      count: function(children) {
        var n = 0;
        mapChildren(children, function() {
          n++;
        });
        return n;
      },
      toArray: function(children) {
        return mapChildren(children, function(child) {
          return child;
        }) || [];
      },
      only: function(children) {
        if (!isValidElement(children))
          throw Error(
            "React.Children.only expected to receive a single React element child."
          );
        return children;
      }
    };
    exports.Activity = REACT_ACTIVITY_TYPE;
    exports.Children = Children;
    exports.Component = Component;
    exports.Fragment = REACT_FRAGMENT_TYPE;
    exports.Profiler = REACT_PROFILER_TYPE;
    exports.PureComponent = PureComponent;
    exports.StrictMode = REACT_STRICT_MODE_TYPE;
    exports.Suspense = REACT_SUSPENSE_TYPE;
    exports.__CLIENT_INTERNALS_DO_NOT_USE_OR_WARN_USERS_THEY_CANNOT_UPGRADE = ReactSharedInternals;
    exports.__COMPILER_RUNTIME = {
      __proto__: null,
      c: function(size) {
        return ReactSharedInternals.H.useMemoCache(size);
      }
    };
    exports.cache = function(fn) {
      return function() {
        return fn.apply(null, arguments);
      };
    };
    exports.cacheSignal = function() {
      return null;
    };
    exports.cloneElement = function(element, config, children) {
      if (null === element || void 0 === element)
        throw Error(
          "The argument must be a React element, but you passed " + element + "."
        );
      var props = assign({}, element.props), key = element.key;
      if (null != config)
        for (propName in void 0 !== config.key && (key = "" + config.key), config)
          !hasOwnProperty.call(config, propName) || "key" === propName || "__self" === propName || "__source" === propName || "ref" === propName && void 0 === config.ref || (props[propName] = config[propName]);
      var propName = arguments.length - 2;
      if (1 === propName) props.children = children;
      else if (1 < propName) {
        for (var childArray = Array(propName), i = 0; i < propName; i++)
          childArray[i] = arguments[i + 2];
        props.children = childArray;
      }
      return ReactElement(element.type, key, props);
    };
    exports.createContext = function(defaultValue) {
      defaultValue = {
        $$typeof: REACT_CONTEXT_TYPE,
        _currentValue: defaultValue,
        _currentValue2: defaultValue,
        _threadCount: 0,
        Provider: null,
        Consumer: null
      };
      defaultValue.Provider = defaultValue;
      defaultValue.Consumer = {
        $$typeof: REACT_CONSUMER_TYPE,
        _context: defaultValue
      };
      return defaultValue;
    };
    exports.createElement = function(type, config, children) {
      var propName, props = {}, key = null;
      if (null != config)
        for (propName in void 0 !== config.key && (key = "" + config.key), config)
          hasOwnProperty.call(config, propName) && "key" !== propName && "__self" !== propName && "__source" !== propName && (props[propName] = config[propName]);
      var childrenLength = arguments.length - 2;
      if (1 === childrenLength) props.children = children;
      else if (1 < childrenLength) {
        for (var childArray = Array(childrenLength), i = 0; i < childrenLength; i++)
          childArray[i] = arguments[i + 2];
        props.children = childArray;
      }
      if (type && type.defaultProps)
        for (propName in childrenLength = type.defaultProps, childrenLength)
          void 0 === props[propName] && (props[propName] = childrenLength[propName]);
      return ReactElement(type, key, props);
    };
    exports.createRef = function() {
      return { current: null };
    };
    exports.forwardRef = function(render) {
      return { $$typeof: REACT_FORWARD_REF_TYPE, render };
    };
    exports.isValidElement = isValidElement;
    exports.lazy = function(ctor) {
      return {
        $$typeof: REACT_LAZY_TYPE,
        _payload: { _status: -1, _result: ctor },
        _init: lazyInitializer
      };
    };
    exports.memo = function(type, compare) {
      return {
        $$typeof: REACT_MEMO_TYPE,
        type,
        compare: void 0 === compare ? null : compare
      };
    };
    exports.startTransition = function(scope) {
      var prevTransition = ReactSharedInternals.T, currentTransition = {};
      ReactSharedInternals.T = currentTransition;
      try {
        var returnValue = scope(), onStartTransitionFinish = ReactSharedInternals.S;
        null !== onStartTransitionFinish && onStartTransitionFinish(currentTransition, returnValue);
        "object" === typeof returnValue && null !== returnValue && "function" === typeof returnValue.then && returnValue.then(noop, reportGlobalError);
      } catch (error) {
        reportGlobalError(error);
      } finally {
        null !== prevTransition && null !== currentTransition.types && (prevTransition.types = currentTransition.types), ReactSharedInternals.T = prevTransition;
      }
    };
    exports.unstable_useCacheRefresh = function() {
      return ReactSharedInternals.H.useCacheRefresh();
    };
    exports.use = function(usable) {
      return ReactSharedInternals.H.use(usable);
    };
    exports.useActionState = function(action, initialState, permalink) {
      return ReactSharedInternals.H.useActionState(action, initialState, permalink);
    };
    exports.useCallback = function(callback, deps) {
      return ReactSharedInternals.H.useCallback(callback, deps);
    };
    exports.useContext = function(Context) {
      return ReactSharedInternals.H.useContext(Context);
    };
    exports.useDebugValue = function() {
    };
    exports.useDeferredValue = function(value, initialValue) {
      return ReactSharedInternals.H.useDeferredValue(value, initialValue);
    };
    exports.useEffect = function(create, deps) {
      return ReactSharedInternals.H.useEffect(create, deps);
    };
    exports.useEffectEvent = function(callback) {
      return ReactSharedInternals.H.useEffectEvent(callback);
    };
    exports.useId = function() {
      return ReactSharedInternals.H.useId();
    };
    exports.useImperativeHandle = function(ref, create, deps) {
      return ReactSharedInternals.H.useImperativeHandle(ref, create, deps);
    };
    exports.useInsertionEffect = function(create, deps) {
      return ReactSharedInternals.H.useInsertionEffect(create, deps);
    };
    exports.useLayoutEffect = function(create, deps) {
      return ReactSharedInternals.H.useLayoutEffect(create, deps);
    };
    exports.useMemo = function(create, deps) {
      return ReactSharedInternals.H.useMemo(create, deps);
    };
    exports.useOptimistic = function(passthrough, reducer) {
      return ReactSharedInternals.H.useOptimistic(passthrough, reducer);
    };
    exports.useReducer = function(reducer, initialArg, init) {
      return ReactSharedInternals.H.useReducer(reducer, initialArg, init);
    };
    exports.useRef = function(initialValue) {
      return ReactSharedInternals.H.useRef(initialValue);
    };
    exports.useState = function(initialState) {
      return ReactSharedInternals.H.useState(initialState);
    };
    exports.useSyncExternalStore = function(subscribe, getSnapshot, getServerSnapshot) {
      return ReactSharedInternals.H.useSyncExternalStore(
        subscribe,
        getSnapshot,
        getServerSnapshot
      );
    };
    exports.useTransition = function() {
      return ReactSharedInternals.H.useTransition();
    };
    exports.version = "19.2.6";
  }
});

// ../../node_modules/react/index.js
var require_react = __commonJS({
  "../../node_modules/react/index.js"(exports, module) {
    "use strict";
    if (true) {
      module.exports = require_react_production();
    } else {
      module.exports = null;
    }
  }
});

// ../../node_modules/react/cjs/react-jsx-runtime.production.js
var require_react_jsx_runtime_production = __commonJS({
  "../../node_modules/react/cjs/react-jsx-runtime.production.js"(exports) {
    "use strict";
    var REACT_ELEMENT_TYPE = Symbol.for("react.transitional.element");
    var REACT_FRAGMENT_TYPE = Symbol.for("react.fragment");
    function jsxProd(type, config, maybeKey) {
      var key = null;
      void 0 !== maybeKey && (key = "" + maybeKey);
      void 0 !== config.key && (key = "" + config.key);
      if ("key" in config) {
        maybeKey = {};
        for (var propName in config)
          "key" !== propName && (maybeKey[propName] = config[propName]);
      } else maybeKey = config;
      config = maybeKey.ref;
      return {
        $$typeof: REACT_ELEMENT_TYPE,
        type,
        key,
        ref: void 0 !== config ? config : null,
        props: maybeKey
      };
    }
    exports.Fragment = REACT_FRAGMENT_TYPE;
    exports.jsx = jsxProd;
    exports.jsxs = jsxProd;
  }
});

// ../../node_modules/react/jsx-runtime.js
var require_jsx_runtime = __commonJS({
  "../../node_modules/react/jsx-runtime.js"(exports, module) {
    "use strict";
    if (true) {
      module.exports = require_react_jsx_runtime_production();
    } else {
      module.exports = null;
    }
  }
});

// dashboard/src/lib/kiraPrompts.ts
var MONOTONIE_REGEL = [
  "Monotonie-Regel: Ein einzelner hoher Monotonie-Wert ist meist kein Problem. Ma\xDFgeblich ist die rollierende 7-Tage-Monotonie (\xD8/\u03C3 der Tageslasten) im Verlauf.",
  "Zonen: < 1,5 gute Variation; 1,5\u20132,0 normal; > 2,0 aufmerksam (vor allem bei hoher Wochenlast); > 2,5 ung\xFCnstig, wenn es l\xE4nger anh\xE4lt; > 3,0 Warnzone.",
  "Muster: 1\u20132 \xE4hnliche Tage normal; 3\u20134 \xE4hnliche, fordernde Tage beobachten (bei schlechterem Schlaf, HRV, Ruhepuls, schweren Beinen oder weniger Motivation reduzieren); 5\u20137 Tage ohne deutliche Entlastung beobachten (nur Hinweis, senkt das Votum nicht \u2013 in Marcs Historie ohne messbare Folgen); mehrere Wochen hohe Monotonie = Risiko f\xFCr Erm\xFCdung, Stagnation, \xDCberlastung, Infekte.",
  "Steuerung: Ein hoher Wert an einem Tag ist praktisch bedeutungslos. Liegt die 7-Tage-Monotonie an 2 Tagen in Folge \xFCber etwa 2,0 und sind schon 4\u20135 fordernde Tage gesammelt \u2192 n\xE4chster Tag sehr locker oder frei. Als Entlastung z\xE4hlt ein lockerer oder freier Tag (h\xF6chstens 40 ESS bzw. 60 % der \xFCblichen Tageslast). Bleibt der Wert \xFCber 2,5 \u2192 innerhalb 24\u201348 h echter Ruhetag oder sehr kurzer Regenerationsreiz. \xDCber 3,0 \u2192 mindestens ein echter Entlastungstag oder deutliche Reduktion. Hohe Monotonie plus schlechte Erholungswerte \u2192 keine Qualit\xE4tseinheit durchdr\xFCcken.",
  "Entscheidend ist die Kombination mit der Gesamtlast (Strain = Wochenlast \xD7 Monotonie). Monotonie senkst du durch klaren Rhythmus: harte Reize klar hart, lockere Tage wirklich locker, mindestens ein echter Erholungstag bei hoher Gesamtlast \u2013 nicht durch weniger Gesamtlast."
].join("\n");
function erfahrungBlock(text) {
  const t = String(text || "").trim();
  if (!t) return "";
  return "\n---\n**ERFAHRUNG AUS NICHTBEFOLGUNG (aus Marcs eigenen Daten, fest berechnet \u2013 ber\xFCcksichtigen):**\n" + t + "\n---\n";
}
function monotonieBlock() {
  return "\n---\n**MONOTONIE-REGEL (verbindlich f\xFCr Bewertung und Plan):**\n" + MONOTONIE_REGEL + "\nBewerte die Prognose-Monotonie oben nach dieser Regel: ein einzelner hoher Tageswert ist kein Grund zum K\xFCrzen; entscheidend sind fordernde Tage ohne Entlastung, anhaltend hohe Werte und die Wochenlast.\n---\n";
}
var WISSENSBLOCK_V105 = `
--- KERNWISSEN (Firstbeat Whitepapers) ---
Du (Kira) musst diese Definitionen als "Quelle der Wahrheit" f\xFCr deine Analysen verwenden.

1.  **Training Effect (TE) Skala (0.0-5.0)[cite: 1799, 2123]:**
    * **Basis:** Basiert auf EPOC (Excess Post-Exercise Oxygen Consumption)[cite: 1667, 2102].
    * **EPOC:** Misst die St\xF6rung der Hom\xF6ostase (also den Erholungsbedarf) durch das Training[cite: 1721, 2102, 3460].
    * **TE 1.0-1.9 (Minor):** Erh\xE4lt die Grundlagenausdauer oder f\xF6rdert die Erholung[cite: 1809, 2123].
    * **TE 2.0-2.9 (Maintaining):** H\xE4lt die aktuelle kardiorespiratorische Fitness[cite: 1811, 2123].
    * **TE 3.0-3.9 (Improving):** Verbessert die kardiorespiratorische Fitness (empfohlen 2-4x pro Woche)[cite: 1813, 2123].
    * **TE 4.0-4.9 (Highly Improving):** Verbessert stark (empfohlen 1-2x pro Woche), erfordert mehr Aufmerksamkeit bei der Erholung[cite: 1815, 2123].
    * **TE 5.0 (Overreaching):** Extremer Reiz ("Overreaching")[cite: 1817, 2123]. F\xFChrt nur mit ausreichender Erholung zur Verbesserung[cite: 2123].


2.  **Training Load & ACWR (Belastungssteuerung):**
    * **Acute Load (ATL):** Die Summe des TRIMP der letzten 7 Tage.
    * **Chronic Load (CTL):** Der Durchschnitts-TRIMP der letzten 28 Tage.
    * **ACWR (Acute:Chronic Workload Ratio):** Das Verh\xE4ltnis von Akut zu Chronisch.
    * **Zielzone ("Sweet Spot"):** ACWR zwischen **0.8 und 1.2**.
    * **Untertraining (ACWR < 0.8):** Deutet auf **Formverlust / Detraining** hin (NICHT auf \xDCberforderung!).
    * **\xDCberlastung (ACWR > 1.5):** Gefahrenzone! Verletzungsrisiko steigt signifikant.
    * **WICHTIG:** Es gibt **keine** mathematische Korrelation zwischen ACWR (Last) und HRV (Stress). Behandle sie als separate Signale.
// ...

3.  **HRV (Stress & Erholung)[cite: 1026]:**
    * **Grundlage:** Die HRV-Analyse spiegelt das Autonome Nervensystem (ANS) wider[cite: 1044, 1164, 1435].
    * **Stress-Reaktion:** Wird durch Dominanz des **Sympathischen Nervensystems** (fight or flight) angezeigt (HR steigt, HRV sinkt)[cite: 1033, 1154, 1304].
    * **Erholung:** Wird durch Dominanz des **Parasympathischen (Vagalen) Nervensystems** (rest and digest) angezeigt (HR sinkt, HRV steigt)[cite: 1034, 1158, 1298].
    * **Schlechte Erholung / Erm\xFCdung:** Zeigt sich, wenn die sympathische Dominanz (Stress) *auch in Ruhephasen* (z.B. Schlaf) anh\xE4lt[cite: 2669, 2898, 2901].

4.  **VO2max:**
    * **Definition:** Die "goldene Standardmessung" f\xFCr die aerobe Fitness[cite: 3042, 3047].
--- ENDE KERNWISSEN ---
`;
function wissensText(wissen) {
  if (!Array.isArray(wissen)) return WISSENSBLOCK_V105;
  let t = "--- EXPERTEN-WISSENSDATENBANK (KIRA SYSTEM) ---\n";
  wissen.forEach((c) => {
    if (c) t += c + "\n\n";
  });
  return t + "--- ENDE WISSENSDATENBANK ---\n";
}
function trefferBlock(tr) {
  if (!tr || !tr.text) return "";
  return "\n---\n**KIRA-TREFFSICHERHEIT (R\xFCckblick, Stand " + tr.datum + "):**\nSo hat Marc in den letzten Wochen auf deine Empfehlungen reagiert. Ber\xFCcksichtige das bei Plan und Empfehlung.\n" + String(tr.text).slice(0, 1500) + "\n---\n";
}
function planungsBlock(p) {
  if (!p) return "";
  const heute = p.heute;
  let out = "";
  const ziele = (p.ziele || []).filter((z) => z.datum >= heute).sort((a, b) => a.datum < b.datum ? -1 : 1);
  if (ziele.length) out += "**ZIELE / TOUREN (darauf hinplanen):**\n" + ziele.slice(0, 5).map((r) => "- " + r.datum + ": " + String(r.text).slice(0, 300)).join("\n") + "\n";
  const wochen = p.wochen || [];
  if (wochen.length) {
    const w = wochen.filter((x) => x.start <= heute).pop() || wochen[0];
    const n = wochen[wochen.indexOf(w) + 1];
    if (w) out += "**SAISONRAHMEN diese Woche (ab " + w.start + "):** Phase " + w.phase + (w.fenster ? " (Zeitfenster: " + w.fenster + ")" : "") + ", Ziel-Wochenlast ca. " + w.ziel_ess + " ESS, Fokus: " + (w.fokus || "-") + (n ? " \xB7 n\xE4chste Woche: " + n.phase + (n.fenster ? " (Zeitfenster: " + n.fenster + ")" : "") + ", ca. " + n.ziel_ess + " ESS" : "") + "\n";
  }
  return out ? "\n---\n" + out + "Richte Plan und Empfehlung an Zielen und Saisonrahmen aus, ohne die Monotonie- und ACWR-Regeln zu verletzen.\n---\n" : "";
}
function parseGermanFloat_(value) {
  if (value === null || value === void 0) return NaN;
  const s = String(value).trim();
  if (!s) return NaN;
  const cleaned = s.replace(/\s/g, "").replace("%", "").replace(/\.(?=\d{3}(\D|$))/g, "").replace(",", ".");
  const n = parseFloat(cleaned);
  return Number.isFinite(n) ? n : NaN;
}
function parseGermanFloat(value) {
  if (typeof value === "number") return value;
  if (typeof value === "string") {
    const n = parseGermanFloat_(value);
    return Number.isFinite(n) ? n : value;
  }
  return value;
}
function getBaselineNumber_(baselineObj, aliases) {
  const b = baselineObj || {};
  const keys = Object.keys(b);
  const norm = (x) => String(x || "").trim().toLowerCase().replace(/\s+/g, "").replace(/[^a-z0-9]/g, "");
  const aliasNorm = (aliases || []).map(norm);
  for (let i = 0; i < keys.length; i++) {
    const kNorm = norm(keys[i]);
    if (!kNorm) continue;
    if (aliasNorm.indexOf(kNorm) !== -1) {
      const n = parseGermanFloat_(b[keys[i]]);
      if (Number.isFinite(n)) return n;
    }
  }
  for (let i = 0; i < keys.length; i++) {
    const kNorm = norm(keys[i]);
    if (!kNorm) continue;
    for (let a = 0; a < aliasNorm.length; a++) {
      const wanted = aliasNorm[a];
      if (!wanted) continue;
      if (kNorm.indexOf(wanted) !== -1 || wanted.indexOf(kNorm) !== -1) {
        const n = parseGermanFloat_(b[keys[i]]);
        if (Number.isFinite(n)) return n;
      }
    }
  }
  return NaN;
}
var SUPERVISOR_SYSTEM = "Du bist ein pr\xE4ziser Daten-Analyst. Deine Aufgabe ist es, die Daten zu analysieren und NUR das angeforderte JSON zu f\xFCllen.";
var SUPERVISOR_SCHEMA = {
  "name": "fitness_report",
  "strict": true,
  "schema": {
    "type": "object",
    "properties": {
      "empfehlung_zukunft": {
        "type": "object",
        "properties": {
          "plan_status": { "type": "string", "enum": ["FREIGABE", "BEOBACHTUNG", "INTERVENTION"] },
          "text": { "type": "string", "description": "Ausf\xFChrliche Gesamtanalyse" }
        },
        "required": ["plan_status", "text"],
        "additionalProperties": false
      },
      "bewertung_ernaehrung_7d": {
        "type": "object",
        "properties": {
          "text": { "type": "string", "description": "Kurzbewertung Ern\xE4hrung" }
        },
        "required": ["text"],
        "additionalProperties": false
      },
      "einzelscore_kommentare": {
        "type": "array",
        "description": "Liste der Kommentare f\xFCr jeden einzelnen Score",
        "items": {
          "type": "object",
          "properties": {
            "metrik": { "type": "string" },
            "text_info": { "type": "string", "description": "Kurzer Kommentar (1 Satz)" }
          },
          "required": ["metrik", "text_info"],
          "additionalProperties": false
        }
      },
      "empfohlener_plan": {
        "type": "array",
        "items": {
          "type": "object",
          "properties": {
            "datum": { "type": "string" },
            "tag": { "type": "string" },
            "original_load_ess": { "type": "number" },
            "empfohlener_load_ess": { "type": "number" },
            "empfohlene_zone": { "type": "string" },
            "beispiel_training_1": { "type": "string" },
            "beispiel_training_2": { "type": "string" },
            "prognostizierte_auswirkung": {
              "type": "string",
              "description": "Detaillierte Erkl\xE4rung."
            },
            "wetterempfehlung": { "type": "string" }
          },
          "required": ["datum", "tag", "original_load_ess", "empfohlener_load_ess", "empfohlene_zone", "beispiel_training_1", "beispiel_training_2", "prognostizierte_auswirkung", "wetterempfehlung"],
          "additionalProperties": false
        }
      }
    },
    "required": ["empfehlung_zukunft", "bewertung_ernaehrung_7d", "einzelscore_kommentare", "empfohlener_plan"],
    "additionalProperties": false
  }
};
function supervisorPrompt(d) {
  return { system: SUPERVISOR_SYSTEM, prompt: supervisorText(d), schema: SUPERVISOR_SCHEMA.schema };
}
function supervisorText(d) {
  const data = { heute: d.heute || {}, zukunft: { plan: d.zukunft_plan || [] }, baseline: d.baseline || {}, history: d.history || "", monotony_varianz: Number(d.monotony_varianz) || 0, rtp_status: String(d.rtp_status || "") };
  const fitnessScores = d.fitnessScores || [];
  const nutritionScore = d.nutritionScore || { raw_wert: 0, num_score: 0 };
  const gesamtScoreNum = d.gesamtScoreNum;
  const gesamtScoreAmpel = d.gesamtScoreAmpel;
  const zoneFactors = d.zoneFactors;
  const elevFactors = d.elevFactors;
  const weekConfigCSV = d.weekConfigCSV;
  const isActivityDone = !!d.isActivityDone;
  const wetterDaten = d.wetterDaten;
  const alleScores = fitnessScores.concat([nutritionScore]);
  const { heute, zukunft, baseline, history, monotony_varianz, rtp_status } = data;
  const acwrRawHeute = heute?.coachE_ACWR_forecast ?? heute?.coache_acwr_forecast;
  const acwrForecastNum = parseGermanFloat_(acwrRawHeute);
  const acwrFix = Number.isFinite(acwrForecastNum) ? acwrForecastNum.toFixed(2).replace(".", ",") : String(acwrRawHeute ?? "").trim();
  const acwrFixDot = acwrFix.replace(",", ".");
  const fallbackTeBalanceHeute = (monotony_varianz * 100).toFixed(1).replace(".", ",") + "%";
  const teBalanceHeute = d.teBalanceHeute || fallbackTeBalanceHeute;
  const dynamicKnowledge = wissensText(d.wissen);
  const cleanDate = String(d.heute_datum || "");
  const cleanTag = String(d.heute_tag || "");
  let heuteString = `- ${cleanDate} (${cleanTag}): `;
  const heutigerIstLoad = parseGermanFloat(heute["load_fb_day"]);
  const isActivityDone_Text = isActivityDone ? "true" : "false";
  const heutigerSollLoad = parseGermanFloat(heute["coachE_ESS_day"]);
  let statusHeuteHeader = isActivityDone ? `\u2705 TRAINING HEUTE BEREITS ABSOLVIERT (Ist-Load: ${heutigerIstLoad})` : `\u23F3 TRAINING HEUTE NOCH OFFEN (Soll-Load: ${heutigerSollLoad})`;
  let scoresText = "";
  alleScores.forEach((s) => {
    let cleanVal = String(s.raw_wert).replace(",", ".");
    let scoreInfo = s.num_score;
    if (s.metrik.includes("ACWR") && parseFloat(cleanVal) > 0.8 && s.num_score < 10) {
      scoreInfo = `${s.num_score} (ACHTUNG: HOCHRISIKO-BEREICH!)`;
    }
    scoresText += `- ${s.metrik}: ${cleanVal} (Score: ${scoreInfo}, ${s.ampel})
`;
  });
  const metricsChecklist = alleScores.map((s) => `"${s.metrik}"`).join(", ");
  const keiMetricFromFitnessScores = fitnessScores.find((s) => s && s.metrik === "KEI");
  const keiPromptValue = keiMetricFromFitnessScores ? String(keiMetricFromFitnessScores.raw_wert || "\u2014") : "\u2014";
  const isRTP = rtp_status && rtp_status.includes("RTP");
  let daysSinceHealthy = 0;
  if (isRTP) {
    try {
      daysSinceHealthy = parseInt(rtp_status.split("_")[2]);
    } catch (e) {
      daysSinceHealthy = 0;
    }
  }
  let sickDaysCount = 0;
  if (isRTP && history) {
    const rows = history.split("\n");
    for (let i = rows.length - 1; i >= 0; i--) {
      const rowStr = rows[i].toLowerCase();
      if (rowStr.includes("krank") || rowStr.includes("sick") || rowStr.includes("infekt")) {
        sickDaysCount++;
      } else if (sickDaysCount > 0 && rowStr.includes(",")) {
        break;
      }
    }
  }
  if (isRTP && sickDaysCount === 0) {
    sickDaysCount = 7;
  }
  let protocolOffset = 0;
  if (sickDaysCount < 9) {
    protocolOffset = 9 - sickDaysCount;
    if (protocolOffset < 0) protocolOffset = 0;
  }
  const getEffectiveDay = (dayIndexFromNow) => {
    if (!isRTP) return 0;
    const arrayIndex = daysSinceHealthy - 1 + protocolOffset + dayIndexFromNow;
    const effectiveDay = arrayIndex + 1;
    return effectiveDay <= 9 ? effectiveDay : 0;
  };
  const rtpRamp = [10, 30, 60, 0, 60, 0, 70, 80, 0];
  let todayPlan_LoadRec = 0;
  let todayPlan_Zone = "Ruhetag (Fixed)";
  let todayPlan_Sport1 = "Geplanter Ruhetag";
  let todayPlan_Sport2 = "Yoga / Stretching";
  let todayPlan_Text = "Regeneration.";
  let heuteZusatzInfo = "";
  const todayEffectiveDay = isRTP ? getEffectiveDay(0) : 0;
  if (todayEffectiveDay > 0) {
    const rtpLoad = rtpRamp[todayEffectiveDay - 1];
    if (isActivityDone) {
      heuteZusatzInfo = ` -> \u2705 RTP TAG ${todayEffectiveDay}/9 ERLEDIGT.`;
      todayPlan_LoadRec = heutigerIstLoad;
      todayPlan_Text = `RTP (Tag ${todayEffectiveDay}/9). Einheit absolviert.`;
      todayPlan_Zone = "Erledigt";
      todayPlan_Sport1 = "Absolviert";
    } else {
      heuteZusatzInfo = ` -> \u26A0\uFE0F RTP TAG ${todayEffectiveDay}/9 (Smart Entry). Load: ${rtpLoad}.`;
      todayPlan_LoadRec = rtpLoad;
      todayPlan_Text = `RTP-Schutz (Tag ${todayEffectiveDay}/9). Protokoll aktiv.`;
      switch (todayEffectiveDay) {
        case 1:
          todayPlan_Zone = "Z1/Rek";
          todayPlan_Sport1 = "Spaziergang";
          break;
        case 2:
          todayPlan_Zone = "Z1 (Bike)";
          todayPlan_Sport1 = "Diagnose-Rolle";
          break;
        case 3:
          todayPlan_Zone = "Z2 (Bike)";
          todayPlan_Sport1 = "Rolle Intensiv";
          break;
        case 4:
          todayPlan_Zone = "Ruhetag";
          todayPlan_Sport1 = "Integration (Ruhe)";
          break;
        case 5:
          todayPlan_Zone = "Z2 (Run)";
          todayPlan_Sport1 = "Lauf (kurz)";
          break;
        case 6:
          todayPlan_Zone = "Ruhetag";
          todayPlan_Sport1 = "Integration (Ruhe)";
          break;
        case 7:
          todayPlan_Zone = "Z2 (Run)";
          todayPlan_Sport1 = "Lauf (Steigerung)";
          break;
        case 8:
          todayPlan_Zone = "Z2 (Bike)";
          todayPlan_Sport1 = "Lange Rolle";
          break;
        case 9:
          todayPlan_Zone = "Ruhetag";
          todayPlan_Sport1 = "Abschluss";
          break;
      }
    }
  }
  let heutigerFixMarker = (heute["fix"] || " ").toString().trim();
  let prognoseLoadText = "Original Load (ESS):\n";
  prognoseLoadText += `${heuteString}${heutigerSollLoad}${heuteZusatzInfo}
`;
  let prognoseFixText = `Fixed Tage (fix='x'):
${heuteString}${heutigerFixMarker}
`;
  let prognoseMonoText = `Prognose Monotonie (Load):
${heuteString}${heute["Monotony7"]} (Belastungsfokus Ratio: ${(monotony_varianz * 100).toFixed(1)}%)
`;
  let prognoseAcwrText = `Prognose ACWR:
${heuteString}${acwrFix}
`;
  let peakCTL = 0;
  try {
    if (history) {
      const histLines = history.split("\n");
      for (let i = histLines.length - 1; i > 0; i--) {
        const parts = histLines[i].split(",").map((s) => s.replace(/"/g, "").trim());
        const histCTL = parseFloat(parts[5]);
        if (!isNaN(histCTL) && histCTL > peakCTL) peakCTL = histCTL;
      }
    }
  } catch (e) {
  }
  const currentCTL = parseGermanFloat(heute["coachE_CTL_forecast"]) || parseGermanFloat(heute["fbCTL_obs"]) || 60;
  const anchorCTL = peakCTL > 0 ? peakCTL : currentCTL;
  const targetRecoveryLoad = Math.round(anchorCTL * 0.6);
  const resting_calories_raw = baseline?.["Resting_Calories (kcal)"];
  const resting_calories_num = parseGermanFloat(resting_calories_raw);
  const resting_calories = Number.isFinite(resting_calories_num) ? resting_calories_num : 0;
  const rhr_default = getBaselineNumber_(baseline, ["RHR_default (bpm)", "RHR_default(bpm)", "rhr_default (bpm)", "rhr_default_bpm", "rhr_default"]);
  const vo2max = parseGermanFloat(baseline["VO2max (mg/kg/min)"]);
  const ftp_rad = parseGermanFloat(baseline["FTP_Rad (W/kg)"]);
  const hill_score = parseGermanFloat(baseline["Hill_Score (0-100)"]);
  const run_econ = parseGermanFloat(baseline["Running_Economy (184-225)"]);
  const allgemeineZiele = baseline["Allgemeine Ziele"] || "Keine Ziele definiert.";
  let ernaehrungLabel = "";
  const rawDeficit = parseFloat(nutritionScore.raw_wert);
  if (rawDeficit > 0) ernaehrungLabel = "DEFIZIT";
  else if (rawDeficit < 0) ernaehrungLabel = "\xDCBERSCHUSS";
  else ernaehrungLabel = "AUSGEGLICHEN";
  const problemScores = alleScores.filter((s) => s.num_score < 80).map((s) => `${s.metrik}: ${s.raw_wert} (${s.ampel}/${s.num_score})`).join("; ");
  const isRecoveryCritical = alleScores.some(
    (s) => (s.metrik === "HRV Status" || s.metrik === "Training Readiness" || s.metrik === "RHR") && s.num_score < 40
  );
  const isRecoveryBad = alleScores.some(
    (s) => (s.metrik === "HRV Status" || s.metrik === "Training Readiness" || s.metrik === "RHR") && s.num_score < 80
  );
  const isRecoverySuper = alleScores.every(
    (s) => s.metrik !== "HRV Status" && s.metrik !== "Training Readiness" || s.num_score >= 90
  );
  zukunft.plan.forEach((tag, i) => {
    let tagString = `${tag.datum} (${tag.tag})`;
    let phase = tag.week_phase ? String(tag.week_phase).trim().toUpperCase() : "A";
    let phaseMarker = "\u{1F680} [PHASE A: AUFBAU]";
    let loadInstruction = `Geplanter Load [ORIGINAL-PLAN: ${tag.original_load_ess}]`;
    if (phase === "E") {
      phaseMarker = "\u{1F6D1} [PHASE E: ENTLASTUNG]";
      loadInstruction = `Original Plan: ${tag.original_load_ess} -> ZIEL: Reduzieren auf ca. ${targetRecoveryLoad}`;
    }
    if (phase === "E") {
      phaseMarker = "\u{1F6D1} [PHASE E: ENTLASTUNG]";
      loadInstruction = `Original Plan: ${tag.original_load_ess} -> ZIEL: Reduzieren auf ca. ${targetRecoveryLoad}`;
    } else if (phase.startsWith("A") && phase !== "A") {
      phaseMarker = `\u{1F680} [PHASE ${phase}: AUFBAU]`;
    }
    if (isRTP) {
      const futureEffDay = getEffectiveDay(i + 1);
      if (futureEffDay > 0) {
        phaseMarker = `\u{1F691} [RTP TAG ${futureEffDay}/9]`;
        const isDone = tag.original_act_load && tag.original_act_load > 5;
        if (isDone) {
          loadInstruction = `\u2705 ERLEDIGT (Ist: ${tag.original_act_load}). [ORIGINAL-PLAN: ${tag.original_load_ess}]`;
        } else {
          const baseInst = `[ORIGINAL-PLAN: ${tag.original_load_ess}] (IGNORIEREN!) -> \u26A0\uFE0F RTP VORGABE: `;
          switch (futureEffDay) {
            case 1:
              loadInstruction = `${baseInst}Ziel ~10 (Walk). Zone: 'Z1/Rek'.`;
              break;
            case 2:
              loadInstruction = `${baseInst}Ziel ~30 (Bike). Zone: 'Z1 (Bike)'.`;
              break;
            case 3:
              loadInstruction = `${baseInst}Ziel ~60 (Bike). Zone: 'Z2 (Bike)'.`;
              break;
            case 4:
              loadInstruction = `${baseInst}RUHETAG. Zone: 'Ruhetag'.`;
              break;
            case 5:
              loadInstruction = `${baseInst}Ziel ~60 (Run). Zone: 'Z2 (Run)'.`;
              break;
            case 6:
              loadInstruction = `${baseInst}RUHETAG. Zone: 'Ruhetag'.`;
              break;
            case 7:
              loadInstruction = `${baseInst}Ziel ~70 (Run). Zone: 'Z2 (Run)'.`;
              break;
            case 8:
              loadInstruction = `${baseInst}Ziel ~80 (Bike). Zone: 'Z2 (Bike)'.`;
              break;
            case 9:
              loadInstruction = `${baseInst}RUHETAG. Zone: 'Ruhetag'.`;
              break;
            default:
              loadInstruction = `RTP Fehler.`;
          }
        }
      }
    }
    prognoseLoadText += `${phaseMarker} - ${tagString}: ${loadInstruction}
`;
    prognoseFixText += `- ${tagString}: ${tag.original_fix_marker || " "}
`;
    prognoseMonoText += `- ${tagString}: ${tag.original_monotony}
`;
    prognoseAcwrText += `- ${tagString}: ${tag.original_acwr}
`;
  });
  let prompt = `Du bist Coach Kira, eine erfahrene KI-Sportwissenschaftlerin und Fitnessexpertin.
Analysiere **deine (Marcs)** Daten. Ich habe die Scores (0-100) berechnet.
Deine Aufgabe ist es, im JSON-Format Folgendes zu liefern:
1. 'empfehlung_zukunft': Eine ausf\xFChrliche textliche Gesamtbewertung ("text") und einen Plan-Status ("plan_status").
2. 'bewertung_ernaehrung_7d': Eine **kurze (max. 2 S\xE4tze)** textliche Bewertung der Ern\xE4hrungsbilanz ("text").
3. 'empfohlener_plan': Einen angepassten 8-TAGE-PLAN (HEUTE + 7 TAGE ZUKUNFT).

4. 'einzelscore_kommentare': Ein Array von Objekten f\xFCr JEDEN der 10 FITNESS-SCORES (NICHT Ern\xE4hrung), mit "metrik" und "text_info".
**WICHTIG: Formuliere alle deine Analysen sehr detailliert und kontextbezogen ('text', 'text_info', 'prognostizierte_auswirkung', 'wetterempfehlung') direkt an **dich (Marc)** und verwende die 'Du'-Anrede.**

${dynamicKnowledge}

--- \u{1F3AF} DER ABSOLUTE DATEN-ANKER F\xDCR HEUTE (${cleanDate}) ---
! WICHTIG: Falls Historien-Tabellen abweichende Werte enthalten: IGNORIEREN !
! NUR DIESE WERTE SIND DIE WAHRHEIT F\xDCR DEINEN TEXT !
- Ruhepuls (RHR) heute: ${heute["rhr_bpm"]} bpm
- HRV-Status heute: ${heute["hrv_status"]} ms (Bereich: ${heute["hrv_threshholds"]})
- Schlaf: ${heute["sleep_hours"]}h (Score: ${heute["sleep_score_0_100"]}/100)
- Trainingszustand: ${heute["trainingszustand"]}
- Aktueller ACWR (Forecast): ${acwrFix}  (WERT IST FIX, genau so verwenden)
- Heutige Belastung (IST): ${heutigerIstLoad} ESS
- Heutige Belastung (SOLL laut Plan): ${heutigerSollLoad} ESS
- TE-Balance (% Intensiv): ${teBalanceHeute} %
- KEI (aus fitnessScores): ${keiPromptValue}
- Protein-Zufuhr: ${heute["protein_g"]}g
- Energie-Bilanz (\xD8 7 Tage): ${nutritionScore.raw_wert} kcal
- RTP-STATUS (System-Schutz): ${rtp_status}
-----------------------------------------------------------

---
**STATUS HEUTE: ${statusHeuteHeader}**
(Nutze diese Info zwingend f\xFCr die Perspektive deiner Analyse!)
---

---
DEINE ALLGEMEINEN ZIELE (KONTEXT):
**${allgemeineZiele}**
---

---
**DEIN IDEALER WOCHENPLAN (ALS HINWEIS, NIEDRIGE PRIORIT\xC4T)**
(CSV: day_of_week,load,sport,zone)
**${weekConfigCSV}**
---

---
**NEU (V42): WETTER PROGNOSE (HEUTE & MORGEN):**
${wetterDaten}
---
${trefferBlock(d.treffer)}
${monotonieBlock()}${erfahrungBlock(d && d.erfahrung)}
${planungsBlock(d.planung)}


---
**DEINE SCH\xC4TZ-FAKTOREN (ZUM BERECHNEN VON BEISPIELEN):**
Zonen-Faktoren (Punkte pro Stunde): ${JSON.stringify(zoneFactors)}
H\xF6hen-Faktoren (Punkte pro 100hm): ${JSON.stringify(elevFactors)}
**FORMEL: Load = (Minuten / 60) * ZonenFaktor + (H\xF6henmeter / 100) * H\xF6henFaktor**
---

DATEN-BASELINE (PHYSIOLOGISCHES PROFIL)
- Ruheumsatz: ${resting_calories.toFixed(0)} kcal/Tag
- RHR-Basis: ${rhr_default} bpm
- VO2max: ${vo2max} (mg/kg/min)
- FTP (Rad): ${ftp_rad} (W/kg)
- Hill Score: ${hill_score} / 100
- Running Economy: ${run_econ} (niedriger ist besser)

---
**VERGANGENE 14 TAGE (HISTORIE & KONTEXT):**
(Zeigt den IST-Load (Actual_Load) und die Training Effect (TE) Verteilung)
${history}
---

BERECHNETER GESAMTSTATUS (HEUTE: ${cleanDate})
- Gesamtscore: ${gesamtScoreNum} / 100 (${gesamtScoreAmpel})
- HEUTIGER Original-Load (ESS): ${heutigerSollLoad}
- HEUTIGER IST-Load (load_fb_day): ${heutigerIstLoad}
- HEUTIGER Wochentag: ${cleanTag}
- HEUTIGE Monotonie (Load): ${heute["Monotony7"]}
- HEUTIGE TE Balance: ${(monotony_varianz * 100).toFixed(1)}% (Dies ist das Ratio Intensiv/Basis)
- V106-INFO: isRecoveryCritical == ${isRecoveryCritical} (mindestens ein Bio-Marker < 40/ROT)
- V106-INFO: isRecoveryBad == ${isRecoveryBad} (mindestens ein Bio-Marker < 80/GELB)
- V106-INFO: isRecoverySuper == ${isRecoverySuper} (Bio-Marker Top > 90)
- **RTP-STATUS (Krankheit): ${rtp_status}** (Wenn != HEALTHY -> ALARM!)
- **RTP-SCHWEREGRAD:** ${sickDaysCount} Tage krank -> Start bei Tag ${todayEffectiveDay}/9.
- Problem-Scores (heute): ${problemScores || "Keine"}

---
BERECHNETE EINZELSCORES (HEUTE):
${scoresText}
---
GEPLANTER TRAININGSPLAN & PROGNOSE (HEUTE + 7 TAGE)
(Dies ist der *alte* Plan. Du sollst ihn verbessern. Achte auf [FIXED] Marker!)
${prognoseLoadText}
${prognoseFixText}
${prognoseMonoText}
${prognoseAcwrText}
---
DEINE AUFGABE (MAXIMALER KONTEXT & SPORTWISSENSCHAFTLICHE ANALYSE):

**Ber\xFCcksichtige bei allen Empfehlungen die 'Allgemeinen Ziele' UND die 'Vergangene 14 Tage (Historie)', um konsistent zu planen.**

1.  **ANALYSIERE ERN\xC4HRUNG ('bewertung_ernaehrung_7d'):**
    * Bewerte die 7-Tage-Bilanz (Score ${nutritionScore.num_score}, Wert ${nutritionScore.raw_wert} kcal) **in maximal 2 S\xE4tzen**.
    * (REGEL: Der Wert ${nutritionScore.raw_wert} kcal ist ein ${ernaehrungLabel}).

2.  **ANALYSIERE GESAMTSTATUS & LEITE PLANSTATUS AB ('empfehlung_zukunft'):**
        * **AUFTRAG:** Schreibe eine **umfassende, tiefgehende Analyse (mindestens 150-200 W\xF6rter)**.
        * **STRUKTUR (Halte dich grob daran):**
          A) **Status Quo (Der Ist-Zustand):** Wie stehen deine Bio-Marker (HRV/RHR/Schlaf) im Vergleich zur Basis? Was ist der Haupttreiber f\xFCr den aktuellen Gesamtscore (${gesamtScoreAmpel})?
          B) **Die Verbindung (Korrelation):** Erkl\xE4re das "Warum". 
          C) **Strategie & Ausblick:** Was bedeutet das konkret f\xFCr die n\xE4chsten 3-4 Tage? Beziehe die **Wochen-Phase** (Aufbau vs. Entlastung) und die **TE-Balance** mit ein.
          D) **ACWR-Kontext-Check (WICHTIG):**
             - Bewerte einen niedrigen ACWR (< 0.8) **NIEMALS** als "\xDCberforderung". Das ist technisch falsch.
             - **Szenario 1 (Schlechte Bio-Marker / Krank):** Ein niedriger ACWR ist hier **GUT und ERWARTET** (notwendige Erholung). Lobe die Disziplin zur Pause.
             - **Szenario 2 (Gute Bio-Marker / Fit):** Ein niedriger ACWR ist hier **SCHLECHT** (Formverlust/Untertraining). Warne davor, dass wir Potenziale verschenken.

        * **STIL:** Sprich wie ein erfahrener Performance-Coach. Analytisch, pr\xE4zise, aber motivierend. Keine Floskeln, sondern Fakten aus den Daten.
    
    **WICHTIG - KONTEXT-WEICHE:**
    * **FALL A (TRAINING BEREITS ABSOLVIERT):** Bewerte die durchgef\xFChrte Einheit r\xFCckblickend! War sie f\xFCr deinen heutigen Zustand (HRV/Readiness) angemessen? War sie vielleicht zu hart? Oder genau richtig? Beachte, dass dadurch auch die Training Readiness naturgem\xE4\xDF niedrig ist. Richte den Blick dann auf die Regeneration f\xFCr morgen.
    * **FALL B (TRAINING NOCH OFFEN):**
      Gib eine klare Handlungsanweisung und Motivation f\xFCr das anstehende Training.

3.  **PASSE 8-TAGE-PLAN AN & VALIDIERE ('empfohlener_plan'):**
    * Erstelle einen **VOLLST\xC4NDIGEN 8-TAGE-PLAN** (heute + 7 Folgetage).
    * **Entferne Kcal/Makros.** Der Plan enth\xE4lt nur die Felder aus dem JSON-Beispiel unten.
    * **WICHTIG (V56-FIX):** Das Feld 'original_load_ess' in deinem JSON MUSS exakt mit dem 'Original Load (ESS)' aus dem Prognose-Block \xFCbereinstimmen.
    
    * **(V136-MOD) UNVERHANDELBARE REGELN (HIERARCHIE IST WICHTIG):**

      *** WICHTIGSTE REGEL (PRIORIT\xC4T 1): PERIODISIERUNG & PHASEN-DISZIPLIN ***
      Schau dir f\xFCr JEDEN EINZELNEN TAG im Abschnitt 'GEPLANTER TRAININGSPLAN' den Marker an.
      Der Marker entscheidet \xFCber Load-Strategie UND Text-Label.
      
      * **WENN '\u{1F680} [PHASE A: AUFBAU - REIZ SETZEN!]':**
        * **TEXT-REGEL (STRENG):** Du darfst das Wort "Entlastungswoche" NIEMALS verwenden. Auch nicht, wenn der Gesamtstatus "VORSICHT" ist.
        * **LOAD-REGEL:** Du sollst "normale" Reize (Z3, Z4, Intervalle) planen.
        * **CAUTION-EXCEPTION:** Falls der Gesamtstatus "ANPASSUNG VORSICHT" ist, darfst du den Load leicht reduzieren (max. -20%), aber du darfst NICHT in den "Tapering/Recovery"-Modus wechseln. Nenne es "Angepasster Aufbau" oder "Moderater Reiz", aber NICHT Entlastung.
        
      * **WENN '\u{1F6D1} [PHASE E: ENTLASTUNG - F\xDCSSE HOCH!]':**
        * **TEXT-REGEL:** Der Text MUSS zwingend mit "Entlastungswoche." beginnen.
        * **LOAD-REGEL:** Du MUSST den Load drastisch senken (Zielwert im Schnitt ca. ${targetRecoveryLoad}).
        * **VERBOT:** Plane KEINE hohen Reize (Z4/Z5) und keine langen Einheiten, egal wie gut die Readiness ist.

      * **REGEL 0: FIXED DAYS (PRIO 0):**
      * Tage mit 'x' sind unver\xE4nderlich. \xDCbernimm den 'original_load_ess'.

      * **REGEL 0.8: KRANKHEITS-R\xDCCKKEHR (RETURN-TO-PLAY - H\xD6CHSTE SICHERHEIT):**
        * Status: **${rtp_status}**
        * **WICHTIG:** Wenn der Status 'RTP_PHASE_X' ist, gilt das Protokoll. Wir steigen SMART ein (basierend auf Krankheitsdauer).
        * **DEINE PFLICHT:**
          1. **DATEN-INTEGRIT\xC4T (EXTREM WICHTIG):** Das Feld 'original_load_ess' darf NIEMALS ver\xE4ndert werden! Es MUSS exakt der Wert sein, der oben im Text hinter "[ORIGINAL-PLAN: ...]" steht.
          2. Schreibe den RTP-Zielwert AUSSCHLIESSLICH in das Feld 'empfohlener_load_ess'.
          3. **RAMP-UP LOGIK (ZWINGEND):** Halte dich strikt an die Anweisungen im Abschnitt 'GEPLANTER TRAININGSPLAN' oben (z.B. RTP Tag 8 -> Bike Load 80).
          4. **WICHTIG:** Wenn in der Anweisung oben steht "Setze Zone X", dann DARFST du in 'empfohlene_zone' NICHT 'Ruhetag' schreiben (au\xDFer es ist wirklich ein Ruhetag)!
          5. Begr\xFCndung MUSS lauten: "RTP-Schutz (Tag X/9). Protokoll aktiv."
        * Ein Plan mit Z4 oder hohem Load in dieser Phase ist ein FEHLER von dir!

      * **REGEL 1: HEUTIGER STATUS (isActivityDone == ${isActivityDone_Text}):**
      * WENN 'true': Setze 'empfohlener_load_ess' auf **${heutigerIstLoad}**. Tag ist vorbei.

      * **REGEL 2: PROBLEML\xD6SUNG (PRIO 4):**
      * -> **PRIO A (PHYSIO-VETO - KRITISCH):**
        * Gilt, wenn 'isRecoveryCritical == true' (siehe Info oben). 
        * **AKTION:** Load drastisch senken (Max 30% des Originals). Begr\xFCndung: "Physio-Veto".
      
      * -> **PRIO B (TE BALANCE ZU NIEDRIG < 30%):**
        * Wenn Veto inaktiv: Schlage Intensit\xE4t vor (Z4/Z5).

      * -> **PRIO C (ACWR ZU HOCH > 1.5):**
        * Reduziere Load.

      * **REGEL 7: DETAIL-TIEFE & BEGR\xDCNDUNG ('prognostizierte_auswirkung'):**
        * Sei nicht wortkarg!
        * **Inhalt:** Erkl\xE4re den physiologischen Nutzen (z.B. Mitochondrien, Laktattoleranz).
        
        * **ABWEICHUNGS-CHECK (WICHTIG):**
          * Vergleiche deinen 'empfohlener_load_ess' mit dem 'original_load_ess'.
          * Ist dein Vorschlag **> 20% niedriger** als das Original?
          * **DANN MUSS** der Text mit der Begr\xFCndung beginnen!
        
        * **In PHASE E:** Text MUSS mit "Entlastungswoche." beginnen.
        * **In PHASE A:** Text darf NICHT "Entlastungswoche" enthalten.

4.  **ANALYSIERE EINZELSCORES ('einzelscore_kommentare'):**
      * Erstelle ein Array von Objekten f\xFCr **JEDEN** der 10 Fitness-Scores.
      * **WICHTIG (V110):** Du MUSST f\xFCr **JEDEN** der folgenden Keys einen Kommentar liefern (exakte Schreibweise):
      * [${metricsChecklist}]
      
      *** SPEZIAL-BRIEFING "KEI" (Der Wahrheits-Detektor): ***
    * Dieser Wert misst: "Fitness-Gewinn (CTL-Trend) MINUS Aufwand (Strain)".
    * **INTERPRETATIONSHILFE F\xDCR DICH (KEI-LEGENDEN-SKALA):**
      * **Wert >= 8:** "High Efficiency" (sehr effizient, loben).
      * **Wert 3 bis < 8:** "Productive" (solider Aufbau).
      * **Wert 0 bis < 3:** "Maintenance / Transition" (Erhalt/\xDCbergang).
      * **Wert < 0:** Ineffizienter Reiz / Warnsignal.
        -> Schreib klar: "Aufwand hoch, Ertrag niedrig. Plan anpassen."
    * Gib f\xFCr jeden Score einen **aussagekr\xE4ftigen Kommentar (1-2 S\xE4tze)**.

REGELN F\xDCR 'plan_status':
- 'INTERVENTION': Wenn 'isRecoveryCritical == true' ODER RTP_STATUS aktiv.
- 'BEOBACHTUNG': Wenn 'isRecoveryBad == true' ODER ACWR > 1.3.
- 'FREIGABE': Wenn alle Scores GR\xDCN sind.

ANTWORTE AUSSCHLIESSLICH IM FOLGENDEN JSON-FORMAT (ohne Markdown):
{
  "empfehlung_zukunft": {
    "plan_status": "FREIGABE | BEOBACHTUNG | INTERVENTION", 
    "text": "DEINE AUSF\xDCHRLICHE GESAMTBEWERTUNG..."
  },
  "bewertung_ernaehrung_7d": {
    "text": "DEINE KURZE (MAX 2 S\xC4TZE) BEWERTUNG DER ERN\xC4HRUNG..."
  },
  "einzelscore_kommentare": [
    { "metrik": "RHR", "text_info": "Dein Kommentar zu RHR..." },
    { "metrik": "ACWR (Forecast)", "text_info": "Dein Kommentar zum ACWR..." }
    /* ... und ALLE anderen aus der Liste ... */
  ],
  "empfohlener_plan": [
    {
      "datum": "${cleanDate}",
      "tag": "${cleanTag}",
      "original_load_ess": 0, // KOPIERE HIER STUMPFE DEN WERT AUS "[ORIGINAL-PLAN: ...]"! NICHT \xC4NDERN!
      "empfohlener_load_ess": ${todayPlan_LoadRec},
      "empfohlene_zone": "${todayPlan_Zone}", 
      "beispiel_training_1": "${todayPlan_Sport1}",
      "beispiel_training_2": "${todayPlan_Sport2}",
      "prognostizierte_auswirkung": "${todayPlan_Text}",
      "wetterempfehlung": "Wetter egal."
    }
    /* ... 7 weitere Tage ... */
  ]
}


`;
  return prompt;
}
function briefingMessages(type, daten) {
  const res = briefingRoh(type, daten);
  if (res.messages.length && res.messages[0].role === "system") res.messages[0].content = String(res.messages[0].content) + "\n\n" + MONOTONIE_REGEL + erfahrungBlock(daten && daten.erfahrung);
  return res;
}
function briefingRoh(type, daten) {
  if (type === "status") {
    const status = daten.status;
    return {
      format: "text",
      messages: [
        { role: "system", content: "Du bist ein pr\xE4ziser deutschsprachiger Ausdauer-, Trainings- und Recovery-Coach." },
        { role: "user", content: [
          "Du bist Coach Kira.",
          "Analysiere die folgenden Status-Items knapp, konkret und handlungsorientiert.",
          "Sprich Marc direkt per Du an.",
          "Fokus: Training, Recovery, Schlaf, ACWR, Readiness, Ern\xE4hrung.",
          "Gib maximal 6 Bulletpoints plus eine klare Empfehlung f\xFCr heute.",
          "",
          JSON.stringify(status, null, 2)
        ].join("\n") }
      ]
    };
  }
  if (type === "plan") {
    const status = daten.status;
    const sim = daten.sim;
    return {
      format: "json",
      messages: [
        { role: "system", content: [
          "Du bist Coach Kira, eine deutschsprachige Sportwissenschaftlerin, Ausdauercoach und Recovery-Analystin.",
          "Du gibst pr\xE4zise, konservative, handlungsorientierte Empfehlungen f\xFCr Marc.",
          "Du nutzt ACWR, ATL, CTL, Training Readiness, Schlaf, HRV, RHR, TE und KEI/Smart Gains.",
          "Du sprichst Marc direkt per Du an.",
          "Du antwortest als valides JSON ohne Markdown-Codeblock."
        ].join(" ") },
        { role: "user", content: [
          "Erstelle ein Plan-Briefing f\xFCr das Plan Cockpit.",
          "Nutze Command-Center-Status, aktuelle PlanApp-Daten und Garmin-/Coach-Kira-Kontext.",
          "",
          "Output-Format exakt:",
          "{",
          '  "briefing": "mehrzeiliger deutscher Text mit konkreten Empfehlungen",',
          '  "recommendations": [',
          '    {"date":"YYYY-MM-DD", "day":"Mo", "recommended_load_ess":130, "recommended_zone":"Z2", "rationale":"kurze Begr\xFCndung"}',
          "  ]",
          "}",
          "",
          "Regeln:",
          "- recommendations muss genau die sichtbaren 14 Tage enthalten.",
          "- ESS nur realistische Werte 0 bis 300.",
          "- Wenn Recovery/Readiness kritisch ist: Last reduzieren.",
          "- Wenn ACWR > 1.3: keine zus\xE4tzliche Steigerung.",
          "- Wenn Schlafdauer oder Readiness schlecht ist: Z2 oder Off bevorzugen.",
          "- Keine medizinischen Diagnosen.",
          "",
          "STATUS:",
          JSON.stringify(status, null, 2),
          "",
          "PLAN_CONTEXT:",
          JSON.stringify(sim, null, 2)
        ].join("\n") }
      ]
    };
  }
  if (type === "consolidated") {
    const context = daten.context;
    const skillPerspectives = [
      { name: "/garmin-gesamtanalyse", role: "Gesamtbild aus Trainings-, Recovery-, Schlaf-, HRV-, Ern\xE4hrungs- und Belastungsdaten." },
      { name: "/garmin-anomalie-erkennung", role: "Auff\xE4lligkeiten, Datenbr\xFCche, unerwartete Spr\xFCnge, m\xF6gliche \xDCberlastung oder Recovery-Risiken." },
      { name: "/garmin-ernaehrung-trainingskorrelation", role: "Zusammenhang zwischen Energie-/Makroaufnahme, Defizit/\xDCberschuss, Load, Recovery und Schlaf." },
      { name: "/garmin-wochenplanung", role: "Adaptive 7- bis 14-Tage-Planung mit Load-, Zone- und Ruhetagslogik." },
      { name: "/garmin-trainingsstatus", role: "Trainingszustand, ATL/CTL/ACWR, TE Balance, KEI, Monotony/Strain und Fortschrittslogik." },
      { name: "/garmin-erholungsstatus", role: "Schlafdauer, Schlafscore, HRV, RHR, Training Readiness, Stress und Erholungsbedarf." },
      { name: "/garmin-gesundheitsstatus", role: "Gesundheitsnahe, nicht-diagnostische Risikoeinsch\xE4tzung und Watchlist." }
    ];
    return {
      format: "text",
      messages: [
        { role: "system", content: [
          "Du bist Coach Kira: eine deutschsprachige Sportwissenschaftlerin, Ausdauercoach, Recovery-Analystin und Garmin-Datenexpertin.",
          "Du synthetisierst mehrere Analyseperspektiven, statt nur Statuswerte nachzuerz\xE4hlen.",
          "Du sprichst Marc direkt per Du an.",
          "Keine medizinischen Diagnosen. Bei Gesundheitsrisiken vorsichtig formulieren.",
          "Nutze klare \xDCberschriften und Bulletpoints. Keine JSON-Antwort."
        ].join(" ") },
        { role: "user", content: [
          "Erstelle eine konsolidierte Coach-Kira-Lagebeurteilung.",
          "",
          "Die folgenden Garmin-Skill-Perspektiven sollen gedanklich verwendet, synthetisiert und zusammengefasst werden:",
          JSON.stringify(skillPerspectives, null, 2),
          "",
          "Pflichtstruktur:",
          "## Gesamtbild",
          "## Auff\xE4lligkeiten und Anomalien",
          "## Erholung und Gesundheitsstatus",
          "## Training und Belastungssteuerung",
          "## Ern\xE4hrung und Trainingskorrelation",
          "## Plan 14 Tage",
          "## Empfehlung f\xFCr heute",
          "## Watchlist",
          "",
          "Konkrete Regeln:",
          "- Nenne konkrete Werte, wenn sie entscheidungsrelevant sind.",
          "- Gib klare Handlungspriorit\xE4ten.",
          "- Weise auf Widerspr\xFCche hin, z.B. guter KEI aber niedrige Readiness.",
          "- Verdichte, nicht ausufern.",
          "- Maximal 900 W\xF6rter.",
          "",
          "DATENKONTEXT:",
          JSON.stringify(context, null, 2)
        ].join("\n") }
      ]
    };
  }
  throw new Error("Unbekannter Briefing-Typ: " + type);
}
function telegramMessages(frage, ctx) {
  const sys = [
    "Du bist Coach Kira, Marcs Sportwissenschaftlerin und Trainingscoach \u2013 wie eine gute Freundin, direkt und per Du.",
    "Du antwortest auf Telegram: kurz (max. 6 S\xE4tze), konkret, auf Deutsch, ohne Markdown-Tabellen und ohne Emojis.",
    "Nutze die Daten unten. Erfinde keine Werte. Keine medizinischen Diagnosen.",
    MONOTONIE_REGEL
  ].join("\n") + erfahrungBlock(ctx.erfahrung);
  const daten = [
    ctx.status ? "STATUS (Cockpit):\n" + JSON.stringify(ctx.status).slice(0, 6e3) : "",
    ctx.vorausblick ? "VORAUSBLICK:\n" + String(ctx.vorausblick).slice(0, 3e3) : ""
  ].filter(Boolean).join("\n\n");
  const msgs = [{ role: "system", content: sys + (daten ? "\n\n" + daten : "") }];
  (ctx.verlauf || []).slice(-3).forEach((v) => {
    msgs.push({ role: "user", content: v.frage });
    msgs.push({ role: "assistant", content: v.antwort });
  });
  msgs.push({ role: "user", content: frage });
  return msgs;
}

// dashboard/src/screens/RecoveryPanel.tsx
var import_react7 = __toESM(require_react(), 1);

// dashboard/src/components/BefindenHinweis.tsx
var import_react4 = __toESM(require_react(), 1);

// dashboard/src/lib/engine.ts
var import_react3 = __toESM(require_react(), 1);

// dashboard/src/lib/garmin.ts
var import_react = __toESM(require_react(), 1);
var CACHE_MS = 10 * 60 * 1e3;

// dashboard/src/lib/useErfahrung.ts
var import_react2 = __toESM(require_react(), 1);

// dashboard/src/lib/monotony.ts
var MONO = { GUT: 1.5, NORMAL: 2, UNGUENSTIG: 2.5, WARNUNG: 3 };
function monoStufe(v) {
  if (v == null || !Number.isFinite(v)) return null;
  if (v > MONO.WARNUNG) return "warnung";
  if (v > MONO.UNGUENSTIG) return "unguenstig";
  if (v > MONO.NORMAL) return "aufmerksam";
  if (v >= MONO.GUT) return "normal";
  return "gut";
}
function refLoad(loads) {
  const xs = loads.slice(-28).filter((x) => Number.isFinite(x));
  return xs.length ? xs.reduce((a, b) => a + b, 0) / xs.length : 0;
}
function isDemanding(load, ref) {
  return load >= Math.max(60, 0.7 * ref);
}
var RELIEF_ANTEIL = 0.6;
function isRelief(load, ref) {
  return load <= Math.max(40, RELIEF_ANTEIL * ref);
}
var f2 = (v) => v.toFixed(2).replace(".", ",");
function assessMonotony(inp) {
  const loads = inp.loads.filter((x) => Number.isFinite(x));
  const ref = refLoad(loads);
  let streak = 0;
  let demanding = 0;
  for (let i = loads.length - 1; i >= 0; i--) {
    if (isRelief(loads[i], ref)) break;
    streak++;
    if (isDemanding(loads[i], ref)) demanding++;
  }
  const hist = [...inp.monoHistory || [], inp.mono];
  let above25 = 0;
  for (let i = hist.length - 1; i >= 0; i--) {
    const v = hist[i];
    if (v == null || v <= MONO.UNGUENSTIG) break;
    above25++;
  }
  const weekLoad = loads.slice(-7).reduce((a, b) => a + b, 0);
  const four = loads.slice(-28);
  const avgWeek = four.length >= 14 ? four.reduce((a, b) => a + b, 0) / four.length * 7 : weekLoad;
  const weekLoadHigh = avgWeek > 0 && weekLoad >= 1.1 * avgWeek;
  const m = inp.mono;
  const prevMono = inp.monoHistory && inp.monoHistory.length ? inp.monoHistory[inp.monoHistory.length - 1] : null;
  const over20 = m != null && m > MONO.NORMAL && prevMono != null && prevMono > MONO.NORMAL;
  const stufe = monoStufe(m);
  const strain = m != null ? Math.round(weekLoad * m) : null;
  const base = { stufe, streakNoRelief: streak, demandingInStreak: demanding, daysAbove25: above25, weekLoad, strain, weekLoadHigh };
  const hinweis = streak >= 5 && demanding >= 3 ? `Hinweis: ${streak} Tage ohne deutliche Entlastung \u2013 keine Einschr\xE4nkung, auf Schlaf, HRV, Ruhepuls und Beine achten.` : void 0;
  const res = (cap, rest, text) => ({ ...base, cap, restWithin48h: rest, text, ...hinweis ? { hinweis } : {} });
  if (m == null) {
    return res(null, false, "");
  }
  const reliefDone = streak < 2;
  if (m > MONO.WARNUNG && !reliefDone)
    return res("EASY", true, `Monotonie ${f2(m)} > 3,0 (Warnzone) \u2192 echten Entlastungstag oder deutliche Reduktion einplanen.`);
  if (m > MONO.UNGUENSTIG && (above25 >= 2 || demanding >= 4) && !reliefDone)
    return res("EASY", true, `Monotonie bleibt \xFCber 2,5 (${above25 >= 2 ? `${above25} Tage` : `${demanding} fordernde Tage`}) \u2192 innerhalb 24\u201348 h Ruhetag oder sehr kurzer Regenerationsreiz.`);
  if (m > MONO.UNGUENSTIG && reliefDone && (above25 >= 2 || m > MONO.WARNUNG))
    return res("TRAIN", false, `Monotonie ${f2(m)} noch hoch, Entlastung in den letzten 48 h ist erfolgt \u2192 keine Qualit\xE4tseinheit, Rhythmus hart/locker beibehalten.`);
  if (over20 && demanding >= 4)
    return res("EASY", false, `Monotonie seit 2 Tagen \xFCber 2,0 (${f2(m)}) nach ${demanding} fordernden Tagen \u2192 n\xE4chster Tag sehr locker oder frei.`);
  if (over20 && inp.recoveryBad)
    return res("TRAIN", false, `Monotonie seit 2 Tagen \xFCber 2,0 (${f2(m)}) und Erholungswerte schlechter \u2192 keine Qualit\xE4tseinheit durchdr\xFCcken.`);
  if (demanding >= 3 && inp.recoveryBad)
    return res("EASY", false, `${demanding} fordernde Tage in Folge und Erholungswerte schlechter \u2192 Belastung reduzieren.`);
  if (over20 && weekLoadHigh)
    return res("TRAIN", false, `Monotonie seit 2 Tagen \xFCber 2,0 (${f2(m)}) bei hoher Wochenlast (Strain ${strain}) \u2192 keine zus\xE4tzliche Qualit\xE4t.`);
  if (m > MONO.NORMAL)
    return res(null, false, `Monotonie ${f2(m)}${over20 ? " seit 2 Tagen \xFCber 2,0" : " \u2013 erster Tag \xFCber 2,0"} \xB7 ${demanding === 1 ? "1 fordernder Tag" : `${demanding} fordernde Tage`} in Folge: beobachten, keine Einschr\xE4nkung.`);
  return res(null, false, stufe === "normal" ? `Monotonie ${f2(m)} normal.` : "");
}
var MONOTONIE_REGEL_TEXT = [
  "Monotonie-Regel: Ein einzelner hoher Monotonie-Wert ist meist kein Problem. Ma\xDFgeblich ist die rollierende 7-Tage-Monotonie (\xD8/\u03C3 der Tageslasten) im Verlauf.",
  "Zonen: < 1,5 gute Variation; 1,5\u20132,0 normal; > 2,0 aufmerksam (vor allem bei hoher Wochenlast); > 2,5 ung\xFCnstig, wenn es l\xE4nger anh\xE4lt; > 3,0 Warnzone.",
  "Muster: 1\u20132 \xE4hnliche Tage normal; 3\u20134 \xE4hnliche, fordernde Tage beobachten (bei schlechterem Schlaf, HRV, Ruhepuls, schweren Beinen oder weniger Motivation reduzieren); 5\u20137 Tage ohne deutliche Entlastung beobachten (nur Hinweis, senkt das Votum nicht \u2013 in Marcs Historie ohne messbare Folgen); mehrere Wochen hohe Monotonie = Risiko f\xFCr Erm\xFCdung, Stagnation, \xDCberlastung, Infekte.",
  "Steuerung: Ein hoher Wert an einem Tag ist praktisch bedeutungslos. Liegt die 7-Tage-Monotonie an 2 Tagen in Folge \xFCber etwa 2,0 und hast du schon 4\u20135 fordernde Tage gesammelt \u2192 n\xE4chster Tag sehr locker oder frei. Als Entlastung z\xE4hlt ein lockerer oder freier Tag (h\xF6chstens 40 ESS bzw. 60 % deiner \xFCblichen Tageslast). Bleibt der Wert \xFCber 2,5 \u2192 innerhalb 24\u201348 h echter Ruhetag oder sehr kurzer Regenerationsreiz. \xDCber 3,0 \u2192 mindestens ein echter Entlastungstag oder deutliche Reduktion. Hohe Monotonie plus schlechte Erholungswerte \u2192 keine Qualit\xE4tseinheit durchdr\xFCcken.",
  "Entscheidend ist die Kombination mit der Gesamtlast (Strain = Wochenlast \xD7 Monotonie). Monotonie senkst du durch klaren Rhythmus: harte Reize klar hart, lockere Tage wirklich locker, mindestens ein echter Erholungstag bei hoher Gesamtlast \u2013 nicht durch weniger Gesamtlast."
].join("\n");

// dashboard/src/lib/chatContext.ts
var KIRA_SYSTEM = [
  "Du bist Kira, Sportwissenschaftlerin und pers\xF6nliche Ausdauer-Trainerin von Marc (Jahrgang 1975; Rennrad, Gravel, Laufen, Bergsport).",
  "Sprich Marc per Du an, antworte auf Deutsch, klar und konkret, ohne Emojis.",
  "St\xFCtze dich zuerst auf die mitgelieferten Cockpit-Daten (JSON unten). Nenne die Zahlen, auf die du dich beziehst.",
  "Erfinde keine Werte. Wenn Daten fehlen oder unklar sind, sag das offen.",
  'Halte dich an die Regeln aus dem Feld "regeln" (Tagesentscheid, Zielbereiche).',
  "Gib am Ende eine klare Empfehlung (z. B. Einheit, Dauer, Zone oder Pause).",
  "Bei Schmerzen, Krankheit oder medizinischen Fragen: auf Arzt bzw. Physiotherapie verweisen."
].join("\n");

// dashboard/src/lib/engine.ts
var KEY = "ck_kira_engine_v2";
function read() {
  try {
    const v = localStorage.getItem(KEY);
    return v === "pplx" ? "pplx" : "local";
  } catch {
    return "local";
  }
}
var store = read();

// dashboard/src/lib/api.ts
function dedupeByDateKeepLast(rows) {
  const map = /* @__PURE__ */ new Map();
  rows.forEach((row) => map.set(row.date, row));
  return Array.from(map.values()).sort((a, b) => a.date.localeCompare(b.date));
}
function toNum(v) {
  if (v === null || v === void 0 || v === "") return null;
  const n = typeof v === "number" ? v : Number(v);
  return Number.isFinite(n) ? n : null;
}

// dashboard/src/components/BefindenHinweis.tsx
var import_jsx_runtime = __toESM(require_jsx_runtime(), 1);

// dashboard/src/lib/format.ts
function fmtNum(v, digits = 0) {
  if (v === null || v === void 0 || !Number.isFinite(v)) return "\u2014";
  return v.toLocaleString("de-DE", {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits
  });
}

// dashboard/src/components/UI.tsx
var import_jsx_runtime2 = __toESM(require_jsx_runtime(), 1);

// dashboard/src/screens/CorrelationPanel.tsx
var import_react5 = __toESM(require_react(), 1);
var import_jsx_runtime3 = __toESM(require_jsx_runtime(), 1);
var VARS = [
  { key: "readiness", label: "Readiness", short: "Ready", unit: "" },
  { key: "hrv", label: "HRV", short: "HRV", unit: "ms" },
  { key: "rhr", label: "RHR", short: "RHR", unit: "bpm" },
  { key: "sleepH", label: "Schlafdauer", short: "Schlaf h", unit: "h" },
  { key: "sleepScore", label: "Schlafscore", short: "Score", unit: "" },
  { key: "essDay", label: "Last", short: "Last", unit: "ESS" }
];
var VAR = Object.fromEntries(VARS.map((v) => [v.key, v]));

// dashboard/src/components/GarminSportProfilTile.tsx
var import_react6 = __toESM(require_react(), 1);
var import_jsx_runtime4 = __toESM(require_jsx_runtime(), 1);

// dashboard/src/screens/RecoveryPanel.tsx
var import_jsx_runtime5 = __toESM(require_jsx_runtime(), 1);
var SLEEP_GOAL_H = 7.5;
var RHR_BASELINE_DAYS = 28;
var RHR_BASELINE_MIN_VALUES = 7;
var RECOVERY_FETCH_METRICS = [
  "Garmin_Training_Readiness",
  "rhr_bpm",
  "sleep_hours",
  "sleep_score_0_100",
  "hrv_status",
  "hrv_threshholds",
  "load_fb_day",
  "Monotony7",
  "befinden_1_5",
  "fbACWR_obs",
  "coachE_ACWR_forecast",
  "coachE_ESS_day",
  "Sport_x",
  "Zone",
  "Target_Aerobic_TE",
  "Target_Anaerobic_TE",
  "activity_done",
  "coachE_ATL_forecast",
  "coachE_CTL_forecast",
  "Aerobic_TE",
  "Anaerobic_TE"
];
var ACWR_WARN = 1.3;
var ACWR_HIGH = 1.5;
var VOTE_ORDER = ["REST", "EASY", "TRAIN", "QUALITY"];
var VOTE_LABEL = {
  QUALITY: "Quality",
  TRAIN: "Train",
  EASY: "Easy",
  REST: "Rest"
};
function parseThresholds(v) {
  if (typeof v !== "string" || !v.includes(";")) return [null, null];
  const [a, b] = v.split(";").map((x) => Number(String(x).trim().replace(",", ".")));
  return [Number.isFinite(a) ? a : null, Number.isFinite(b) ? b : null];
}
function mean(values) {
  return values.length ? values.reduce((a, b) => a + b, 0) / values.length : null;
}
function levelReadiness(v) {
  if (v == null) return "leer";
  if (v >= 75) return "gruen";
  if (v >= 55) return "gelb";
  if (v >= 40) return "orange";
  return "rot";
}
function levelRhrDelta(d) {
  if (d == null) return "leer";
  if (d <= -5) return "grau";
  if (d <= 1) return "gruen";
  if (d < 3) return "gelb";
  if (d < 5) return "orange";
  return "rot";
}
function levelSleep(h) {
  if (h == null) return "leer";
  const ratio = h / SLEEP_GOAL_H;
  if (ratio >= 0.95) return "gruen";
  if (ratio >= 0.85) return "gelb";
  if (ratio >= 0.8) return "orange";
  return "rot";
}
function levelHrv(v, low) {
  if (v == null || low == null) return "leer";
  if (v < low - 2) return "rot";
  if (v < low) return "orange";
  return "gruen";
}
function levelBefinden(v) {
  if (v == null) return "leer";
  if (v >= 4) return "gruen";
  if (v >= 3) return "gelb";
  if (v >= 2) return "orange";
  return "rot";
}
function capVote(vote, max) {
  return VOTE_ORDER.indexOf(vote) > VOTE_ORDER.indexOf(max) ? max : vote;
}
function decideVote(d) {
  if (d.readiness == null) return { vote: null, reasons: ["Readiness fehlt (Morgenwerte noch nicht erfasst)."], bind: null };
  const reasons = [];
  let vote = d.readiness >= 75 ? "QUALITY" : d.readiness >= 55 ? "TRAIN" : d.readiness >= 40 ? "EASY" : "REST";
  reasons.push(`Readiness ${fmtNum(d.readiness)} \u2192 Basis ${VOTE_LABEL[vote]}.`);
  let bind = "readiness";
  const senke = (neu, grund) => {
    if (VOTE_ORDER.indexOf(neu) < VOTE_ORDER.indexOf(vote)) bind = grund;
    vote = neu;
  };
  const { rhr, sleep, hrv } = d.levels;
  const warn = [rhr, sleep, hrv].filter((l) => l === "orange" || l === "rot").length;
  if ([rhr, sleep, hrv].includes("rot")) {
    const before = vote;
    senke(capVote(vote, "TRAIN"), "rot");
    if (vote !== before) reasons.push("Ein Einzelsignal ist rot \u2192 keine Qualit\xE4tseinheit.");
  }
  const sleepRatio = d.sleepH != null ? d.sleepH / SLEEP_GOAL_H : null;
  if (d.rhrDelta != null && d.rhrDelta >= 5 && sleepRatio != null && sleepRatio < 0.8) {
    senke(capVote(vote, "EASY"), "rhr_schlaf");
    reasons.push("RHR \u2265 +5 bpm und Schlaf < 80 % des Ziels \u2192 maximal Easy.");
  }
  if (hrv === "rot") {
    senke(capVote(vote, "EASY"), "hrv");
    reasons.push("HRV unter dem Normalband (HRV-Alarm) \u2192 maximal Easy.");
  }
  if (warn >= 2) {
    senke(capVote(vote, "EASY"), "warn2");
    reasons.push(`${warn} Warnsignale gleichzeitig \u2192 maximal Easy.`);
  }
  if (d.rhrDelta != null && d.rhrDelta >= 5 && d.readiness < 55) {
    senke("REST", "rhr_ready");
    reasons.push("RHR \u2265 +5 bpm bei niedriger Readiness \u2192 Rest.");
  }
  if (d.acwr != null && d.acwr > ACWR_HIGH) {
    const before = vote;
    senke(capVote(vote, "EASY"), "acwr");
    if (vote !== before) reasons.push(`ACWR ${fmtNum(d.acwr, 2)} > ${fmtNum(ACWR_HIGH, 1)} \u2192 maximal Easy.`);
  } else if (d.acwr != null && d.acwr > ACWR_WARN) {
    const before = vote;
    senke(capVote(vote, "TRAIN"), "acwr");
    if (vote !== before) reasons.push(`ACWR ${fmtNum(d.acwr, 2)} > ${fmtNum(ACWR_WARN, 1)} \u2192 keine Qualit\xE4tseinheit.`);
  }
  const ma = d.monoAssess;
  if (ma && ma.cap) {
    const before = vote;
    senke(capVote(vote, ma.cap), "mono");
    if (vote !== before || ma.restWithin48h) reasons.push(ma.text);
  } else if (ma && ma.text && ma.stufe && ma.stufe !== "gut" && ma.stufe !== "normal") {
    reasons.push(ma.text);
  }
  if (ma && ma.hinweis) reasons.push(ma.hinweis);
  if (d.befinden != null) {
    if (d.befinden <= 1) {
      senke("REST", "befinden");
      reasons.push("Befinden 1/5 \u2192 Rest.");
    } else if (d.befinden <= 2) {
      senke(capVote(vote, "EASY"), "befinden");
      reasons.push("Befinden 2/5 \u2192 maximal Easy.");
    } else if (d.befinden <= 3) {
      const before = vote;
      senke(capVote(vote, "TRAIN"), "befinden");
      if (vote !== before) reasons.push("Befinden 3/5 \u2192 keine Qualit\xE4tseinheit.");
    }
  }
  return { vote, reasons, bind };
}
function buildDays(raw) {
  const rows = dedupeByDateKeepLast(
    raw.filter((r) => r && typeof r.date === "string").map((r) => ({ ...r, date: String(r.date).slice(0, 10) }))
  ).sort((a, b) => a.date < b.date ? -1 : 1);
  return rows.map((r, idx) => {
    const readiness = toNum(r.Garmin_Training_Readiness);
    const rhr = toNum(r.rhr_bpm);
    const sleepH = toNum(r.sleep_hours);
    const sleepScore = toNum(r.sleep_score_0_100);
    const hrvRaw = toNum(r.hrv_status);
    const hrv = hrvRaw != null && hrvRaw >= 10 && hrvRaw <= 200 ? hrvRaw : null;
    const [hrvLow, hrvHigh] = parseThresholds(r.hrv_threshholds);
    const prior = rows.slice(Math.max(0, idx - RHR_BASELINE_DAYS), idx).map((p) => toNum(p.rhr_bpm)).filter((v) => v != null && v > 0);
    const rhrBaseline = prior.length >= RHR_BASELINE_MIN_VALUES ? mean(prior) : null;
    const rhrDelta = rhr != null && rhrBaseline != null ? rhr - rhrBaseline : null;
    const sleepWindow = rows.slice(Math.max(0, idx - 6), idx + 1).map((p) => toNum(p.sleep_hours)).filter((v) => v != null && v > 0);
    const sleepAvg7 = sleepWindow.length >= 3 ? mean(sleepWindow) : null;
    const befindenRaw = toNum(r.befinden_1_5);
    const befinden = befindenRaw != null && befindenRaw >= 1 && befindenRaw <= 5 ? Math.round(befindenRaw) : null;
    const base = {
      date: r.date,
      readiness,
      rhr,
      rhrBaseline,
      rhrDelta,
      sleepH,
      sleepScore,
      sleepAvg7,
      hrv,
      hrvLow,
      hrvHigh,
      load: toNum(r.load_fb_day),
      // Belastung bis GESTERN (heutige Zeile enthält bereits die heutige Einheit)
      monotony: idx > 0 ? toNum(rows[idx - 1].Monotony7) : null,
      monoLoads: rows.slice(Math.max(0, idx - 28), idx).map((p) => toNum(p.load_fb_day) ?? toNum(p.coachE_ESS_day) ?? 0),
      monoHist: rows.slice(Math.max(0, idx - 8), Math.max(0, idx - 1)).map((p) => toNum(p.Monotony7)),
      monoAssess: null,
      acwr: idx > 0 ? toNum(rows[idx - 1].fbACWR_obs) ?? toNum(rows[idx - 1].coachE_ACWR_forecast) : null,
      activityDone: String(r.activity_done || "").trim().toLowerCase() === "x",
      essDay: toNum(r.coachE_ESS_day),
      atlPost: toNum(r.coachE_ATL_forecast),
      ctlPost: toNum(r.coachE_CTL_forecast),
      monotonyPost: toNum(r.Monotony7),
      acwrPost: toNum(r.fbACWR_obs) ?? toNum(r.coachE_ACWR_forecast),
      actual: {
        teAe: toNum(r.Aerobic_TE),
        teAn: toNum(r.Anaerobic_TE),
        load: toNum(r.coachE_ESS_day),
        sport: String(r.Sport_x || ""),
        zone: String(r.Zone || "")
      },
      befinden,
      hasMorningData: readiness != null || rhr != null || sleepH != null,
      levels: {
        readiness: levelReadiness(readiness),
        rhr: levelRhrDelta(rhrDelta),
        sleep: levelSleep(sleepH),
        hrv: levelHrv(hrv, hrvLow),
        befinden: levelBefinden(befinden)
      }
    };
    const recoveryBad = [base.levels.rhr, base.levels.sleep, base.levels.hrv].some((l) => l === "orange" || l === "rot") || base.befinden != null && base.befinden <= 2;
    base.monoAssess = assessMonotony({ mono: base.monotony, loads: base.monoLoads, monoHistory: base.monoHist, recoveryBad });
    const { vote, reasons, bind } = decideVote(base);
    return { ...base, vote, reasons, bind };
  });
}
function plannedDemand(p) {
  const zone = (p.zone || "").toUpperCase();
  const sport = (p.sport || "").toLowerCase();
  if (p.load <= 0 || sport === "off" || zone === "OFF") return "REST";
  if (p.teAn >= 1 || p.teAe >= 4.5 || /Z4|Z5/.test(zone) || p.load >= 200) return "QUALITY";
  if (p.teAe >= 3 || /Z3/.test(zone) || p.load >= 120) return "TRAIN";
  return "EASY";
}
function demandOf(load, sport, zone, teAe, teAn) {
  return plannedDemand({
    date: "",
    load: load ?? 0,
    sport,
    zone,
    teAe: teAe ?? 0,
    teAn: teAn ?? 0,
    locked: false
  });
}

// dashboard/src/lib/votumErfahrung.ts
var GRUND_LABEL = {
  readiness: "niedriger Bereitschaft",
  rot: "rotes Einzelsignal",
  rhr_schlaf: "Ruhepuls hoch + wenig Schlaf",
  hrv: "HRV-Alarm",
  warn2: "zwei Warnsignale",
  rhr_ready: "Ruhepuls hoch + niedrige Bereitschaft",
  acwr: "ACWR",
  mono: "Monotonie",
  befinden: "Befinden"
};
var SCHUTZ = ["hrv", "rhr_ready", "befinden", "rot"];
var MIN_N = 8;
var MIN_N_ALLE = 12;
var pos = (v) => typeof v === "number" && Number.isFinite(v) && v > 0 ? v : null;
var mean2 = (xs) => xs.length ? xs.reduce((a, b) => a + b, 0) / xs.length : 0;
var r2 = (v) => Math.round(v * 100) / 100;
function zeilen(days) {
  const feld = (d, f) => d ? pos(d[f]) : null;
  const score = (i, k) => {
    const n = days[i + k];
    if (!n || feld(n, "readiness") == null) return null;
    const win = days.slice(Math.max(0, i - 6), i + 1);
    const dl = (f) => {
      const w = feld(n, f);
      const b = win.map((x) => feld(x, f)).filter((v) => v != null);
      return w != null && b.length ? w - mean2(b) : 0;
    };
    return dl("readiness") / 10 + dl("hrv") / 5 - dl("rhr") / 3 + dl("sleepScore") / 10;
  };
  const out = [];
  days.forEach((d, i) => {
    if (!d.vote || d.readiness == null) return;
    const next = days[i + 1];
    if (!next || Date.parse(next.date) - Date.parse(d.date) !== 864e5) return;
    const last = d.actual.load ?? d.load;
    if (last == null) return;
    const s1 = score(i, 1);
    if (s1 == null) return;
    const A = demandOf(last, d.actual.sport, d.actual.zone, d.actual.teAe, d.actual.teAn);
    const folge = days.slice(i + 1, i + 4);
    out.push({
      i,
      datum: d.date,
      V: d.vote,
      A,
      bind: d.bind || "readiness",
      ready: d.readiness,
      s: [s1, score(i, 2), score(i, 3)],
      tief: folge.some((x) => x.readiness != null && x.readiness < 55)
    });
  });
  return out;
}
function rng(seed) {
  let s = seed % 2147483647;
  return () => (s = s * 16807 % 2147483647) / 2147483647;
}
function bootstrap(xs, seed = 7, n = 1500) {
  if (xs.length < 3) return [-Infinity, Infinity];
  const r = rng(seed);
  const m = [];
  for (let b = 0; b < n; b++) {
    let s = 0;
    for (let j = 0; j < xs.length; j++) s += xs[Math.floor(r() * xs.length)];
    m.push(s / xs.length);
  }
  m.sort((a, b) => a - b);
  return [m[Math.floor(0.05 * n)], m[Math.floor(0.95 * n)]];
}
function klassifiziere(n, k, kLo, kHi, tiefDiff, minN = MIN_N) {
  if (n < minN) return "unklar";
  if (k <= -0.5 || kHi < 0 || tiefDiff >= 0.25) return "kostet";
  if (k >= 0.3 && kLo > -0.5 && tiefDiff <= 0.15) return "vertraeglich";
  return "unklar";
}
function erfahrungAusTagen(days) {
  const z = zeilen(days);
  const ref = z.filter((r) => VOTE_ORDER.indexOf(r.A) <= VOTE_ORDER.indexOf(r.V));
  const haerter = z.filter((r) => VOTE_ORDER.indexOf(r.A) > VOTE_ORDER.indexOf(r.V));
  const faelle = [];
  for (const r of haerter) {
    let c = [];
    for (const w of [7, 10, 14]) {
      c = ref.filter((x) => x.A === r.A && Math.abs(x.ready - r.ready) <= w && x.datum !== r.datum);
      if (c.length >= 4) break;
    }
    if (!c.length) continue;
    const d = [0, 1, 2].map((k) => {
      const cs = c.map((x) => x.s[k]).filter((v) => v != null);
      return r.s[k] != null && cs.length ? r2(r.s[k] - mean2(cs)) : null;
    });
    const ds = d.filter((v) => v != null);
    faelle.push({
      datum: r.datum,
      votum: r.V,
      anspruch: r.A,
      bind: r.bind,
      readiness: r.ready,
      d,
      k: r2(mean2(ds)),
      tief: r.tief,
      refTief: r2(mean2(c.map((x) => x.tief ? 1 : 0))),
      nRef: c.length
    });
  }
  const gruppe = (key, bind, von, nach, fs, minN) => {
    const ks = fs.map((f) => f.k);
    const [kLo, kHi] = bootstrap(ks, key.length * 31 + fs.length);
    const k = mean2(ks);
    const tief = mean2(fs.map((f) => f.tief ? 1 : 0));
    const refTief = mean2(fs.map((f) => f.refTief));
    const dm = (j) => {
      const v = fs.map((f) => f.d[j]).filter((x) => x != null);
      return v.length ? r2(mean2(v)) : null;
    };
    return { key, bind, von, nach, n: fs.length, d1: dm(0), d2: dm(1), d3: dm(2), k: r2(k), kLo: r2(kLo), kHi: r2(kHi), tief: r2(tief), refTief: r2(refTief), klasse: klassifiziere(fs.length, k, kLo, kHi, tief - refTief, minN) };
  };
  const gruppen = [];
  const keys = /* @__PURE__ */ new Map();
  const keysAlle = /* @__PURE__ */ new Map();
  for (const f of faelle) {
    const k = `${f.bind}|${f.votum}>${f.anspruch}`;
    keys.set(k, [...keys.get(k) || [], f]);
    if (SCHUTZ.includes(f.bind)) continue;
    const ka = `*|${f.votum}>${f.anspruch}`;
    keysAlle.set(ka, [...keysAlle.get(ka) || [], f]);
  }
  for (const [k, fs] of keys) gruppen.push(gruppe(k, fs[0].bind, fs[0].votum, fs[0].anspruch, fs, MIN_N));
  for (const [k, fs] of keysAlle) gruppen.push(gruppe(k, "*", fs[0].votum, fs[0].anspruch, fs, MIN_N_ALLE));
  gruppen.sort((a, b) => b.n - a.n);
  return { stand: days.length ? days[days.length - 1].date : "", faelle, gruppen, nTage: z.length, nHaerter: faelle.length };
}
function erfahrungAusRohdaten(raw) {
  return erfahrungAusTagen(buildDays(raw));
}
var fz = (v, d = 1) => v == null ? "\u2013" : `${v > 0 ? "+" : v < 0 ? "\u2212" : "\xB1"}${Math.abs(v).toLocaleString("de-DE", { maximumFractionDigits: d, minimumFractionDigits: d })}`;
var pct = (v) => `${Math.round(v * 100)} %`;
var gName = (g) => `${VOTE_LABEL[g.von]}\u2192${VOTE_LABEL[g.nach]}${g.bind === "*" ? "" : ` bei ${GRUND_LABEL[g.bind]}`}`;
function gruppenText(g) {
  const urteil = g.klasse === "vertraeglich" ? "vertr\xE4glich" : g.klasse === "kostet" ? "hat Erholung gekostet" : "kein klares Muster";
  return `${gName(g)} (${g.n}\xD7): ${urteil} \u2013 Erholung Tag 1/2/3 ${fz(g.d1)}/${fz(g.d2)}/${fz(g.d3)} gegen\xFCber vergleichbaren erlaubten Tagen, Tief unter 55 in ${pct(g.tief)} der F\xE4lle (sonst ${pct(g.refTief)}).`;
}
function erfahrungText(e, max = 8) {
  if (!e) return "";
  const gs = e.gruppen.filter((g) => g.klasse !== "unklar").sort((a, b) => Math.abs(b.k) * b.n - Math.abs(a.k) * a.n).slice(0, max);
  if (!gs.length) return "";
  return [
    `Stand ${e.stand}: ${e.nHaerter} Tage, an denen Marc h\xE4rter trainiert hat als das Regel-Votum, verglichen mit Tagen gleicher Einheitsstufe und \xE4hnlicher Ausgangsbereitschaft, an denen das Votum sie erlaubte.`,
    ...gs.map((g) => "- " + gruppenText(g)),
    "Nutze das f\xFCr Bewertung und Empfehlung: Wo H\xE4rter bei Marc vertr\xE4glich war, darfst du Spielraum nennen; wo es gekostet hat, klar warnen und die Folgetage einplanen. Schutzsignale (HRV-Alarm, Ruhepuls hoch bei niedriger Bereitschaft, Befinden \u2264 2, rotes Einzelsignal) bleiben bindend."
  ].join("\n");
}

// tools/mac-runner/kira_prompts_cli.ts
var APPS_SCRIPT = "https://script.google.com/macros/s/AKfycbxCEN11KRlFaLL7uVJyeBLCrRJVmfBWagSmqvyJ8Ci7nwxi8HbolzTy23Z-G2mivC2h/exec";
async function erfahrung() {
  try {
    const ctl = new AbortController();
    const to = setTimeout(() => ctl.abort(), 3e4);
    const url = APPS_SCRIPT + "?mode=chartData&range=9999d&metrics=" + RECOVERY_FETCH_METRICS.join(",");
    const j = await (await fetch(url, { redirect: "follow", signal: ctl.signal })).json();
    clearTimeout(to);
    const e = erfahrungAusRohdaten(j.data || []);
    return { text: erfahrungText(e), gruppen: e.gruppen.filter((g) => g.n >= 8) };
  } catch {
    return { text: "", gruppen: [] };
  }
}
async function lesen() {
  const teile = [];
  for await (const t of process.stdin) teile.push(t);
  return Buffer.concat(teile).toString("utf8");
}
(async () => {
  const art = process.argv[2];
  const d = JSON.parse(await lesen() || "{}");
  let out;
  const erf = art === "supervisor" || art === "briefing" || art === "telegram" || art === "erfahrung" ? await erfahrung() : { text: "", gruppen: [] };
  if (art === "supervisor") out = supervisorPrompt({ ...d, erfahrung: d.erfahrung ?? erf.text });
  else if (art === "briefing") out = briefingMessages(d.type, { ...d.daten || {}, erfahrung: (d.daten || {}).erfahrung ?? erf.text });
  else if (art === "telegram") out = { messages: telegramMessages(String(d.frage || ""), { ...d, erfahrung: d.erfahrung ?? erf.text }) };
  else if (art === "erfahrung") out = erf;
  else throw new Error("Unbekannte Art: " + art + " (supervisor | briefing | telegram | erfahrung)");
  process.stdout.write(JSON.stringify(out));
})().catch((e) => {
  process.stderr.write(String(e && e.message || e));
  process.exit(1);
});
/*! Bundled license information:

react/cjs/react.production.js:
  (**
   * @license React
   * react.production.js
   *
   * Copyright (c) Meta Platforms, Inc. and affiliates.
   *
   * This source code is licensed under the MIT license found in the
   * LICENSE file in the root directory of this source tree.
   *)

react/cjs/react-jsx-runtime.production.js:
  (**
   * @license React
   * react-jsx-runtime.production.js
   *
   * Copyright (c) Meta Platforms, Inc. and affiliates.
   *
   * This source code is licensed under the MIT license found in the
   * LICENSE file in the root directory of this source tree.
   *)
*/

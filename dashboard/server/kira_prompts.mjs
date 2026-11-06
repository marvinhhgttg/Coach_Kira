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

// node_modules/react/cjs/react.production.min.js
var require_react_production_min = __commonJS({
  "node_modules/react/cjs/react.production.min.js"(exports) {
    "use strict";
    var l = Symbol.for("react.element");
    var n = Symbol.for("react.portal");
    var p = Symbol.for("react.fragment");
    var q = Symbol.for("react.strict_mode");
    var r = Symbol.for("react.profiler");
    var t = Symbol.for("react.provider");
    var u = Symbol.for("react.context");
    var v = Symbol.for("react.forward_ref");
    var w = Symbol.for("react.suspense");
    var x = Symbol.for("react.memo");
    var y = Symbol.for("react.lazy");
    var z = Symbol.iterator;
    function A(a) {
      if (null === a || "object" !== typeof a) return null;
      a = z && a[z] || a["@@iterator"];
      return "function" === typeof a ? a : null;
    }
    var B = { isMounted: function() {
      return false;
    }, enqueueForceUpdate: function() {
    }, enqueueReplaceState: function() {
    }, enqueueSetState: function() {
    } };
    var C = Object.assign;
    var D = {};
    function E(a, b, e) {
      this.props = a;
      this.context = b;
      this.refs = D;
      this.updater = e || B;
    }
    E.prototype.isReactComponent = {};
    E.prototype.setState = function(a, b) {
      if ("object" !== typeof a && "function" !== typeof a && null != a) throw Error("setState(...): takes an object of state variables to update or a function which returns an object of state variables.");
      this.updater.enqueueSetState(this, a, b, "setState");
    };
    E.prototype.forceUpdate = function(a) {
      this.updater.enqueueForceUpdate(this, a, "forceUpdate");
    };
    function F() {
    }
    F.prototype = E.prototype;
    function G(a, b, e) {
      this.props = a;
      this.context = b;
      this.refs = D;
      this.updater = e || B;
    }
    var H = G.prototype = new F();
    H.constructor = G;
    C(H, E.prototype);
    H.isPureReactComponent = true;
    var I = Array.isArray;
    var J = Object.prototype.hasOwnProperty;
    var K = { current: null };
    var L = { key: true, ref: true, __self: true, __source: true };
    function M(a, b, e) {
      var d, c = {}, k = null, h = null;
      if (null != b) for (d in void 0 !== b.ref && (h = b.ref), void 0 !== b.key && (k = "" + b.key), b) J.call(b, d) && !L.hasOwnProperty(d) && (c[d] = b[d]);
      var g = arguments.length - 2;
      if (1 === g) c.children = e;
      else if (1 < g) {
        for (var f = Array(g), m = 0; m < g; m++) f[m] = arguments[m + 2];
        c.children = f;
      }
      if (a && a.defaultProps) for (d in g = a.defaultProps, g) void 0 === c[d] && (c[d] = g[d]);
      return { $$typeof: l, type: a, key: k, ref: h, props: c, _owner: K.current };
    }
    function N(a, b) {
      return { $$typeof: l, type: a.type, key: b, ref: a.ref, props: a.props, _owner: a._owner };
    }
    function O(a) {
      return "object" === typeof a && null !== a && a.$$typeof === l;
    }
    function escape(a) {
      var b = { "=": "=0", ":": "=2" };
      return "$" + a.replace(/[=:]/g, function(a2) {
        return b[a2];
      });
    }
    var P = /\/+/g;
    function Q(a, b) {
      return "object" === typeof a && null !== a && null != a.key ? escape("" + a.key) : b.toString(36);
    }
    function R(a, b, e, d, c) {
      var k = typeof a;
      if ("undefined" === k || "boolean" === k) a = null;
      var h = false;
      if (null === a) h = true;
      else switch (k) {
        case "string":
        case "number":
          h = true;
          break;
        case "object":
          switch (a.$$typeof) {
            case l:
            case n:
              h = true;
          }
      }
      if (h) return h = a, c = c(h), a = "" === d ? "." + Q(h, 0) : d, I(c) ? (e = "", null != a && (e = a.replace(P, "$&/") + "/"), R(c, b, e, "", function(a2) {
        return a2;
      })) : null != c && (O(c) && (c = N(c, e + (!c.key || h && h.key === c.key ? "" : ("" + c.key).replace(P, "$&/") + "/") + a)), b.push(c)), 1;
      h = 0;
      d = "" === d ? "." : d + ":";
      if (I(a)) for (var g = 0; g < a.length; g++) {
        k = a[g];
        var f = d + Q(k, g);
        h += R(k, b, e, f, c);
      }
      else if (f = A(a), "function" === typeof f) for (a = f.call(a), g = 0; !(k = a.next()).done; ) k = k.value, f = d + Q(k, g++), h += R(k, b, e, f, c);
      else if ("object" === k) throw b = String(a), Error("Objects are not valid as a React child (found: " + ("[object Object]" === b ? "object with keys {" + Object.keys(a).join(", ") + "}" : b) + "). If you meant to render a collection of children, use an array instead.");
      return h;
    }
    function S(a, b, e) {
      if (null == a) return a;
      var d = [], c = 0;
      R(a, d, "", "", function(a2) {
        return b.call(e, a2, c++);
      });
      return d;
    }
    function T(a) {
      if (-1 === a._status) {
        var b = a._result;
        b = b();
        b.then(function(b2) {
          if (0 === a._status || -1 === a._status) a._status = 1, a._result = b2;
        }, function(b2) {
          if (0 === a._status || -1 === a._status) a._status = 2, a._result = b2;
        });
        -1 === a._status && (a._status = 0, a._result = b);
      }
      if (1 === a._status) return a._result.default;
      throw a._result;
    }
    var U = { current: null };
    var V = { transition: null };
    var W = { ReactCurrentDispatcher: U, ReactCurrentBatchConfig: V, ReactCurrentOwner: K };
    function X() {
      throw Error("act(...) is not supported in production builds of React.");
    }
    exports.Children = { map: S, forEach: function(a, b, e) {
      S(a, function() {
        b.apply(this, arguments);
      }, e);
    }, count: function(a) {
      var b = 0;
      S(a, function() {
        b++;
      });
      return b;
    }, toArray: function(a) {
      return S(a, function(a2) {
        return a2;
      }) || [];
    }, only: function(a) {
      if (!O(a)) throw Error("React.Children.only expected to receive a single React element child.");
      return a;
    } };
    exports.Component = E;
    exports.Fragment = p;
    exports.Profiler = r;
    exports.PureComponent = G;
    exports.StrictMode = q;
    exports.Suspense = w;
    exports.__SECRET_INTERNALS_DO_NOT_USE_OR_YOU_WILL_BE_FIRED = W;
    exports.act = X;
    exports.cloneElement = function(a, b, e) {
      if (null === a || void 0 === a) throw Error("React.cloneElement(...): The argument must be a React element, but you passed " + a + ".");
      var d = C({}, a.props), c = a.key, k = a.ref, h = a._owner;
      if (null != b) {
        void 0 !== b.ref && (k = b.ref, h = K.current);
        void 0 !== b.key && (c = "" + b.key);
        if (a.type && a.type.defaultProps) var g = a.type.defaultProps;
        for (f in b) J.call(b, f) && !L.hasOwnProperty(f) && (d[f] = void 0 === b[f] && void 0 !== g ? g[f] : b[f]);
      }
      var f = arguments.length - 2;
      if (1 === f) d.children = e;
      else if (1 < f) {
        g = Array(f);
        for (var m = 0; m < f; m++) g[m] = arguments[m + 2];
        d.children = g;
      }
      return { $$typeof: l, type: a.type, key: c, ref: k, props: d, _owner: h };
    };
    exports.createContext = function(a) {
      a = { $$typeof: u, _currentValue: a, _currentValue2: a, _threadCount: 0, Provider: null, Consumer: null, _defaultValue: null, _globalName: null };
      a.Provider = { $$typeof: t, _context: a };
      return a.Consumer = a;
    };
    exports.createElement = M;
    exports.createFactory = function(a) {
      var b = M.bind(null, a);
      b.type = a;
      return b;
    };
    exports.createRef = function() {
      return { current: null };
    };
    exports.forwardRef = function(a) {
      return { $$typeof: v, render: a };
    };
    exports.isValidElement = O;
    exports.lazy = function(a) {
      return { $$typeof: y, _payload: { _status: -1, _result: a }, _init: T };
    };
    exports.memo = function(a, b) {
      return { $$typeof: x, type: a, compare: void 0 === b ? null : b };
    };
    exports.startTransition = function(a) {
      var b = V.transition;
      V.transition = {};
      try {
        a();
      } finally {
        V.transition = b;
      }
    };
    exports.unstable_act = X;
    exports.useCallback = function(a, b) {
      return U.current.useCallback(a, b);
    };
    exports.useContext = function(a) {
      return U.current.useContext(a);
    };
    exports.useDebugValue = function() {
    };
    exports.useDeferredValue = function(a) {
      return U.current.useDeferredValue(a);
    };
    exports.useEffect = function(a, b) {
      return U.current.useEffect(a, b);
    };
    exports.useId = function() {
      return U.current.useId();
    };
    exports.useImperativeHandle = function(a, b, e) {
      return U.current.useImperativeHandle(a, b, e);
    };
    exports.useInsertionEffect = function(a, b) {
      return U.current.useInsertionEffect(a, b);
    };
    exports.useLayoutEffect = function(a, b) {
      return U.current.useLayoutEffect(a, b);
    };
    exports.useMemo = function(a, b) {
      return U.current.useMemo(a, b);
    };
    exports.useReducer = function(a, b, e) {
      return U.current.useReducer(a, b, e);
    };
    exports.useRef = function(a) {
      return U.current.useRef(a);
    };
    exports.useState = function(a) {
      return U.current.useState(a);
    };
    exports.useSyncExternalStore = function(a, b, e) {
      return U.current.useSyncExternalStore(a, b, e);
    };
    exports.useTransition = function() {
      return U.current.useTransition();
    };
    exports.version = "18.3.1";
  }
});

// node_modules/react/cjs/react.development.js
var require_react_development = __commonJS({
  "node_modules/react/cjs/react.development.js"(exports, module) {
    "use strict";
    if (process.env.NODE_ENV !== "production") {
      (function() {
        "use strict";
        if (typeof __REACT_DEVTOOLS_GLOBAL_HOOK__ !== "undefined" && typeof __REACT_DEVTOOLS_GLOBAL_HOOK__.registerInternalModuleStart === "function") {
          __REACT_DEVTOOLS_GLOBAL_HOOK__.registerInternalModuleStart(new Error());
        }
        var ReactVersion = "18.3.1";
        var REACT_ELEMENT_TYPE = Symbol.for("react.element");
        var REACT_PORTAL_TYPE = Symbol.for("react.portal");
        var REACT_FRAGMENT_TYPE = Symbol.for("react.fragment");
        var REACT_STRICT_MODE_TYPE = Symbol.for("react.strict_mode");
        var REACT_PROFILER_TYPE = Symbol.for("react.profiler");
        var REACT_PROVIDER_TYPE = Symbol.for("react.provider");
        var REACT_CONTEXT_TYPE = Symbol.for("react.context");
        var REACT_FORWARD_REF_TYPE = Symbol.for("react.forward_ref");
        var REACT_SUSPENSE_TYPE = Symbol.for("react.suspense");
        var REACT_SUSPENSE_LIST_TYPE = Symbol.for("react.suspense_list");
        var REACT_MEMO_TYPE = Symbol.for("react.memo");
        var REACT_LAZY_TYPE = Symbol.for("react.lazy");
        var REACT_OFFSCREEN_TYPE = Symbol.for("react.offscreen");
        var MAYBE_ITERATOR_SYMBOL = Symbol.iterator;
        var FAUX_ITERATOR_SYMBOL = "@@iterator";
        function getIteratorFn(maybeIterable) {
          if (maybeIterable === null || typeof maybeIterable !== "object") {
            return null;
          }
          var maybeIterator = MAYBE_ITERATOR_SYMBOL && maybeIterable[MAYBE_ITERATOR_SYMBOL] || maybeIterable[FAUX_ITERATOR_SYMBOL];
          if (typeof maybeIterator === "function") {
            return maybeIterator;
          }
          return null;
        }
        var ReactCurrentDispatcher = {
          /**
           * @internal
           * @type {ReactComponent}
           */
          current: null
        };
        var ReactCurrentBatchConfig = {
          transition: null
        };
        var ReactCurrentActQueue = {
          current: null,
          // Used to reproduce behavior of `batchedUpdates` in legacy mode.
          isBatchingLegacy: false,
          didScheduleLegacyUpdate: false
        };
        var ReactCurrentOwner = {
          /**
           * @internal
           * @type {ReactComponent}
           */
          current: null
        };
        var ReactDebugCurrentFrame = {};
        var currentExtraStackFrame = null;
        function setExtraStackFrame(stack) {
          {
            currentExtraStackFrame = stack;
          }
        }
        {
          ReactDebugCurrentFrame.setExtraStackFrame = function(stack) {
            {
              currentExtraStackFrame = stack;
            }
          };
          ReactDebugCurrentFrame.getCurrentStack = null;
          ReactDebugCurrentFrame.getStackAddendum = function() {
            var stack = "";
            if (currentExtraStackFrame) {
              stack += currentExtraStackFrame;
            }
            var impl = ReactDebugCurrentFrame.getCurrentStack;
            if (impl) {
              stack += impl() || "";
            }
            return stack;
          };
        }
        var enableScopeAPI = false;
        var enableCacheElement = false;
        var enableTransitionTracing = false;
        var enableLegacyHidden = false;
        var enableDebugTracing = false;
        var ReactSharedInternals = {
          ReactCurrentDispatcher,
          ReactCurrentBatchConfig,
          ReactCurrentOwner
        };
        {
          ReactSharedInternals.ReactDebugCurrentFrame = ReactDebugCurrentFrame;
          ReactSharedInternals.ReactCurrentActQueue = ReactCurrentActQueue;
        }
        function warn(format) {
          {
            {
              for (var _len = arguments.length, args = new Array(_len > 1 ? _len - 1 : 0), _key = 1; _key < _len; _key++) {
                args[_key - 1] = arguments[_key];
              }
              printWarning("warn", format, args);
            }
          }
        }
        function error(format) {
          {
            {
              for (var _len2 = arguments.length, args = new Array(_len2 > 1 ? _len2 - 1 : 0), _key2 = 1; _key2 < _len2; _key2++) {
                args[_key2 - 1] = arguments[_key2];
              }
              printWarning("error", format, args);
            }
          }
        }
        function printWarning(level, format, args) {
          {
            var ReactDebugCurrentFrame2 = ReactSharedInternals.ReactDebugCurrentFrame;
            var stack = ReactDebugCurrentFrame2.getStackAddendum();
            if (stack !== "") {
              format += "%s";
              args = args.concat([stack]);
            }
            var argsWithFormat = args.map(function(item) {
              return String(item);
            });
            argsWithFormat.unshift("Warning: " + format);
            Function.prototype.apply.call(console[level], console, argsWithFormat);
          }
        }
        var didWarnStateUpdateForUnmountedComponent = {};
        function warnNoop(publicInstance, callerName) {
          {
            var _constructor = publicInstance.constructor;
            var componentName = _constructor && (_constructor.displayName || _constructor.name) || "ReactClass";
            var warningKey = componentName + "." + callerName;
            if (didWarnStateUpdateForUnmountedComponent[warningKey]) {
              return;
            }
            error("Can't call %s on a component that is not yet mounted. This is a no-op, but it might indicate a bug in your application. Instead, assign to `this.state` directly or define a `state = {};` class property with the desired state in the %s component.", callerName, componentName);
            didWarnStateUpdateForUnmountedComponent[warningKey] = true;
          }
        }
        var ReactNoopUpdateQueue = {
          /**
           * Checks whether or not this composite component is mounted.
           * @param {ReactClass} publicInstance The instance we want to test.
           * @return {boolean} True if mounted, false otherwise.
           * @protected
           * @final
           */
          isMounted: function(publicInstance) {
            return false;
          },
          /**
           * Forces an update. This should only be invoked when it is known with
           * certainty that we are **not** in a DOM transaction.
           *
           * You may want to call this when you know that some deeper aspect of the
           * component's state has changed but `setState` was not called.
           *
           * This will not invoke `shouldComponentUpdate`, but it will invoke
           * `componentWillUpdate` and `componentDidUpdate`.
           *
           * @param {ReactClass} publicInstance The instance that should rerender.
           * @param {?function} callback Called after component is updated.
           * @param {?string} callerName name of the calling function in the public API.
           * @internal
           */
          enqueueForceUpdate: function(publicInstance, callback, callerName) {
            warnNoop(publicInstance, "forceUpdate");
          },
          /**
           * Replaces all of the state. Always use this or `setState` to mutate state.
           * You should treat `this.state` as immutable.
           *
           * There is no guarantee that `this.state` will be immediately updated, so
           * accessing `this.state` after calling this method may return the old value.
           *
           * @param {ReactClass} publicInstance The instance that should rerender.
           * @param {object} completeState Next state.
           * @param {?function} callback Called after component is updated.
           * @param {?string} callerName name of the calling function in the public API.
           * @internal
           */
          enqueueReplaceState: function(publicInstance, completeState, callback, callerName) {
            warnNoop(publicInstance, "replaceState");
          },
          /**
           * Sets a subset of the state. This only exists because _pendingState is
           * internal. This provides a merging strategy that is not available to deep
           * properties which is confusing. TODO: Expose pendingState or don't use it
           * during the merge.
           *
           * @param {ReactClass} publicInstance The instance that should rerender.
           * @param {object} partialState Next partial state to be merged with state.
           * @param {?function} callback Called after component is updated.
           * @param {?string} Name of the calling function in the public API.
           * @internal
           */
          enqueueSetState: function(publicInstance, partialState, callback, callerName) {
            warnNoop(publicInstance, "setState");
          }
        };
        var assign = Object.assign;
        var emptyObject = {};
        {
          Object.freeze(emptyObject);
        }
        function Component(props, context, updater) {
          this.props = props;
          this.context = context;
          this.refs = emptyObject;
          this.updater = updater || ReactNoopUpdateQueue;
        }
        Component.prototype.isReactComponent = {};
        Component.prototype.setState = function(partialState, callback) {
          if (typeof partialState !== "object" && typeof partialState !== "function" && partialState != null) {
            throw new Error("setState(...): takes an object of state variables to update or a function which returns an object of state variables.");
          }
          this.updater.enqueueSetState(this, partialState, callback, "setState");
        };
        Component.prototype.forceUpdate = function(callback) {
          this.updater.enqueueForceUpdate(this, callback, "forceUpdate");
        };
        {
          var deprecatedAPIs = {
            isMounted: ["isMounted", "Instead, make sure to clean up subscriptions and pending requests in componentWillUnmount to prevent memory leaks."],
            replaceState: ["replaceState", "Refactor your code to use setState instead (see https://github.com/facebook/react/issues/3236)."]
          };
          var defineDeprecationWarning = function(methodName, info) {
            Object.defineProperty(Component.prototype, methodName, {
              get: function() {
                warn("%s(...) is deprecated in plain JavaScript React classes. %s", info[0], info[1]);
                return void 0;
              }
            });
          };
          for (var fnName in deprecatedAPIs) {
            if (deprecatedAPIs.hasOwnProperty(fnName)) {
              defineDeprecationWarning(fnName, deprecatedAPIs[fnName]);
            }
          }
        }
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
        function createRef() {
          var refObject = {
            current: null
          };
          {
            Object.seal(refObject);
          }
          return refObject;
        }
        var isArrayImpl = Array.isArray;
        function isArray(a) {
          return isArrayImpl(a);
        }
        function typeName(value) {
          {
            var hasToStringTag = typeof Symbol === "function" && Symbol.toStringTag;
            var type = hasToStringTag && value[Symbol.toStringTag] || value.constructor.name || "Object";
            return type;
          }
        }
        function willCoercionThrow(value) {
          {
            try {
              testStringCoercion(value);
              return false;
            } catch (e) {
              return true;
            }
          }
        }
        function testStringCoercion(value) {
          return "" + value;
        }
        function checkKeyStringCoercion(value) {
          {
            if (willCoercionThrow(value)) {
              error("The provided key is an unsupported type %s. This value must be coerced to a string before before using it here.", typeName(value));
              return testStringCoercion(value);
            }
          }
        }
        function getWrappedName(outerType, innerType, wrapperName) {
          var displayName = outerType.displayName;
          if (displayName) {
            return displayName;
          }
          var functionName = innerType.displayName || innerType.name || "";
          return functionName !== "" ? wrapperName + "(" + functionName + ")" : wrapperName;
        }
        function getContextName(type) {
          return type.displayName || "Context";
        }
        function getComponentNameFromType(type) {
          if (type == null) {
            return null;
          }
          {
            if (typeof type.tag === "number") {
              error("Received an unexpected object in getComponentNameFromType(). This is likely a bug in React. Please file an issue.");
            }
          }
          if (typeof type === "function") {
            return type.displayName || type.name || null;
          }
          if (typeof type === "string") {
            return type;
          }
          switch (type) {
            case REACT_FRAGMENT_TYPE:
              return "Fragment";
            case REACT_PORTAL_TYPE:
              return "Portal";
            case REACT_PROFILER_TYPE:
              return "Profiler";
            case REACT_STRICT_MODE_TYPE:
              return "StrictMode";
            case REACT_SUSPENSE_TYPE:
              return "Suspense";
            case REACT_SUSPENSE_LIST_TYPE:
              return "SuspenseList";
          }
          if (typeof type === "object") {
            switch (type.$$typeof) {
              case REACT_CONTEXT_TYPE:
                var context = type;
                return getContextName(context) + ".Consumer";
              case REACT_PROVIDER_TYPE:
                var provider = type;
                return getContextName(provider._context) + ".Provider";
              case REACT_FORWARD_REF_TYPE:
                return getWrappedName(type, type.render, "ForwardRef");
              case REACT_MEMO_TYPE:
                var outerName = type.displayName || null;
                if (outerName !== null) {
                  return outerName;
                }
                return getComponentNameFromType(type.type) || "Memo";
              case REACT_LAZY_TYPE: {
                var lazyComponent = type;
                var payload = lazyComponent._payload;
                var init = lazyComponent._init;
                try {
                  return getComponentNameFromType(init(payload));
                } catch (x) {
                  return null;
                }
              }
            }
          }
          return null;
        }
        var hasOwnProperty = Object.prototype.hasOwnProperty;
        var RESERVED_PROPS = {
          key: true,
          ref: true,
          __self: true,
          __source: true
        };
        var specialPropKeyWarningShown, specialPropRefWarningShown, didWarnAboutStringRefs;
        {
          didWarnAboutStringRefs = {};
        }
        function hasValidRef(config) {
          {
            if (hasOwnProperty.call(config, "ref")) {
              var getter = Object.getOwnPropertyDescriptor(config, "ref").get;
              if (getter && getter.isReactWarning) {
                return false;
              }
            }
          }
          return config.ref !== void 0;
        }
        function hasValidKey(config) {
          {
            if (hasOwnProperty.call(config, "key")) {
              var getter = Object.getOwnPropertyDescriptor(config, "key").get;
              if (getter && getter.isReactWarning) {
                return false;
              }
            }
          }
          return config.key !== void 0;
        }
        function defineKeyPropWarningGetter(props, displayName) {
          var warnAboutAccessingKey = function() {
            {
              if (!specialPropKeyWarningShown) {
                specialPropKeyWarningShown = true;
                error("%s: `key` is not a prop. Trying to access it will result in `undefined` being returned. If you need to access the same value within the child component, you should pass it as a different prop. (https://reactjs.org/link/special-props)", displayName);
              }
            }
          };
          warnAboutAccessingKey.isReactWarning = true;
          Object.defineProperty(props, "key", {
            get: warnAboutAccessingKey,
            configurable: true
          });
        }
        function defineRefPropWarningGetter(props, displayName) {
          var warnAboutAccessingRef = function() {
            {
              if (!specialPropRefWarningShown) {
                specialPropRefWarningShown = true;
                error("%s: `ref` is not a prop. Trying to access it will result in `undefined` being returned. If you need to access the same value within the child component, you should pass it as a different prop. (https://reactjs.org/link/special-props)", displayName);
              }
            }
          };
          warnAboutAccessingRef.isReactWarning = true;
          Object.defineProperty(props, "ref", {
            get: warnAboutAccessingRef,
            configurable: true
          });
        }
        function warnIfStringRefCannotBeAutoConverted(config) {
          {
            if (typeof config.ref === "string" && ReactCurrentOwner.current && config.__self && ReactCurrentOwner.current.stateNode !== config.__self) {
              var componentName = getComponentNameFromType(ReactCurrentOwner.current.type);
              if (!didWarnAboutStringRefs[componentName]) {
                error('Component "%s" contains the string ref "%s". Support for string refs will be removed in a future major release. This case cannot be automatically converted to an arrow function. We ask you to manually fix this case by using useRef() or createRef() instead. Learn more about using refs safely here: https://reactjs.org/link/strict-mode-string-ref', componentName, config.ref);
                didWarnAboutStringRefs[componentName] = true;
              }
            }
          }
        }
        var ReactElement = function(type, key, ref, self, source, owner, props) {
          var element = {
            // This tag allows us to uniquely identify this as a React Element
            $$typeof: REACT_ELEMENT_TYPE,
            // Built-in properties that belong on the element
            type,
            key,
            ref,
            props,
            // Record the component responsible for creating this element.
            _owner: owner
          };
          {
            element._store = {};
            Object.defineProperty(element._store, "validated", {
              configurable: false,
              enumerable: false,
              writable: true,
              value: false
            });
            Object.defineProperty(element, "_self", {
              configurable: false,
              enumerable: false,
              writable: false,
              value: self
            });
            Object.defineProperty(element, "_source", {
              configurable: false,
              enumerable: false,
              writable: false,
              value: source
            });
            if (Object.freeze) {
              Object.freeze(element.props);
              Object.freeze(element);
            }
          }
          return element;
        };
        function createElement(type, config, children) {
          var propName;
          var props = {};
          var key = null;
          var ref = null;
          var self = null;
          var source = null;
          if (config != null) {
            if (hasValidRef(config)) {
              ref = config.ref;
              {
                warnIfStringRefCannotBeAutoConverted(config);
              }
            }
            if (hasValidKey(config)) {
              {
                checkKeyStringCoercion(config.key);
              }
              key = "" + config.key;
            }
            self = config.__self === void 0 ? null : config.__self;
            source = config.__source === void 0 ? null : config.__source;
            for (propName in config) {
              if (hasOwnProperty.call(config, propName) && !RESERVED_PROPS.hasOwnProperty(propName)) {
                props[propName] = config[propName];
              }
            }
          }
          var childrenLength = arguments.length - 2;
          if (childrenLength === 1) {
            props.children = children;
          } else if (childrenLength > 1) {
            var childArray = Array(childrenLength);
            for (var i = 0; i < childrenLength; i++) {
              childArray[i] = arguments[i + 2];
            }
            {
              if (Object.freeze) {
                Object.freeze(childArray);
              }
            }
            props.children = childArray;
          }
          if (type && type.defaultProps) {
            var defaultProps = type.defaultProps;
            for (propName in defaultProps) {
              if (props[propName] === void 0) {
                props[propName] = defaultProps[propName];
              }
            }
          }
          {
            if (key || ref) {
              var displayName = typeof type === "function" ? type.displayName || type.name || "Unknown" : type;
              if (key) {
                defineKeyPropWarningGetter(props, displayName);
              }
              if (ref) {
                defineRefPropWarningGetter(props, displayName);
              }
            }
          }
          return ReactElement(type, key, ref, self, source, ReactCurrentOwner.current, props);
        }
        function cloneAndReplaceKey(oldElement, newKey) {
          var newElement = ReactElement(oldElement.type, newKey, oldElement.ref, oldElement._self, oldElement._source, oldElement._owner, oldElement.props);
          return newElement;
        }
        function cloneElement(element, config, children) {
          if (element === null || element === void 0) {
            throw new Error("React.cloneElement(...): The argument must be a React element, but you passed " + element + ".");
          }
          var propName;
          var props = assign({}, element.props);
          var key = element.key;
          var ref = element.ref;
          var self = element._self;
          var source = element._source;
          var owner = element._owner;
          if (config != null) {
            if (hasValidRef(config)) {
              ref = config.ref;
              owner = ReactCurrentOwner.current;
            }
            if (hasValidKey(config)) {
              {
                checkKeyStringCoercion(config.key);
              }
              key = "" + config.key;
            }
            var defaultProps;
            if (element.type && element.type.defaultProps) {
              defaultProps = element.type.defaultProps;
            }
            for (propName in config) {
              if (hasOwnProperty.call(config, propName) && !RESERVED_PROPS.hasOwnProperty(propName)) {
                if (config[propName] === void 0 && defaultProps !== void 0) {
                  props[propName] = defaultProps[propName];
                } else {
                  props[propName] = config[propName];
                }
              }
            }
          }
          var childrenLength = arguments.length - 2;
          if (childrenLength === 1) {
            props.children = children;
          } else if (childrenLength > 1) {
            var childArray = Array(childrenLength);
            for (var i = 0; i < childrenLength; i++) {
              childArray[i] = arguments[i + 2];
            }
            props.children = childArray;
          }
          return ReactElement(element.type, key, ref, self, source, owner, props);
        }
        function isValidElement(object) {
          return typeof object === "object" && object !== null && object.$$typeof === REACT_ELEMENT_TYPE;
        }
        var SEPARATOR = ".";
        var SUBSEPARATOR = ":";
        function escape(key) {
          var escapeRegex = /[=:]/g;
          var escaperLookup = {
            "=": "=0",
            ":": "=2"
          };
          var escapedString = key.replace(escapeRegex, function(match) {
            return escaperLookup[match];
          });
          return "$" + escapedString;
        }
        var didWarnAboutMaps = false;
        var userProvidedKeyEscapeRegex = /\/+/g;
        function escapeUserProvidedKey(text) {
          return text.replace(userProvidedKeyEscapeRegex, "$&/");
        }
        function getElementKey(element, index) {
          if (typeof element === "object" && element !== null && element.key != null) {
            {
              checkKeyStringCoercion(element.key);
            }
            return escape("" + element.key);
          }
          return index.toString(36);
        }
        function mapIntoArray(children, array, escapedPrefix, nameSoFar, callback) {
          var type = typeof children;
          if (type === "undefined" || type === "boolean") {
            children = null;
          }
          var invokeCallback = false;
          if (children === null) {
            invokeCallback = true;
          } else {
            switch (type) {
              case "string":
              case "number":
                invokeCallback = true;
                break;
              case "object":
                switch (children.$$typeof) {
                  case REACT_ELEMENT_TYPE:
                  case REACT_PORTAL_TYPE:
                    invokeCallback = true;
                }
            }
          }
          if (invokeCallback) {
            var _child = children;
            var mappedChild = callback(_child);
            var childKey = nameSoFar === "" ? SEPARATOR + getElementKey(_child, 0) : nameSoFar;
            if (isArray(mappedChild)) {
              var escapedChildKey = "";
              if (childKey != null) {
                escapedChildKey = escapeUserProvidedKey(childKey) + "/";
              }
              mapIntoArray(mappedChild, array, escapedChildKey, "", function(c) {
                return c;
              });
            } else if (mappedChild != null) {
              if (isValidElement(mappedChild)) {
                {
                  if (mappedChild.key && (!_child || _child.key !== mappedChild.key)) {
                    checkKeyStringCoercion(mappedChild.key);
                  }
                }
                mappedChild = cloneAndReplaceKey(
                  mappedChild,
                  // Keep both the (mapped) and old keys if they differ, just as
                  // traverseAllChildren used to do for objects as children
                  escapedPrefix + // $FlowFixMe Flow incorrectly thinks React.Portal doesn't have a key
                  (mappedChild.key && (!_child || _child.key !== mappedChild.key) ? (
                    // $FlowFixMe Flow incorrectly thinks existing element's key can be a number
                    // eslint-disable-next-line react-internal/safe-string-coercion
                    escapeUserProvidedKey("" + mappedChild.key) + "/"
                  ) : "") + childKey
                );
              }
              array.push(mappedChild);
            }
            return 1;
          }
          var child;
          var nextName;
          var subtreeCount = 0;
          var nextNamePrefix = nameSoFar === "" ? SEPARATOR : nameSoFar + SUBSEPARATOR;
          if (isArray(children)) {
            for (var i = 0; i < children.length; i++) {
              child = children[i];
              nextName = nextNamePrefix + getElementKey(child, i);
              subtreeCount += mapIntoArray(child, array, escapedPrefix, nextName, callback);
            }
          } else {
            var iteratorFn = getIteratorFn(children);
            if (typeof iteratorFn === "function") {
              var iterableChildren = children;
              {
                if (iteratorFn === iterableChildren.entries) {
                  if (!didWarnAboutMaps) {
                    warn("Using Maps as children is not supported. Use an array of keyed ReactElements instead.");
                  }
                  didWarnAboutMaps = true;
                }
              }
              var iterator = iteratorFn.call(iterableChildren);
              var step;
              var ii = 0;
              while (!(step = iterator.next()).done) {
                child = step.value;
                nextName = nextNamePrefix + getElementKey(child, ii++);
                subtreeCount += mapIntoArray(child, array, escapedPrefix, nextName, callback);
              }
            } else if (type === "object") {
              var childrenString = String(children);
              throw new Error("Objects are not valid as a React child (found: " + (childrenString === "[object Object]" ? "object with keys {" + Object.keys(children).join(", ") + "}" : childrenString) + "). If you meant to render a collection of children, use an array instead.");
            }
          }
          return subtreeCount;
        }
        function mapChildren(children, func, context) {
          if (children == null) {
            return children;
          }
          var result = [];
          var count = 0;
          mapIntoArray(children, result, "", "", function(child) {
            return func.call(context, child, count++);
          });
          return result;
        }
        function countChildren(children) {
          var n = 0;
          mapChildren(children, function() {
            n++;
          });
          return n;
        }
        function forEachChildren(children, forEachFunc, forEachContext) {
          mapChildren(children, function() {
            forEachFunc.apply(this, arguments);
          }, forEachContext);
        }
        function toArray(children) {
          return mapChildren(children, function(child) {
            return child;
          }) || [];
        }
        function onlyChild(children) {
          if (!isValidElement(children)) {
            throw new Error("React.Children.only expected to receive a single React element child.");
          }
          return children;
        }
        function createContext(defaultValue) {
          var context = {
            $$typeof: REACT_CONTEXT_TYPE,
            // As a workaround to support multiple concurrent renderers, we categorize
            // some renderers as primary and others as secondary. We only expect
            // there to be two concurrent renderers at most: React Native (primary) and
            // Fabric (secondary); React DOM (primary) and React ART (secondary).
            // Secondary renderers store their context values on separate fields.
            _currentValue: defaultValue,
            _currentValue2: defaultValue,
            // Used to track how many concurrent renderers this context currently
            // supports within in a single renderer. Such as parallel server rendering.
            _threadCount: 0,
            // These are circular
            Provider: null,
            Consumer: null,
            // Add these to use same hidden class in VM as ServerContext
            _defaultValue: null,
            _globalName: null
          };
          context.Provider = {
            $$typeof: REACT_PROVIDER_TYPE,
            _context: context
          };
          var hasWarnedAboutUsingNestedContextConsumers = false;
          var hasWarnedAboutUsingConsumerProvider = false;
          var hasWarnedAboutDisplayNameOnConsumer = false;
          {
            var Consumer = {
              $$typeof: REACT_CONTEXT_TYPE,
              _context: context
            };
            Object.defineProperties(Consumer, {
              Provider: {
                get: function() {
                  if (!hasWarnedAboutUsingConsumerProvider) {
                    hasWarnedAboutUsingConsumerProvider = true;
                    error("Rendering <Context.Consumer.Provider> is not supported and will be removed in a future major release. Did you mean to render <Context.Provider> instead?");
                  }
                  return context.Provider;
                },
                set: function(_Provider) {
                  context.Provider = _Provider;
                }
              },
              _currentValue: {
                get: function() {
                  return context._currentValue;
                },
                set: function(_currentValue) {
                  context._currentValue = _currentValue;
                }
              },
              _currentValue2: {
                get: function() {
                  return context._currentValue2;
                },
                set: function(_currentValue2) {
                  context._currentValue2 = _currentValue2;
                }
              },
              _threadCount: {
                get: function() {
                  return context._threadCount;
                },
                set: function(_threadCount) {
                  context._threadCount = _threadCount;
                }
              },
              Consumer: {
                get: function() {
                  if (!hasWarnedAboutUsingNestedContextConsumers) {
                    hasWarnedAboutUsingNestedContextConsumers = true;
                    error("Rendering <Context.Consumer.Consumer> is not supported and will be removed in a future major release. Did you mean to render <Context.Consumer> instead?");
                  }
                  return context.Consumer;
                }
              },
              displayName: {
                get: function() {
                  return context.displayName;
                },
                set: function(displayName) {
                  if (!hasWarnedAboutDisplayNameOnConsumer) {
                    warn("Setting `displayName` on Context.Consumer has no effect. You should set it directly on the context with Context.displayName = '%s'.", displayName);
                    hasWarnedAboutDisplayNameOnConsumer = true;
                  }
                }
              }
            });
            context.Consumer = Consumer;
          }
          {
            context._currentRenderer = null;
            context._currentRenderer2 = null;
          }
          return context;
        }
        var Uninitialized = -1;
        var Pending = 0;
        var Resolved = 1;
        var Rejected = 2;
        function lazyInitializer(payload) {
          if (payload._status === Uninitialized) {
            var ctor = payload._result;
            var thenable = ctor();
            thenable.then(function(moduleObject2) {
              if (payload._status === Pending || payload._status === Uninitialized) {
                var resolved = payload;
                resolved._status = Resolved;
                resolved._result = moduleObject2;
              }
            }, function(error2) {
              if (payload._status === Pending || payload._status === Uninitialized) {
                var rejected = payload;
                rejected._status = Rejected;
                rejected._result = error2;
              }
            });
            if (payload._status === Uninitialized) {
              var pending = payload;
              pending._status = Pending;
              pending._result = thenable;
            }
          }
          if (payload._status === Resolved) {
            var moduleObject = payload._result;
            {
              if (moduleObject === void 0) {
                error("lazy: Expected the result of a dynamic import() call. Instead received: %s\n\nYour code should look like: \n  const MyComponent = lazy(() => import('./MyComponent'))\n\nDid you accidentally put curly braces around the import?", moduleObject);
              }
            }
            {
              if (!("default" in moduleObject)) {
                error("lazy: Expected the result of a dynamic import() call. Instead received: %s\n\nYour code should look like: \n  const MyComponent = lazy(() => import('./MyComponent'))", moduleObject);
              }
            }
            return moduleObject.default;
          } else {
            throw payload._result;
          }
        }
        function lazy(ctor) {
          var payload = {
            // We use these fields to store the result.
            _status: Uninitialized,
            _result: ctor
          };
          var lazyType = {
            $$typeof: REACT_LAZY_TYPE,
            _payload: payload,
            _init: lazyInitializer
          };
          {
            var defaultProps;
            var propTypes;
            Object.defineProperties(lazyType, {
              defaultProps: {
                configurable: true,
                get: function() {
                  return defaultProps;
                },
                set: function(newDefaultProps) {
                  error("React.lazy(...): It is not supported to assign `defaultProps` to a lazy component import. Either specify them where the component is defined, or create a wrapping component around it.");
                  defaultProps = newDefaultProps;
                  Object.defineProperty(lazyType, "defaultProps", {
                    enumerable: true
                  });
                }
              },
              propTypes: {
                configurable: true,
                get: function() {
                  return propTypes;
                },
                set: function(newPropTypes) {
                  error("React.lazy(...): It is not supported to assign `propTypes` to a lazy component import. Either specify them where the component is defined, or create a wrapping component around it.");
                  propTypes = newPropTypes;
                  Object.defineProperty(lazyType, "propTypes", {
                    enumerable: true
                  });
                }
              }
            });
          }
          return lazyType;
        }
        function forwardRef(render) {
          {
            if (render != null && render.$$typeof === REACT_MEMO_TYPE) {
              error("forwardRef requires a render function but received a `memo` component. Instead of forwardRef(memo(...)), use memo(forwardRef(...)).");
            } else if (typeof render !== "function") {
              error("forwardRef requires a render function but was given %s.", render === null ? "null" : typeof render);
            } else {
              if (render.length !== 0 && render.length !== 2) {
                error("forwardRef render functions accept exactly two parameters: props and ref. %s", render.length === 1 ? "Did you forget to use the ref parameter?" : "Any additional parameter will be undefined.");
              }
            }
            if (render != null) {
              if (render.defaultProps != null || render.propTypes != null) {
                error("forwardRef render functions do not support propTypes or defaultProps. Did you accidentally pass a React component?");
              }
            }
          }
          var elementType = {
            $$typeof: REACT_FORWARD_REF_TYPE,
            render
          };
          {
            var ownName;
            Object.defineProperty(elementType, "displayName", {
              enumerable: false,
              configurable: true,
              get: function() {
                return ownName;
              },
              set: function(name) {
                ownName = name;
                if (!render.name && !render.displayName) {
                  render.displayName = name;
                }
              }
            });
          }
          return elementType;
        }
        var REACT_MODULE_REFERENCE;
        {
          REACT_MODULE_REFERENCE = Symbol.for("react.module.reference");
        }
        function isValidElementType(type) {
          if (typeof type === "string" || typeof type === "function") {
            return true;
          }
          if (type === REACT_FRAGMENT_TYPE || type === REACT_PROFILER_TYPE || enableDebugTracing || type === REACT_STRICT_MODE_TYPE || type === REACT_SUSPENSE_TYPE || type === REACT_SUSPENSE_LIST_TYPE || enableLegacyHidden || type === REACT_OFFSCREEN_TYPE || enableScopeAPI || enableCacheElement || enableTransitionTracing) {
            return true;
          }
          if (typeof type === "object" && type !== null) {
            if (type.$$typeof === REACT_LAZY_TYPE || type.$$typeof === REACT_MEMO_TYPE || type.$$typeof === REACT_PROVIDER_TYPE || type.$$typeof === REACT_CONTEXT_TYPE || type.$$typeof === REACT_FORWARD_REF_TYPE || // This needs to include all possible module reference object
            // types supported by any Flight configuration anywhere since
            // we don't know which Flight build this will end up being used
            // with.
            type.$$typeof === REACT_MODULE_REFERENCE || type.getModuleId !== void 0) {
              return true;
            }
          }
          return false;
        }
        function memo(type, compare) {
          {
            if (!isValidElementType(type)) {
              error("memo: The first argument must be a component. Instead received: %s", type === null ? "null" : typeof type);
            }
          }
          var elementType = {
            $$typeof: REACT_MEMO_TYPE,
            type,
            compare: compare === void 0 ? null : compare
          };
          {
            var ownName;
            Object.defineProperty(elementType, "displayName", {
              enumerable: false,
              configurable: true,
              get: function() {
                return ownName;
              },
              set: function(name) {
                ownName = name;
                if (!type.name && !type.displayName) {
                  type.displayName = name;
                }
              }
            });
          }
          return elementType;
        }
        function resolveDispatcher() {
          var dispatcher = ReactCurrentDispatcher.current;
          {
            if (dispatcher === null) {
              error("Invalid hook call. Hooks can only be called inside of the body of a function component. This could happen for one of the following reasons:\n1. You might have mismatching versions of React and the renderer (such as React DOM)\n2. You might be breaking the Rules of Hooks\n3. You might have more than one copy of React in the same app\nSee https://reactjs.org/link/invalid-hook-call for tips about how to debug and fix this problem.");
            }
          }
          return dispatcher;
        }
        function useContext(Context) {
          var dispatcher = resolveDispatcher();
          {
            if (Context._context !== void 0) {
              var realContext = Context._context;
              if (realContext.Consumer === Context) {
                error("Calling useContext(Context.Consumer) is not supported, may cause bugs, and will be removed in a future major release. Did you mean to call useContext(Context) instead?");
              } else if (realContext.Provider === Context) {
                error("Calling useContext(Context.Provider) is not supported. Did you mean to call useContext(Context) instead?");
              }
            }
          }
          return dispatcher.useContext(Context);
        }
        function useState9(initialState) {
          var dispatcher = resolveDispatcher();
          return dispatcher.useState(initialState);
        }
        function useReducer(reducer, initialArg, init) {
          var dispatcher = resolveDispatcher();
          return dispatcher.useReducer(reducer, initialArg, init);
        }
        function useRef3(initialValue) {
          var dispatcher = resolveDispatcher();
          return dispatcher.useRef(initialValue);
        }
        function useEffect7(create, deps) {
          var dispatcher = resolveDispatcher();
          return dispatcher.useEffect(create, deps);
        }
        function useInsertionEffect(create, deps) {
          var dispatcher = resolveDispatcher();
          return dispatcher.useInsertionEffect(create, deps);
        }
        function useLayoutEffect(create, deps) {
          var dispatcher = resolveDispatcher();
          return dispatcher.useLayoutEffect(create, deps);
        }
        function useCallback(callback, deps) {
          var dispatcher = resolveDispatcher();
          return dispatcher.useCallback(callback, deps);
        }
        function useMemo4(create, deps) {
          var dispatcher = resolveDispatcher();
          return dispatcher.useMemo(create, deps);
        }
        function useImperativeHandle(ref, create, deps) {
          var dispatcher = resolveDispatcher();
          return dispatcher.useImperativeHandle(ref, create, deps);
        }
        function useDebugValue(value, formatterFn) {
          {
            var dispatcher = resolveDispatcher();
            return dispatcher.useDebugValue(value, formatterFn);
          }
        }
        function useTransition() {
          var dispatcher = resolveDispatcher();
          return dispatcher.useTransition();
        }
        function useDeferredValue(value) {
          var dispatcher = resolveDispatcher();
          return dispatcher.useDeferredValue(value);
        }
        function useId() {
          var dispatcher = resolveDispatcher();
          return dispatcher.useId();
        }
        function useSyncExternalStore(subscribe, getSnapshot, getServerSnapshot) {
          var dispatcher = resolveDispatcher();
          return dispatcher.useSyncExternalStore(subscribe, getSnapshot, getServerSnapshot);
        }
        var disabledDepth = 0;
        var prevLog;
        var prevInfo;
        var prevWarn;
        var prevError;
        var prevGroup;
        var prevGroupCollapsed;
        var prevGroupEnd;
        function disabledLog() {
        }
        disabledLog.__reactDisabledLog = true;
        function disableLogs() {
          {
            if (disabledDepth === 0) {
              prevLog = console.log;
              prevInfo = console.info;
              prevWarn = console.warn;
              prevError = console.error;
              prevGroup = console.group;
              prevGroupCollapsed = console.groupCollapsed;
              prevGroupEnd = console.groupEnd;
              var props = {
                configurable: true,
                enumerable: true,
                value: disabledLog,
                writable: true
              };
              Object.defineProperties(console, {
                info: props,
                log: props,
                warn: props,
                error: props,
                group: props,
                groupCollapsed: props,
                groupEnd: props
              });
            }
            disabledDepth++;
          }
        }
        function reenableLogs() {
          {
            disabledDepth--;
            if (disabledDepth === 0) {
              var props = {
                configurable: true,
                enumerable: true,
                writable: true
              };
              Object.defineProperties(console, {
                log: assign({}, props, {
                  value: prevLog
                }),
                info: assign({}, props, {
                  value: prevInfo
                }),
                warn: assign({}, props, {
                  value: prevWarn
                }),
                error: assign({}, props, {
                  value: prevError
                }),
                group: assign({}, props, {
                  value: prevGroup
                }),
                groupCollapsed: assign({}, props, {
                  value: prevGroupCollapsed
                }),
                groupEnd: assign({}, props, {
                  value: prevGroupEnd
                })
              });
            }
            if (disabledDepth < 0) {
              error("disabledDepth fell below zero. This is a bug in React. Please file an issue.");
            }
          }
        }
        var ReactCurrentDispatcher$1 = ReactSharedInternals.ReactCurrentDispatcher;
        var prefix;
        function describeBuiltInComponentFrame(name, source, ownerFn) {
          {
            if (prefix === void 0) {
              try {
                throw Error();
              } catch (x) {
                var match = x.stack.trim().match(/\n( *(at )?)/);
                prefix = match && match[1] || "";
              }
            }
            return "\n" + prefix + name;
          }
        }
        var reentry = false;
        var componentFrameCache;
        {
          var PossiblyWeakMap = typeof WeakMap === "function" ? WeakMap : Map;
          componentFrameCache = new PossiblyWeakMap();
        }
        function describeNativeComponentFrame(fn, construct) {
          if (!fn || reentry) {
            return "";
          }
          {
            var frame = componentFrameCache.get(fn);
            if (frame !== void 0) {
              return frame;
            }
          }
          var control;
          reentry = true;
          var previousPrepareStackTrace = Error.prepareStackTrace;
          Error.prepareStackTrace = void 0;
          var previousDispatcher;
          {
            previousDispatcher = ReactCurrentDispatcher$1.current;
            ReactCurrentDispatcher$1.current = null;
            disableLogs();
          }
          try {
            if (construct) {
              var Fake = function() {
                throw Error();
              };
              Object.defineProperty(Fake.prototype, "props", {
                set: function() {
                  throw Error();
                }
              });
              if (typeof Reflect === "object" && Reflect.construct) {
                try {
                  Reflect.construct(Fake, []);
                } catch (x) {
                  control = x;
                }
                Reflect.construct(fn, [], Fake);
              } else {
                try {
                  Fake.call();
                } catch (x) {
                  control = x;
                }
                fn.call(Fake.prototype);
              }
            } else {
              try {
                throw Error();
              } catch (x) {
                control = x;
              }
              fn();
            }
          } catch (sample) {
            if (sample && control && typeof sample.stack === "string") {
              var sampleLines = sample.stack.split("\n");
              var controlLines = control.stack.split("\n");
              var s = sampleLines.length - 1;
              var c = controlLines.length - 1;
              while (s >= 1 && c >= 0 && sampleLines[s] !== controlLines[c]) {
                c--;
              }
              for (; s >= 1 && c >= 0; s--, c--) {
                if (sampleLines[s] !== controlLines[c]) {
                  if (s !== 1 || c !== 1) {
                    do {
                      s--;
                      c--;
                      if (c < 0 || sampleLines[s] !== controlLines[c]) {
                        var _frame = "\n" + sampleLines[s].replace(" at new ", " at ");
                        if (fn.displayName && _frame.includes("<anonymous>")) {
                          _frame = _frame.replace("<anonymous>", fn.displayName);
                        }
                        {
                          if (typeof fn === "function") {
                            componentFrameCache.set(fn, _frame);
                          }
                        }
                        return _frame;
                      }
                    } while (s >= 1 && c >= 0);
                  }
                  break;
                }
              }
            }
          } finally {
            reentry = false;
            {
              ReactCurrentDispatcher$1.current = previousDispatcher;
              reenableLogs();
            }
            Error.prepareStackTrace = previousPrepareStackTrace;
          }
          var name = fn ? fn.displayName || fn.name : "";
          var syntheticFrame = name ? describeBuiltInComponentFrame(name) : "";
          {
            if (typeof fn === "function") {
              componentFrameCache.set(fn, syntheticFrame);
            }
          }
          return syntheticFrame;
        }
        function describeFunctionComponentFrame(fn, source, ownerFn) {
          {
            return describeNativeComponentFrame(fn, false);
          }
        }
        function shouldConstruct(Component2) {
          var prototype = Component2.prototype;
          return !!(prototype && prototype.isReactComponent);
        }
        function describeUnknownElementTypeFrameInDEV(type, source, ownerFn) {
          if (type == null) {
            return "";
          }
          if (typeof type === "function") {
            {
              return describeNativeComponentFrame(type, shouldConstruct(type));
            }
          }
          if (typeof type === "string") {
            return describeBuiltInComponentFrame(type);
          }
          switch (type) {
            case REACT_SUSPENSE_TYPE:
              return describeBuiltInComponentFrame("Suspense");
            case REACT_SUSPENSE_LIST_TYPE:
              return describeBuiltInComponentFrame("SuspenseList");
          }
          if (typeof type === "object") {
            switch (type.$$typeof) {
              case REACT_FORWARD_REF_TYPE:
                return describeFunctionComponentFrame(type.render);
              case REACT_MEMO_TYPE:
                return describeUnknownElementTypeFrameInDEV(type.type, source, ownerFn);
              case REACT_LAZY_TYPE: {
                var lazyComponent = type;
                var payload = lazyComponent._payload;
                var init = lazyComponent._init;
                try {
                  return describeUnknownElementTypeFrameInDEV(init(payload), source, ownerFn);
                } catch (x) {
                }
              }
            }
          }
          return "";
        }
        var loggedTypeFailures = {};
        var ReactDebugCurrentFrame$1 = ReactSharedInternals.ReactDebugCurrentFrame;
        function setCurrentlyValidatingElement(element) {
          {
            if (element) {
              var owner = element._owner;
              var stack = describeUnknownElementTypeFrameInDEV(element.type, element._source, owner ? owner.type : null);
              ReactDebugCurrentFrame$1.setExtraStackFrame(stack);
            } else {
              ReactDebugCurrentFrame$1.setExtraStackFrame(null);
            }
          }
        }
        function checkPropTypes(typeSpecs, values, location, componentName, element) {
          {
            var has = Function.call.bind(hasOwnProperty);
            for (var typeSpecName in typeSpecs) {
              if (has(typeSpecs, typeSpecName)) {
                var error$1 = void 0;
                try {
                  if (typeof typeSpecs[typeSpecName] !== "function") {
                    var err = Error((componentName || "React class") + ": " + location + " type `" + typeSpecName + "` is invalid; it must be a function, usually from the `prop-types` package, but received `" + typeof typeSpecs[typeSpecName] + "`.This often happens because of typos such as `PropTypes.function` instead of `PropTypes.func`.");
                    err.name = "Invariant Violation";
                    throw err;
                  }
                  error$1 = typeSpecs[typeSpecName](values, typeSpecName, componentName, location, null, "SECRET_DO_NOT_PASS_THIS_OR_YOU_WILL_BE_FIRED");
                } catch (ex) {
                  error$1 = ex;
                }
                if (error$1 && !(error$1 instanceof Error)) {
                  setCurrentlyValidatingElement(element);
                  error("%s: type specification of %s `%s` is invalid; the type checker function must return `null` or an `Error` but returned a %s. You may have forgotten to pass an argument to the type checker creator (arrayOf, instanceOf, objectOf, oneOf, oneOfType, and shape all require an argument).", componentName || "React class", location, typeSpecName, typeof error$1);
                  setCurrentlyValidatingElement(null);
                }
                if (error$1 instanceof Error && !(error$1.message in loggedTypeFailures)) {
                  loggedTypeFailures[error$1.message] = true;
                  setCurrentlyValidatingElement(element);
                  error("Failed %s type: %s", location, error$1.message);
                  setCurrentlyValidatingElement(null);
                }
              }
            }
          }
        }
        function setCurrentlyValidatingElement$1(element) {
          {
            if (element) {
              var owner = element._owner;
              var stack = describeUnknownElementTypeFrameInDEV(element.type, element._source, owner ? owner.type : null);
              setExtraStackFrame(stack);
            } else {
              setExtraStackFrame(null);
            }
          }
        }
        var propTypesMisspellWarningShown;
        {
          propTypesMisspellWarningShown = false;
        }
        function getDeclarationErrorAddendum() {
          if (ReactCurrentOwner.current) {
            var name = getComponentNameFromType(ReactCurrentOwner.current.type);
            if (name) {
              return "\n\nCheck the render method of `" + name + "`.";
            }
          }
          return "";
        }
        function getSourceInfoErrorAddendum(source) {
          if (source !== void 0) {
            var fileName = source.fileName.replace(/^.*[\\\/]/, "");
            var lineNumber = source.lineNumber;
            return "\n\nCheck your code at " + fileName + ":" + lineNumber + ".";
          }
          return "";
        }
        function getSourceInfoErrorAddendumForProps(elementProps) {
          if (elementProps !== null && elementProps !== void 0) {
            return getSourceInfoErrorAddendum(elementProps.__source);
          }
          return "";
        }
        var ownerHasKeyUseWarning = {};
        function getCurrentComponentErrorInfo(parentType) {
          var info = getDeclarationErrorAddendum();
          if (!info) {
            var parentName = typeof parentType === "string" ? parentType : parentType.displayName || parentType.name;
            if (parentName) {
              info = "\n\nCheck the top-level render call using <" + parentName + ">.";
            }
          }
          return info;
        }
        function validateExplicitKey(element, parentType) {
          if (!element._store || element._store.validated || element.key != null) {
            return;
          }
          element._store.validated = true;
          var currentComponentErrorInfo = getCurrentComponentErrorInfo(parentType);
          if (ownerHasKeyUseWarning[currentComponentErrorInfo]) {
            return;
          }
          ownerHasKeyUseWarning[currentComponentErrorInfo] = true;
          var childOwner = "";
          if (element && element._owner && element._owner !== ReactCurrentOwner.current) {
            childOwner = " It was passed a child from " + getComponentNameFromType(element._owner.type) + ".";
          }
          {
            setCurrentlyValidatingElement$1(element);
            error('Each child in a list should have a unique "key" prop.%s%s See https://reactjs.org/link/warning-keys for more information.', currentComponentErrorInfo, childOwner);
            setCurrentlyValidatingElement$1(null);
          }
        }
        function validateChildKeys(node, parentType) {
          if (typeof node !== "object") {
            return;
          }
          if (isArray(node)) {
            for (var i = 0; i < node.length; i++) {
              var child = node[i];
              if (isValidElement(child)) {
                validateExplicitKey(child, parentType);
              }
            }
          } else if (isValidElement(node)) {
            if (node._store) {
              node._store.validated = true;
            }
          } else if (node) {
            var iteratorFn = getIteratorFn(node);
            if (typeof iteratorFn === "function") {
              if (iteratorFn !== node.entries) {
                var iterator = iteratorFn.call(node);
                var step;
                while (!(step = iterator.next()).done) {
                  if (isValidElement(step.value)) {
                    validateExplicitKey(step.value, parentType);
                  }
                }
              }
            }
          }
        }
        function validatePropTypes(element) {
          {
            var type = element.type;
            if (type === null || type === void 0 || typeof type === "string") {
              return;
            }
            var propTypes;
            if (typeof type === "function") {
              propTypes = type.propTypes;
            } else if (typeof type === "object" && (type.$$typeof === REACT_FORWARD_REF_TYPE || // Note: Memo only checks outer props here.
            // Inner props are checked in the reconciler.
            type.$$typeof === REACT_MEMO_TYPE)) {
              propTypes = type.propTypes;
            } else {
              return;
            }
            if (propTypes) {
              var name = getComponentNameFromType(type);
              checkPropTypes(propTypes, element.props, "prop", name, element);
            } else if (type.PropTypes !== void 0 && !propTypesMisspellWarningShown) {
              propTypesMisspellWarningShown = true;
              var _name = getComponentNameFromType(type);
              error("Component %s declared `PropTypes` instead of `propTypes`. Did you misspell the property assignment?", _name || "Unknown");
            }
            if (typeof type.getDefaultProps === "function" && !type.getDefaultProps.isReactClassApproved) {
              error("getDefaultProps is only used on classic React.createClass definitions. Use a static property named `defaultProps` instead.");
            }
          }
        }
        function validateFragmentProps(fragment) {
          {
            var keys = Object.keys(fragment.props);
            for (var i = 0; i < keys.length; i++) {
              var key = keys[i];
              if (key !== "children" && key !== "key") {
                setCurrentlyValidatingElement$1(fragment);
                error("Invalid prop `%s` supplied to `React.Fragment`. React.Fragment can only have `key` and `children` props.", key);
                setCurrentlyValidatingElement$1(null);
                break;
              }
            }
            if (fragment.ref !== null) {
              setCurrentlyValidatingElement$1(fragment);
              error("Invalid attribute `ref` supplied to `React.Fragment`.");
              setCurrentlyValidatingElement$1(null);
            }
          }
        }
        function createElementWithValidation(type, props, children) {
          var validType = isValidElementType(type);
          if (!validType) {
            var info = "";
            if (type === void 0 || typeof type === "object" && type !== null && Object.keys(type).length === 0) {
              info += " You likely forgot to export your component from the file it's defined in, or you might have mixed up default and named imports.";
            }
            var sourceInfo = getSourceInfoErrorAddendumForProps(props);
            if (sourceInfo) {
              info += sourceInfo;
            } else {
              info += getDeclarationErrorAddendum();
            }
            var typeString;
            if (type === null) {
              typeString = "null";
            } else if (isArray(type)) {
              typeString = "array";
            } else if (type !== void 0 && type.$$typeof === REACT_ELEMENT_TYPE) {
              typeString = "<" + (getComponentNameFromType(type.type) || "Unknown") + " />";
              info = " Did you accidentally export a JSX literal instead of a component?";
            } else {
              typeString = typeof type;
            }
            {
              error("React.createElement: type is invalid -- expected a string (for built-in components) or a class/function (for composite components) but got: %s.%s", typeString, info);
            }
          }
          var element = createElement.apply(this, arguments);
          if (element == null) {
            return element;
          }
          if (validType) {
            for (var i = 2; i < arguments.length; i++) {
              validateChildKeys(arguments[i], type);
            }
          }
          if (type === REACT_FRAGMENT_TYPE) {
            validateFragmentProps(element);
          } else {
            validatePropTypes(element);
          }
          return element;
        }
        var didWarnAboutDeprecatedCreateFactory = false;
        function createFactoryWithValidation(type) {
          var validatedFactory = createElementWithValidation.bind(null, type);
          validatedFactory.type = type;
          {
            if (!didWarnAboutDeprecatedCreateFactory) {
              didWarnAboutDeprecatedCreateFactory = true;
              warn("React.createFactory() is deprecated and will be removed in a future major release. Consider using JSX or use React.createElement() directly instead.");
            }
            Object.defineProperty(validatedFactory, "type", {
              enumerable: false,
              get: function() {
                warn("Factory.type is deprecated. Access the class directly before passing it to createFactory.");
                Object.defineProperty(this, "type", {
                  value: type
                });
                return type;
              }
            });
          }
          return validatedFactory;
        }
        function cloneElementWithValidation(element, props, children) {
          var newElement = cloneElement.apply(this, arguments);
          for (var i = 2; i < arguments.length; i++) {
            validateChildKeys(arguments[i], newElement.type);
          }
          validatePropTypes(newElement);
          return newElement;
        }
        function startTransition(scope, options) {
          var prevTransition = ReactCurrentBatchConfig.transition;
          ReactCurrentBatchConfig.transition = {};
          var currentTransition = ReactCurrentBatchConfig.transition;
          {
            ReactCurrentBatchConfig.transition._updatedFibers = /* @__PURE__ */ new Set();
          }
          try {
            scope();
          } finally {
            ReactCurrentBatchConfig.transition = prevTransition;
            {
              if (prevTransition === null && currentTransition._updatedFibers) {
                var updatedFibersCount = currentTransition._updatedFibers.size;
                if (updatedFibersCount > 10) {
                  warn("Detected a large number of updates inside startTransition. If this is due to a subscription please re-write it to use React provided hooks. Otherwise concurrent mode guarantees are off the table.");
                }
                currentTransition._updatedFibers.clear();
              }
            }
          }
        }
        var didWarnAboutMessageChannel = false;
        var enqueueTaskImpl = null;
        function enqueueTask(task) {
          if (enqueueTaskImpl === null) {
            try {
              var requireString = ("require" + Math.random()).slice(0, 7);
              var nodeRequire = module && module[requireString];
              enqueueTaskImpl = nodeRequire.call(module, "timers").setImmediate;
            } catch (_err) {
              enqueueTaskImpl = function(callback) {
                {
                  if (didWarnAboutMessageChannel === false) {
                    didWarnAboutMessageChannel = true;
                    if (typeof MessageChannel === "undefined") {
                      error("This browser does not have a MessageChannel implementation, so enqueuing tasks via await act(async () => ...) will fail. Please file an issue at https://github.com/facebook/react/issues if you encounter this warning.");
                    }
                  }
                }
                var channel = new MessageChannel();
                channel.port1.onmessage = callback;
                channel.port2.postMessage(void 0);
              };
            }
          }
          return enqueueTaskImpl(task);
        }
        var actScopeDepth = 0;
        var didWarnNoAwaitAct = false;
        function act(callback) {
          {
            var prevActScopeDepth = actScopeDepth;
            actScopeDepth++;
            if (ReactCurrentActQueue.current === null) {
              ReactCurrentActQueue.current = [];
            }
            var prevIsBatchingLegacy = ReactCurrentActQueue.isBatchingLegacy;
            var result;
            try {
              ReactCurrentActQueue.isBatchingLegacy = true;
              result = callback();
              if (!prevIsBatchingLegacy && ReactCurrentActQueue.didScheduleLegacyUpdate) {
                var queue = ReactCurrentActQueue.current;
                if (queue !== null) {
                  ReactCurrentActQueue.didScheduleLegacyUpdate = false;
                  flushActQueue(queue);
                }
              }
            } catch (error2) {
              popActScope(prevActScopeDepth);
              throw error2;
            } finally {
              ReactCurrentActQueue.isBatchingLegacy = prevIsBatchingLegacy;
            }
            if (result !== null && typeof result === "object" && typeof result.then === "function") {
              var thenableResult = result;
              var wasAwaited = false;
              var thenable = {
                then: function(resolve, reject) {
                  wasAwaited = true;
                  thenableResult.then(function(returnValue2) {
                    popActScope(prevActScopeDepth);
                    if (actScopeDepth === 0) {
                      recursivelyFlushAsyncActWork(returnValue2, resolve, reject);
                    } else {
                      resolve(returnValue2);
                    }
                  }, function(error2) {
                    popActScope(prevActScopeDepth);
                    reject(error2);
                  });
                }
              };
              {
                if (!didWarnNoAwaitAct && typeof Promise !== "undefined") {
                  Promise.resolve().then(function() {
                  }).then(function() {
                    if (!wasAwaited) {
                      didWarnNoAwaitAct = true;
                      error("You called act(async () => ...) without await. This could lead to unexpected testing behaviour, interleaving multiple act calls and mixing their scopes. You should - await act(async () => ...);");
                    }
                  });
                }
              }
              return thenable;
            } else {
              var returnValue = result;
              popActScope(prevActScopeDepth);
              if (actScopeDepth === 0) {
                var _queue = ReactCurrentActQueue.current;
                if (_queue !== null) {
                  flushActQueue(_queue);
                  ReactCurrentActQueue.current = null;
                }
                var _thenable = {
                  then: function(resolve, reject) {
                    if (ReactCurrentActQueue.current === null) {
                      ReactCurrentActQueue.current = [];
                      recursivelyFlushAsyncActWork(returnValue, resolve, reject);
                    } else {
                      resolve(returnValue);
                    }
                  }
                };
                return _thenable;
              } else {
                var _thenable2 = {
                  then: function(resolve, reject) {
                    resolve(returnValue);
                  }
                };
                return _thenable2;
              }
            }
          }
        }
        function popActScope(prevActScopeDepth) {
          {
            if (prevActScopeDepth !== actScopeDepth - 1) {
              error("You seem to have overlapping act() calls, this is not supported. Be sure to await previous act() calls before making a new one. ");
            }
            actScopeDepth = prevActScopeDepth;
          }
        }
        function recursivelyFlushAsyncActWork(returnValue, resolve, reject) {
          {
            var queue = ReactCurrentActQueue.current;
            if (queue !== null) {
              try {
                flushActQueue(queue);
                enqueueTask(function() {
                  if (queue.length === 0) {
                    ReactCurrentActQueue.current = null;
                    resolve(returnValue);
                  } else {
                    recursivelyFlushAsyncActWork(returnValue, resolve, reject);
                  }
                });
              } catch (error2) {
                reject(error2);
              }
            } else {
              resolve(returnValue);
            }
          }
        }
        var isFlushing = false;
        function flushActQueue(queue) {
          {
            if (!isFlushing) {
              isFlushing = true;
              var i = 0;
              try {
                for (; i < queue.length; i++) {
                  var callback = queue[i];
                  do {
                    callback = callback(true);
                  } while (callback !== null);
                }
                queue.length = 0;
              } catch (error2) {
                queue = queue.slice(i + 1);
                throw error2;
              } finally {
                isFlushing = false;
              }
            }
          }
        }
        var createElement$1 = createElementWithValidation;
        var cloneElement$1 = cloneElementWithValidation;
        var createFactory = createFactoryWithValidation;
        var Children = {
          map: mapChildren,
          forEach: forEachChildren,
          count: countChildren,
          toArray,
          only: onlyChild
        };
        exports.Children = Children;
        exports.Component = Component;
        exports.Fragment = REACT_FRAGMENT_TYPE;
        exports.Profiler = REACT_PROFILER_TYPE;
        exports.PureComponent = PureComponent;
        exports.StrictMode = REACT_STRICT_MODE_TYPE;
        exports.Suspense = REACT_SUSPENSE_TYPE;
        exports.__SECRET_INTERNALS_DO_NOT_USE_OR_YOU_WILL_BE_FIRED = ReactSharedInternals;
        exports.act = act;
        exports.cloneElement = cloneElement$1;
        exports.createContext = createContext;
        exports.createElement = createElement$1;
        exports.createFactory = createFactory;
        exports.createRef = createRef;
        exports.forwardRef = forwardRef;
        exports.isValidElement = isValidElement;
        exports.lazy = lazy;
        exports.memo = memo;
        exports.startTransition = startTransition;
        exports.unstable_act = act;
        exports.useCallback = useCallback;
        exports.useContext = useContext;
        exports.useDebugValue = useDebugValue;
        exports.useDeferredValue = useDeferredValue;
        exports.useEffect = useEffect7;
        exports.useId = useId;
        exports.useImperativeHandle = useImperativeHandle;
        exports.useInsertionEffect = useInsertionEffect;
        exports.useLayoutEffect = useLayoutEffect;
        exports.useMemo = useMemo4;
        exports.useReducer = useReducer;
        exports.useRef = useRef3;
        exports.useState = useState9;
        exports.useSyncExternalStore = useSyncExternalStore;
        exports.useTransition = useTransition;
        exports.version = ReactVersion;
        if (typeof __REACT_DEVTOOLS_GLOBAL_HOOK__ !== "undefined" && typeof __REACT_DEVTOOLS_GLOBAL_HOOK__.registerInternalModuleStop === "function") {
          __REACT_DEVTOOLS_GLOBAL_HOOK__.registerInternalModuleStop(new Error());
        }
      })();
    }
  }
});

// node_modules/react/index.js
var require_react = __commonJS({
  "node_modules/react/index.js"(exports, module) {
    "use strict";
    if (process.env.NODE_ENV === "production") {
      module.exports = require_react_production_min();
    } else {
      module.exports = require_react_development();
    }
  }
});

// node_modules/react/cjs/react-jsx-runtime.production.min.js
var require_react_jsx_runtime_production_min = __commonJS({
  "node_modules/react/cjs/react-jsx-runtime.production.min.js"(exports) {
    "use strict";
    var f = require_react();
    var k = Symbol.for("react.element");
    var l = Symbol.for("react.fragment");
    var m = Object.prototype.hasOwnProperty;
    var n = f.__SECRET_INTERNALS_DO_NOT_USE_OR_YOU_WILL_BE_FIRED.ReactCurrentOwner;
    var p = { key: true, ref: true, __self: true, __source: true };
    function q(c, a, g) {
      var b, d = {}, e = null, h = null;
      void 0 !== g && (e = "" + g);
      void 0 !== a.key && (e = "" + a.key);
      void 0 !== a.ref && (h = a.ref);
      for (b in a) m.call(a, b) && !p.hasOwnProperty(b) && (d[b] = a[b]);
      if (c && c.defaultProps) for (b in a = c.defaultProps, a) void 0 === d[b] && (d[b] = a[b]);
      return { $$typeof: k, type: c, key: e, ref: h, props: d, _owner: n.current };
    }
    exports.Fragment = l;
    exports.jsx = q;
    exports.jsxs = q;
  }
});

// node_modules/react/cjs/react-jsx-runtime.development.js
var require_react_jsx_runtime_development = __commonJS({
  "node_modules/react/cjs/react-jsx-runtime.development.js"(exports) {
    "use strict";
    if (process.env.NODE_ENV !== "production") {
      (function() {
        "use strict";
        var React = require_react();
        var REACT_ELEMENT_TYPE = Symbol.for("react.element");
        var REACT_PORTAL_TYPE = Symbol.for("react.portal");
        var REACT_FRAGMENT_TYPE = Symbol.for("react.fragment");
        var REACT_STRICT_MODE_TYPE = Symbol.for("react.strict_mode");
        var REACT_PROFILER_TYPE = Symbol.for("react.profiler");
        var REACT_PROVIDER_TYPE = Symbol.for("react.provider");
        var REACT_CONTEXT_TYPE = Symbol.for("react.context");
        var REACT_FORWARD_REF_TYPE = Symbol.for("react.forward_ref");
        var REACT_SUSPENSE_TYPE = Symbol.for("react.suspense");
        var REACT_SUSPENSE_LIST_TYPE = Symbol.for("react.suspense_list");
        var REACT_MEMO_TYPE = Symbol.for("react.memo");
        var REACT_LAZY_TYPE = Symbol.for("react.lazy");
        var REACT_OFFSCREEN_TYPE = Symbol.for("react.offscreen");
        var MAYBE_ITERATOR_SYMBOL = Symbol.iterator;
        var FAUX_ITERATOR_SYMBOL = "@@iterator";
        function getIteratorFn(maybeIterable) {
          if (maybeIterable === null || typeof maybeIterable !== "object") {
            return null;
          }
          var maybeIterator = MAYBE_ITERATOR_SYMBOL && maybeIterable[MAYBE_ITERATOR_SYMBOL] || maybeIterable[FAUX_ITERATOR_SYMBOL];
          if (typeof maybeIterator === "function") {
            return maybeIterator;
          }
          return null;
        }
        var ReactSharedInternals = React.__SECRET_INTERNALS_DO_NOT_USE_OR_YOU_WILL_BE_FIRED;
        function error(format) {
          {
            {
              for (var _len2 = arguments.length, args = new Array(_len2 > 1 ? _len2 - 1 : 0), _key2 = 1; _key2 < _len2; _key2++) {
                args[_key2 - 1] = arguments[_key2];
              }
              printWarning("error", format, args);
            }
          }
        }
        function printWarning(level, format, args) {
          {
            var ReactDebugCurrentFrame2 = ReactSharedInternals.ReactDebugCurrentFrame;
            var stack = ReactDebugCurrentFrame2.getStackAddendum();
            if (stack !== "") {
              format += "%s";
              args = args.concat([stack]);
            }
            var argsWithFormat = args.map(function(item) {
              return String(item);
            });
            argsWithFormat.unshift("Warning: " + format);
            Function.prototype.apply.call(console[level], console, argsWithFormat);
          }
        }
        var enableScopeAPI = false;
        var enableCacheElement = false;
        var enableTransitionTracing = false;
        var enableLegacyHidden = false;
        var enableDebugTracing = false;
        var REACT_MODULE_REFERENCE;
        {
          REACT_MODULE_REFERENCE = Symbol.for("react.module.reference");
        }
        function isValidElementType(type) {
          if (typeof type === "string" || typeof type === "function") {
            return true;
          }
          if (type === REACT_FRAGMENT_TYPE || type === REACT_PROFILER_TYPE || enableDebugTracing || type === REACT_STRICT_MODE_TYPE || type === REACT_SUSPENSE_TYPE || type === REACT_SUSPENSE_LIST_TYPE || enableLegacyHidden || type === REACT_OFFSCREEN_TYPE || enableScopeAPI || enableCacheElement || enableTransitionTracing) {
            return true;
          }
          if (typeof type === "object" && type !== null) {
            if (type.$$typeof === REACT_LAZY_TYPE || type.$$typeof === REACT_MEMO_TYPE || type.$$typeof === REACT_PROVIDER_TYPE || type.$$typeof === REACT_CONTEXT_TYPE || type.$$typeof === REACT_FORWARD_REF_TYPE || // This needs to include all possible module reference object
            // types supported by any Flight configuration anywhere since
            // we don't know which Flight build this will end up being used
            // with.
            type.$$typeof === REACT_MODULE_REFERENCE || type.getModuleId !== void 0) {
              return true;
            }
          }
          return false;
        }
        function getWrappedName(outerType, innerType, wrapperName) {
          var displayName = outerType.displayName;
          if (displayName) {
            return displayName;
          }
          var functionName = innerType.displayName || innerType.name || "";
          return functionName !== "" ? wrapperName + "(" + functionName + ")" : wrapperName;
        }
        function getContextName(type) {
          return type.displayName || "Context";
        }
        function getComponentNameFromType(type) {
          if (type == null) {
            return null;
          }
          {
            if (typeof type.tag === "number") {
              error("Received an unexpected object in getComponentNameFromType(). This is likely a bug in React. Please file an issue.");
            }
          }
          if (typeof type === "function") {
            return type.displayName || type.name || null;
          }
          if (typeof type === "string") {
            return type;
          }
          switch (type) {
            case REACT_FRAGMENT_TYPE:
              return "Fragment";
            case REACT_PORTAL_TYPE:
              return "Portal";
            case REACT_PROFILER_TYPE:
              return "Profiler";
            case REACT_STRICT_MODE_TYPE:
              return "StrictMode";
            case REACT_SUSPENSE_TYPE:
              return "Suspense";
            case REACT_SUSPENSE_LIST_TYPE:
              return "SuspenseList";
          }
          if (typeof type === "object") {
            switch (type.$$typeof) {
              case REACT_CONTEXT_TYPE:
                var context = type;
                return getContextName(context) + ".Consumer";
              case REACT_PROVIDER_TYPE:
                var provider = type;
                return getContextName(provider._context) + ".Provider";
              case REACT_FORWARD_REF_TYPE:
                return getWrappedName(type, type.render, "ForwardRef");
              case REACT_MEMO_TYPE:
                var outerName = type.displayName || null;
                if (outerName !== null) {
                  return outerName;
                }
                return getComponentNameFromType(type.type) || "Memo";
              case REACT_LAZY_TYPE: {
                var lazyComponent = type;
                var payload = lazyComponent._payload;
                var init = lazyComponent._init;
                try {
                  return getComponentNameFromType(init(payload));
                } catch (x) {
                  return null;
                }
              }
            }
          }
          return null;
        }
        var assign = Object.assign;
        var disabledDepth = 0;
        var prevLog;
        var prevInfo;
        var prevWarn;
        var prevError;
        var prevGroup;
        var prevGroupCollapsed;
        var prevGroupEnd;
        function disabledLog() {
        }
        disabledLog.__reactDisabledLog = true;
        function disableLogs() {
          {
            if (disabledDepth === 0) {
              prevLog = console.log;
              prevInfo = console.info;
              prevWarn = console.warn;
              prevError = console.error;
              prevGroup = console.group;
              prevGroupCollapsed = console.groupCollapsed;
              prevGroupEnd = console.groupEnd;
              var props = {
                configurable: true,
                enumerable: true,
                value: disabledLog,
                writable: true
              };
              Object.defineProperties(console, {
                info: props,
                log: props,
                warn: props,
                error: props,
                group: props,
                groupCollapsed: props,
                groupEnd: props
              });
            }
            disabledDepth++;
          }
        }
        function reenableLogs() {
          {
            disabledDepth--;
            if (disabledDepth === 0) {
              var props = {
                configurable: true,
                enumerable: true,
                writable: true
              };
              Object.defineProperties(console, {
                log: assign({}, props, {
                  value: prevLog
                }),
                info: assign({}, props, {
                  value: prevInfo
                }),
                warn: assign({}, props, {
                  value: prevWarn
                }),
                error: assign({}, props, {
                  value: prevError
                }),
                group: assign({}, props, {
                  value: prevGroup
                }),
                groupCollapsed: assign({}, props, {
                  value: prevGroupCollapsed
                }),
                groupEnd: assign({}, props, {
                  value: prevGroupEnd
                })
              });
            }
            if (disabledDepth < 0) {
              error("disabledDepth fell below zero. This is a bug in React. Please file an issue.");
            }
          }
        }
        var ReactCurrentDispatcher = ReactSharedInternals.ReactCurrentDispatcher;
        var prefix;
        function describeBuiltInComponentFrame(name, source, ownerFn) {
          {
            if (prefix === void 0) {
              try {
                throw Error();
              } catch (x) {
                var match = x.stack.trim().match(/\n( *(at )?)/);
                prefix = match && match[1] || "";
              }
            }
            return "\n" + prefix + name;
          }
        }
        var reentry = false;
        var componentFrameCache;
        {
          var PossiblyWeakMap = typeof WeakMap === "function" ? WeakMap : Map;
          componentFrameCache = new PossiblyWeakMap();
        }
        function describeNativeComponentFrame(fn, construct) {
          if (!fn || reentry) {
            return "";
          }
          {
            var frame = componentFrameCache.get(fn);
            if (frame !== void 0) {
              return frame;
            }
          }
          var control;
          reentry = true;
          var previousPrepareStackTrace = Error.prepareStackTrace;
          Error.prepareStackTrace = void 0;
          var previousDispatcher;
          {
            previousDispatcher = ReactCurrentDispatcher.current;
            ReactCurrentDispatcher.current = null;
            disableLogs();
          }
          try {
            if (construct) {
              var Fake = function() {
                throw Error();
              };
              Object.defineProperty(Fake.prototype, "props", {
                set: function() {
                  throw Error();
                }
              });
              if (typeof Reflect === "object" && Reflect.construct) {
                try {
                  Reflect.construct(Fake, []);
                } catch (x) {
                  control = x;
                }
                Reflect.construct(fn, [], Fake);
              } else {
                try {
                  Fake.call();
                } catch (x) {
                  control = x;
                }
                fn.call(Fake.prototype);
              }
            } else {
              try {
                throw Error();
              } catch (x) {
                control = x;
              }
              fn();
            }
          } catch (sample) {
            if (sample && control && typeof sample.stack === "string") {
              var sampleLines = sample.stack.split("\n");
              var controlLines = control.stack.split("\n");
              var s = sampleLines.length - 1;
              var c = controlLines.length - 1;
              while (s >= 1 && c >= 0 && sampleLines[s] !== controlLines[c]) {
                c--;
              }
              for (; s >= 1 && c >= 0; s--, c--) {
                if (sampleLines[s] !== controlLines[c]) {
                  if (s !== 1 || c !== 1) {
                    do {
                      s--;
                      c--;
                      if (c < 0 || sampleLines[s] !== controlLines[c]) {
                        var _frame = "\n" + sampleLines[s].replace(" at new ", " at ");
                        if (fn.displayName && _frame.includes("<anonymous>")) {
                          _frame = _frame.replace("<anonymous>", fn.displayName);
                        }
                        {
                          if (typeof fn === "function") {
                            componentFrameCache.set(fn, _frame);
                          }
                        }
                        return _frame;
                      }
                    } while (s >= 1 && c >= 0);
                  }
                  break;
                }
              }
            }
          } finally {
            reentry = false;
            {
              ReactCurrentDispatcher.current = previousDispatcher;
              reenableLogs();
            }
            Error.prepareStackTrace = previousPrepareStackTrace;
          }
          var name = fn ? fn.displayName || fn.name : "";
          var syntheticFrame = name ? describeBuiltInComponentFrame(name) : "";
          {
            if (typeof fn === "function") {
              componentFrameCache.set(fn, syntheticFrame);
            }
          }
          return syntheticFrame;
        }
        function describeFunctionComponentFrame(fn, source, ownerFn) {
          {
            return describeNativeComponentFrame(fn, false);
          }
        }
        function shouldConstruct(Component) {
          var prototype = Component.prototype;
          return !!(prototype && prototype.isReactComponent);
        }
        function describeUnknownElementTypeFrameInDEV(type, source, ownerFn) {
          if (type == null) {
            return "";
          }
          if (typeof type === "function") {
            {
              return describeNativeComponentFrame(type, shouldConstruct(type));
            }
          }
          if (typeof type === "string") {
            return describeBuiltInComponentFrame(type);
          }
          switch (type) {
            case REACT_SUSPENSE_TYPE:
              return describeBuiltInComponentFrame("Suspense");
            case REACT_SUSPENSE_LIST_TYPE:
              return describeBuiltInComponentFrame("SuspenseList");
          }
          if (typeof type === "object") {
            switch (type.$$typeof) {
              case REACT_FORWARD_REF_TYPE:
                return describeFunctionComponentFrame(type.render);
              case REACT_MEMO_TYPE:
                return describeUnknownElementTypeFrameInDEV(type.type, source, ownerFn);
              case REACT_LAZY_TYPE: {
                var lazyComponent = type;
                var payload = lazyComponent._payload;
                var init = lazyComponent._init;
                try {
                  return describeUnknownElementTypeFrameInDEV(init(payload), source, ownerFn);
                } catch (x) {
                }
              }
            }
          }
          return "";
        }
        var hasOwnProperty = Object.prototype.hasOwnProperty;
        var loggedTypeFailures = {};
        var ReactDebugCurrentFrame = ReactSharedInternals.ReactDebugCurrentFrame;
        function setCurrentlyValidatingElement(element) {
          {
            if (element) {
              var owner = element._owner;
              var stack = describeUnknownElementTypeFrameInDEV(element.type, element._source, owner ? owner.type : null);
              ReactDebugCurrentFrame.setExtraStackFrame(stack);
            } else {
              ReactDebugCurrentFrame.setExtraStackFrame(null);
            }
          }
        }
        function checkPropTypes(typeSpecs, values, location, componentName, element) {
          {
            var has = Function.call.bind(hasOwnProperty);
            for (var typeSpecName in typeSpecs) {
              if (has(typeSpecs, typeSpecName)) {
                var error$1 = void 0;
                try {
                  if (typeof typeSpecs[typeSpecName] !== "function") {
                    var err = Error((componentName || "React class") + ": " + location + " type `" + typeSpecName + "` is invalid; it must be a function, usually from the `prop-types` package, but received `" + typeof typeSpecs[typeSpecName] + "`.This often happens because of typos such as `PropTypes.function` instead of `PropTypes.func`.");
                    err.name = "Invariant Violation";
                    throw err;
                  }
                  error$1 = typeSpecs[typeSpecName](values, typeSpecName, componentName, location, null, "SECRET_DO_NOT_PASS_THIS_OR_YOU_WILL_BE_FIRED");
                } catch (ex) {
                  error$1 = ex;
                }
                if (error$1 && !(error$1 instanceof Error)) {
                  setCurrentlyValidatingElement(element);
                  error("%s: type specification of %s `%s` is invalid; the type checker function must return `null` or an `Error` but returned a %s. You may have forgotten to pass an argument to the type checker creator (arrayOf, instanceOf, objectOf, oneOf, oneOfType, and shape all require an argument).", componentName || "React class", location, typeSpecName, typeof error$1);
                  setCurrentlyValidatingElement(null);
                }
                if (error$1 instanceof Error && !(error$1.message in loggedTypeFailures)) {
                  loggedTypeFailures[error$1.message] = true;
                  setCurrentlyValidatingElement(element);
                  error("Failed %s type: %s", location, error$1.message);
                  setCurrentlyValidatingElement(null);
                }
              }
            }
          }
        }
        var isArrayImpl = Array.isArray;
        function isArray(a) {
          return isArrayImpl(a);
        }
        function typeName(value) {
          {
            var hasToStringTag = typeof Symbol === "function" && Symbol.toStringTag;
            var type = hasToStringTag && value[Symbol.toStringTag] || value.constructor.name || "Object";
            return type;
          }
        }
        function willCoercionThrow(value) {
          {
            try {
              testStringCoercion(value);
              return false;
            } catch (e) {
              return true;
            }
          }
        }
        function testStringCoercion(value) {
          return "" + value;
        }
        function checkKeyStringCoercion(value) {
          {
            if (willCoercionThrow(value)) {
              error("The provided key is an unsupported type %s. This value must be coerced to a string before before using it here.", typeName(value));
              return testStringCoercion(value);
            }
          }
        }
        var ReactCurrentOwner = ReactSharedInternals.ReactCurrentOwner;
        var RESERVED_PROPS = {
          key: true,
          ref: true,
          __self: true,
          __source: true
        };
        var specialPropKeyWarningShown;
        var specialPropRefWarningShown;
        var didWarnAboutStringRefs;
        {
          didWarnAboutStringRefs = {};
        }
        function hasValidRef(config) {
          {
            if (hasOwnProperty.call(config, "ref")) {
              var getter = Object.getOwnPropertyDescriptor(config, "ref").get;
              if (getter && getter.isReactWarning) {
                return false;
              }
            }
          }
          return config.ref !== void 0;
        }
        function hasValidKey(config) {
          {
            if (hasOwnProperty.call(config, "key")) {
              var getter = Object.getOwnPropertyDescriptor(config, "key").get;
              if (getter && getter.isReactWarning) {
                return false;
              }
            }
          }
          return config.key !== void 0;
        }
        function warnIfStringRefCannotBeAutoConverted(config, self) {
          {
            if (typeof config.ref === "string" && ReactCurrentOwner.current && self && ReactCurrentOwner.current.stateNode !== self) {
              var componentName = getComponentNameFromType(ReactCurrentOwner.current.type);
              if (!didWarnAboutStringRefs[componentName]) {
                error('Component "%s" contains the string ref "%s". Support for string refs will be removed in a future major release. This case cannot be automatically converted to an arrow function. We ask you to manually fix this case by using useRef() or createRef() instead. Learn more about using refs safely here: https://reactjs.org/link/strict-mode-string-ref', getComponentNameFromType(ReactCurrentOwner.current.type), config.ref);
                didWarnAboutStringRefs[componentName] = true;
              }
            }
          }
        }
        function defineKeyPropWarningGetter(props, displayName) {
          {
            var warnAboutAccessingKey = function() {
              if (!specialPropKeyWarningShown) {
                specialPropKeyWarningShown = true;
                error("%s: `key` is not a prop. Trying to access it will result in `undefined` being returned. If you need to access the same value within the child component, you should pass it as a different prop. (https://reactjs.org/link/special-props)", displayName);
              }
            };
            warnAboutAccessingKey.isReactWarning = true;
            Object.defineProperty(props, "key", {
              get: warnAboutAccessingKey,
              configurable: true
            });
          }
        }
        function defineRefPropWarningGetter(props, displayName) {
          {
            var warnAboutAccessingRef = function() {
              if (!specialPropRefWarningShown) {
                specialPropRefWarningShown = true;
                error("%s: `ref` is not a prop. Trying to access it will result in `undefined` being returned. If you need to access the same value within the child component, you should pass it as a different prop. (https://reactjs.org/link/special-props)", displayName);
              }
            };
            warnAboutAccessingRef.isReactWarning = true;
            Object.defineProperty(props, "ref", {
              get: warnAboutAccessingRef,
              configurable: true
            });
          }
        }
        var ReactElement = function(type, key, ref, self, source, owner, props) {
          var element = {
            // This tag allows us to uniquely identify this as a React Element
            $$typeof: REACT_ELEMENT_TYPE,
            // Built-in properties that belong on the element
            type,
            key,
            ref,
            props,
            // Record the component responsible for creating this element.
            _owner: owner
          };
          {
            element._store = {};
            Object.defineProperty(element._store, "validated", {
              configurable: false,
              enumerable: false,
              writable: true,
              value: false
            });
            Object.defineProperty(element, "_self", {
              configurable: false,
              enumerable: false,
              writable: false,
              value: self
            });
            Object.defineProperty(element, "_source", {
              configurable: false,
              enumerable: false,
              writable: false,
              value: source
            });
            if (Object.freeze) {
              Object.freeze(element.props);
              Object.freeze(element);
            }
          }
          return element;
        };
        function jsxDEV(type, config, maybeKey, source, self) {
          {
            var propName;
            var props = {};
            var key = null;
            var ref = null;
            if (maybeKey !== void 0) {
              {
                checkKeyStringCoercion(maybeKey);
              }
              key = "" + maybeKey;
            }
            if (hasValidKey(config)) {
              {
                checkKeyStringCoercion(config.key);
              }
              key = "" + config.key;
            }
            if (hasValidRef(config)) {
              ref = config.ref;
              warnIfStringRefCannotBeAutoConverted(config, self);
            }
            for (propName in config) {
              if (hasOwnProperty.call(config, propName) && !RESERVED_PROPS.hasOwnProperty(propName)) {
                props[propName] = config[propName];
              }
            }
            if (type && type.defaultProps) {
              var defaultProps = type.defaultProps;
              for (propName in defaultProps) {
                if (props[propName] === void 0) {
                  props[propName] = defaultProps[propName];
                }
              }
            }
            if (key || ref) {
              var displayName = typeof type === "function" ? type.displayName || type.name || "Unknown" : type;
              if (key) {
                defineKeyPropWarningGetter(props, displayName);
              }
              if (ref) {
                defineRefPropWarningGetter(props, displayName);
              }
            }
            return ReactElement(type, key, ref, self, source, ReactCurrentOwner.current, props);
          }
        }
        var ReactCurrentOwner$1 = ReactSharedInternals.ReactCurrentOwner;
        var ReactDebugCurrentFrame$1 = ReactSharedInternals.ReactDebugCurrentFrame;
        function setCurrentlyValidatingElement$1(element) {
          {
            if (element) {
              var owner = element._owner;
              var stack = describeUnknownElementTypeFrameInDEV(element.type, element._source, owner ? owner.type : null);
              ReactDebugCurrentFrame$1.setExtraStackFrame(stack);
            } else {
              ReactDebugCurrentFrame$1.setExtraStackFrame(null);
            }
          }
        }
        var propTypesMisspellWarningShown;
        {
          propTypesMisspellWarningShown = false;
        }
        function isValidElement(object) {
          {
            return typeof object === "object" && object !== null && object.$$typeof === REACT_ELEMENT_TYPE;
          }
        }
        function getDeclarationErrorAddendum() {
          {
            if (ReactCurrentOwner$1.current) {
              var name = getComponentNameFromType(ReactCurrentOwner$1.current.type);
              if (name) {
                return "\n\nCheck the render method of `" + name + "`.";
              }
            }
            return "";
          }
        }
        function getSourceInfoErrorAddendum(source) {
          {
            if (source !== void 0) {
              var fileName = source.fileName.replace(/^.*[\\\/]/, "");
              var lineNumber = source.lineNumber;
              return "\n\nCheck your code at " + fileName + ":" + lineNumber + ".";
            }
            return "";
          }
        }
        var ownerHasKeyUseWarning = {};
        function getCurrentComponentErrorInfo(parentType) {
          {
            var info = getDeclarationErrorAddendum();
            if (!info) {
              var parentName = typeof parentType === "string" ? parentType : parentType.displayName || parentType.name;
              if (parentName) {
                info = "\n\nCheck the top-level render call using <" + parentName + ">.";
              }
            }
            return info;
          }
        }
        function validateExplicitKey(element, parentType) {
          {
            if (!element._store || element._store.validated || element.key != null) {
              return;
            }
            element._store.validated = true;
            var currentComponentErrorInfo = getCurrentComponentErrorInfo(parentType);
            if (ownerHasKeyUseWarning[currentComponentErrorInfo]) {
              return;
            }
            ownerHasKeyUseWarning[currentComponentErrorInfo] = true;
            var childOwner = "";
            if (element && element._owner && element._owner !== ReactCurrentOwner$1.current) {
              childOwner = " It was passed a child from " + getComponentNameFromType(element._owner.type) + ".";
            }
            setCurrentlyValidatingElement$1(element);
            error('Each child in a list should have a unique "key" prop.%s%s See https://reactjs.org/link/warning-keys for more information.', currentComponentErrorInfo, childOwner);
            setCurrentlyValidatingElement$1(null);
          }
        }
        function validateChildKeys(node, parentType) {
          {
            if (typeof node !== "object") {
              return;
            }
            if (isArray(node)) {
              for (var i = 0; i < node.length; i++) {
                var child = node[i];
                if (isValidElement(child)) {
                  validateExplicitKey(child, parentType);
                }
              }
            } else if (isValidElement(node)) {
              if (node._store) {
                node._store.validated = true;
              }
            } else if (node) {
              var iteratorFn = getIteratorFn(node);
              if (typeof iteratorFn === "function") {
                if (iteratorFn !== node.entries) {
                  var iterator = iteratorFn.call(node);
                  var step;
                  while (!(step = iterator.next()).done) {
                    if (isValidElement(step.value)) {
                      validateExplicitKey(step.value, parentType);
                    }
                  }
                }
              }
            }
          }
        }
        function validatePropTypes(element) {
          {
            var type = element.type;
            if (type === null || type === void 0 || typeof type === "string") {
              return;
            }
            var propTypes;
            if (typeof type === "function") {
              propTypes = type.propTypes;
            } else if (typeof type === "object" && (type.$$typeof === REACT_FORWARD_REF_TYPE || // Note: Memo only checks outer props here.
            // Inner props are checked in the reconciler.
            type.$$typeof === REACT_MEMO_TYPE)) {
              propTypes = type.propTypes;
            } else {
              return;
            }
            if (propTypes) {
              var name = getComponentNameFromType(type);
              checkPropTypes(propTypes, element.props, "prop", name, element);
            } else if (type.PropTypes !== void 0 && !propTypesMisspellWarningShown) {
              propTypesMisspellWarningShown = true;
              var _name = getComponentNameFromType(type);
              error("Component %s declared `PropTypes` instead of `propTypes`. Did you misspell the property assignment?", _name || "Unknown");
            }
            if (typeof type.getDefaultProps === "function" && !type.getDefaultProps.isReactClassApproved) {
              error("getDefaultProps is only used on classic React.createClass definitions. Use a static property named `defaultProps` instead.");
            }
          }
        }
        function validateFragmentProps(fragment) {
          {
            var keys = Object.keys(fragment.props);
            for (var i = 0; i < keys.length; i++) {
              var key = keys[i];
              if (key !== "children" && key !== "key") {
                setCurrentlyValidatingElement$1(fragment);
                error("Invalid prop `%s` supplied to `React.Fragment`. React.Fragment can only have `key` and `children` props.", key);
                setCurrentlyValidatingElement$1(null);
                break;
              }
            }
            if (fragment.ref !== null) {
              setCurrentlyValidatingElement$1(fragment);
              error("Invalid attribute `ref` supplied to `React.Fragment`.");
              setCurrentlyValidatingElement$1(null);
            }
          }
        }
        var didWarnAboutKeySpread = {};
        function jsxWithValidation(type, props, key, isStaticChildren, source, self) {
          {
            var validType = isValidElementType(type);
            if (!validType) {
              var info = "";
              if (type === void 0 || typeof type === "object" && type !== null && Object.keys(type).length === 0) {
                info += " You likely forgot to export your component from the file it's defined in, or you might have mixed up default and named imports.";
              }
              var sourceInfo = getSourceInfoErrorAddendum(source);
              if (sourceInfo) {
                info += sourceInfo;
              } else {
                info += getDeclarationErrorAddendum();
              }
              var typeString;
              if (type === null) {
                typeString = "null";
              } else if (isArray(type)) {
                typeString = "array";
              } else if (type !== void 0 && type.$$typeof === REACT_ELEMENT_TYPE) {
                typeString = "<" + (getComponentNameFromType(type.type) || "Unknown") + " />";
                info = " Did you accidentally export a JSX literal instead of a component?";
              } else {
                typeString = typeof type;
              }
              error("React.jsx: type is invalid -- expected a string (for built-in components) or a class/function (for composite components) but got: %s.%s", typeString, info);
            }
            var element = jsxDEV(type, props, key, source, self);
            if (element == null) {
              return element;
            }
            if (validType) {
              var children = props.children;
              if (children !== void 0) {
                if (isStaticChildren) {
                  if (isArray(children)) {
                    for (var i = 0; i < children.length; i++) {
                      validateChildKeys(children[i], type);
                    }
                    if (Object.freeze) {
                      Object.freeze(children);
                    }
                  } else {
                    error("React.jsx: Static children should always be an array. You are likely explicitly calling React.jsxs or React.jsxDEV. Use the Babel transform instead.");
                  }
                } else {
                  validateChildKeys(children, type);
                }
              }
            }
            {
              if (hasOwnProperty.call(props, "key")) {
                var componentName = getComponentNameFromType(type);
                var keys = Object.keys(props).filter(function(k) {
                  return k !== "key";
                });
                var beforeExample = keys.length > 0 ? "{key: someKey, " + keys.join(": ..., ") + ": ...}" : "{key: someKey}";
                if (!didWarnAboutKeySpread[componentName + beforeExample]) {
                  var afterExample = keys.length > 0 ? "{" + keys.join(": ..., ") + ": ...}" : "{}";
                  error('A props object containing a "key" prop is being spread into JSX:\n  let props = %s;\n  <%s {...props} />\nReact keys must be passed directly to JSX without using spread:\n  let props = %s;\n  <%s key={someKey} {...props} />', beforeExample, componentName, afterExample, componentName);
                  didWarnAboutKeySpread[componentName + beforeExample] = true;
                }
              }
            }
            if (type === REACT_FRAGMENT_TYPE) {
              validateFragmentProps(element);
            } else {
              validatePropTypes(element);
            }
            return element;
          }
        }
        function jsxWithValidationStatic(type, props, key) {
          {
            return jsxWithValidation(type, props, key, true);
          }
        }
        function jsxWithValidationDynamic(type, props, key) {
          {
            return jsxWithValidation(type, props, key, false);
          }
        }
        var jsx7 = jsxWithValidationDynamic;
        var jsxs7 = jsxWithValidationStatic;
        exports.Fragment = REACT_FRAGMENT_TYPE;
        exports.jsx = jsx7;
        exports.jsxs = jsxs7;
      })();
    }
  }
});

// node_modules/react/jsx-runtime.js
var require_jsx_runtime = __commonJS({
  "node_modules/react/jsx-runtime.js"(exports, module) {
    "use strict";
    if (process.env.NODE_ENV === "production") {
      module.exports = require_react_jsx_runtime_production_min();
    } else {
      module.exports = require_react_jsx_runtime_development();
    }
  }
});

// src/lib/kiraPrompts.ts
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

// src/screens/RecoveryPanel.tsx
var import_react8 = __toESM(require_react(), 1);

// src/components/ZeitraumWahl.tsx
var import_react = __toESM(require_react(), 1);
var import_jsx_runtime = __toESM(require_jsx_runtime(), 1);

// src/components/BefindenHinweis.tsx
var import_react5 = __toESM(require_react(), 1);

// src/lib/engine.ts
var import_react4 = __toESM(require_react(), 1);

// src/lib/garmin.ts
var import_react2 = __toESM(require_react(), 1);
var CACHE_MS = 10 * 60 * 1e3;

// src/lib/useErfahrung.ts
var import_react3 = __toESM(require_react(), 1);

// src/lib/monotony.ts
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

// src/lib/chatContext.ts
var OLLAMA_HTTP = "http://localhost:11434";
var OLLAMA_HTTPS = "https://localhost:11435";
var IST_SAFARI = typeof navigator !== "undefined" && /Safari\//.test(navigator.userAgent) && !/Chrome|Chromium|CriOS|Edg|Comet/i.test(navigator.userAgent);
var KANDIDATEN = IST_SAFARI ? [OLLAMA_HTTPS, OLLAMA_HTTP] : [OLLAMA_HTTP, OLLAMA_HTTPS];
var ollamaBasis = KANDIDATEN[0];
var KIRA_SYSTEM = [
  "Du bist Kira, Sportwissenschaftlerin und pers\xF6nliche Ausdauer-Trainerin von Marc (Jahrgang 1975; Rennrad, Gravel, Laufen, Bergsport).",
  "Sprich Marc per Du an, antworte auf Deutsch, klar und konkret, ohne Emojis.",
  "St\xFCtze dich zuerst auf die mitgelieferten Cockpit-Daten (JSON unten). Nenne die Zahlen, auf die du dich beziehst.",
  "Erfinde keine Werte. Wenn Daten fehlen oder unklar sind, sag das offen.",
  'Halte dich an die Regeln aus dem Feld "regeln" (Tagesentscheid, Zielbereiche).',
  "Gib am Ende eine klare Empfehlung (z. B. Einheit, Dauer, Zone oder Pause).",
  "Bei Schmerzen, Krankheit oder medizinischen Fragen: auf Arzt bzw. Physiotherapie verweisen."
].join("\n");

// src/lib/engine.ts
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

// src/lib/api.ts
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

// src/components/BefindenHinweis.tsx
var import_jsx_runtime2 = __toESM(require_jsx_runtime(), 1);

// src/lib/format.ts
function fmtNum(v, digits = 0) {
  if (v === null || v === void 0 || !Number.isFinite(v)) return "\u2014";
  return v.toLocaleString("de-DE", {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits
  });
}

// src/components/UI.tsx
var import_jsx_runtime3 = __toESM(require_jsx_runtime(), 1);

// src/screens/CorrelationPanel.tsx
var import_react6 = __toESM(require_react(), 1);
var import_jsx_runtime4 = __toESM(require_jsx_runtime(), 1);
var VARS = [
  { key: "readiness", label: "Readiness", short: "Ready", unit: "" },
  { key: "hrv", label: "HRV", short: "HRV", unit: "ms" },
  { key: "rhr", label: "RHR", short: "RHR", unit: "bpm" },
  { key: "sleepH", label: "Schlafdauer", short: "Schlaf h", unit: "h" },
  { key: "sleepScore", label: "Schlafscore", short: "Score", unit: "" },
  { key: "essDay", label: "Last", short: "Last", unit: "ESS" }
];
var VAR = Object.fromEntries(VARS.map((v) => [v.key, v]));

// src/components/GarminSportProfilTile.tsx
var import_react7 = __toESM(require_react(), 1);
var import_jsx_runtime5 = __toESM(require_jsx_runtime(), 1);

// src/screens/RecoveryPanel.tsx
var import_jsx_runtime6 = __toESM(require_jsx_runtime(), 1);
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

// src/lib/votumErfahrung.ts
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

// server/kira_prompts_cli.ts
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

react/cjs/react.production.min.js:
  (**
   * @license React
   * react.production.min.js
   *
   * Copyright (c) Facebook, Inc. and its affiliates.
   *
   * This source code is licensed under the MIT license found in the
   * LICENSE file in the root directory of this source tree.
   *)

react/cjs/react.development.js:
  (**
   * @license React
   * react.development.js
   *
   * Copyright (c) Facebook, Inc. and its affiliates.
   *
   * This source code is licensed under the MIT license found in the
   * LICENSE file in the root directory of this source tree.
   *)

react/cjs/react-jsx-runtime.production.min.js:
  (**
   * @license React
   * react-jsx-runtime.production.min.js
   *
   * Copyright (c) Facebook, Inc. and its affiliates.
   *
   * This source code is licensed under the MIT license found in the
   * LICENSE file in the root directory of this source tree.
   *)

react/cjs/react-jsx-runtime.development.js:
  (**
   * @license React
   * react-jsx-runtime.development.js
   *
   * Copyright (c) Facebook, Inc. and its affiliates.
   *
   * This source code is licensed under the MIT license found in the
   * LICENSE file in the root directory of this source tree.
   *)
*/

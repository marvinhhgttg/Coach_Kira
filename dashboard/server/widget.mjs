var __create = Object.create;
var __defProp = Object.defineProperty;
var __getOwnPropDesc = Object.getOwnPropertyDescriptor;
var __getOwnPropNames = Object.getOwnPropertyNames;
var __getProtoOf = Object.getPrototypeOf;
var __hasOwnProp = Object.prototype.hasOwnProperty;
var __commonJS = (cb2, mod) => function __require() {
  return mod || (0, cb2[__getOwnPropNames(cb2)[0]])((mod = { exports: {} }).exports, mod), mod.exports;
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

// ../node_modules/react/cjs/react.production.min.js
var require_react_production_min = __commonJS({
  "../node_modules/react/cjs/react.production.min.js"(exports) {
    "use strict";
    var l = Symbol.for("react.element");
    var n = Symbol.for("react.portal");
    var p = Symbol.for("react.fragment");
    var q = Symbol.for("react.strict_mode");
    var r = Symbol.for("react.profiler");
    var t = Symbol.for("react.provider");
    var u2 = Symbol.for("react.context");
    var v = Symbol.for("react.forward_ref");
    var w = Symbol.for("react.suspense");
    var x = Symbol.for("react.memo");
    var y = Symbol.for("react.lazy");
    var z2 = Symbol.iterator;
    function A(a) {
      if (null === a || "object" !== typeof a) return null;
      a = z2 && a[z2] || a["@@iterator"];
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
    function G2(a, b, e) {
      this.props = a;
      this.context = b;
      this.refs = D;
      this.updater = e || B;
    }
    var H = G2.prototype = new F();
    H.constructor = G2;
    C(H, E.prototype);
    H.isPureReactComponent = true;
    var I = Array.isArray;
    var J = Object.prototype.hasOwnProperty;
    var K = { current: null };
    var L = { key: true, ref: true, __self: true, __source: true };
    function M(a, b, e) {
      var d, c = {}, k = null, h2 = null;
      if (null != b) for (d in void 0 !== b.ref && (h2 = b.ref), void 0 !== b.key && (k = "" + b.key), b) J.call(b, d) && !L.hasOwnProperty(d) && (c[d] = b[d]);
      var g = arguments.length - 2;
      if (1 === g) c.children = e;
      else if (1 < g) {
        for (var f3 = Array(g), m = 0; m < g; m++) f3[m] = arguments[m + 2];
        c.children = f3;
      }
      if (a && a.defaultProps) for (d in g = a.defaultProps, g) void 0 === c[d] && (c[d] = g[d]);
      return { $$typeof: l, type: a, key: k, ref: h2, props: c, _owner: K.current };
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
      var h2 = false;
      if (null === a) h2 = true;
      else switch (k) {
        case "string":
        case "number":
          h2 = true;
          break;
        case "object":
          switch (a.$$typeof) {
            case l:
            case n:
              h2 = true;
          }
      }
      if (h2) return h2 = a, c = c(h2), a = "" === d ? "." + Q(h2, 0) : d, I(c) ? (e = "", null != a && (e = a.replace(P, "$&/") + "/"), R(c, b, e, "", function(a2) {
        return a2;
      })) : null != c && (O(c) && (c = N(c, e + (!c.key || h2 && h2.key === c.key ? "" : ("" + c.key).replace(P, "$&/") + "/") + a)), b.push(c)), 1;
      h2 = 0;
      d = "" === d ? "." : d + ":";
      if (I(a)) for (var g = 0; g < a.length; g++) {
        k = a[g];
        var f3 = d + Q(k, g);
        h2 += R(k, b, e, f3, c);
      }
      else if (f3 = A(a), "function" === typeof f3) for (a = f3.call(a), g = 0; !(k = a.next()).done; ) k = k.value, f3 = d + Q(k, g++), h2 += R(k, b, e, f3, c);
      else if ("object" === k) throw b = String(a), Error("Objects are not valid as a React child (found: " + ("[object Object]" === b ? "object with keys {" + Object.keys(a).join(", ") + "}" : b) + "). If you meant to render a collection of children, use an array instead.");
      return h2;
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
    exports.PureComponent = G2;
    exports.StrictMode = q;
    exports.Suspense = w;
    exports.__SECRET_INTERNALS_DO_NOT_USE_OR_YOU_WILL_BE_FIRED = W;
    exports.act = X;
    exports.cloneElement = function(a, b, e) {
      if (null === a || void 0 === a) throw Error("React.cloneElement(...): The argument must be a React element, but you passed " + a + ".");
      var d = C({}, a.props), c = a.key, k = a.ref, h2 = a._owner;
      if (null != b) {
        void 0 !== b.ref && (k = b.ref, h2 = K.current);
        void 0 !== b.key && (c = "" + b.key);
        if (a.type && a.type.defaultProps) var g = a.type.defaultProps;
        for (f3 in b) J.call(b, f3) && !L.hasOwnProperty(f3) && (d[f3] = void 0 === b[f3] && void 0 !== g ? g[f3] : b[f3]);
      }
      var f3 = arguments.length - 2;
      if (1 === f3) d.children = e;
      else if (1 < f3) {
        g = Array(f3);
        for (var m = 0; m < f3; m++) g[m] = arguments[m + 2];
        d.children = g;
      }
      return { $$typeof: l, type: a.type, key: c, ref: k, props: d, _owner: h2 };
    };
    exports.createContext = function(a) {
      a = { $$typeof: u2, _currentValue: a, _currentValue2: a, _threadCount: 0, Provider: null, Consumer: null, _defaultValue: null, _globalName: null };
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

// ../node_modules/react/cjs/react.development.js
var require_react_development = __commonJS({
  "../node_modules/react/cjs/react.development.js"(exports, module) {
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

// ../node_modules/react/index.js
var require_react = __commonJS({
  "../node_modules/react/index.js"(exports, module) {
    "use strict";
    if (process.env.NODE_ENV === "production") {
      module.exports = require_react_production_min();
    } else {
      module.exports = require_react_development();
    }
  }
});

// ../node_modules/react/cjs/react-jsx-runtime.production.min.js
var require_react_jsx_runtime_production_min = __commonJS({
  "../node_modules/react/cjs/react-jsx-runtime.production.min.js"(exports) {
    "use strict";
    var f3 = require_react();
    var k = Symbol.for("react.element");
    var l = Symbol.for("react.fragment");
    var m = Object.prototype.hasOwnProperty;
    var n = f3.__SECRET_INTERNALS_DO_NOT_USE_OR_YOU_WILL_BE_FIRED.ReactCurrentOwner;
    var p = { key: true, ref: true, __self: true, __source: true };
    function q(c, a, g) {
      var b, d = {}, e = null, h2 = null;
      void 0 !== g && (e = "" + g);
      void 0 !== a.key && (e = "" + a.key);
      void 0 !== a.ref && (h2 = a.ref);
      for (b in a) m.call(a, b) && !p.hasOwnProperty(b) && (d[b] = a[b]);
      if (c && c.defaultProps) for (b in a = c.defaultProps, a) void 0 === d[b] && (d[b] = a[b]);
      return { $$typeof: k, type: c, key: e, ref: h2, props: d, _owner: n.current };
    }
    exports.Fragment = l;
    exports.jsx = q;
    exports.jsxs = q;
  }
});

// ../node_modules/react/cjs/react-jsx-runtime.development.js
var require_react_jsx_runtime_development = __commonJS({
  "../node_modules/react/cjs/react-jsx-runtime.development.js"(exports) {
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

// ../node_modules/react/jsx-runtime.js
var require_jsx_runtime = __commonJS({
  "../node_modules/react/jsx-runtime.js"(exports, module) {
    "use strict";
    if (process.env.NODE_ENV === "production") {
      module.exports = require_react_jsx_runtime_production_min();
    } else {
      module.exports = require_react_jsx_runtime_development();
    }
  }
});

// ../src/lib/engine.ts
var import_react8 = __toESM(require_react(), 1);

// ../src/lib/garmin.ts
var import_react = __toESM(require_react(), 1);

// ../src/lib/extras.ts
async function getJson(url) {
  const res = await fetch(url, { redirect: "follow", cache: "no-store" });
  const t = await res.text();
  try {
    return JSON.parse(t);
  } catch {
    throw new Error("Die Datenquelle antwortet gerade nicht wie erwartet.");
  }
}
async function fetchExtras(typ, n = 30) {
  const t = Array.isArray(typ) ? typ.join(",") : typ;
  const j = await getJson(`${BASE_URL}?mode=extras&typ=${encodeURIComponent(t)}&n=${n}&cb=${Date.now()}`);
  if (!j || !j.ok) throw new Error(j && j.error || "Extras nicht lesbar");
  return j.items || [];
}

// ../src/lib/garmin.ts
var cache = null;
var CACHE_MS = 10 * 60 * 1e3;
function fetchGarminTage() {
  if (cache && Date.now() - cache.at < CACHE_MS) return cache.p;
  const p = fetchExtras("garmin_tag", 1e3).then((items) => {
    const m = /* @__PURE__ */ new Map();
    for (const it of items) if (it.data && typeof it.data === "object") m.set(it.datum, it.data);
    return m;
  }).catch(() => /* @__PURE__ */ new Map());
  cache = { at: Date.now(), p };
  return p;
}

// ../src/kern/formeln.ts
var ATL_MODEL = { a: -8.819849, b: 0.83645661, c: 1.24059231 };
var CTL_MODEL = { a: -3.95042, b: 0.97283824, c: 0.2336734, g: 0.7, k: 7.1, win: 28 };
function atlStep(atl, last) {
  return ATL_MODEL.a + ATL_MODEL.b * atl + ATL_MODEL.c * last;
}
function ctlStep(ctl, last, mittel28) {
  const m = CTL_MODEL;
  return m.g * (m.a + m.b * ctl + m.c * last) + (1 - m.g) * m.k * mittel28;
}
function windowMean(reihe, win = CTL_MODEL.win) {
  const w = reihe.slice(-win);
  return w.length ? w.reduce((a, b) => a + b, 0) / w.length : 0;
}

// ../src/screens/RecoveryPanel.tsx
var import_react7 = __toESM(require_react(), 1);

// ../src/components/ZeitraumWahl.tsx
var import_react2 = __toESM(require_react(), 1);

// ../src/lib/windows.ts
var ALL_RANGE = "9999d";
var isCustom = (w) => typeof w === "object" && w !== null;
var heute = () => (/* @__PURE__ */ new Date()).toLocaleDateString("sv-SE");
var tageZwischen = (a, b) => Math.round((Date.parse(b + "T12:00:00Z") - Date.parse(a + "T12:00:00Z")) / 864e5);
var STEPS = [
  [14, "14d"],
  [30, "30d"],
  [60, "60d"],
  [90, "90d"],
  [120, "120d"],
  [180, "180d"],
  [240, "240d"],
  [360, "360d"]
];
function rangeFor(w, lead = 0, min = 0) {
  if (!isCustom(w) && w === 0) return ALL_RANGE;
  const tage = isCustom(w) ? tageZwischen(w.von, heute()) + 1 : w;
  const n = Math.max(tage + lead, min);
  for (const [d, r] of STEPS) if (n <= d) return r;
  return ALL_RANGE;
}

// ../src/components/ZeitraumWahl.tsx
var import_jsx_runtime = __toESM(require_jsx_runtime(), 1);

// ../src/components/BefindenHinweis.tsx
var import_react3 = __toESM(require_react(), 1);
var import_jsx_runtime2 = __toESM(require_jsx_runtime(), 1);

// ../src/lib/useErfahrung.ts
var import_react4 = __toESM(require_react(), 1);

// ../src/kern/soll.ts
var BAND = { REST: [0, 0.25], EASY: [0.3, 0.8], TRAIN: [0.7, 1.25], QUALITY: [1, 1.7] };
var BAND_MIN_N = 8;
var BAND_FENSTER = 0.4;
function halteLast(yesterday) {
  const ctl0 = yesterday?.ctlPost ?? null;
  if (ctl0 == null || ctl0 <= 0) return null;
  const m0 = ctl0 / 7.04;
  for (let L = 0; L <= 500; L += 1) if (ctlStep(ctl0, L, (27 * m0 + L) / 28) >= ctl0) return L;
  return null;
}
function persoenlicheBaender(days) {
  const pts = [];
  for (let i = 1; i + 1 < days.length; i++) {
    const d = days[i], n = days[i + 1];
    if (!d.vote || d.readiness == null || n.readiness == null || !d.activityDone) continue;
    const ist = d.actual?.load ?? d.load ?? 0;
    const h2 = halteLast(days[i - 1]);
    if (!h2 || ist <= 0) continue;
    pts.push({ v: d.vote, q: ist / h2, d1: n.readiness - d.readiness });
  }
  const out = {};
  for (const v of ["EASY", "TRAIN"]) {
    const xs = pts.filter((p) => p.v === v);
    let best = null;
    for (let edge = 0.8; edge <= 2.4 + 1e-9; edge += 0.2) {
      const g = xs.filter((p) => p.q > edge - BAND_FENSTER && p.q <= edge);
      if (g.length < BAND_MIN_N) continue;
      const d1 = g.reduce((a, p) => a + p.d1, 0) / g.length;
      if (d1 < 0) break;
      best = { faktor: Math.round(edge * 10) / 10, n: g.length, d1: Math.round(d1 * 10) / 10 };
    }
    if (best) out[v] = best;
  }
  return out;
}
function lastEmpfehlung(today, yesterday, baender) {
  const vote = today.vote;
  const atl0 = yesterday?.atlPost ?? null;
  const ctl0 = yesterday?.ctlPost ?? null;
  if (!vote || atl0 == null || ctl0 == null || ctl0 <= 0) return null;
  const m0 = ctl0 / 7.04;
  const ctlNach = (L) => ctlStep(ctl0, L, (27 * m0 + L) / 28);
  const acwrNach = (L) => atlStep(atl0, L) / ctlNach(L);
  let halte = 0;
  for (let L = 0; L <= 500; L += 1) if (ctlNach(L) >= ctl0) {
    halte = L;
    break;
  }
  let acwrGrenze = null;
  for (let L = 0; L <= 600; L += 1) if (acwrNach(L) > 1.3) {
    acwrGrenze = L - 1;
    break;
  }
  const r5 = (x) => Math.max(0, Math.round(x / 5) * 5);
  let [lo, hi] = BAND[vote].map((f3) => r5(f3 * halte));
  const pb = baender?.[vote];
  const persoenlich = pb && pb.faktor > BAND[vote][1] ? { faktor: pb.faktor, standard: hi, n: pb.n, d1: pb.d1 } : null;
  if (persoenlich) hi = r5(pb.faktor * halte);
  let gedeckelt = false;
  if (acwrGrenze != null && hi > acwrGrenze) {
    hi = r5(Math.max(0, acwrGrenze));
    gedeckelt = true;
    if (lo > hi) lo = r5(hi * 0.6);
  }
  return { min: lo, max: hi, halte: r5(halte), acwrGrenze: acwrGrenze != null ? r5(acwrGrenze) : null, gedeckelt, persoenlich };
}

// ../src/lib/cloudLLM.ts
var SENTINEL = "__PORT_5000__";
var SERVER_BASE = SENTINEL.startsWith("__PORT") ? "" : "/" + SENTINEL.replace(/^\/+/, "");

// ../src/lib/format.ts
function fmtNum(v, digits = 0) {
  if (v === null || v === void 0 || !Number.isFinite(v)) return "\u2014";
  return v.toLocaleString("de-DE", {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits
  });
}

// ../src/components/UI.tsx
var import_jsx_runtime3 = __toESM(require_jsx_runtime(), 1);

// ../src/screens/CorrelationPanel.tsx
var import_react5 = __toESM(require_react(), 1);
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

// ../src/lib/trFarbe.ts
var TR_BAENDER = [
  { von: 0, bis: 25, farbe: "#D9433A", name: "Schlecht" },
  { von: 25, bis: 50, farbe: "#E5822F", name: "Niedrig" },
  { von: 50, bis: 75, farbe: "#5DB75B", name: "M\xE4\xDFig" },
  { von: 75, bis: 95, farbe: "#3A78CF", name: "Hoch" },
  { von: 95, bis: 101, farbe: "#6E45E9", name: "Optimal" }
];

// ../src/lib/garminSkala.ts
var G = { rot: "#D9433A", orange: "#E5822F", gruen: "#5DB75B", blau: "#3A78CF", lila: "#6E45E9", grau: "#9aa5b1" };
var TR_SKALA = { min: 0, max: 100, baender: TR_BAENDER.map((b) => ({ ...b })) };
var BB_SKALA = {
  min: 0,
  max: 100,
  baender: [
    { von: 0, bis: 26, farbe: G.rot, name: "Sehr niedrig" },
    { von: 26, bis: 51, farbe: G.orange, name: "Niedrig" },
    { von: 51, bis: 76, farbe: G.gruen, name: "Mittel" },
    { von: 76, bis: 101, farbe: G.blau, name: "Hoch" }
  ]
};
var SCHLAF_SCORE_SKALA = {
  min: 0,
  max: 100,
  baender: [
    { von: 0, bis: 60, farbe: G.rot, name: "Schlecht" },
    { von: 60, bis: 80, farbe: G.orange, name: "Ordentlich" },
    { von: 80, bis: 90, farbe: G.gruen, name: "Gut" },
    { von: 90, bis: 101, farbe: G.blau, name: "Ausgezeichnet" }
  ]
};
var SCORE_SKALA = {
  min: 0,
  max: 100,
  baender: [
    { von: 0, bis: 30, farbe: G.rot, name: "Schlecht" },
    { von: 30, bis: 50, farbe: G.orange, name: "Niedrig" },
    { von: 50, bis: 70, farbe: G.gruen, name: "M\xE4\xDFig" },
    { von: 70, bis: 101, farbe: G.blau, name: "Gut" }
  ]
};
var RHR_DELTA_SKALA = {
  min: -8,
  max: 8,
  baender: [
    { von: -99, bis: -5, farbe: G.lila, name: "Auff\xE4llig niedrig" },
    { von: -5, bis: 1, farbe: G.blau, name: "Normal" },
    { von: 1, bis: 3, farbe: G.gruen, name: "Leicht erh\xF6ht" },
    { von: 3, bis: 5, farbe: G.orange, name: "Erh\xF6ht" },
    { von: 5, bis: 99, farbe: G.rot, name: "Deutlich erh\xF6ht" }
  ]
};
var BEFINDEN_SKALA = {
  min: 1,
  max: 5,
  baender: [
    { von: 0, bis: 2, farbe: G.rot, name: "Schlecht" },
    { von: 2, bis: 3, farbe: G.orange, name: "M\xE4\xDFig" },
    { von: 3, bis: 4, farbe: G.gruen, name: "Ok" },
    { von: 4, bis: 9, farbe: G.blau, name: "Gut" }
  ]
};
var TOURFORM_SKALA = {
  min: 0,
  max: 100,
  baender: [
    { von: 0, bis: 60, farbe: G.orange, name: "Aufbau" },
    { von: 60, bis: 80, farbe: G.gruen, name: "Auf Kurs" },
    { von: 80, bis: 101, farbe: G.blau, name: "Bereit" }
  ]
};

// ../src/components/GarminSportProfilTile.tsx
var import_react6 = __toESM(require_react(), 1);
var import_jsx_runtime5 = __toESM(require_jsx_runtime(), 1);

// ../src/lib/monotony.ts
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

// ../src/screens/RecoveryPanel.tsx
var import_jsx_runtime6 = __toESM(require_jsx_runtime(), 1);

// ../src/lib/chatContext.ts
var OLLAMA_HTTP = "http://localhost:11434";
var OLLAMA_HTTPS = "https://localhost:11435";
var IST_MAC_APP = typeof window !== "undefined" && !!window.webkit?.messageHandlers?.kiraMac;
var IST_SAFARI = IST_MAC_APP || typeof navigator !== "undefined" && /Safari\//.test(navigator.userAgent) && !/Chrome|Chromium|CriOS|Edg|Comet/i.test(navigator.userAgent);
var KANDIDATEN = IST_SAFARI ? [OLLAMA_HTTPS, OLLAMA_HTTP] : [OLLAMA_HTTP, OLLAMA_HTTPS];
var IST_APP = typeof window !== "undefined" && !!window.webkit?.messageHandlers?.kira;
var UEBER_MAC = "mac";
var ollamaBasis = IST_APP ? UEBER_MAC : KANDIDATEN[0];
var KIRA_SYSTEM = [
  "Du bist Kira, Sportwissenschaftlerin und pers\xF6nliche Ausdauer-Trainerin von Marc (Jahrgang 1975; Rennrad, Gravel, Laufen, Bergsport).",
  "Sprich Marc per Du an, antworte auf Deutsch, klar und konkret, ohne Emojis.",
  "St\xFCtze dich zuerst auf die mitgelieferten Cockpit-Daten (JSON unten). Nenne die Zahlen, auf die du dich beziehst.",
  "Erfinde keine Werte. Wenn Daten fehlen oder unklar sind, sag das offen.",
  'Halte dich an die Regeln aus dem Feld "regeln" (Tagesentscheid, Zielbereiche).',
  "Gib am Ende eine klare Empfehlung (z. B. Einheit, Dauer, Zone oder Pause).",
  "Bei Schmerzen, Krankheit oder medizinischen Fragen: auf Arzt bzw. Physiotherapie verweisen."
].join("\n");

// ../src/lib/engine.ts
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

// ../src/lib/api.ts
var BASE_URL = "https://script.google.com/macros/s/AKfycbxCEN11KRlFaLL7uVJyeBLCrRJVmfBWagSmqvyJ8Ci7nwxi8HbolzTy23Z-G2mivC2h/exec";
function cb() {
  return Date.now().toString();
}
function buildUrl(params) {
  const usp = new URLSearchParams();
  for (const [k, v] of Object.entries(params)) {
    if (v === void 0 || v === null || v === "") continue;
    usp.set(k, String(v));
  }
  usp.set("cb", cb());
  return `${BASE_URL}?${usp.toString()}`;
}
var ApiError = class extends Error {
  constructor(message, status) {
    super(message);
    this.status = status;
    this.name = "ApiError";
  }
};
var TimeoutError = class extends ApiError {
  constructor(message = "Zeit\xFCberschreitung") {
    super(message);
    this.name = "TimeoutError";
  }
};
function withTimeout(ms, factory) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), ms);
  return factory(controller.signal).catch((err) => {
    if (controller.signal.aborted) throw new TimeoutError(`Timeout nach ${ms} ms`);
    throw err;
  }).finally(() => clearTimeout(timer));
}
async function getJson2(url, signal) {
  const isRead = /[?&]mode=/.test(url) && !/[?&]action=/.test(url);
  try {
    return await getJsonOnce(url, signal);
  } catch (err) {
    const retryable = isRead && !signal?.aborted && err instanceof ApiError && (err.status == null || err.status === 404 || err.status >= 500);
    if (!retryable) throw err;
    await new Promise((r) => setTimeout(r, 1500));
    return getJsonOnce(url.replace(/([?&]cb=)\d+/, `$1${Date.now()}`), signal);
  }
}
async function getJsonOnce(url, signal) {
  let res;
  try {
    res = await fetch(url, {
      method: "GET",
      redirect: "follow",
      mode: "cors",
      cache: "no-store",
      signal
    });
  } catch (err) {
    throw new ApiError(`Netzwerkfehler: ${err?.message || "unbekannt"}`);
  }
  if (!res.ok) {
    throw new ApiError(`HTTP ${res.status}`, res.status);
  }
  const text = await res.text();
  try {
    return JSON.parse(text);
  } catch {
    throw new ApiError("Antwort ist kein g\xFCltiges JSON");
  }
}
function zeitreiheUrl() {
  let env;
  try {
    env = typeof process !== "undefined" ? process.env?.CK_ZEITREIHE_URL : void 0;
  } catch {
    env = void 0;
  }
  if (env === "aus") return null;
  if (env) return env;
  if (typeof window !== "undefined") return `${SERVER_BASE}/api/rechenkern/chartData`;
  return "https://coachkira.pplx.app/port/5000/api/rechenkern/chartData";
}
async function fetchChartData(range, metrics, signal) {
  const z2 = zeitreiheUrl();
  if (z2 && !signal?.aborted) {
    const ctl = new AbortController();
    const ab = () => ctl.abort();
    signal?.addEventListener("abort", ab);
    const to = setTimeout(ab, 9e4);
    try {
      const usp = new URLSearchParams({ range: String(range), metrics: metrics.join(","), cb: cb() });
      const r = await fetch(`${z2}?${usp}`, { signal: ctl.signal, cache: "no-store" });
      const j = await r.json();
      if (j && j.ok && Array.isArray(j.data) && j.data.length) return j;
    } catch {
    } finally {
      clearTimeout(to);
      signal?.removeEventListener("abort", ab);
    }
    if (signal?.aborted) throw new ApiError("abgebrochen");
  }
  return getJson2(
    buildUrl({ mode: "chartData", range, metrics: metrics.join(",") }),
    signal
  );
}
function fetchPlanSimulation(signal) {
  return getJson2(buildUrl({ mode: "planSimulationJson" }), signal);
}
function fetchPlanSimulationWithTimeout(ms = 8e3) {
  return withTimeout(ms, (signal) => fetchPlanSimulation(signal));
}
function dedupeByDateKeepLast2(rows) {
  const map = /* @__PURE__ */ new Map();
  rows.forEach((row) => map.set(row.date, row));
  return Array.from(map.values()).sort((a, b) => a.date.localeCompare(b.date));
}
function toNum(v) {
  if (v === null || v === void 0 || v === "") return null;
  const n = typeof v === "number" ? v : Number(v);
  return Number.isFinite(n) ? n : null;
}

// ../src/kern/votumRegeln.ts
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
var VOTE_LEVEL = {
  QUALITY: "blau",
  // Garmin-Spektrum: Blau über Grün
  TRAIN: "gruen",
  EASY: "gelb",
  REST: "rot"
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
function levelSleep(h2) {
  if (h2 == null) return "leer";
  const ratio = h2 / SLEEP_GOAL_H;
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
  const rows = dedupeByDateKeepLast2(
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

// ../src/lib/planning.ts
function zielArt(z2) {
  if (z2.art === "sonstiges" && /marsch|wander|walk|hike|trekking|pilger/i.test(`${z2.name} ${z2.notiz || ""}`)) return "marsch";
  return z2.art;
}
function tourLast(z2, faktoren) {
  if (zielArt(z2) === "marsch") return Math.round(2 * z2.dauer_h + z2.hm / 100 * 1);
  const typ = z2.art === "rad" ? "Bike" : z2.art === "lauf" ? "Run" : z2.art === "skitour" ? "Skitour" : "Other";
  const f3 = (faktoren?.load || []).find((x) => String(x.ActivityType) === typ && /Z2|Default/.test(String(x.Zone)));
  const perH = Number(f3?.Factor) || (typ === "Bike" ? 85 : typ === "Run" ? 120 : 50);
  const ef = (faktoren?.elev || []).find((x) => String(x.ActivityType) === (typ === "Skitour" ? "Other" : typ));
  const per100 = Number(ef?.ElevFactor_per_100m) || (typ === "Run" ? 10 : typ === "Bike" ? 4 : 6);
  const bikeHike = z2.art === "bike_hike" ? 0.5 * 85 * z2.dauer_h * 0.5 : 0;
  const base = z2.art === "bike_hike" ? 50 * z2.dauer_h * 0.5 + bikeHike : perH * z2.dauer_h;
  return Math.round(base + z2.hm / 100 * per100);
}

// ../src/lib/paket4.ts
function toNum2(v) {
  if (v === null || v === void 0 || v === "" || String(v).trim() === "") return null;
  const s = String(v).trim();
  if (s.toLowerCase() === "nan" || s.toLowerCase() === "#ref!") return null;
  const n = Number(s.replace(",", "."));
  return Number.isFinite(n) ? n : null;
}
var CSV_KEYS = ["load_fb_day", "Sport_x", "Zone", "sleep_hours", "Garmin_Training_Readiness", "hrv_status", "rhr_bpm", "fbCTL_obs", "garminEnduranceScore", "distance_km_day", "elev_m_day"];
var p4Cache = null;
var p4Promise = null;
async function ladeP4() {
  if (p4Cache) return p4Cache;
  if (!p4Promise) {
    p4Promise = (async () => {
      try {
        const r = await fetchChartData("9999d", CSV_KEYS);
        const data = r && r.data;
        if (!data || !Array.isArray(data)) {
          p4Cache = [];
          return [];
        }
        const tags = [];
        const seenDates = /* @__PURE__ */ new Map();
        for (const row of data) {
          const date = String(row.date || "").split("T")[0];
          if (date === "" || date > (/* @__PURE__ */ new Date()).toISOString().split("T")[0]) continue;
          if (seenDates.has(date)) continue;
          seenDates.set(date, true);
          const load = toNum2(row.load_fb_day);
          const sportRaw = String(row.Sport_x || "").trim();
          const zoneRaw = String(row.Zone || "").trim();
          const sleep = toNum2(row.sleep_hours);
          const readiness = toNum2(row.Garmin_Training_Readiness);
          const hrv = toNum2(row.hrv_status);
          const rhr = toNum2(row.rhr_bpm);
          const ctl = toNum2(row.fbCTL_obs);
          const es = toNum2(row.garminEnduranceScore);
          const km = toNum2(row.distance_km_day);
          const hm = toNum2(row.elev_m_day);
          tags.push({ date, load, sport: sportRaw, zone: zoneRaw, sleep, readiness, hrv, rhr, ctl, es, km, hm });
        }
        tags.sort((a, b) => a.date.localeCompare(b.date));
        p4Cache = tags;
        return tags;
      } catch {
        p4Cache = [];
        return [];
      } finally {
        p4Promise = null;
      }
    })();
  }
  return p4Promise;
}
function tourform(t, z2, persTourLast) {
  const sorted = [...t].sort((a, b) => a.date.localeCompare(b.date));
  const last8 = sorted.slice(-56);
  let weeksCount = 0;
  if (last8.length >= 7) {
    const dayCount = last8.length;
    weeksCount = Math.max(1, dayCount / 7);
  }
  let wochenLastIst = 0;
  if (weeksCount > 0) {
    let weekSum = 0;
    for (const tag of last8) {
      if (tag.load != null) weekSum += tag.load;
    }
    wochenLastIst = weekSum / weeksCount;
  }
  const tl = persTourLast && persTourLast > 0 ? persTourLast : tourLast(z2, void 0);
  const wochenLastZiel = 2 * tl;
  const volumen = Math.min(1, wochenLastIst / (2 * tl));
  const last10 = sorted.slice(-70);
  const langeZielWert = z2.prioritaet === "A" ? 6 : 4;
  let langeTage = 0;
  const threshold = 0.6 * tl;
  for (const tag of last10) {
    if (tag.load != null && tag.load >= threshold) {
      langeTage++;
    }
  }
  const lange = Math.min(1, langeTage / langeZielWert);
  let maxHm = null;
  for (const tag of last10) {
    if (tag.hm != null && tag.hm > 0) {
      if (maxHm == null || tag.hm > maxHm) {
        maxHm = tag.hm;
      }
    }
  }
  const hmZiel = 0.7 * z2.hm;
  let hoehe = null;
  if (maxHm != null) {
    hoehe = hmZiel > 0 ? Math.min(1, maxHm / hmZiel) : null;
  }
  let prozent = 0;
  if (hoehe == null) {
    prozent = Math.round(100 * (0.6 * volumen + 0.4 * lange));
  } else {
    prozent = Math.round(100 * (0.5 * volumen + 0.35 * lange + 0.15 * hoehe));
  }
  const hinweise = [];
  if (hoehe == null) {
    const volScore = 0.6 * volumen;
    const langScore = 0.4 * lange;
    if (volumen < 0.7) {
      hinweise.push(`Wochenlast auf etwa ${Math.round(wochenLastZiel)} ESS steigern (aktuell ca. ${Math.round(wochenLastIst)}).`);
    }
    if (lange < 0.7) {
      const needed = Math.max(0, langeZielWert - langeTage);
      hinweise.push(`Noch ${needed} lange Tage ab ${Math.round(0.6 * tl)} ESS in den n\xE4chsten Wochen.`);
    }
    hinweise.push("H\xF6henmeter werden kaum erfasst \u2013 Wert nicht ber\xFCcksichtigt.");
  } else {
    const volScore = 0.5 * volumen;
    const langScore = 0.35 * lange;
    const hoeheScore = 0.15 * hoehe;
    const comps = [
      { name: "Volumen", score: volScore, value: volumen },
      { name: "Lange Tage", score: langScore, value: lange },
      { name: "H\xF6he", score: hoeheScore, value: hoehe }
    ];
    comps.sort((a, b) => a.score - b.score);
    const weakest = comps[0];
    if (weakest.name === "Volumen") {
      hinweise.push(`Wochenlast auf etwa ${Math.round(wochenLastZiel)} ESS steigern (aktuell ca. ${Math.round(wochenLastIst)}).`);
    } else if (weakest.name === "Lange Tage") {
      const needed = Math.max(0, langeZielWert - langeTage);
      hinweise.push(`Noch ${needed} lange Tage ab ${Math.round(0.6 * tl)} ESS in den n\xE4chsten Wochen.`);
    } else if (weakest.name === "H\xF6he") {
      hinweise.push(`H\xF6henmeter: zuletzt h\xF6chstens ${maxHm} Hm an einem Tag \u2013 Ziel etwa ${Math.round(hmZiel)}.`);
    }
  }
  return {
    prozent,
    volumen,
    lange,
    hoehe,
    tourLast: tl,
    wochenLastIst,
    wochenLastZiel,
    langeTage,
    langeZiel: langeZielWert,
    maxHm,
    hmZiel,
    hinweise
  };
}

// ../src/lib/essRechner.ts
var PROFILE = [
  { id: "run", label: "Laufen", sport: "Run", typen: ["running", "treadmill_running", "trail_running"], tempo: "pace" },
  { id: "bike_road", label: "Rennrad", sport: "Bike", typen: ["road_biking", "cycling"], tempo: "kmh" },
  { id: "bike_gravel", label: "Gravel/MTB", sport: "Bike", typen: ["gravel_cycling", "mountain_biking"], tempo: "kmh" },
  { id: "bike_indoor", label: "Rad indoor", sport: "Bike", typen: ["indoor_cycling", "virtual_ride"], tempo: "none" },
  { id: "hike", label: "Wandern/Bergtour", sport: "Hike", typen: ["hiking", "walking", "mountaineering"], tempo: "kmh" },
  { id: "row", label: "Rudern indoor", sport: "Row", typen: ["indoor_rowing", "rowing_v2"], tempo: "row" }
];
var LN2 = Math.log(2);
function quantil(xs, q) {
  if (!xs.length) return NaN;
  const s = [...xs].sort((a, b) => a - b);
  const i = (s.length - 1) * q;
  const lo = Math.floor(i);
  return s[lo] + (s[Math.min(s.length - 1, lo + 1)] - s[lo]) * (i - lo);
}
function wls(X, y, w, ridge = 1e-6) {
  const k = X[0].length + 1;
  const A = Array.from({ length: k }, () => new Array(k).fill(0));
  const b = new Array(k).fill(0);
  for (let i = 0; i < y.length; i++) {
    const r = [1, ...X[i]];
    for (let p = 0; p < k; p++) {
      b[p] += w[i] * r[p] * y[i];
      for (let q = 0; q < k; q++) A[p][q] += w[i] * r[p] * r[q];
    }
  }
  for (let p = 1; p < k; p++) A[p][p] += ridge;
  for (let c = 0; c < k; c++) {
    let piv = c;
    for (let r = c + 1; r < k; r++) if (Math.abs(A[r][c]) > Math.abs(A[piv][c])) piv = r;
    if (Math.abs(A[piv][c]) < 1e-12) return null;
    [A[c], A[piv]] = [A[piv], A[c]];
    [b[c], b[piv]] = [b[piv], b[c]];
    for (let r = 0; r < k; r++) {
      if (r === c) continue;
      const f3 = A[r][c] / A[c][c];
      for (let q = c; q < k; q++) A[r][q] -= f3 * A[c][q];
      b[r] -= f3 * b[c];
    }
  }
  return b.map((v, i) => v / A[i][i]);
}
var dot = (beta, x) => beta[0] + x.reduce((s, v, i) => s + v * beta[i + 1], 0);
function plausibel(id, a) {
  const v = a.v_kmh || 0;
  if (id === "run") return v >= 5 && v <= 20;
  if (id === "bike_road" || id === "bike_gravel") return v >= 8 && v <= 45;
  if (id === "hike") return v > 0.8 && v <= 8 && (a.km || 0) < 60;
  if (id === "row") return v >= 6 && v <= 25;
  return true;
}
function lerneModelle(tage, heute2 = /* @__PURE__ */ new Date()) {
  const t0 = heute2.getTime();
  const out = [];
  for (const p of PROFILE) {
    const proben = [];
    tage.forEach((tag, datum) => {
      for (const a of tag.akt || []) {
        if (!a.typ || !p.typen.includes(a.typ)) continue;
        const min = (a.dauer_s || 0) / 60;
        if (min < 10 || !a.last || a.last < 3 || !a.hf_avg || a.hf_avg < 60) continue;
        if (!plausibel(p.id, a)) continue;
        const alter = (t0 - (/* @__PURE__ */ new Date(datum + "T12:00:00")).getTime()) / 864e5;
        proben.push({
          datum,
          name: a.name,
          min,
          km: a.km || 0,
          hm: a.hm || 0,
          hf: a.hf_avg,
          np: a.watt_np || void 0,
          v: a.v_kmh || 0,
          ess: a.last,
          w: Math.exp(-LN2 * Math.max(0, alter) / 365)
        });
      }
    });
    if (proben.length < 8) continue;
    const y = proben.map((s) => Math.log(s.ess));
    const w = proben.map((s) => s.w);
    let base;
    if (p.id === "hike") {
      const hb = wls(proben.map((s) => [s.hf]), proben.map((s) => Math.log(s.ess) - Math.log(s.min)), proben.map((s) => s.w * Math.sqrt(s.ess)));
      base = hb ? [hb[0], 1, hb[1]] : null;
    } else base = wls(proben.map((s) => [Math.log(s.min), s.hf]), y, w);
    if (!base || base[2] <= 0) continue;
    const relErr = (pred, ps) => ps.reduce((sum, s) => sum + Math.abs(pred(s) - s.ess) / s.ess, 0) / ps.length;
    const m = {
      id: p.id,
      n: proben.length,
      base,
      fehler: relErr((s) => Math.exp(dot(base, [Math.log(s.min), s.hf])), p.id === "hike" && proben.filter((s) => s.ess >= 30).length >= 5 ? proben.filter((s) => s.ess >= 30) : proben),
      hf: { p5: quantil(proben.map((s) => s.hf), 0.05), p25: quantil(proben.map((s) => s.hf), 0.25), p50: quantil(proben.map((s) => s.hf), 0.5), p75: quantil(proben.map((s) => s.hf), 0.75), p95: quantil(proben.map((s) => s.hf), 0.95) },
      min: { p5: quantil(proben.map((s) => s.min), 0.05), p50: quantil(proben.map((s) => s.min), 0.5), p95: quantil(proben.map((s) => s.min), 0.95) },
      hpk: { p50: quantil(proben.filter((s) => s.km > 0.5).map((s) => s.hm / s.km), 0.5) || 0, p90: quantil(proben.filter((s) => s.km > 0.5).map((s) => s.hm / s.km), 0.9) || 0 },
      proben
    };
    if (p.sport === "Bike") {
      const pn = proben.filter((s) => s.np && s.np > 50);
      if (pn.length >= 8) {
        const ext = wls(pn.map((s) => [Math.log(s.min), s.hf, s.np]), pn.map((s) => Math.log(s.ess)), pn.map((s) => s.w));
        const npLink = wls(pn.map((s) => [s.hf]), pn.map((s) => s.np), pn.map((s) => s.w));
        if (ext) {
          m.ext = ext;
          m.extKey = "np";
          m.fehlerExt = relErr((s) => Math.exp(dot(ext, [Math.log(s.min), s.hf, s.np])), pn);
        }
        if (npLink && npLink[1] > 0) m.npLink = npLink;
      }
    }
    if (p.id === "run") {
      const ext = wls(proben.map((s) => [Math.log(s.min), s.hf, s.v]), y, w);
      if (ext) {
        m.ext = ext;
        m.extKey = "v";
        m.fehlerExt = relErr((s) => Math.exp(dot(ext, [Math.log(s.min), s.hf, s.v])), proben);
      }
    }
    if (p.id === "run" || p.id === "bike_road" || p.id === "bike_gravel" || p.id === "row") {
      const pv = proben.filter((s) => s.km > 0.5);
      const vl = wls(pv.map((s) => [s.hf, s.hm / s.km]), pv.map((s) => s.v), pv.map((s) => s.w));
      if (vl && vl[1] > 0) m.vLink = vl;
    }
    if (m.vLink) {
      const pv = proben.filter((s) => s.km > 0.5);
      const o = wls(pv.map((s) => [Math.log(s.min), s.v, s.hm / s.km]), pv.map((s) => Math.log(s.ess)), pv.map((s) => s.w));
      if (o) {
        m.ohneHfV = o;
        m.fehlerOhneHf = relErr((s) => Math.exp(dot(o, [Math.log(s.min), s.v, s.km > 0.5 ? s.hm / s.km : 0])), pv);
      }
    }
    if (p.sport === "Bike") {
      const pn = proben.filter((s) => s.np && s.np > 50);
      if (pn.length >= 8) {
        const o = wls(pn.map((s) => [Math.log(s.min), s.np]), pn.map((s) => Math.log(s.ess)), pn.map((s) => s.w));
        if (o) m.ohneHfNp = o;
      }
    }
    if (p.id === "hike") {
      const hl = wls(proben.map((s) => [s.hm / (s.min / 60), s.v]), proben.map((s) => s.hf), w);
      if (hl) m.hfLink = hl;
      const ph = proben.filter((s) => s.km > 1);
      let sxx = 0, sxy = 0, syy = 0, sx = 0, sy = 0;
      for (const s of ph) {
        const h2 = s.min / 60;
        sxx += s.w * s.km * s.km;
        sxy += s.w * s.km * s.hm;
        syy += s.w * s.hm * s.hm;
        sx += s.w * s.km * h2;
        sy += s.w * s.hm * h2;
      }
      const det = sxx * syy - sxy * sxy;
      let x = det ? (sx * syy - sy * sxy) / det : 0.25;
      let yv = det ? (sxx * sy - sxy * sx) / det : 1 / 400;
      if (!(x > 0.08 && x < 0.5)) x = 0.22;
      if (!(yv > 1 / 1200 && yv < 1 / 150)) yv = 1 / 400;
      m.hike = { vFlach: 1 / x, hmProH: 1 / yv };
    }
    out.push(m);
  }
  return out;
}
var essBase = (m, min, hf) => Math.exp(dot(m.base, [Math.log(min), hf]));
var hfAusEss = (m, ess, min) => (Math.log(ess) - m.base[0] - m.base[1] * Math.log(min)) / m.base[2];
var vAusHf = (m, hf, hpk) => m.vLink ? dot(m.vLink, [hf, hpk]) : void 0;
var npAusHf = (m, hf) => m.npLink ? dot(m.npLink, [hf]) : void 0;
var hfAusNp = (m, np) => m.npLink ? (np - m.npLink[0]) / m.npLink[1] : void 0;
var hikeMin = (m, km, hm) => m.hike ? (km / m.hike.vFlach + hm / m.hike.hmProH) * 60 : void 0;
function pruefe(m, w) {
  const h2 = [];
  if (w.hf > m.hf.p95 + 3) h2.push(`\xD8 HF ${Math.round(w.hf)} liegt \xFCber deinem \xFCblichen Bereich (bis ${Math.round(m.hf.p95)}) \u2013 Sch\xE4tzung unsicher.`);
  if (w.hf < m.hf.p5 - 3) h2.push(`\xD8 HF ${Math.round(w.hf)} liegt unter deinem \xFCblichen Bereich (ab ${Math.round(m.hf.p5)}).`);
  if (w.min > m.min.p95 * 1.3) h2.push(`So lange Einheiten (${Math.round(w.min)} min) gibt es in deinen Daten kaum \u2013 Sch\xE4tzung unsicher.`);
  if (w.min < m.min.p5 * 0.7) h2.push(`Sehr kurze Einheit \u2013 Sch\xE4tzung unsicher.`);
  return h2;
}
function vorwaerts(m, ein) {
  const g = [];
  const w = { ...ein };
  if (w.km && w.hm == null) {
    w.hm = Math.round(w.km * m.hpk.p50);
    g.push("hm");
  }
  const hpk = w.km && w.km > 0.5 ? (w.hm || 0) / w.km : m.hpk.p50;
  if (!w.v && w.km && w.min) w.v = w.km / (w.min / 60);
  if (!w.min && m.id === "hike" && w.km) {
    w.min = hikeMin(m, w.km, w.hm || 0);
    g.push("min");
  }
  const hfGegeben = !!w.hf;
  let essOhneHf;
  let fehlerOhneHf;
  if (!w.hf) {
    let hf;
    if (w.min && w.np && m.ohneHfNp) {
      essOhneHf = Math.exp(dot(m.ohneHfNp, [Math.log(w.min), w.np]));
      fehlerOhneHf = (m.fehlerExt ?? m.fehler) * 1.4;
    } else if (w.min && w.v && m.ohneHfV) {
      essOhneHf = Math.exp(dot(m.ohneHfV, [Math.log(w.min), w.v, hpk]));
      fehlerOhneHf = m.fehlerOhneHf ?? m.fehler;
    }
    if (essOhneHf && w.min) hf = hfAusEss(m, essOhneHf, w.min);
    else if (m.id === "hike" && m.hfLink && w.min && w.km) hf = dot(m.hfLink, [(w.hm || 0) / (w.min / 60), w.km / (w.min / 60)]);
    else if (w.np) hf = hfAusNp(m, w.np);
    w.hf = hf != null && isFinite(hf) ? hf : m.hf.p50;
    g.push("hf");
  }
  if (!w.v && m.id !== "bike_indoor") {
    const v = m.id === "hike" ? void 0 : vAusHf(m, w.hf, hpk);
    if (v) {
      w.v = v;
      g.push("v");
    }
  }
  if (!w.min && w.km && w.v) {
    w.min = w.km / w.v * 60;
    g.push("min");
  }
  if (!w.min) return null;
  if (!w.km && w.v) {
    w.km = w.v * w.min / 60;
    g.push("km");
  }
  if (!w.np && m.npLink) {
    w.np = npAusHf(m, w.hf);
    g.push("np");
  }
  let ess = essBase(m, w.min, w.hf);
  let typ = "basis";
  let fehler = hfGegeben ? m.fehler : m.fehler * 1.5;
  if (essOhneHf) {
    ess = essOhneHf;
    typ = ein.np ? "np" : "v";
    fehler = fehlerOhneHf ?? fehler;
  } else if (!hfGegeben) {
  } else if (m.ext && m.extKey === "np" && ein.np) {
    ess = Math.exp(dot(m.ext, [Math.log(w.min), w.hf, ein.np]));
    typ = "np";
    fehler = m.fehlerExt ?? fehler;
  } else if (m.ext && m.extKey === "v" && w.v && !g.includes("v")) {
    ess = Math.exp(dot(m.ext, [Math.log(w.min), w.hf, w.v]));
    typ = "v";
    fehler = m.fehlerExt ?? fehler;
  }
  const hinweise = pruefe(m, w);
  if (!hfGegeben) hinweise.push("Mit deiner \xD8 HF wird die Sch\xE4tzung deutlich genauer \u2013 Garmin rechnet die Last aus der Herzfrequenz.");
  return { ess, lo: ess * (1 - fehler), hi: ess * (1 + fehler), werte: w, geschaetzt: g, modell: typ, hinweise };
}

// ../src/data/skitouren_garmin.json
var skitouren_garmin_default = {
  quelle: "Garmin Connect API (activitylist-service), abgerufen 29.10.2026",
  hinweis: "12/2022\u201312/2023: Garmin-Last etwa doppelt so hoch wie davor und mit der aktuellen Uhr \u2013 nicht vergleichbar",
  touren: [
    {
      datum: "2020-12-29",
      name: "Blaichach Ski-/Snowboardtour",
      min: 228.4,
      km: 13.8,
      hm: 773,
      hf: 132,
      ess: 113,
      vergleichbar: true
    },
    {
      datum: "2020-12-31",
      name: "Nesselwang Ski-/Snowboardtour",
      min: 211.2,
      km: 10.5,
      hm: 888,
      hf: 127,
      ess: 114,
      vergleichbar: true
    },
    {
      datum: "2021-01-09",
      name: "Mittelberg Ski-/Snowboardtour",
      min: 219.7,
      km: 14.5,
      hm: 1393,
      hf: 120,
      ess: 70,
      vergleichbar: true
    },
    {
      datum: "2021-01-20",
      name: "Immenstadt i. Allg\xE4u Ski-/Snowboardtour",
      min: 132.9,
      km: 9.9,
      hm: 723,
      hf: 134,
      ess: 103,
      vergleichbar: true
    },
    {
      datum: "2021-01-24",
      name: "Immenstadt i. Allg\xE4u Ski-/Snowboardtour",
      min: 95.5,
      km: 8.1,
      hm: 718,
      hf: 139,
      ess: 103,
      vergleichbar: true
    },
    {
      datum: "2021-01-26",
      name: "Blaichach Ski-/Snowboardtour",
      min: 178.5,
      km: 11.3,
      hm: 825,
      hf: 139,
      ess: 163,
      vergleichbar: true
    },
    {
      datum: "2021-03-17",
      name: "Immenstadt i. Allg\xE4u Ski-/Snowboardtour",
      min: 193.9,
      km: 12.2,
      hm: 1042,
      hf: 140,
      ess: 152,
      vergleichbar: true
    },
    {
      datum: "2021-03-20",
      name: "Hochgrat Skitour",
      min: 140.2,
      km: 10,
      hm: 984,
      hf: 140,
      ess: 127,
      vergleichbar: true
    },
    {
      datum: "2021-03-26",
      name: "Eineguntkopf Skitour",
      min: 171.5,
      km: 9.6,
      hm: 764,
      hf: 131,
      ess: 111,
      vergleichbar: true
    },
    {
      datum: "2021-11-29",
      name: "Skitour Ofterschwanger Horn",
      min: 139.8,
      km: 7.8,
      hm: 531,
      hf: 144,
      ess: 209,
      vergleichbar: true
    },
    {
      datum: "2021-11-30",
      name: "Skitour Ofterschwanger Horn",
      min: 90.7,
      km: 4.9,
      hm: 518,
      hf: 145,
      ess: 148,
      vergleichbar: true
    },
    {
      datum: "2021-12-03",
      name: "Skitour Rangiswanger Horn von Sigiswang ",
      min: 131.1,
      km: 6.7,
      hm: 724,
      hf: 140,
      ess: 137,
      vergleichbar: true
    },
    {
      datum: "2021-12-12",
      name: "Skitour Immenstadt - Mittag - B\xE4renk\xF6pfle - D\xFCrrehorn und zur\xFCck",
      min: 174.5,
      km: 10.8,
      hm: 1035,
      hf: 143,
      ess: 189,
      vergleichbar: true
    },
    {
      datum: "2021-12-15",
      name: "Skitour nach Hause ",
      min: 164,
      km: 11.9,
      hm: 772,
      hf: 139,
      ess: 166,
      vergleichbar: true
    },
    {
      datum: "2021-12-27",
      name: "Skitour Fellhorn und Kanzelwand",
      min: 195.4,
      km: 14.1,
      hm: 1315,
      hf: 151,
      ess: 276,
      vergleichbar: true
    },
    {
      datum: "2022-01-06",
      name: "Skitour Riedberger Horn und Bolgental",
      min: 120.3,
      km: 7.1,
      hm: 633,
      hf: 137,
      ess: 113,
      vergleichbar: true
    },
    {
      datum: "2022-01-12",
      name: "Skitour Kanzelwand und Fellhorn",
      min: 217.7,
      km: 14.5,
      hm: 1349,
      hf: 136,
      ess: 137,
      vergleichbar: true
    },
    {
      datum: "2022-01-13",
      name: "Skitour Gschwender Horn",
      min: 96.4,
      km: 5.9,
      hm: 618,
      hf: 137,
      ess: 97,
      vergleichbar: true
    },
    {
      datum: "2022-01-14",
      name: "Skitour Wannenkopf",
      min: 152.1,
      km: 9.5,
      hm: 814,
      hf: 133,
      ess: 100,
      vergleichbar: true
    },
    {
      datum: "2022-01-18",
      name: "Skitour G\xFCntlespitze",
      min: 217.3,
      km: 9.8,
      hm: 939,
      hf: 135,
      ess: 147,
      vergleichbar: true
    },
    {
      datum: "2022-01-22",
      name: "Skitour nach Hause",
      min: 153.3,
      km: 11.8,
      hm: 769,
      hf: 130,
      ess: 101,
      vergleichbar: true
    },
    {
      datum: "2022-01-25",
      name: "Skitour nach Hause ",
      min: 145,
      km: 11.2,
      hm: 760,
      hf: 125,
      ess: 69,
      vergleichbar: true
    },
    {
      datum: "2022-01-29",
      name: "Blitzskitour auf den Mittag",
      min: 86.3,
      km: 6.2,
      hm: 658,
      hf: 146,
      ess: 159,
      vergleichbar: true
    },
    {
      datum: "2022-02-04",
      name: "Skitour nach Hause ",
      min: 133.6,
      km: 8.4,
      hm: 719,
      hf: 139,
      ess: 124,
      vergleichbar: true
    },
    {
      datum: "2022-02-09",
      name: "Skitour Feuerst\xE4tterkopf",
      min: 125.9,
      km: 9.4,
      hm: 657,
      hf: 129,
      ess: 82,
      vergleichbar: true
    },
    {
      datum: "2022-02-12",
      name: "Skitour Steinmannl und H\xE4hlekopf",
      min: 232.3,
      km: 16.5,
      hm: 1145,
      hf: 140,
      ess: 215,
      vergleichbar: true
    },
    {
      datum: "2022-02-19",
      name: "Skitour Ochsenkopf und Rangiswanger Horn",
      min: 166.1,
      km: 10.5,
      hm: 880,
      hf: 132,
      ess: 107,
      vergleichbar: true
    },
    {
      datum: "2022-02-23",
      name: "Skitour Mittag - B\xE4renk\xF6pfle - D\xFCrrehorn - B\xE4renk\xF6pfle - Mittag",
      min: 155.8,
      km: 11.5,
      hm: 1066,
      hf: 147,
      ess: 210,
      vergleichbar: true
    },
    {
      datum: "2022-02-24",
      name: "Skitour Riedberger Horn ",
      min: 55.9,
      km: 4.1,
      hm: 329,
      hf: 132,
      ess: 59,
      vergleichbar: true
    },
    {
      datum: "2022-02-26",
      name: "Skitour Seekopf",
      min: 256.9,
      km: 15.2,
      hm: 1509,
      hf: 136,
      ess: 186,
      vergleichbar: true
    },
    {
      datum: "2022-03-04",
      name: "Skitour Hochgrat",
      min: 106.1,
      km: 7.4,
      hm: 842,
      hf: 145,
      ess: 188,
      vergleichbar: true
    },
    {
      datum: "2022-03-05",
      name: "Skitour Ponten",
      min: 140,
      km: 8.6,
      hm: 831,
      hf: 137,
      ess: 117,
      vergleichbar: true
    },
    {
      datum: "2022-03-09",
      name: "Skitour Gr\xFCnhorn",
      min: 177.8,
      km: 10.8,
      hm: 970,
      hf: 134,
      ess: 166,
      vergleichbar: true
    },
    {
      datum: "2022-03-20",
      name: "Skitour Schochen",
      min: 229.6,
      km: 13.4,
      hm: 1162,
      hf: 128,
      ess: 116,
      vergleichbar: true
    },
    {
      datum: "2022-12-17",
      name: "Riedberger Horn ",
      min: 70.7,
      km: 3.8,
      hm: 337,
      hf: 156,
      ess: 271,
      vergleichbar: false
    },
    {
      datum: "2022-12-18",
      name: "Ochsenkopf und Rangiswanger Horn ",
      min: 217.8,
      km: 10.7,
      hm: 926,
      hf: 143,
      ess: 288,
      vergleichbar: false
    },
    {
      datum: "2023-01-28",
      name: "Rangiswanger Horn ",
      min: 112.3,
      km: 6.2,
      hm: 734,
      hf: 153,
      ess: 331,
      vergleichbar: false
    },
    {
      datum: "2023-02-10",
      name: "Riedberger Horn - 3 Aufstiege",
      min: 160.6,
      km: 8.5,
      hm: 826,
      hf: 154,
      ess: 489,
      vergleichbar: false
    },
    {
      datum: "2023-02-26",
      name: "Oberstaufen Skitour",
      min: 152.1,
      km: 9,
      hm: 941,
      hf: 148,
      ess: 365,
      vergleichbar: false
    },
    {
      datum: "2023-04-21",
      name: "Austoben im Skigebiet Fellhorn - Kanzelwand",
      min: 219.3,
      km: 15.3,
      hm: 1496,
      hf: 142,
      ess: 272,
      vergleichbar: false
    },
    {
      datum: "2023-04-26",
      name: "Hahnenk\xF6pfle",
      min: 158.4,
      km: 9.2,
      hm: 1041,
      hf: 148,
      ess: 330,
      vergleichbar: false
    },
    {
      datum: "2023-11-29",
      name: "Ofterschwang Skitouren",
      min: 169.5,
      km: 10.5,
      hm: 954,
      hf: 143,
      ess: 333,
      vergleichbar: false
    },
    {
      datum: "2023-12-30",
      name: "Pistenskitour Muttelkopf",
      min: 134.5,
      km: 8.3,
      hm: 751,
      hf: 134,
      ess: 164,
      vergleichbar: false
    },
    {
      datum: "2025-11-29",
      name: "Skitour - Ochsenkopf und Weiherkopf",
      min: 211.4,
      km: 13.7,
      hm: 1171,
      hf: 136,
      ess: 170,
      vergleichbar: true
    }
  ]
};

// ../src/data/lange_maersche_garmin.json
var lange_maersche_garmin_default = {
  quelle: "Garmin Connect API, abgerufen 29.10.2026",
  auswahl: "Gehen/Wandern \u2265 4 h, h\xF6chstens 40 Hm/km",
  touren: [
    {
      datum: "2020-09-27",
      name: "Immenstadt i. Allg\xE4u Wandern",
      min: 269.2,
      km: 21.23,
      hm: 482,
      hf: 112,
      ess: 56
    },
    {
      datum: "2020-11-29",
      name: "F\xFCssen Wandern",
      min: 305.8,
      km: 20.61,
      hm: 356,
      hf: 89,
      ess: 20
    },
    {
      datum: "2021-02-13",
      name: "Immenstadt i. Allg\xE4u Wandern",
      min: 247.2,
      km: 20.01,
      hm: 472,
      hf: 102,
      ess: 30
    },
    {
      datum: "2021-04-04",
      name: "Immenstadt i. Allg\xE4u Wandern",
      min: 279.1,
      km: 22.39,
      hm: 663,
      hf: 93,
      ess: 33
    },
    {
      datum: "2021-06-19",
      name: "Scharnitz Wandern",
      min: 344.1,
      km: 24.38,
      hm: 700,
      hf: 92,
      ess: 43
    },
    {
      datum: "2022-02-27",
      name: "Siedelalpe mit Tanja",
      min: 247.4,
      km: 17.4,
      hm: 416,
      hf: 96,
      ess: 19
    },
    {
      datum: "2022-06-18",
      name: "Tag 2 - \xDE\xF3rsm\xF6rk - Emstrur",
      min: 312.5,
      km: 18.59,
      hm: 458,
      hf: 102,
      ess: 26
    },
    {
      datum: "2022-06-19",
      name: "Tag 3 - Emstrur - \xC1lftavatn",
      min: 326.7,
      km: 17.9,
      hm: 194,
      hf: 98,
      ess: 24
    },
    {
      datum: "2022-06-20",
      name: "Tag 4 - \xC1lftavatn - Landmannalaugar",
      min: 396.2,
      km: 22.03,
      hm: 661,
      hf: 108,
      ess: 27
    },
    {
      datum: "2022-06-27",
      name: "Reykjavegur ",
      min: 401.4,
      km: 28.05,
      hm: 827,
      hf: 88,
      ess: 10
    },
    {
      datum: "2022-07-16",
      name: "Immenstadt i. Allg\xE4u Wandern",
      min: 316.6,
      km: 19.08,
      hm: 487,
      hf: 99,
      ess: 10
    },
    {
      datum: "2022-08-15",
      name: "Zur Siedelalpe",
      min: 257.7,
      km: 17.44,
      hm: 419,
      hf: 88,
      ess: 5
    },
    {
      datum: "2022-10-28",
      name: "High Peak Wandern",
      min: 248.5,
      km: 16.11,
      hm: 552,
      hf: 105,
      ess: 58
    },
    {
      datum: "2023-03-28",
      name: "Illerschleifen",
      min: 243.4,
      km: 20.47,
      hm: 706,
      hf: 98,
      ess: 16
    },
    {
      datum: "2023-04-10",
      name: "Monte Luppia - Garda",
      min: 269.5,
      km: 18,
      hm: 528,
      hf: 90,
      ess: 7
    },
    {
      datum: "2024-07-20",
      name: "Hasenstabweg - Spessart",
      min: 912.3,
      km: 62.8,
      hm: 1661,
      hf: 96,
      ess: 14
    },
    {
      datum: "2025-09-13",
      name: "Oberg\xFCnzburg Wandern",
      min: 243.2,
      km: 12.37,
      hm: 388,
      hf: 80,
      ess: 3
    },
    {
      datum: "2025-09-27",
      name: "Altm\xFChltal Teil 1",
      min: 351,
      km: 24.45,
      hm: 607,
      hf: 87,
      ess: 7
    },
    {
      datum: "2025-09-28",
      name: "Altm\xFChltal Teil 2",
      min: 344.2,
      km: 25.19,
      hm: 552,
      hf: 85,
      ess: 8
    }
  ]
};

// ../src/lib/tourEss.ts
var MIN_N = 8;
function eigeneSkitouren(tage) {
  const m = /* @__PURE__ */ new Map();
  for (const t of skitouren_garmin_default.touren) m.set(t.datum + "|" + t.min, t);
  const daten = new Set([...m.values()].map((t) => t.datum));
  if (tage) {
    for (const [datum, t] of tage)
      for (const a of t.akt || [])
        if (/backcountry_ski|skitour|ski_touring/i.test(String(a.typ || "")) && a.dauer_s && a.dauer_s >= 2700 && a.hf_avg && (a.last || 0) >= 40 && !daten.has(datum))
          m.set(datum + "|" + a.dauer_s, { datum, name: a.name, min: a.dauer_s / 60, km: a.km, hm: a.hm || 0, hf: a.hf_avg, ess: a.last, vergleichbar: true });
  }
  return [...m.values()].sort((a, b) => a.datum.localeCompare(b.datum));
}
function eigeneBergtouren(tage) {
  const out = [];
  if (!tage) return out;
  for (const [datum, t] of tage)
    for (const a of t.akt || [])
      if (/^(hiking|mountaineering)$/i.test(String(a.typ || "")) && (a.hm || 0) >= 700 && (a.dauer_s || 0) >= 7200 && (a.last || 0) > 0)
        out.push({ datum, name: a.name, min: a.dauer_s / 60, km: a.km, hm: a.hm, hf: a.hf_avg, ess: a.last });
  return out.sort((a, b) => a.datum.localeCompare(b.datum));
}
var tageCache = null;
function setzeGarminTage(t) {
  tageCache = t;
}
function eigeneMaersche(tage) {
  const m = /* @__PURE__ */ new Map();
  for (const t of lange_maersche_garmin_default.touren) m.set(t.datum, t);
  if (tage) {
    for (const [datum, t] of tage)
      for (const a of t.akt || [])
        if (/^(hiking|walking|speed_walking|mountaineering)$/i.test(String(a.typ || "")) && (a.dauer_s || 0) >= 14400 && a.km && (a.hm || 0) / a.km <= 40 && a.last != null && !m.has(datum))
          m.set(datum, { datum, name: a.name, min: a.dauer_s / 60, km: a.km, hm: a.hm || 0, hf: a.hf_avg, ess: a.last });
  }
  return [...m.values()].sort((a, b) => a.datum.localeCompare(b.datum));
}
var median = (x) => {
  const s = [...x].sort((a, b) => a - b);
  return s.length % 2 ? s[(s.length - 1) / 2] : (s[s.length / 2 - 1] + s[s.length / 2]) / 2;
};
var de = (iso2) => `${iso2.slice(8, 10)}.${iso2.slice(5, 7)}.${iso2.slice(2, 4)}`;
var finde = (modelle, ...ids) => {
  for (const id of ids) {
    const m = modelle.find((x) => x.id === id && x.n >= MIN_N);
    if (m) return m;
  }
  return null;
};
function tourLastPersoenlich(z2, modelle, skitouren = [], bergtouren = []) {
  skitouren = [...skitouren];
  if (!z2 || !(z2.dauer_h > 0) || !modelle?.length) return null;
  const hinweise = [];
  const km = z2.distanz_km > 0 ? z2.distanz_km : void 0;
  const hm = Math.max(0, z2.hm || 0);
  const min = z2.dauer_h * 60;
  let plan = [];
  const art = zielArt(z2);
  if (art === "marsch") {
    const ms = eigeneMaersche(tageCache).filter((t) => t.datum >= "2024-01-01");
    if (ms.length >= 3) {
      const ph = ms.map((t) => t.ess / (t.min / 60)).sort((a, b) => a - b);
      const med = median(ph);
      const ess = Math.round(med * z2.dauer_h), lo = Math.round(ph[0] * z2.dauer_h), hi = Math.round(ph[ph.length - 1] * z2.dauer_h);
      const lang = ms.reduce((a, t) => t.min > a.min ? t : a, ms[0]);
      hinweise.push(
        `Marsch aus ${ms.length} eigenen langen Touren seit 2024 (\u2265 4 h, flach): Garmin vergab ${ph[0].toLocaleString("de-DE", { maximumFractionDigits: 1 })}\u2013${ph[ph.length - 1].toLocaleString("de-DE", { maximumFractionDigits: 1 })} ESS je Stunde \u2013 z. B. ${de(lang.datum)}: ${(Math.round(lang.min / 6) / 10).toLocaleString("de-DE")} h, ${Math.round(lang.km || 0)} km, ${lang.hm} Hm \u2192 ${lang.ess} ESS. Garmin misst die Last \xFCber die Intensit\xE4t, nicht \xFCber die Dauer \u2013 f\xFCr einen Marsch z\xE4hlen deshalb Strecke und Zeit am St\xFCck.`
      );
      return { ess, lo, hi, teile: [{ profil: "hike", min: Math.round(min), km, hm: Math.round(hm), ess, lo, hi, n: ms.length }], hinweise };
    }
    const m = finde(modelle, "hike");
    if (!m) return null;
    const e = vorwaerts(m, { min, km, hm });
    if (!e) return null;
    hinweise.push("Marsch \xFCber dein Wandermodell gesch\xE4tzt \u2013 noch zu wenige eigene lange M\xE4rsche f\xFCr einen direkten Abgleich.");
    return { ess: Math.round(e.ess), lo: Math.round(e.lo), hi: Math.round(e.hi), teile: [{ profil: "hike", min: Math.round(min), km, hm: Math.round(hm), ess: Math.round(e.ess), lo: Math.round(e.lo), hi: Math.round(e.hi), n: m.n }], hinweise };
  }
  if (art === "lauf") {
    const m = finde(modelle, "run");
    if (!m) return null;
    plan = [{ m, min, km, hm }];
  } else if (art === "rad") {
    const m = finde(modelle, "bike_road", "bike_gravel");
    if (!m) return null;
    plan = [{ m, min, km, hm }];
  } else if (art === "bergtour" || art === "skitour") {
    const m = finde(modelle, "hike");
    if (!m) return null;
    if (art === "skitour") {
      const alleSki = skitouren;
      skitouren = skitouren.filter((t) => t.vergleichbar);
      const ausgeschl = alleSki.length - skitouren.length;
      if (!skitouren.length) {
        hinweise.push("Skitour \xFCber dein Wandermodell gesch\xE4tzt \u2013 eigene Skitouren fehlen, die Herzfrequenz beim Skitourengehen liegt meist h\xF6her (Sch\xE4tzung eher zu niedrig).");
        plan = [{ m, min, km, hm }];
      } else {
        const hf = median(skitouren.map((t) => t.hf));
        const ratios = skitouren.map((t) => ({ t, e: vorwaerts(m, { min: t.min, km: t.km, hm: t.hm, hf: t.hf }) })).filter((x) => x.e && x.e.ess > 0).map((x) => x.t.ess / x.e.ess);
        const k = ratios.length ? Math.min(1.5, Math.max(0.7, median(ratios))) : 1;
        const e = vorwaerts(m, { min, km, hm, hf });
        if (!e) return null;
        const f3 = (v) => Math.round(v * k);
        const jahre = `${skitouren[0].datum.slice(0, 4)}\u2013${skitouren[skitouren.length - 1].datum.slice(0, 4)}`;
        const letzte = skitouren[skitouren.length - 1];
        hinweise.push(
          `Skitour mit deiner Skitour-Herzfrequenz (\xD8 HF ${Math.round(hf)}) gerechnet und an ${skitouren.length} eigene${skitouren.length === 1 ? "r" : "n"} Skitour${skitouren.length === 1 ? "" : "en"} aus Garmin abgeglichen (${jahre}, Faktor ${k.toLocaleString("de-DE", { maximumFractionDigits: 2 })}; zuletzt ${de(letzte.datum)}: ${(Math.round(letzte.min / 6) / 10).toLocaleString("de-DE")} h, ${letzte.hm} Hm, ${letzte.ess} ESS).` + (ausgeschl ? ` ${ausgeschl} Touren von 12/2022\u201312/2023 nicht verwendet: Garmin-Last dort etwa doppelt so hoch, anderer Ma\xDFstab.` : "") + (skitouren.length < 3 ? " Wenige eigene Skitouren \u2013 Sch\xE4tzung wird mit jeder Tour genauer." : "")
        );
        const teile2 = [{ profil: "hike", min: Math.round(min), km, hm: Math.round(hm), ess: f3(e.ess), lo: f3(e.lo), hi: f3(e.hi), n: skitouren.length }];
        return { ess: f3(e.ess), lo: f3(e.lo), hi: f3(e.hi), teile: teile2, hinweise };
      }
    } else {
      const e = vorwaerts(m, { min, km, hm });
      if (!e) return null;
      const ratios = bergtouren.map((t) => ({ t, p: vorwaerts(m, { min: t.min, km: t.km, hm: t.hm }) })).filter((x) => x.p && x.p.ess > 0).map((x) => x.t.ess / x.p.ess);
      const k = ratios.length >= 3 ? Math.min(1.5, Math.max(0.7, median(ratios))) : 1;
      if (ratios.length >= 3) {
        const letzte = bergtouren[bergtouren.length - 1];
        hinweise.push(
          `Bergtour an ${ratios.length} eigenen Bergtouren seit 08/2025 abgeglichen (\u2265 700 Hm, Faktor ${k.toLocaleString("de-DE", { maximumFractionDigits: 2 })}; zuletzt ${de(letzte.datum)}: ${(Math.round(letzte.min / 6) / 10).toLocaleString("de-DE")} h, ${letzte.hm} Hm, ${letzte.ess} ESS). \xC4ltere Bergtouren (2020\u201307/2025) hatten bei gleicher Tour rund 30 % weniger Garmin-Last und z\xE4hlen nicht.`
        );
      } else hinweise.push("Bergtour \xFCber dein Wandermodell \u2013 noch zu wenige eigene Bergtouren ab 700 Hm f\xFCr einen Abgleich.");
      const f3 = (v) => Math.round(v * k);
      return { ess: f3(e.ess), lo: f3(e.lo), hi: f3(e.hi), teile: [{ profil: "hike", min: Math.round(min), km, hm: Math.round(hm), ess: f3(e.ess), lo: f3(e.lo), hi: f3(e.hi), n: m.n }], hinweise: [...hinweise, ...e.hinweise.filter((h2) => !hinweise.includes(h2))] };
    }
  } else if (art === "bike_hike") {
    const rad = finde(modelle, "bike_gravel", "bike_road");
    const geh = finde(modelle, "hike");
    if (!rad || !geh) return null;
    const h2 = (v) => v == null ? void 0 : v / 2;
    plan = [
      { m: rad, min: min / 2, km: h2(km), hm: hm / 2 },
      { m: geh, min: min / 2, km: h2(km), hm: hm / 2 }
    ];
    hinweise.push("Bike & Hike: je halbe Zeit und halbe H\xF6henmeter auf Rad und zu Fu\xDF angenommen.");
  } else return null;
  const teile = [];
  for (const p of plan) {
    const e = vorwaerts(p.m, { min: p.min, km: p.km, hm: p.hm });
    if (!e || !Number.isFinite(e.ess)) return null;
    teile.push({ profil: p.m.id, min: Math.round(p.min), km: p.km, hm: Math.round(p.hm), ess: e.ess, lo: e.lo, hi: e.hi, n: p.m.n });
    for (const t of e.hinweise) if (!hinweise.includes(t)) hinweise.push(t);
  }
  const sum = (k) => Math.round(teile.reduce((a, t) => a + t[k], 0));
  return { ess: sum("ess"), lo: sum("lo"), hi: sum("hi"), teile, hinweise };
}

// widget_cli.ts
var baenderJetzt = null;
async function baenderLaden() {
  if (baenderJetzt) return baenderJetzt;
  try {
    const all = await fetchChartData("9999d", RECOVERY_FETCH_METRICS);
    baenderJetzt = persoenlicheBaender(buildDays(all?.data || []));
  } catch {
    baenderJetzt = null;
  }
  return baenderJetzt;
}
var iso = (d = /* @__PURE__ */ new Date()) => d.toLocaleDateString("sv-SE", { timeZone: "Europe/Berlin" });
var num = (v) => {
  const n = typeof v === "number" ? v : Number(String(v ?? "").replace(",", "."));
  return Number.isFinite(n) ? n : null;
};
var STATUS_TEXT = { QUALITY: "Qualit\xE4t m\xF6glich", TRAIN: "normal trainieren", EASY: "locker trainieren", REST: "Erholung priorisieren" };
var tageMerk = [];
var plan7Merk = null;
var neuWert = (r, post, morgen) => {
  const p = r ? num(r[post]) : null;
  if (p != null && p > 0) return p;
  const m = r ? num(r[morgen]) : null;
  return m != null && m > 0 ? m : null;
};
async function heuteTeil() {
  const heute2 = iso();
  const res = await fetchChartData(rangeFor(14, 28, 90), [...RECOVERY_FETCH_METRICS, "fb_TR_obs", "fbACWR_obs", "garminACWR"]);
  const days = buildDays(res.data || []).filter((d) => d.date <= heute2);
  const acwrJetzt = (() => {
    const r = (res.data || []).find((x) => String(x.date).slice(0, 10) === heute2);
    return neuWert(r, "fbACWR_obs", "garminACWR");
  })();
  const trJetzt = (() => {
    const r = (res.data || []).find((x) => String(x.date).slice(0, 10) === heute2);
    const v = r ? num(r.fb_TR_obs) : null;
    return v != null && v > 0 ? v : null;
  })();
  tageMerk = days;
  const t = days.length && days[days.length - 1].date === heute2 ? days[days.length - 1] : null;
  let plan = null;
  let naechste = [];
  try {
    const sim = await fetchPlanSimulationWithTimeout(45e3);
    plan = sim?.ok ? (sim.days || []).find((x) => String(x.date).slice(0, 10) === heute2) || null : null;
    if (sim?.ok)
      naechste = (sim.days || []).filter((x) => String(x.date).slice(0, 10) > heute2).slice(0, 7).map((x) => ({ d: String(x.date).slice(0, 10), ess: num(x.sim_load ?? x.load), sport: String(x.sport || ""), zone: String(x.zone || "") }));
  } catch {
  }
  try {
    if (t && t.atlPost != null && t.ctlPost != null && naechste.length) {
      let atl = t.atlPost, ctl = t.ctlPost;
      const loads = days.map((d) => d.essDay ?? 0);
      const p7 = { start: naechste[0].d, ess: [], atl: [], ctl: [], acwr: [] };
      for (const x of naechste.slice(0, 7)) {
        const l = x.ess ?? 0;
        loads.push(l);
        atl = atlStep(atl, l);
        ctl = ctlStep(ctl, l, windowMean(loads));
        p7.ess.push(Math.round(l));
        p7.atl.push(Math.round(atl));
        p7.ctl.push(Math.round(ctl));
        p7.acwr.push(ctl > 0 ? Math.round(atl / ctl * 100) / 100 : 0);
      }
      plan7Merk = p7;
    }
  } catch {
  }
  const ti = t ? days.length - 1 : -1;
  const rahmen = t && ti > 0 ? lastEmpfehlung(t, days[ti - 1], await baenderLaden()) : null;
  const planLoad = plan ? num(plan.load) : t?.essDay ?? null;
  const woche = days.slice(-7).map((d) => ({ d: d.date, ist: d.activityDone ? d.actual?.load ?? d.load ?? 0 : 0 }));
  const montag = (() => {
    const d = /* @__PURE__ */ new Date(heute2 + "T12:00:00Z");
    const wd = (d.getUTCDay() + 6) % 7;
    d.setUTCDate(d.getUTCDate() - wd);
    return d.toISOString().slice(0, 10);
  })();
  const sonntag = (() => {
    const d = /* @__PURE__ */ new Date(montag + "T12:00:00Z");
    d.setUTCDate(d.getUTCDate() + 6);
    return d.toISOString().slice(0, 10);
  })();
  const istTag = (d) => d.activityDone ? d.actual?.load ?? d.load ?? 0 : 0;
  const wIstVorher = days.filter((d) => d.date >= montag && d.date < heute2).reduce((a, d) => a + istTag(d), 0);
  const wIstHeute = t ? istTag(t) : 0;
  const wPlanRest = naechste.filter((x) => x.d <= sonntag).reduce((a, x) => a + (x.ess || 0), 0);
  const wocheTeil = { kw: montag, istVorher: Math.round(wIstVorher), ist: Math.round(wIstVorher + wIstHeute), budgetBasis: Math.round(wIstVorher + Math.max(wIstHeute, planLoad || 0) + wPlanRest) };
  return {
    datum: heute2,
    woche: wocheTeil,
    votum: t?.vote ? { stufe: VOTE_LEVEL[t.vote], vote: t.vote, label: VOTE_LABEL[t.vote], text: STATUS_TEXT[t.vote], gruende: t.reasons.slice(0, 3) } : { stufe: "grau", vote: null, label: "\u2013", text: t ? "Morgenwerte fehlen" : "noch keine Daten f\xFCr heute", gruende: t?.reasons?.slice(0, 1) || [] },
    morgen: t ? {
      eingetragen: !!t.hasMorningData,
      // 22.11.2026: Anzeige immer der neueste Wert (nach Aktivität), Votum bleibt beim Morgenwert
      readiness: trJetzt ?? t.readiness,
      readinessMorgen: t.readiness,
      readinessQuelle: trJetzt != null && trJetzt !== t.readiness ? "aktuell" : "morgens",
      schlaf: t.sleepScore,
      schlafH: t.sleepH,
      hrv: t.hrv,
      hrvStufe: t.levels.hrv,
      rhr: t.rhr,
      rhrStufe: t.levels.rhr,
      rhrDelta: t.rhrDelta,
      befinden: t.befinden
    } : null,
    last: {
      // sollHeute bleibt für ältere App-/Uhr-Versionen: Obergrenze des Rahmens
      sollHeute: rahmen ? rahmen.max : planLoad,
      sollMin: rahmen ? rahmen.min : null,
      sollMax: rahmen ? rahmen.max : null,
      sollText: rahmen ? `${rahmen.min}\u2013${rahmen.max}` : planLoad != null ? String(Math.round(planLoad)) : null,
      halten: rahmen ? rahmen.halte : null,
      planHeute: planLoad,
      planLage: rahmen && planLoad != null && planLoad > 0 ? planLoad > rahmen.max ? "darueber" : planLoad < rahmen.min ? "darunter" : "im Rahmen" : null,
      sportHeute: plan ? String(plan.sport || "") : "",
      zoneHeute: plan ? String(plan.zone || "") : "",
      istHeute: t?.activityDone ? t.actual?.load ?? t.load ?? null : 0,
      erledigt: !!t?.activityDone,
      // 23.11.2026: Anzeige neuester Garmin-Wert von heute, Votum-Basis (gestern nach Aktivität) getrennt
      acwr: acwrJetzt ?? t?.acwr ?? null,
      acwrVotum: t?.acwr ?? null,
      acwrBand: (() => {
        const v = acwrJetzt ?? t?.acwr;
        return v == null ? null : v > 1.5 ? "hoch" : v > 1.3 ? "erhoeht" : v < 0.8 ? "niedrig" : "korridor";
      })(),
      woche
    },
    naechste,
    _rahmen: rahmen
  };
}
async function formTeil() {
  const heute2 = iso();
  const res = await fetchChartData("35d", ["fbCTL_obs", "fbATL_obs", "garminCTL", "garminATL", "garminEnduranceScore", "Monotony7"]);
  const rows = (res?.data || []).filter((r) => String(r.date).slice(0, 10) <= heute2);
  const letzter = (k) => {
    for (let i = rows.length - 1; i >= 0; i--) {
      const v = num(rows[i][k]);
      if (v != null && v !== 0) return { v, d: String(rows[i].date).slice(0, 10), i };
    }
    return null;
  };
  const es = letzter("garminEnduranceScore");
  let es28 = null;
  if (es) {
    const ziel = /* @__PURE__ */ new Date(es.d + "T12:00:00Z");
    ziel.setUTCDate(ziel.getUTCDate() - 28);
    const zIso = ziel.toISOString().slice(0, 10);
    const alt = rows.find((r) => String(r.date).slice(0, 10) >= zIso && num(r.garminEnduranceScore));
    es28 = alt ? Math.round(es.v - num(alt.garminEnduranceScore)) : null;
  }
  let tag = null;
  try {
    const g = await fetchGarminTage();
    const t = g.get(heute2);
    if (t) tag = { bbAufwachen: t.bb_aufwachen ?? null, bbMax: t.tag?.bb_max ?? null, stress: t.tag?.stress_avg ?? null, schritte: t.tag?.schritte ?? null };
  } catch {
  }
  const letzterNeu = (post, morgen) => {
    for (let i = rows.length - 1; i >= 0; i--) {
      const v = neuWert(rows[i], post, morgen);
      if (v != null) return v;
    }
    return null;
  };
  return { ctl: letzterNeu("fbCTL_obs", "garminCTL"), atl: letzterNeu("fbATL_obs", "garminATL"), monotonie: letzter("Monotony7")?.v ?? null, es: es?.v ?? null, es28, tag };
}
async function zielTeil() {
  const heute2 = iso();
  const items = await fetchExtras("ziel", 60);
  const ziele = items.map((it) => it.data?.ziel).filter((z3) => !!z3 && !!z3.datum && z3.datum >= heute2).sort((a, b) => a.datum.localeCompare(b.datum));
  const z2 = ziele[0];
  if (!z2) return null;
  const tage = Math.round(((/* @__PURE__ */ new Date(z2.datum + "T12:00:00Z")).getTime() - (/* @__PURE__ */ new Date(heute2 + "T12:00:00Z")).getTime()) / 864e5);
  let prozent = null;
  try {
    const [p4, garmin] = await Promise.all([ladeP4(), fetchGarminTage()]);
    setzeGarminTage(garmin);
    const modelle = garmin ? lerneModelle(garmin) : [];
    const pers = modelle.length ? tourLastPersoenlich(z2, modelle, eigeneSkitouren(garmin), eigeneBergtouren(garmin)) : null;
    const wf0 = tourform(p4, z2, pers?.ess);
    prozent = wf0.prozent;
    if (zielArt(z2) === "marsch" && garmin?.size) {
      const ab = new Date(Date.now() - 56 * 864e5).toLocaleDateString("sv-SE");
      let km = 0, h2 = 0;
      for (const [d, t] of garmin) {
        if (d < ab) continue;
        km = Math.max(km, (t.akt || []).filter((x) => /hiking|walking|mountaineering/i.test(String(x.typ || ""))).reduce((a, x) => a + (x.km || 0), 0));
        h2 = Math.max(h2, (t.summe?.dauer_s ?? (t.akt || []).reduce((a, x) => a + (x.dauer_s || 0), 0)) / 3600);
      }
      const strecke = z2.distanz_km > 0 ? Math.min(1, km / (z2.distanz_km * 0.6)) : null;
      const zeit = z2.dauer_h > 0 ? Math.min(1, h2 / (z2.dauer_h * 0.7)) : null;
      prozent = Math.round(100 * (0.2 * wf0.volumen + 0.45 * (strecke ?? wf0.lange) + 0.35 * (zeit ?? 1)));
    }
  } catch {
  }
  return { name: z2.name, datum: z2.datum, tage, prio: z2.prio ?? null, tourform: prozent, weitere: ziele.length - 1 };
}
var GRUPPE = (typ) => /cycling|biking|virtual_ride/.test(typ) ? "rad" : /running/.test(typ) ? "lauf" : /hiking|mountaineering|walking/.test(typ) ? "berg" : /ski/.test(typ) ? "ski" : "sonst";
function zonenRaten(xs) {
  if (xs.length < 12) return null;
  const prior = [20, 45, 80, 130, 190], lam = 0.5;
  const r = prior.slice();
  for (let it = 0; it < 300; it++) {
    for (let z2 = 0; z2 < 5; z2++) {
      let num0 = lam * prior[z2], den = lam;
      for (const x of xs) {
        const rest = x.last - x.h.reduce((a, h2, k) => k === z2 ? a : a + h2 * r[k], 0);
        num0 += x.h[z2] * rest;
        den += x.h[z2] * x.h[z2];
      }
      r[z2] = Math.max(0, num0 / den);
    }
  }
  for (let z2 = 1; z2 < 5; z2++) if (r[z2] < r[z2 - 1]) r[z2] = r[z2 - 1];
  return r.map((v) => Math.round(v));
}
async function feldTeil(rahmen, t) {
  const g = await fetchGarminTage();
  const pro = { rad: [], lauf: [], berg: [], ski: [], sonst: [], alle: [] };
  for (const [, tag] of g) for (const a of tag.akt || []) {
    const zs = a.zonen_s;
    if (!zs || zs.length < 5 || !a.last || (a.dauer_s || 0) < 900) continue;
    const x = { h: zs.slice(0, 5).map((v) => (v || 0) / 3600), last: a.last };
    pro[GRUPPE(String(a.typ || ""))].push(x);
    pro.alle.push(x);
  }
  const alle = zonenRaten(pro.alle) || [20, 45, 80, 130, 190];
  const raten = {};
  for (const k of ["rad", "lauf", "berg", "ski", "sonst"]) raten[k] = zonenRaten(pro[k]) || alle;
  return {
    datum: iso(),
    vote: t?.vote ?? null,
    label: t?.vote ? VOTE_LABEL[t.vote] : "\u2013",
    min: rahmen ? rahmen.min : null,
    max: rahmen ? rahmen.max : null,
    raten,
    // 23.11.2026: ESS heute schon gespeicherter Einheiten – das Datenfeld rechnet Tagesstand = vorher + laufend
    vorher: Math.round((g.get(iso())?.akt || []).reduce((a, x) => a + ((x.last || 0) > 0 ? x.last : 0), 0)),
    n: Object.fromEntries(Object.entries(pro).map(([k, v]) => [k, v.length]))
  };
}
var LV_G = { gruen: "b", gelb: "G", orange: "O", rot: "R", grau: "v" };
var LV_HRV = { gruen: "G", gelb: "O", orange: "O", rot: "R" };
async function uhrTeil(rahmen) {
  const heute2 = iso();
  const days = tageMerk;
  const t = days.length && days[days.length - 1].date === heute2 ? days[days.length - 1] : null;
  const b = await baenderLaden();
  const l14 = { von: null, ist: [], lo: [], hi: [] };
  const ab = Math.max(1, days.length - 14);
  for (let i = ab; i < days.length; i++) {
    const d = days[i];
    const e = lastEmpfehlung(d, days[i - 1], b);
    if (!l14.von) l14.von = d.date;
    l14.ist.push(Math.round(d.activityDone ? d.actual?.load ?? d.load ?? 0 : 0));
    l14.lo.push(e ? e.min : -1);
    l14.hi.push(e ? e.max : -1);
  }
  const m = days.slice(-7);
  const m7 = {
    rd: m.map((d) => d.readiness ?? -1),
    hrv: m.map((d) => d.hrv ?? -1),
    rhr: m.map((d) => d.rhr ?? -1),
    sl: m.map((d) => d.sleepScore ?? -1),
    // 23.11.2026: Bereitschaft im Garmin-Muster (R <25 · O <50 · G <75 · b <95 · v Optimal)
    rdS: m.map((d) => d.readiness == null ? "x" : d.readiness >= 95 ? "v" : d.readiness >= 75 ? "b" : d.readiness >= 50 ? "G" : d.readiness >= 25 ? "O" : "R").join(""),
    hrvS: m.map((d) => LV_HRV[d.levels.hrv] || "x").join(""),
    rhrS: m.map((d) => LV_G[d.levels.rhr] || "x").join(""),
    slS: m.map((d) => LV_G[d.levels.sleep] || "x").join("")
  };
  let form = null;
  try {
    const r = await fetchChartData("100d", ["fbCTL_obs", "fbATL_obs", "garminCTL", "garminATL", "garminEnduranceScore"]);
    const rows = (r?.data || []).filter((x) => String(x.date).slice(0, 10) <= heute2);
    const reihe = (k, n) => rows.slice(-n).map((x) => {
      const v = num(x[k]);
      return v == null || v <= 0 ? -1 : Math.round(v);
    });
    const es = reihe("garminEnduranceScore", 90).filter((_, i, a) => (a.length - 1 - i) % 3 === 0);
    const reihe2 = (post, morgen, n) => rows.slice(-n).map((x) => {
      const v = neuWert(x, post, morgen);
      return v == null ? -1 : Math.round(v);
    });
    form = { ctl: reihe2("fbCTL_obs", "garminCTL", 42), atl: reihe2("fbATL_obs", "garminATL", 42), es };
  } catch {
  }
  let letzte = null;
  try {
    const g = await fetchGarminTage();
    const ds = [...g.keys()].filter((d) => d <= heute2).sort();
    for (let i = ds.length - 1; i >= 0 && !letzte; i--) {
      const akt = (g.get(ds[i])?.akt || []).slice().sort((x, y) => String(y.start || "").localeCompare(String(x.start || "")));
      const a = akt.find((x) => (x.dauer_s || 0) >= 600 && (x.last || 0) > 0);
      if (!a) continue;
      const di = days.findIndex((d) => d.date === ds[i]);
      const e = di > 0 ? lastEmpfehlung(days[di], days[di - 1], b) : null;
      const tagLast = (g.get(ds[i])?.akt || []).reduce((s2, x) => s2 + (x.last || 0), 0);
      const echte = akt.filter((x) => (x.dauer_s || 0) >= 600 && (x.last || 0) > 0).reverse();
      const zSum = [0, 0, 0, 0, 0];
      for (const x of echte) (x.zonen_s || []).slice(0, 5).forEach((v, k) => {
        zSum[k] += v || 0;
      });
      letzte = {
        einheiten: echte.slice(0, 4).map((x) => ({ name: String(x.name || "").slice(0, 28), sport: String(x.typ || ""), min: Math.round((x.dauer_s || 0) / 60), ess: Math.round(x.last || 0) })),
        tagMin: Math.round(echte.reduce((s2, x) => s2 + (x.dauer_s || 0), 0) / 60),
        tagKm: Math.round(echte.reduce((s2, x) => s2 + (x.km || 0), 0) * 10) / 10,
        tagZ: zSum.map((v) => Math.round(v / 60)),
        d: ds[i],
        name: String(a.name || "").slice(0, 40),
        sport: String(a.typ || ""),
        min: Math.round((a.dauer_s || 0) / 60),
        ess: Math.round(a.last || 0),
        tagEss: Math.round(tagLast),
        te: a.te_ae ?? null,
        tean: a.te_an ?? null,
        label: a.te_label ?? null,
        z: (a.zonen_s || []).slice(0, 5).map((v) => Math.round((v || 0) / 60)),
        km: a.km ?? null,
        hm: a.hm ?? null,
        lo: e ? e.min : null,
        hi: e ? e.max : null,
        lage: e ? tagLast > e.max ? "darueber" : tagLast < e.min ? "darunter" : "im Rahmen" : null
      };
    }
  } catch {
  }
  let feld = null;
  try {
    feld = await feldTeil(rahmen, t);
  } catch {
  }
  return { l14, m7, form, letzte, feld, plan7: plan7Merk };
}
var [h, z, f] = await Promise.all([heuteTeil(), zielTeil().catch(() => null), formTeil().catch(() => null)]);
var u = null;
try {
  u = await uhrTeil(h._rahmen ?? null);
} catch {
}
var { _rahmen: _r, ...hRest } = h;
process.stdout.write(JSON.stringify({ ok: true, stand: (/* @__PURE__ */ new Date()).toISOString(), ...hRest, ziel: z, form: f, uhr: u }));
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

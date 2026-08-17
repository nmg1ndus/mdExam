var docx = (function(exports) {
  "use strict";var __defProp = Object.defineProperty;
var __defProps = Object.defineProperties;
var __getOwnPropDescs = Object.getOwnPropertyDescriptors;
var __getOwnPropSymbols = Object.getOwnPropertySymbols;
var __hasOwnProp = Object.prototype.hasOwnProperty;
var __propIsEnum = Object.prototype.propertyIsEnumerable;
var __defNormalProp = (obj, key, value) => key in obj ? __defProp(obj, key, { enumerable: true, configurable: true, writable: true, value }) : obj[key] = value;
var __spreadValues = (a, b) => {
  for (var prop in b || (b = {}))
    if (__hasOwnProp.call(b, prop))
      __defNormalProp(a, prop, b[prop]);
  if (__getOwnPropSymbols)
    for (var prop of __getOwnPropSymbols(b)) {
      if (__propIsEnum.call(b, prop))
        __defNormalProp(a, prop, b[prop]);
    }
  return a;
};
var __spreadProps = (a, b) => __defProps(a, __getOwnPropDescs(b));
var __objRest = (source, exclude) => {
  var target = {};
  for (var prop in source)
    if (__hasOwnProp.call(source, prop) && exclude.indexOf(prop) < 0)
      target[prop] = source[prop];
  if (source != null && __getOwnPropSymbols)
    for (var prop of __getOwnPropSymbols(source)) {
      if (exclude.indexOf(prop) < 0 && __propIsEnum.call(source, prop))
        target[prop] = source[prop];
    }
  return target;
};
var __publicField = (obj, key, value) => __defNormalProp(obj, typeof key !== "symbol" ? key + "" : key, value);
var __async = (__this, __arguments, generator) => {
  return new Promise((resolve, reject) => {
    var fulfilled = (value) => {
      try {
        step(generator.next(value));
      } catch (e) {
        reject(e);
      }
    };
    var rejected = (value) => {
      try {
        step(generator.throw(value));
      } catch (e) {
        reject(e);
      }
    };
    var step = (x) => x.done ? resolve(x.value) : Promise.resolve(x.value).then(fulfilled, rejected);
    step((generator = generator.apply(__this, __arguments)).next());
  });
};

  class BaseXmlComponent {
    /**
     * Creates a new BaseXmlComponent with the specified XML element name.
     *
     * @param rootKey - The XML element name (e.g., "w:p", "w:r", "w:t")
     */
    constructor(rootKey) {
      /** The XML element name for this component (e.g., "w:p" for paragraph). */
      __publicField(this, "rootKey");
      this.rootKey = rootKey;
    }
  }
  const EMPTY_OBJECT = Object.seal({});
  class XmlComponent extends BaseXmlComponent {
    /**
     * Creates a new XmlComponent.
     *
     * @param rootKey - The XML element name (e.g., "w:p", "w:r", "w:t")
     */
    constructor(rootKey) {
      super(rootKey);
      /**
       * Array of child components, text nodes, and attributes.
       *
       * This array forms the content of the XML element. It can contain other
       * XmlComponents, string values (text nodes), or attribute components.
       */
      // eslint-disable-next-line functional/prefer-readonly-type, @typescript-eslint/no-explicit-any
      __publicField(this, "root");
      this.root = new Array();
    }
    /**
     * Prepares this component and its children for XML serialization.
     *
     * This method is called by the Formatter to convert the component tree into
     * an object structure compatible with the xml library (https://www.npmjs.com/package/xml).
     * It recursively processes all children and handles special cases like
     * attribute-only elements and empty elements.
     *
     * The method can be overridden by subclasses to customize XML representation
     * or execute side effects during serialization (e.g., creating relationships).
     *
     * @param context - The serialization context containing document state
     * @returns The XML-serializable object, or undefined to exclude from output
     *
     * @example
     * ```typescript
     * // Override to add custom serialization logic
     * prepForXml(context: IContext): IXmlableObject | undefined {
     *   // Custom logic here
     *   return super.prepForXml(context);
     * }
     * ```
     */
    prepForXml(context) {
      var _a;
      context.stack.push(this);
      const children = this.root.map((comp) => {
        if (comp instanceof BaseXmlComponent) {
          return comp.prepForXml(context);
        }
        return comp;
      }).filter((comp) => comp !== void 0);
      context.stack.pop();
      return {
        [this.rootKey]: children.length ? children.length === 1 && ((_a = children[0]) == null ? void 0 : _a._attr) ? children[0] : children : EMPTY_OBJECT
      };
    }
    /**
     * Adds a child element to this component.
     *
     * @deprecated Do not use this method. It is only used internally by the library. It will be removed in a future version.
     * @param child - The child component or text string to add
     * @returns This component (for chaining)
     */
    addChildElement(child) {
      this.root.push(child);
      return this;
    }
  }
  class IgnoreIfEmptyXmlComponent extends XmlComponent {
    constructor(rootKey, includeIfEmpty) {
      super(rootKey);
      __publicField(this, "includeIfEmpty");
      this.includeIfEmpty = includeIfEmpty;
    }
    /**
     * Prepares the component for XML serialization, excluding it if empty.
     *
     * @param context - The serialization context
     * @returns The XML-serializable object, or undefined if empty
     */
    prepForXml(context) {
      const result = super.prepForXml(context);
      if (this.includeIfEmpty) {
        return result;
      }
      if (result && (typeof result[this.rootKey] !== "object" || Object.keys(result[this.rootKey]).length)) {
        return result;
      }
      return void 0;
    }
  }
  class XmlAttributeComponent extends BaseXmlComponent {
    /**
     * Creates a new attribute component.
     *
     * @param root - The attribute data object
     */
    constructor(root) {
      super("_attr");
      /** Optional mapping from property names to XML attribute names. */
      __publicField(this, "xmlKeys");
      this.root = root;
    }
    /**
     * Converts the attribute data to an XML-serializable object.
     *
     * This method transforms the property names using xmlKeys (if defined)
     * and filters out undefined values.
     *
     * @param _ - Context (unused for attributes)
     * @returns Object with _attr key containing the mapped attributes
     */
    prepForXml(_) {
      const attrs = {};
      Object.entries(this.root).forEach(([key, value]) => {
        if (value !== void 0) {
          const newKey = this.xmlKeys && this.xmlKeys[key] || key;
          attrs[newKey] = value;
        }
      });
      return { _attr: attrs };
    }
  }
  class NextAttributeComponent extends BaseXmlComponent {
    /**
     * Creates a new NextAttributeComponent.
     *
     * @param root - Attribute payload with explicit key-value mappings
     */
    constructor(root) {
      super("_attr");
      this.root = root;
    }
    /**
     * Converts the attribute payload to an XML-serializable object.
     *
     * Extracts the key and value from each property and filters out
     * undefined values.
     *
     * @param _ - Context (unused for attributes)
     * @returns Object with _attr key containing the attributes
     */
    prepForXml(_) {
      const attrs = Object.values(this.root).filter(({ value }) => value !== void 0).reduce((acc, { key, value }) => __spreadProps(__spreadValues({}, acc), { [key]: value }), {});
      return { _attr: attrs };
    }
  }
  class Attributes extends XmlAttributeComponent {
    constructor() {
      super(...arguments);
      __publicField(this, "xmlKeys", {
        val: "w:val",
        color: "w:color",
        fill: "w:fill",
        space: "w:space",
        sz: "w:sz",
        type: "w:type",
        rsidR: "w:rsidR",
        rsidRPr: "w:rsidRPr",
        rsidSect: "w:rsidSect",
        w: "w:w",
        h: "w:h",
        top: "w:top",
        right: "w:right",
        bottom: "w:bottom",
        left: "w:left",
        header: "w:header",
        footer: "w:footer",
        gutter: "w:gutter",
        linePitch: "w:linePitch",
        pos: "w:pos"
      });
    }
  }
  var commonjsGlobal = typeof globalThis !== "undefined" ? globalThis : typeof window !== "undefined" ? window : typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : {};
  function getDefaultExportFromCjs$1(x) {
    return x && x.__esModule && Object.prototype.hasOwnProperty.call(x, "default") ? x["default"] : x;
  }
  var sax = {};
  var events = { exports: {} };
  var hasRequiredEvents;
  function requireEvents() {
    if (hasRequiredEvents) return events.exports;
    hasRequiredEvents = 1;
    var R = typeof Reflect === "object" ? Reflect : null;
    var ReflectApply = R && typeof R.apply === "function" ? R.apply : function ReflectApply2(target, receiver, args) {
      return Function.prototype.apply.call(target, receiver, args);
    };
    var ReflectOwnKeys;
    if (R && typeof R.ownKeys === "function") {
      ReflectOwnKeys = R.ownKeys;
    } else if (Object.getOwnPropertySymbols) {
      ReflectOwnKeys = function ReflectOwnKeys2(target) {
        return Object.getOwnPropertyNames(target).concat(Object.getOwnPropertySymbols(target));
      };
    } else {
      ReflectOwnKeys = function ReflectOwnKeys2(target) {
        return Object.getOwnPropertyNames(target);
      };
    }
    function ProcessEmitWarning(warning) {
      if (console && console.warn) console.warn(warning);
    }
    var NumberIsNaN = Number.isNaN || function NumberIsNaN2(value) {
      return value !== value;
    };
    function EventEmitter() {
      EventEmitter.init.call(this);
    }
    events.exports = EventEmitter;
    events.exports.once = once;
    EventEmitter.EventEmitter = EventEmitter;
    EventEmitter.prototype._events = void 0;
    EventEmitter.prototype._eventsCount = 0;
    EventEmitter.prototype._maxListeners = void 0;
    var defaultMaxListeners = 10;
    function checkListener(listener) {
      if (typeof listener !== "function") {
        throw new TypeError('The "listener" argument must be of type Function. Received type ' + typeof listener);
      }
    }
    Object.defineProperty(EventEmitter, "defaultMaxListeners", {
      enumerable: true,
      get: function() {
        return defaultMaxListeners;
      },
      set: function(arg) {
        if (typeof arg !== "number" || arg < 0 || NumberIsNaN(arg)) {
          throw new RangeError('The value of "defaultMaxListeners" is out of range. It must be a non-negative number. Received ' + arg + ".");
        }
        defaultMaxListeners = arg;
      }
    });
    EventEmitter.init = function() {
      if (this._events === void 0 || this._events === Object.getPrototypeOf(this)._events) {
        this._events = /* @__PURE__ */ Object.create(null);
        this._eventsCount = 0;
      }
      this._maxListeners = this._maxListeners || void 0;
    };
    EventEmitter.prototype.setMaxListeners = function setMaxListeners(n) {
      if (typeof n !== "number" || n < 0 || NumberIsNaN(n)) {
        throw new RangeError('The value of "n" is out of range. It must be a non-negative number. Received ' + n + ".");
      }
      this._maxListeners = n;
      return this;
    };
    function _getMaxListeners(that) {
      if (that._maxListeners === void 0)
        return EventEmitter.defaultMaxListeners;
      return that._maxListeners;
    }
    EventEmitter.prototype.getMaxListeners = function getMaxListeners() {
      return _getMaxListeners(this);
    };
    EventEmitter.prototype.emit = function emit(type2) {
      var args = [];
      for (var i = 1; i < arguments.length; i++) args.push(arguments[i]);
      var doError = type2 === "error";
      var events2 = this._events;
      if (events2 !== void 0)
        doError = doError && events2.error === void 0;
      else if (!doError)
        return false;
      if (doError) {
        var er;
        if (args.length > 0)
          er = args[0];
        if (er instanceof Error) {
          throw er;
        }
        var err = new Error("Unhandled error." + (er ? " (" + er.message + ")" : ""));
        err.context = er;
        throw err;
      }
      var handler = events2[type2];
      if (handler === void 0)
        return false;
      if (typeof handler === "function") {
        ReflectApply(handler, this, args);
      } else {
        var len = handler.length;
        var listeners = arrayClone(handler, len);
        for (var i = 0; i < len; ++i)
          ReflectApply(listeners[i], this, args);
      }
      return true;
    };
    function _addListener(target, type2, listener, prepend) {
      var m;
      var events2;
      var existing;
      checkListener(listener);
      events2 = target._events;
      if (events2 === void 0) {
        events2 = target._events = /* @__PURE__ */ Object.create(null);
        target._eventsCount = 0;
      } else {
        if (events2.newListener !== void 0) {
          target.emit(
            "newListener",
            type2,
            listener.listener ? listener.listener : listener
          );
          events2 = target._events;
        }
        existing = events2[type2];
      }
      if (existing === void 0) {
        existing = events2[type2] = listener;
        ++target._eventsCount;
      } else {
        if (typeof existing === "function") {
          existing = events2[type2] = prepend ? [listener, existing] : [existing, listener];
        } else if (prepend) {
          existing.unshift(listener);
        } else {
          existing.push(listener);
        }
        m = _getMaxListeners(target);
        if (m > 0 && existing.length > m && !existing.warned) {
          existing.warned = true;
          var w = new Error("Possible EventEmitter memory leak detected. " + existing.length + " " + String(type2) + " listeners added. Use emitter.setMaxListeners() to increase limit");
          w.name = "MaxListenersExceededWarning";
          w.emitter = target;
          w.type = type2;
          w.count = existing.length;
          ProcessEmitWarning(w);
        }
      }
      return target;
    }
    EventEmitter.prototype.addListener = function addListener(type2, listener) {
      return _addListener(this, type2, listener, false);
    };
    EventEmitter.prototype.on = EventEmitter.prototype.addListener;
    EventEmitter.prototype.prependListener = function prependListener(type2, listener) {
      return _addListener(this, type2, listener, true);
    };
    function onceWrapper() {
      if (!this.fired) {
        this.target.removeListener(this.type, this.wrapFn);
        this.fired = true;
        if (arguments.length === 0)
          return this.listener.call(this.target);
        return this.listener.apply(this.target, arguments);
      }
    }
    function _onceWrap(target, type2, listener) {
      var state2 = { fired: false, wrapFn: void 0, target, type: type2, listener };
      var wrapped = onceWrapper.bind(state2);
      wrapped.listener = listener;
      state2.wrapFn = wrapped;
      return wrapped;
    }
    EventEmitter.prototype.once = function once2(type2, listener) {
      checkListener(listener);
      this.on(type2, _onceWrap(this, type2, listener));
      return this;
    };
    EventEmitter.prototype.prependOnceListener = function prependOnceListener(type2, listener) {
      checkListener(listener);
      this.prependListener(type2, _onceWrap(this, type2, listener));
      return this;
    };
    EventEmitter.prototype.removeListener = function removeListener(type2, listener) {
      var list, events2, position, i, originalListener;
      checkListener(listener);
      events2 = this._events;
      if (events2 === void 0)
        return this;
      list = events2[type2];
      if (list === void 0)
        return this;
      if (list === listener || list.listener === listener) {
        if (--this._eventsCount === 0)
          this._events = /* @__PURE__ */ Object.create(null);
        else {
          delete events2[type2];
          if (events2.removeListener)
            this.emit("removeListener", type2, list.listener || listener);
        }
      } else if (typeof list !== "function") {
        position = -1;
        for (i = list.length - 1; i >= 0; i--) {
          if (list[i] === listener || list[i].listener === listener) {
            originalListener = list[i].listener;
            position = i;
            break;
          }
        }
        if (position < 0)
          return this;
        if (position === 0)
          list.shift();
        else {
          spliceOne(list, position);
        }
        if (list.length === 1)
          events2[type2] = list[0];
        if (events2.removeListener !== void 0)
          this.emit("removeListener", type2, originalListener || listener);
      }
      return this;
    };
    EventEmitter.prototype.off = EventEmitter.prototype.removeListener;
    EventEmitter.prototype.removeAllListeners = function removeAllListeners(type2) {
      var listeners, events2, i;
      events2 = this._events;
      if (events2 === void 0)
        return this;
      if (events2.removeListener === void 0) {
        if (arguments.length === 0) {
          this._events = /* @__PURE__ */ Object.create(null);
          this._eventsCount = 0;
        } else if (events2[type2] !== void 0) {
          if (--this._eventsCount === 0)
            this._events = /* @__PURE__ */ Object.create(null);
          else
            delete events2[type2];
        }
        return this;
      }
      if (arguments.length === 0) {
        var keys = Object.keys(events2);
        var key;
        for (i = 0; i < keys.length; ++i) {
          key = keys[i];
          if (key === "removeListener") continue;
          this.removeAllListeners(key);
        }
        this.removeAllListeners("removeListener");
        this._events = /* @__PURE__ */ Object.create(null);
        this._eventsCount = 0;
        return this;
      }
      listeners = events2[type2];
      if (typeof listeners === "function") {
        this.removeListener(type2, listeners);
      } else if (listeners !== void 0) {
        for (i = listeners.length - 1; i >= 0; i--) {
          this.removeListener(type2, listeners[i]);
        }
      }
      return this;
    };
    function _listeners(target, type2, unwrap) {
      var events2 = target._events;
      if (events2 === void 0)
        return [];
      var evlistener = events2[type2];
      if (evlistener === void 0)
        return [];
      if (typeof evlistener === "function")
        return unwrap ? [evlistener.listener || evlistener] : [evlistener];
      return unwrap ? unwrapListeners(evlistener) : arrayClone(evlistener, evlistener.length);
    }
    EventEmitter.prototype.listeners = function listeners(type2) {
      return _listeners(this, type2, true);
    };
    EventEmitter.prototype.rawListeners = function rawListeners(type2) {
      return _listeners(this, type2, false);
    };
    EventEmitter.listenerCount = function(emitter, type2) {
      if (typeof emitter.listenerCount === "function") {
        return emitter.listenerCount(type2);
      } else {
        return listenerCount.call(emitter, type2);
      }
    };
    EventEmitter.prototype.listenerCount = listenerCount;
    function listenerCount(type2) {
      var events2 = this._events;
      if (events2 !== void 0) {
        var evlistener = events2[type2];
        if (typeof evlistener === "function") {
          return 1;
        } else if (evlistener !== void 0) {
          return evlistener.length;
        }
      }
      return 0;
    }
    EventEmitter.prototype.eventNames = function eventNames() {
      return this._eventsCount > 0 ? ReflectOwnKeys(this._events) : [];
    };
    function arrayClone(arr, n) {
      var copy = new Array(n);
      for (var i = 0; i < n; ++i)
        copy[i] = arr[i];
      return copy;
    }
    function spliceOne(list, index) {
      for (; index + 1 < list.length; index++)
        list[index] = list[index

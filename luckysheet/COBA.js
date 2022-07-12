(function(f){if(typeof exports==="object"&&typeof module!=="undefined"){module.exports=f()}else if(typeof define==="function"&&define.amd){define([],f)}else{var g;if(typeof window!=="undefined"){g=window}else if(typeof global!=="undefined"){g=global}else if(typeof self!=="undefined"){g=self}else{g=this}g.LuckyExcel = f()}})(function(){var define,module,exports;return (function(){function r(e,n,t){function o(i,f){if(!n[i]){if(!e[i]){var c="function"==typeof require&&require;if(!f&&c)return c(i,!0);if(u)return u(i,!0);var a=new Error("Cannot find module '"+i+"'");throw a.code="MODULE_NOT_FOUND",a}var p=n[i]={exports:{}};e[i][0].call(p.exports,function(r){var n=e[i][1][r];return o(n||r)},p,p.exports,r,e,n,t)}return n[i].exports}for(var u="function"==typeof require&&require,i=0;i<t.length;i++)o(t[i]);return o}return r})()({1:[function(require,module,exports){

},{}],2:[function(require,module,exports){

},{}],3:[function(require,module,exports){

},{"base64-js":1,"buffer":3,"ieee754":6}],4:[function(require,module,exports){

},{}],6:[function(require,module,exports){

},{}],7:[function(require,module,exports){

},{}],8:[function(require,module,exports){

},{}],9:[function(require,module,exports){

},{}],10:[function(require,module,exports){

},{}],11:[function(require,module,exports){

},{"./support":40,"./utils":42}],12:[function(require,module,exports){

},{"./external":16,"./stream/Crc32Probe":35,"./stream/DataLengthProbe":36,"./stream/DataWorker":37}],13:[function(require,module,exports){

},{"./flate":17,"./stream/GenericWorker":38}],14:[function(require,module,exports){

},{"./utils":42}],15:[function(require,module,exports){

},{}],16:[function(require,module,exports){

},{"lie":46}],17:[function(require,module,exports){

},{"./stream/GenericWorker":38,"./utils":42,"pako":47}],18:[function(require,module,exports){

},{"../crc32":14,"../signature":33,"../stream/GenericWorker":38,"../utf8":41,"../utils":42}],19:[function(require,module,exports){

},{"../compressions":13,"./ZipFileWorker":18}],20:[function(require,module,exports){

},{"./defaults":15,"./external":16,"./load":21,"./object":25,"./support":40}],21:[function(require,module,exports){

},{"./external":16,"./nodejsUtils":22,"./stream/Crc32Probe":35,"./utf8":41,"./utils":42,"./zipEntries":43}],22:[function(require,module,exports){

},{"buffer":3}],23:[function(require,module,exports){

},{"../stream/GenericWorker":38,"../utils":42}],24:[function(require,module,exports){

},{"../utils":42,"readable-stream":26}],25:[function(require,module,exports){

},{"./compressedObject":12,"./defaults":15,"./generate":19,"./nodejs/NodejsStreamInputAdapter":23,"./nodejsUtils":22,"./stream/GenericWorker":38,"./stream/StreamHelper":39,"./utf8":41,"./utils":42,"./zipObject":45}],26:[function(require,module,exports){

},{"stream":80}],27:[function(require,module,exports){

},{"../utils":42,"./DataReader":28}],28:[function(require,module,exports){

},{"../utils":42}],29:[function(require,module,exports){

},{"../utils":42,"./Uint8ArrayReader":31}],30:[function(require,module,exports){

},{"../utils":42,"./ArrayReader":27}],32:[function(require,module,exports){

},{"../support":40,"../utils":42,"./ArrayReader":27,"./NodeBufferReader":29,"./StringReader":30,"./Uint8ArrayReader":31}],33:[function(require,module,exports){

},{}],34:[function(require,module,exports){

},{"../utils":42,"./GenericWorker":38}],35:[function(require,module,exports){

},{"../crc32":14,"../utils":42,"./GenericWorker":38}],36:[function(require,module,exports){

},{"../utils":42,"./GenericWorker":38}],37:[function(require,module,exports){

},{"../utils":42,"./GenericWorker":38}],38:[function(require,module,exports){

},{}],39:[function(require,module,exports){

},{"../base64":11,"../external":16,"../nodejs/NodejsStreamOutputAdapter":24,"../support":40,"../utils":42,"./ConvertWorker":34,"./GenericWorker":38,"buffer":3}],40:[function(require,module,exports){

},{"buffer":3,"readable-stream":26}],41:[function(require,module,exports){

},{"./nodejsUtils":22,"./stream/GenericWorker":38,"./support":40,"./utils":42}],42:[function(require,module,exports){

},{"./base64":11,"./external":16,"./nodejsUtils":22,"./support":40,"set-immediate-shim":79}],43:[function(require,module,exports){

},{"./reader/readerFor":32,"./signature":33,"./support":40,"./utf8":41,"./utils":42,"./zipEntry":44}],44:[function(require,module,exports){

},{"./compressedObject":12,"./compressions":13,"./crc32":14,"./reader/readerFor":32,"./support":40,"./utf8":41,"./utils":42}],45:[function(require,module,exports){

},{"./compressedObject":12,"./stream/DataWorker":37,"./stream/GenericWorker":38,"./stream/StreamHelper":39,"./utf8":41}],46:[function(require,module,exports){

},{"immediate":7}],47:[function(require,module,exports){

},{"./utils/common":50,"./utils/strings":51,"./zlib/deflate":55,"./zlib/messages":60,"./zlib/zstream":62}],49:[function(require,module,exports){

},{"./utils/common":50,"./utils/strings":51,"./zlib/constants":53,"./zlib/gzheader":56,"./zlib/inflate":58,"./zlib/messages":60,"./zlib/zstream":62}],50:[function(require,module,exports){

},{}],51:[function(require,module,exports){

},{"./common":50}],52:[function(require,module,exports){

},{}],53:[function(require,module,exports){

},{}],54:[function(require,module,exports){

},{}],55:[function(require,module,exports){

},{"../utils/common":50,"./adler32":52,"./crc32":54,"./messages":60,"./trees":61}],56:[function(require,module,exports){

},{}],57:[function(require,module,exports){

},{}],58:[function(require,module,exports){

},{"../utils/common":50,"./adler32":52,"./crc32":54,"./inffast":57,"./inftrees":59}],59:[function(require,module,exports){

},{"../utils/common":50}],60:[function(require,module,exports){

},{}],61:[function(require,module,exports){

},{"../utils/common":50}],62:[function(require,module,exports){

},{}],63:[function(require,module,exports){

},{"_process":64}],64:[function(require,module,exports){

},{}],65:[function(require,module,exports){

},{"./lib/_stream_duplex.js":66}],66:[function(require,module,exports){

},{"./_stream_readable":68,"./_stream_writable":70,"core-util-is":4,"inherits":8,"process-nextick-args":63}],67:[function(require,module,exports){

},{"./_stream_transform":69,"core-util-is":4,"inherits":8}],68:[function(require,module,exports){

},{"./_stream_duplex":66,"./internal/streams/BufferList":71,"./internal/streams/destroy":72,"./internal/streams/stream":73,"_process":64,"core-util-is":4,"events":5,"inherits":8,"isarray":10,"process-nextick-args":63,"safe-buffer":78,"string_decoder/":81,"util":2}],69:[function(require,module,exports){

},{"./_stream_duplex":66,"core-util-is":4,"inherits":8}],70:[function(require,module,exports){

},{"./_stream_duplex":66,"./internal/streams/destroy":72,"./internal/streams/stream":73,"_process":64,"core-util-is":4,"inherits":8,"process-nextick-args":63,"safe-buffer":78,"timers":82,"util-deprecate":83}],71:[function(require,module,exports){

},{"safe-buffer":78,"util":2}],72:[function(require,module,exports){

},{"process-nextick-args":63}],73:[function(require,module,exports){

},{"events":5}],74:[function(require,module,exports){

},{"./readable":75}],75:[function(require,module,exports){

},{"./lib/_stream_duplex.js":66,"./lib/_stream_passthrough.js":67,"./lib/_stream_readable.js":68,"./lib/_stream_transform.js":69,"./lib/_stream_writable.js":70}],76:[function(require,module,exports){

},{"./readable":75}],77:[function(require,module,exports){

},{"./lib/_stream_writable.js":70}],78:[function(require,module,exports){

},{"buffer":3}],79:[function(require,module,exports){

},{"timers":82}],80:[function(require,module,exports){

},{"events":5,"inherits":8,"readable-stream/duplex.js":65,"readable-stream/passthrough.js":74,"readable-stream/readable.js":75,"readable-stream/transform.js":76,"readable-stream/writable.js":77}],81:[function(require,module,exports){

},{"safe-buffer":78}],82:[function(require,module,exports){

},{"process/browser.js":64,"timers":82}],83:[function(require,module,exports){

},{}],84:[function(require,module,exports){

},{"./common/method":93,"jszip":20}],85:[function(require,module,exports){

},{}],86:[function(require,module,exports){

},{"../common/constant":91,"../common/method":93,"./LuckyBase":85,"./ReadXml":90}],87:[function(require,module,exports){

},{"../common/constant":91,"../common/method":93,"./LuckyBase":85,"./LuckyImage":88,"./LuckySheet":89,"./ReadXml":90}],88:[function(require,module,exports){

},{"../common/emf":92,"./LuckyBase":85}],89:[function(require,module,exports){

},{"../common/method":93,"./LuckyBase":85,"./LuckyCell":86,"./ReadXml":90}],90:[function(require,module,exports){

},{"../common/constant":91,"../common/method":93}],91:[function(require,module,exports){

},{}],92:[function(require,module,exports){

},{}],93:[function(require,module,exports){

},{"./constant":91}],94:[function(require,module,exports){

},{"./HandleZip":84,"./ToLuckySheet/LuckyFile":87}],95:[function(require,module,exports){

},{"./main":94}]},{},[95])(95)
});
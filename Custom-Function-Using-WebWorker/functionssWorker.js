!function(){function t(n){return t="function"==typeof Symbol&&"symbol"==typeof Symbol.iterator?function(t){return typeof t}:function(t){return t&&"function"==typeof Symbol&&t.constructor===Symbol&&t!==Symbol.prototype?"symbol":typeof t},t(n)}function n(t){for(var n=0,a=0;a<t;a++){n+=Math.tan(Math.atan(Math.tan(Math.atan(Math.tan(Math.atan(Math.tan(Math.atan(Math.tan(Math.atan(50))))))))));for(var o=0;o<t;o++)n-=Math.tan(Math.atan(Math.tan(Math.atan(Math.tan(Math.atan(Math.tan(Math.atan(Math.tan(Math.atan(50))))))))))}return n}function a(t){return new Promise((function(a,o){setTimeout((function(){a(n(t))}),1e3)}))}function o(t){throw new Error}function r(t){return Promise.reject(new Error)}self.addEventListener("message",(function(e){var u=e.data;"string"==typeof u&&(u=JSON.parse(u));var f=u.jobId;try{var i=function(t,e){if("TEST"==t)return n.apply(null,e);if("TEST_PROMISE"==t)return a.apply(null,e);if("TEST_ERROR"==t)return o.apply(null,e);if("TEST_ERROR_PROMISE"==t)return r.apply(null,e);throw new Error("not supported")}(u.name,u.parameters);"function"==typeof i||"object"==t(i)&&"function"==typeof i.then?i.then((function(t){postMessage({jobId:f,result:t})})).catch((function(t){postMessage({jobId:f,error:!0})})):postMessage({jobId:f,result:i})}catch(t){postMessage({jobId:f,error:!0})}}))}();
//# sourceMappingURL=functionssWorker.js.map
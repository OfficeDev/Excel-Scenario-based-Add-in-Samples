/*! For license information please see 7a5224bbf9840dcded69.js.LICENSE.txt */
function _typeof(e){return _typeof="function"==typeof Symbol&&"symbol"==typeof Symbol.iterator?function(e){return typeof e}:function(e){return e&&"function"==typeof Symbol&&e.constructor===Symbol&&e!==Symbol.prototype?"symbol":typeof e},_typeof(e)}function _regeneratorRuntime(){"use strict";_regeneratorRuntime=function(){return t};var e,t={},r=Object.prototype,n=r.hasOwnProperty,o=Object.defineProperty||function(e,t,r){e[t]=r.value},a="function"==typeof Symbol?Symbol:{},i=a.iterator||"@@iterator",c=a.asyncIterator||"@@asyncIterator",s=a.toStringTag||"@@toStringTag";function u(e,t,r){return Object.defineProperty(e,t,{value:r,enumerable:!0,configurable:!0,writable:!0}),e[t]}try{u({},"")}catch(e){u=function(e,t,r){return e[t]=r}}function l(e,t,r,n){var a=t&&t.prototype instanceof g?t:g,i=Object.create(a.prototype),c=new N(n||[]);return o(i,"_invoke",{value:O(e,r,c)}),i}function f(e,t,r){try{return{type:"normal",arg:e.call(t,r)}}catch(e){return{type:"throw",arg:e}}}t.wrap=l;var h="suspendedStart",p="suspendedYield",m="executing",y="completed",v={};function g(){}function d(){}function b(){}var w={};u(w,i,(function(){return this}));var x=Object.getPrototypeOf,k=x&&x(x(j([])));k&&k!==r&&n.call(k,i)&&(w=k);var _=b.prototype=g.prototype=Object.create(w);function S(e){["next","throw","return"].forEach((function(t){u(e,t,(function(e){return this._invoke(t,e)}))}))}function E(e,t){function r(o,a,i,c){var s=f(e[o],e,a);if("throw"!==s.type){var u=s.arg,l=u.value;return l&&"object"==_typeof(l)&&n.call(l,"__await")?t.resolve(l.__await).then((function(e){r("next",e,i,c)}),(function(e){r("throw",e,i,c)})):t.resolve(l).then((function(e){u.value=e,i(u)}),(function(e){return r("throw",e,i,c)}))}c(s.arg)}var a;o(this,"_invoke",{value:function(e,n){function o(){return new t((function(t,o){r(e,n,t,o)}))}return a=a?a.then(o,o):o()}})}function O(t,r,n){var o=h;return function(a,i){if(o===m)throw new Error("Generator is already running");if(o===y){if("throw"===a)throw i;return{value:e,done:!0}}for(n.method=a,n.arg=i;;){var c=n.delegate;if(c){var s=T(c,n);if(s){if(s===v)continue;return s}}if("next"===n.method)n.sent=n._sent=n.arg;else if("throw"===n.method){if(o===h)throw o=y,n.arg;n.dispatchException(n.arg)}else"return"===n.method&&n.abrupt("return",n.arg);o=m;var u=f(t,r,n);if("normal"===u.type){if(o=n.done?y:p,u.arg===v)continue;return{value:u.arg,done:n.done}}"throw"===u.type&&(o=y,n.method="throw",n.arg=u.arg)}}}function T(t,r){var n=r.method,o=t.iterator[n];if(o===e)return r.delegate=null,"throw"===n&&t.iterator.return&&(r.method="return",r.arg=e,T(t,r),"throw"===r.method)||"return"!==n&&(r.method="throw",r.arg=new TypeError("The iterator does not provide a '"+n+"' method")),v;var a=f(o,t.iterator,r.arg);if("throw"===a.type)return r.method="throw",r.arg=a.arg,r.delegate=null,v;var i=a.arg;return i?i.done?(r[t.resultName]=i.value,r.next=t.nextLoc,"return"!==r.method&&(r.method="next",r.arg=e),r.delegate=null,v):i:(r.method="throw",r.arg=new TypeError("iterator result is not an object"),r.delegate=null,v)}function L(e){var t={tryLoc:e[0]};1 in e&&(t.catchLoc=e[1]),2 in e&&(t.finallyLoc=e[2],t.afterLoc=e[3]),this.tryEntries.push(t)}function P(e){var t=e.completion||{};t.type="normal",delete t.arg,e.completion=t}function N(e){this.tryEntries=[{tryLoc:"root"}],e.forEach(L,this),this.reset(!0)}function j(t){if(t||""===t){var r=t[i];if(r)return r.call(t);if("function"==typeof t.next)return t;if(!isNaN(t.length)){var o=-1,a=function r(){for(;++o<t.length;)if(n.call(t,o))return r.value=t[o],r.done=!1,r;return r.value=e,r.done=!0,r};return a.next=a}}throw new TypeError(_typeof(t)+" is not iterable")}return d.prototype=b,o(_,"constructor",{value:b,configurable:!0}),o(b,"constructor",{value:d,configurable:!0}),d.displayName=u(b,s,"GeneratorFunction"),t.isGeneratorFunction=function(e){var t="function"==typeof e&&e.constructor;return!!t&&(t===d||"GeneratorFunction"===(t.displayName||t.name))},t.mark=function(e){return Object.setPrototypeOf?Object.setPrototypeOf(e,b):(e.__proto__=b,u(e,s,"GeneratorFunction")),e.prototype=Object.create(_),e},t.awrap=function(e){return{__await:e}},S(E.prototype),u(E.prototype,c,(function(){return this})),t.AsyncIterator=E,t.async=function(e,r,n,o,a){void 0===a&&(a=Promise);var i=new E(l(e,r,n,o),a);return t.isGeneratorFunction(r)?i:i.next().then((function(e){return e.done?e.value:i.next()}))},S(_),u(_,s,"Generator"),u(_,i,(function(){return this})),u(_,"toString",(function(){return"[object Generator]"})),t.keys=function(e){var t=Object(e),r=[];for(var n in t)r.push(n);return r.reverse(),function e(){for(;r.length;){var n=r.pop();if(n in t)return e.value=n,e.done=!1,e}return e.done=!0,e}},t.values=j,N.prototype={constructor:N,reset:function(t){if(this.prev=0,this.next=0,this.sent=this._sent=e,this.done=!1,this.delegate=null,this.method="next",this.arg=e,this.tryEntries.forEach(P),!t)for(var r in this)"t"===r.charAt(0)&&n.call(this,r)&&!isNaN(+r.slice(1))&&(this[r]=e)},stop:function(){this.done=!0;var e=this.tryEntries[0].completion;if("throw"===e.type)throw e.arg;return this.rval},dispatchException:function(t){if(this.done)throw t;var r=this;function o(n,o){return c.type="throw",c.arg=t,r.next=n,o&&(r.method="next",r.arg=e),!!o}for(var a=this.tryEntries.length-1;a>=0;--a){var i=this.tryEntries[a],c=i.completion;if("root"===i.tryLoc)return o("end");if(i.tryLoc<=this.prev){var s=n.call(i,"catchLoc"),u=n.call(i,"finallyLoc");if(s&&u){if(this.prev<i.catchLoc)return o(i.catchLoc,!0);if(this.prev<i.finallyLoc)return o(i.finallyLoc)}else if(s){if(this.prev<i.catchLoc)return o(i.catchLoc,!0)}else{if(!u)throw new Error("try statement without catch or finally");if(this.prev<i.finallyLoc)return o(i.finallyLoc)}}}},abrupt:function(e,t){for(var r=this.tryEntries.length-1;r>=0;--r){var o=this.tryEntries[r];if(o.tryLoc<=this.prev&&n.call(o,"finallyLoc")&&this.prev<o.finallyLoc){var a=o;break}}a&&("break"===e||"continue"===e)&&a.tryLoc<=t&&t<=a.finallyLoc&&(a=null);var i=a?a.completion:{};return i.type=e,i.arg=t,a?(this.method="next",this.next=a.finallyLoc,v):this.complete(i)},complete:function(e,t){if("throw"===e.type)throw e.arg;return"break"===e.type||"continue"===e.type?this.next=e.arg:"return"===e.type?(this.rval=this.arg=e.arg,this.method="return",this.next="end"):"normal"===e.type&&t&&(this.next=t),v},finish:function(e){for(var t=this.tryEntries.length-1;t>=0;--t){var r=this.tryEntries[t];if(r.finallyLoc===e)return this.complete(r.completion,r.afterLoc),P(r),v}},catch:function(e){for(var t=this.tryEntries.length-1;t>=0;--t){var r=this.tryEntries[t];if(r.tryLoc===e){var n=r.completion;if("throw"===n.type){var o=n.arg;P(r)}return o}}throw new Error("illegal catch attempt")},delegateYield:function(t,r,n){return this.delegate={iterator:j(t),resultName:r,nextLoc:n},"next"===this.method&&(this.arg=e),v}},t}function asyncGeneratorStep(e,t,r,n,o,a,i){try{var c=e[a](i),s=c.value}catch(e){return void r(e)}c.done?t(s):Promise.resolve(s).then(n,o)}function _asyncToGenerator(e){return function(){var t=this,r=arguments;return new Promise((function(n,o){var a=e.apply(t,r);function i(e){asyncGeneratorStep(a,n,o,i,c,"next",e)}function c(e){asyncGeneratorStep(a,n,o,i,c,"throw",e)}i(void 0)}))}}function _classCallCheck(e,t){if(!(e instanceof t))throw new TypeError("Cannot call a class as a function")}function _defineProperties(e,t){for(var r=0;r<t.length;r++){var n=t[r];n.enumerable=n.enumerable||!1,n.configurable=!0,"value"in n&&(n.writable=!0),Object.defineProperty(e,_toPropertyKey(n.key),n)}}function _createClass(e,t,r){return t&&_defineProperties(e.prototype,t),r&&_defineProperties(e,r),Object.defineProperty(e,"prototype",{writable:!1}),e}function _toPropertyKey(e){var t=_toPrimitive(e,"string");return"symbol"===_typeof(t)?t:String(t)}function _toPrimitive(e,t){if("object"!==_typeof(e)||null===e)return e;var r=e[Symbol.toPrimitive];if(void 0!==r){var n=r.call(e,t||"default");if("object"!==_typeof(n))return n;throw new TypeError("@@toPrimitive must return a primitive value.")}return("string"===t?String:Number)(e)}Office.onReady((function(e){e.host===Office.HostType.Excel&&(document.getElementById("sendEmail").onclick=sendEmail,document.getElementById("createSampleData").onclick=createSampleData)}));var DialogAPIAuthProvider=function(){function e(){_classCallCheck(this,e)}var t,r;return _createClass(e,[{key:"getAccessToken",value:(r=_asyncToGenerator(_regeneratorRuntime().mark((function e(){return _regeneratorRuntime().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:if(!this._accessToken){e.next=4;break}return e.abrupt("return",this._accessToken);case 4:return e.abrupt("return",this.login());case 5:case"end":return e.stop()}}),e,this)}))),function(){return r.apply(this,arguments)})},{key:"login",value:(t=_asyncToGenerator(_regeneratorRuntime().mark((function e(){var t=this;return _regeneratorRuntime().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return e.abrupt("return",new Promise((function(e,r){var n=location.href.substring(0,location.href.lastIndexOf("/"))+"/consent.html";Office.context.ui.displayDialogAsync(n,{height:60,width:60},(function(n){if(n.status===Office.AsyncResultStatus.Failed)r(n.error);else{var o=n.value;o.addEventHandler(Office.EventType.DialogEventReceived,(function(e){r(e.error)})),o.addEventHandler(Office.EventType.DialogMessageReceived,(function(n){var a=JSON.parse(n.message);o.close(),"success"===a.status?(t._accessToken=a.result,e(t._accessToken)):r(a.result)}))}}))})));case 1:case"end":return e.stop()}}),e)}))),function(){return t.apply(this,arguments)})}]),e}(),dialogAPIAuthProvider=new DialogAPIAuthProvider;function showStatus(e,t){$(".status").empty(),$("<div/>",{class:"status-card ms-depth-4 ".concat(t?"error-msg":"success-msg")}).append($("<p/>",{class:"ms-fontSize-24 ms-fontWeight-bold",text:t?"An error occurred":"Success"})).append($("<p/>",{class:"ms-fontSize-16 ms-fontWeight-regular",text:e})).appendTo(".status")}function clearStatus(){$(".status").empty()}function createSampleData(){return _createSampleData.apply(this,arguments)}function _createSampleData(){return _createSampleData=_asyncToGenerator(_regeneratorRuntime().mark((function e(){return _regeneratorRuntime().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return e.prev=0,e.next=3,Excel.run(function(){var e=_asyncToGenerator(_regeneratorRuntime().mark((function e(t){var r,n;return _regeneratorRuntime().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return t.workbook.worksheets.getItemOrNullObject("Sample").delete(),r=t.workbook.worksheets.add("Sample"),(n=r.tables.add("A1:E1",!0)).name="InvoiceTable",n.getRange().numberFormat="@",n.getHeaderRowRange().values=[["Email","Name","Invoice Number","Amount","Due Date"]],n.rows.add(0,[["client1@email.com","John","INV001","$500","2023-11-15"],["client2@email.com","Sarah","INV002","$750","2023-11-20"],["client3@microsoft.com","Michael","INV003","$300","2023-11-10"],["client4@microsoft.com","Lisa","INV004","$900","2023-11-15"]]),n.getRange().format.autofitColumns(),n.getRange().format.autofitRows(),r.activate(),e.next=13,t.sync();case 13:case"end":return e.stop()}}),e)})));return function(t){return e.apply(this,arguments)}}());case 3:e.next=8;break;case 5:e.prev=5,e.t0=e.catch(0),showStatus("Exception when creating sample data: ".concat(JSON.stringify(e.t0)),!0);case 8:case"end":return e.stop()}}),e,null,[[0,5]])}))),_createSampleData.apply(this,arguments)}function getStr(e){var t=e.match(/\<\<(.*?)\>\>/g);if(t){for(var r=0;r<t.length;r++)t[r]=t[r].replace(/\<\<|\>\>/g,"");return t[0]}}function sendEmail(e){return _sendEmail.apply(this,arguments)}function _sendEmail(){return _sendEmail=_asyncToGenerator(_regeneratorRuntime().mark((function e(t){var r;return _regeneratorRuntime().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return t.preventDefault(),clearStatus(),r=MicrosoftGraph.Client.initWithMiddleware({authProvider:dialogAPIAuthProvider}),e.next=5,Excel.run(function(){var e=_asyncToGenerator(_regeneratorRuntime().mark((function e(t){var n,o,a,i,c,s,u,l,f,h,p,m,y,v,g,d,b,w,x,k;return _regeneratorRuntime().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:if(e.prev=0,n=$("#Subject").val(),o=$("#ToLine").val(),a=$("#Content").val(),!(o&&n&&a)){e.next=97;break}if(null==(i=getStr(o))){e.next=94;break}return(c=t.workbook.tables.getItem("InvoiceTable").columns.getItemOrNullObject(i)).load(),e.next=11,t.sync();case 11:if(c.isNullObject){e.next=91;break}s=1,u=c.values,l=n.toString();case 15:if(!(s<u.length)){e.next=88;break}f="",h=0;case 18:if(!(h<l.length)){e.next=45;break}if("<"!=l[h]){e.next=41;break}return p=getStr(l.substring(h)),(m=t.workbook.tables.getItem("InvoiceTable").columns.getItemOrNullObject(p)).load(),e.next=25,t.sync();case 25:if(m.isNullObject){e.next=37;break}y=m.values,f+=y[s][0];case 28:if(!(h<l.length)){e.next=35;break}if(">"!=l[h]){e.next=32;break}return h+=1,e.abrupt("break",35);case 32:h+=1,e.next=28;break;case 35:e.next=39;break;case 37:return showStatus('There is no corresponding column name as "'.concat(p,'" in Subject.'),!0),e.abrupt("return");case 39:e.next=42;break;case 41:f+=l[h];case 42:h++,e.next=18;break;case 45:v=a.toString(),g="",d=0;case 48:if(!(d<v.length)){e.next=75;break}if("<"!=v[d]){e.next=71;break}return b=getStr(v.substring(d)),(w=t.workbook.tables.getItem("InvoiceTable").columns.getItemOrNullObject(b)).load(),e.next=55,t.sync();case 55:if(w.isNullObject){e.next=67;break}x=w.values,g+=x[s][0];case 58:if(!(d<v.length)){e.next=65;break}if(">"!=v[d]){e.next=62;break}return d+=1,e.abrupt("break",65);case 62:d+=1,e.next=58;break;case 65:e.next=69;break;case 67:return showStatus('There is no corresponding column name as "'.concat(b,'" in Content.'),!0),e.abrupt("return");case 69:e.next=72;break;case 71:g+=v[d];case 72:d++,e.next=48;break;case 75:return e.prev=75,k={message:{subject:f,body:{contentType:"Text",content:g},toRecipients:[{emailAddress:{address:u[s][0]}}]}},e.next=79,r.api("me/SendMail").post(k);case 79:e.next=85;break;case 81:e.prev=81,e.t0=e.catch(75),console.log("Error: ".concat(JSON.stringify(e.t0))),showStatus("Exception sending emails via Graph: ".concat(JSON.stringify(e.t0)),!0);case 85:s++,e.next=15;break;case 88:showStatus("Already sent ".concat(s-1," emails via Microsoft Graph."),!1),e.next=92;break;case 91:showStatus('There is no corresponding column name as "'.concat(i,'" in ToLine.'),!0);case 92:e.next=95;break;case 94:showStatus("There is no corresponding column name in ToLine.",!0);case 95:e.next=98;break;case 97:showStatus("Please fill in all the fields.",!0);case 98:e.next=104;break;case 100:e.prev=100,e.t1=e.catch(0),console.log("Error: ".concat(JSON.stringify(e.t1))),showStatus("Exception sending emails via Graph: ".concat(JSON.stringify(e.t1)),!0);case 104:case"end":return e.stop()}}),e,null,[[0,100],[75,81]])})));return function(t){return e.apply(this,arguments)}}());case 5:case"end":return e.stop()}}),e)}))),_sendEmail.apply(this,arguments)}
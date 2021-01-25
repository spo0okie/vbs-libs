URL = {
    encode : function(s){return encodeURIComponent(s).replace(/'/g,"%27").replace(/"/g,"%22")},
    decode : function(s){return decodeURIComponent(s.replace(/\+/g,  " "))},
    addslashes: function (s) {
    if (!s) return '';
    return s.replace(/\\/g, '\\\\').
        replace(/\u0008/g, '\\b').
        replace(/\t/g, '\\t').
        replace(/\n/g, '\\n').
        replace(/\f/g, '\\f').
        replace(/\r/g, '\\r').
        //replace(/'/g, '\\\'').
        replace(/"/g, '\\"');
     }

}
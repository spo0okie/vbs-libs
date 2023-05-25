/*
 * pwgen.js
 *
 * Copyright (C) 2003-2006 KATO Kazuyoshi <kzys@8-p.info>
 * Updates by Frank4DD (C) 2010
 *
 * This program is a JavaScript port of pwgen.
 * The original C source code written by Theodore Ts'o.
 * <http://sourceforge.net/projects/pwgen/>
 * 
 * This file may be distributed under the terms of the GNU General
 * Public License.
 */

var PWGen = Class.create();

PWGen.prototype = {
    initialize: function() {
        this.maxLength = 8;
        this.includeCapitalLetter = true;
        this.includeNumber = true;
    },

    generate0: function() {
        var result = "";
        var prev = 0;
        var isFirst = true;
        
        var requested = 0;
        if (this.includeCapitalLetter) {
            requested |= this.INCLUDE_CAPITAL_LETTER;
        }
        
        if (this.includeNumber) {
            requested |= this.INCLUDE_NUMBER;
        }
        
        if (this.includeSpecial) {
            requested |= this.INCLUDE_SPECIAL;
        }
        
        var shouldBe = (Math.random() < 0.5) ? this.VOWEL : this.CONSONANT;
        
        while (result.length < this.maxLength) {
            i = Math.floor((this.ELEMENTS.length - 1) * Math.random());
            str = this.ELEMENTS[i][0];
            flags = this.ELEMENTS[i][1];

            /* Filter on the basic type of the next element */
            if ((flags & shouldBe) == 0)
                continue;
            /* Handle the NOT_FIRST flag */
            if (isFirst && (flags & this.NOT_FIRST))
                continue;
            /* Don't allow VOWEL followed a Vowel/Dipthong pair */
            if ((prev & this.VOWEL) && (flags & this.VOWEL) && (flags & this.DIPTHONG))
                continue;
            /* Don't allow us to overflow the buffer */
            if ( (result.length + str.length) > this.maxLength)
                continue;
            
            
            if (requested & this.INCLUDE_CAPITAL_LETTER) {
                if ((isFirst || (flags & this.CONSONANT)) &&
                    (Math.random() > 0.3)) {
                    str = str.slice(0, 1).toUpperCase() + str.slice(1, str.length);
                    requested &= ~this.INCLUDE_CAPITAL_LETTER;
                }
            }
            
            /*
             * OK, we found an element which matches our criteria,
             * let's do it!
             */
            result += str;
            
            
            if (requested & this.INCLUDE_NUMBER) {
                if (!isFirst && (Math.random() < 0.3)) {
                    if ( (result.length + str.length) > this.maxLength)
                        result = result.slice(0,-1);
                    result += Math.floor(10 * Math.random()).toString();
                    requested &= ~this.INCLUDE_NUMBER;
                    
                    isFirst = true;
                    prev = 0;
                    shouldBe = (Math.random() < 0.5) ? this.VOWEL : this.CONSONANT;
                    continue;
                }
            }
            

            if (requested & this.INCLUDE_SPECIAL) {
                if (!isFirst && (Math.random() < 0.3)) {
                    if ( (result.length + str.length) > this.maxLength)
                        result = result.slice(0,-1);
                var possible = "!@#$^*()-_+?=./:',";
                result += possible.charAt(Math.floor(Math.random() * possible.length));
                requested &= ~this.INCLUDE_SPECIAL;

                    isFirst = true;
                    prev = 0;
                    shouldBe = (Math.random() < 0.5) ? this.VOWEL : this.CONSONANT;
                    continue;
                }
            }

            /*
             * OK, figure out what the next element should be
             */
            if (shouldBe == this.CONSONANT) {
                shouldBe = this.VOWEL;
            } else { /* should_be == VOWEL */
                if ((prev & this.VOWEL) ||
                    (flags & this.DIPTHONG) || (Math.random() > 0.3)) {
                    shouldBe = this.CONSONANT;
                } else {
                    shouldBe = this.VOWEL;
                }
            }
            prev = flags;
            isFirst = false;
        }
        
        if (requested & (this.INCLUDE_NUMBER | this.INCLUDE_SPECIAL | this.INCLUDE_CAPITAL_LETTER))
            return null;
        
        return result;
    },

    generate: function() {
        var result = null;

        while (! result)
            result = this.generate0();
        
        return result;
    },

    INCLUDE_NUMBER: 1,
    INCLUDE_SPECIAL: 1 << 1 << 1,
    INCLUDE_CAPITAL_LETTER: 1 << 1,

    CONSONANT: 1,
    VOWEL:     1 << 1,
    DIPTHONG:  1 << 2,
    NOT_FIRST: 1 << 3
};

PWGen.prototype.ELEMENTS = [
    [ "a",  PWGen.prototype.VOWEL ],
    [ "ae", PWGen.prototype.VOWEL | PWGen.prototype.DIPTHONG ],
    [ "ah", PWGen.prototype.VOWEL | PWGen.prototype.DIPTHONG ],
    [ "ai", PWGen.prototype.VOWEL | PWGen.prototype.DIPTHONG ],
    [ "b",  PWGen.prototype.CONSONANT ],
    [ "c",  PWGen.prototype.CONSONANT ],
    [ "ch", PWGen.prototype.CONSONANT | PWGen.prototype.DIPTHONG ],
    [ "d",  PWGen.prototype.CONSONANT ],
    [ "e",  PWGen.prototype.VOWEL ],
    [ "ee", PWGen.prototype.VOWEL | PWGen.prototype.DIPTHONG ],
    [ "ei", PWGen.prototype.VOWEL | PWGen.prototype.DIPTHONG ],
    [ "f",  PWGen.prototype.CONSONANT ],
    [ "g",  PWGen.prototype.CONSONANT ],
    [ "gh", PWGen.prototype.CONSONANT | PWGen.prototype.DIPTHONG | PWGen.prototype.NOT_FIRST ],
    [ "h",  PWGen.prototype.CONSONANT ],
    [ "i",  PWGen.prototype.VOWEL ],
    [ "ie", PWGen.prototype.VOWEL | PWGen.prototype.DIPTHONG ],
    [ "j",  PWGen.prototype.CONSONANT ],
    [ "k",  PWGen.prototype.CONSONANT ],
    [ "l",  PWGen.prototype.CONSONANT ],
    [ "m",  PWGen.prototype.CONSONANT ],
    [ "n",  PWGen.prototype.CONSONANT ],
    [ "ng", PWGen.prototype.CONSONANT | PWGen.prototype.DIPTHONG | PWGen.prototype.NOT_FIRST ],
    [ "o",  PWGen.prototype.VOWEL ],
    [ "oh", PWGen.prototype.VOWEL | PWGen.prototype.DIPTHONG ],
    [ "oo", PWGen.prototype.VOWEL | PWGen.prototype.DIPTHONG],
    [ "p",  PWGen.prototype.CONSONANT ],
    [ "ph", PWGen.prototype.CONSONANT | PWGen.prototype.DIPTHONG ],
    [ "qu", PWGen.prototype.CONSONANT | PWGen.prototype.DIPTHONG],
    [ "r",  PWGen.prototype.CONSONANT ],
    [ "s",  PWGen.prototype.CONSONANT ],
    [ "sh", PWGen.prototype.CONSONANT | PWGen.prototype.DIPTHONG],
    [ "t",  PWGen.prototype.CONSONANT ],
    [ "th", PWGen.prototype.CONSONANT | PWGen.prototype.DIPTHONG],
    [ "u",  PWGen.prototype.VOWEL ],
    [ "v",  PWGen.prototype.CONSONANT ],
    [ "w",  PWGen.prototype.CONSONANT ],
    [ "x",  PWGen.prototype.CONSONANT ],
    [ "y",  PWGen.prototype.CONSONANT ],
    [ "z",  PWGen.prototype.CONSONANT ],
];

// SIG // Begin signature block
// SIG // MIIIXwYJKoZIhvcNAQcCoIIIUDCCCEwCAQExDzANBglg
// SIG // hkgBZQMEAgEFADB3BgorBgEEAYI3AgEEoGkwZzAyBgor
// SIG // BgEEAYI3AgEeMCQCAQEEEBDgyQbOONQRoqMAEEvTUJAC
// SIG // AQACAQACAQACAQACAQAwMTANBglghkgBZQMEAgEFAAQg
// SIG // lkyEQ+nigqou5Wd5lrTRo8T5+xkxj/hXvHiXE+AOT7qg
// SIG // ggWcMIIFmDCCA4CgAwIBAgIBAzANBgkqhkiG9w0BAQsF
// SIG // ADBtMQswCQYDVQQGEwJSVTENMAsGA1UECAwEVXJhbDEU
// SIG // MBIGA1UEBwwLQ2hlbHlhYmluc2sxETAPBgNVBAoMCFJl
// SIG // dmlha2luMQswCQYDVQQLDAJJVDEZMBcGA1UEAwwQcmV2
// SIG // aWFraW4tcm9vdC1DQTAeFw0yMzA1MjUxNTM3MDBaFw0y
// SIG // NDA2MDMxNTM3MDBaMGMxCzAJBgNVBAYTAlJVMQ0wCwYD
// SIG // VQQIDARVcmFsMQ0wCwYDVQQHDARDaGVsMREwDwYDVQQK
// SIG // DAhSZXZpYWtpbjELMAkGA1UECwwCSVQxFjAUBgNVBAMM
// SIG // DXJldmlha2luLWNvZGUwggEiMA0GCSqGSIb3DQEBAQUA
// SIG // A4IBDwAwggEKAoIBAQCtsuYd7CVRsLwbN6ybLrnCr72O
// SIG // nqGhfdASM37B9yC8+b5nnbw6EqDEN2IHpy32wOoThAlg
// SIG // zPna/D5/VX/TYuLR/1vjW+vRQPKbJi8m97BMr8PemMWl
// SIG // w6mjl9x4qW0x4irIwXra/Z4R34BgrY8ZACZRah0riiWY
// SIG // GXPvCw3ZjNYMXRJF4rVKJ6c/PNg1bNlML1Q8oHcy3MPC
// SIG // CVCHF/Qf3Bl/l76GKJhylViC5/ZiX34LfzCopdK1xnnY
// SIG // 45cP1c83pQH2IE3ucjGMwzWDYCwTNAeYi69aaK40fGHC
// SIG // Z9EJg6sS1RnEyCpp+Sj23T/GOJyTxM4kaiPmlMDZoCAq
// SIG // UndLk6HVAgMBAAGjggFLMIIBRzAJBgNVHRMEAjAAMBEG
// SIG // CWCGSAGG+EIBAQQEAwIFoDAzBglghkgBhvhCAQ0EJhYk
// SIG // T3BlblNTTCBHZW5lcmF0ZWQgQ2xpZW50IENlcnRpZmlj
// SIG // YXRlMB0GA1UdDgQWBBSXtltT7BkMs4W7USOsFdk+mc0S
// SIG // HjAfBgNVHSMEGDAWgBSNQkTnQD4Z5d3UogsBh0kUyrwl
// SIG // pzAOBgNVHQ8BAf8EBAMCBeAwJwYDVR0lBCAwHgYIKwYB
// SIG // BQUHAwIGCCsGAQUFBwMEBggrBgEFBQcDAzA4BgNVHR8E
// SIG // MTAvMC2gK6AphidodHRwOi8vcGtpLnJldmlha2luLm5l
// SIG // dC9jcmwvcm9vdC1jYS5jcmwwPwYIKwYBBQUHAQEEMzAx
// SIG // MC8GCCsGAQUFBzAChiNodHRwOi8vcGtpLnJldmlha2lu
// SIG // L25ldC9yb290LWNhLmNydDANBgkqhkiG9w0BAQsFAAOC
// SIG // AgEAix6Hc2aULCO6RiT4W5PIiB9zQgA4BGT3W5YdSttn
// SIG // gGhnmWDEfT2bhB/ZnRLkrtZeL/sYDj94FIfKZMvFTsNN
// SIG // CUeDNiV9hQyJrsrI9Gq3nkgcnCOGc/9mqqL7ItS33s1M
// SIG // ltSXVA7sLhoQ65yPrP70kd3681COUsCYOq7hroIR3Th4
// SIG // L8INGLvUR+Xll1sunIHrnuiTD/GZFNemDec0f3n8mNKp
// SIG // 5KiWuYlNYv0Zg//rTvCZfk2Y74Mk/2lCeABVKcQoJai+
// SIG // XiSN0mq1b6RlFmfbiuzU3iudZ3SKHKEd3reGBXZxD7b1
// SIG // QubveA17QKbgzwjT6DX9ISFjbIOuB9HUo3Bl7VLZ4DyH
// SIG // 2mt0z+UC1zpE9DLFzoawf4f5/KN6mixGX9Q7tSQQCOKo
// SIG // Jiyk7Y+0aLXhK7RmJdDK3vIieJkXSx0ip1SXdRYgr0sQ
// SIG // VsNq2D2SYJ0A1r2wWJ4sNuiHnDuxWuxLsAdC0rZTlKis
// SIG // 21i4uOIr3BCj2MFdTTdkeX5xB979r/8MLBdrDlzoVxMz
// SIG // tEWwXdNlqiCQosIMVq44bJF1zjFPD6pYk0JgEF9y8wTd
// SIG // G2LyGFjTqJYyCrKrWFkQa8GX6pazj4EarEpNjdVC6IXJ
// SIG // YRa4vRqUEWfS9WeTGlIR9hJyqtHKAc9N82lwrhTlPhh+
// SIG // lkL15ZPRXnnd5aICNgQpndNfyBIxggIbMIICFwIBATBy
// SIG // MG0xCzAJBgNVBAYTAlJVMQ0wCwYDVQQIDARVcmFsMRQw
// SIG // EgYDVQQHDAtDaGVseWFiaW5zazERMA8GA1UECgwIUmV2
// SIG // aWFraW4xCzAJBgNVBAsMAklUMRkwFwYDVQQDDBByZXZp
// SIG // YWtpbi1yb290LUNBAgEDMA0GCWCGSAFlAwQCAQUAoHww
// SIG // EAYKKwYBBAGCNwIBDDECMAAwGQYJKoZIhvcNAQkDMQwG
// SIG // CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisG
// SIG // AQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIKmUG9bEaz9w
// SIG // nF+/iXrdALBV5IupT+5fsEZkhmH0CuRXMA0GCSqGSIb3
// SIG // DQEBAQUABIIBAHxBt/yyILpuzH/kjchUFS87A5d26T2c
// SIG // WkKhn2clPGtNQix2oBBQDdkc2jb8CzNLrZqXBppwURNI
// SIG // 72Aq7TF8p4fJZie57rcTdg4y0WOgUgMUQwlCkwCwIU8j
// SIG // jifKq1dWi6JRVo4C8fC6wWQQxLJSbqCPoJ70Ui/vGzRK
// SIG // yHvL588rIHveqJmrHRJ9L2q9ORQ3trWAx71Qbo+Odv3N
// SIG // rpHZToLFqxlPpsHSprIBbCiB2ea2A+askElPDVQZqxBY
// SIG // E7sI6wtxNxCyQqxJZHmtkJ9LvsxJutXO4oE6XUkaf95O
// SIG // knL63LLUHQwzFLL50wsFi/Fb2nXM95m+5ZR+ZW+UHtINYjE=
// SIG // End signature block

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
// SIG // MIIH0QYJKoZIhvcNAQcCoIIHwjCCB74CAQExDzANBglg
// SIG // hkgBZQMEAgEFADB3BgorBgEEAYI3AgEEoGkwZzAyBgor
// SIG // BgEEAYI3AgEeMCQCAQEEEBDgyQbOONQRoqMAEEvTUJAC
// SIG // AQACAQACAQACAQACAQAwMTANBglghkgBZQMEAgEFAAQg
// SIG // lkyEQ+nigqou5Wd5lrTRo8T5+xkxj/hXvHiXE+AOT7qg
// SIG // ggUsMIIFKDCCAxCgAwIBAgIBADANBgkqhkiG9w0BAQsF
// SIG // ADBtMQswCQYDVQQGEwJSVTENMAsGA1UECAwEVXJhbDEU
// SIG // MBIGA1UEBwwLQ2hlbHlhYmluc2sxETAPBgNVBAoMCFJl
// SIG // dmlha2luMQswCQYDVQQLDAJJVDEZMBcGA1UEAwwQcmV2
// SIG // aWFraW4tcm9vdC1DQTAeFw0yMzA1MjQwNDU2NTdaFw0y
// SIG // NDA2MDIwNDU2NTdaMGAxCzAJBgNVBAYTAlJVMQ0wCwYD
// SIG // VQQIDARVcmFsMQ0wCwYDVQQHDARDaGVsMREwDwYDVQQK
// SIG // DAhSZXZpYWtpbjELMAkGA1UECwwCSVQxEzARBgNVBAMM
// SIG // CnNjcmlwdHNpZ24wggEiMA0GCSqGSIb3DQEBAQUAA4IB
// SIG // DwAwggEKAoIBAQDBTtnKwGde6qQttu1TOo/JIGTZ2hoa
// SIG // HIGDBFKgexeDT8choad2DXRQzxGyu2y9w7djwuthEODY
// SIG // KLVf12PcofOKnowAoSIqQ7VW77I8I4VLI7hi0VDGZ9V9
// SIG // W4pC/mcJjkaEMSAFj6/CST5tpeczI2KxYpM1f+mEWGiu
// SIG // TkB3K3jVhsaDCuWZYZoszAJkUgp3SevPyqA6JuqzwpHD
// SIG // aDbNG5ohd1MwcwvRKab6HNwkEprYyTiX6uWZ8rBGyIGE
// SIG // 4ZtshlAt6yyn6U/tYREG9+pA9CzoPHfB3gh6taeR0/25
// SIG // oeZ5WYHuGMNeHaHYeeIXKS9gfPh3ND/fJGQaTljVSGX5
// SIG // e3StAgMBAAGjgd8wgdwwHQYDVR0OBBYEFDc+8unMGviq
// SIG // cvfVA1vXi3LqheoJMB8GA1UdIwQYMBaAFKJJoRQ/bOk/
// SIG // S1B2wDmCrQ0ZzJbKMA8GA1UdEwEB/wQFMAMBAf8wDgYD
// SIG // VR0PAQH/BAQDAgGGMDgGA1UdHwQxMC8wLaAroCmGJ2h0
// SIG // dHA6Ly9wa2kucmV2aWFraW4ubmV0L2NybC9yb290LWNh
// SIG // LmNybDA/BggrBgEFBQcBAQQzMDEwLwYIKwYBBQUHMAKG
// SIG // I2h0dHA6Ly9wa2kucmV2aWFraW4vbmV0L3Jvb3QtY2Eu
// SIG // Y3J0MA0GCSqGSIb3DQEBCwUAA4ICAQCyB0c0PKF0ffSX
// SIG // RmTBaqNWVOEpokgkdJbUNhVhKL4d7MR2wF1GX6rTeGTD
// SIG // hF4p1R3N6wRR0AAFVfp63st3w51XoQbJmGInJ7IFgrB2
// SIG // 7G6XzFVkp0llNu/1ygiqHm8v7JZEhdiqCun+JDd0ata4
// SIG // HKz2lca85tg2wnDfm0n3N7cdI56UkB+dKAzMLINVNT9X
// SIG // GSF70kXtCSPeLPDorVge0oWLxDvUiYAzlLvXk2+MTlrJ
// SIG // ka3R/s84X5W6CP9JJptIuzVuSd5ETB+tz/6xid2ELhNK
// SIG // ihkETnTViqdKp0CFGS/tRSDnfQ7Kp+Udr/SL7V/cg6Kh
// SIG // y8tXMCW+EJQBhrAGhudOvnIcFtTrUmhjupqMUuaLqDVY
// SIG // ACSwtmuihx7RAKREee0d8DJ99P3unNqfThtTPfHCzgeU
// SIG // Yk+i505Y8Op7G286bAwMv+m6SvnOT8vexSzJ3c77Vuyv
// SIG // HEU49MkgZAhpajQjTeOq2Kj3o1m+jxQ3OkWgMD6EMoJ8
// SIG // PIQS1XPhXcZ1N81uheeUf9EX13m32CulsDHmOnhQcT56
// SIG // jKt/9dn4jqHodqEodaz2Jb/tu7u6uIHmuaB2g6DTRxAO
// SIG // v33V/0yI40FG0SPAoNsWNVFySO5UwnewXA6H1hWEFezZ
// SIG // UPWWnqWb+F2uNUC8gl7Uguc2q3pJ5RhoJX+TxgBIt3oW
// SIG // SrZ8foMC3jGCAf0wggH5AgEBMHIwbTELMAkGA1UEBhMC
// SIG // UlUxDTALBgNVBAgMBFVyYWwxFDASBgNVBAcMC0NoZWx5
// SIG // YWJpbnNrMREwDwYDVQQKDAhSZXZpYWtpbjELMAkGA1UE
// SIG // CwwCSVQxGTAXBgNVBAMMEHJldmlha2luLXJvb3QtQ0EC
// SIG // AQAwDQYJYIZIAWUDBAIBBQCgXjAQBgorBgEEAYI3AgEM
// SIG // MQIwADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAv
// SIG // BgkqhkiG9w0BCQQxIgQgqZQb1sRrP3CcX7+Jet0AsFXk
// SIG // i6lP7l+wRmSGYfQK5FcwDQYJKoZIhvcNAQEBBQAEggEA
// SIG // B5Y+UnC5VBi7/Gegc7t1htXVE8fNeBXecRl1pxwHRhmb
// SIG // cZdZJUG/CmyykRSUEqUJ1U56C7QWwJ1z+u0huK76Nwqj
// SIG // xX1wTsbqFA24YvvrIco103tcPoklEdKVPP0zvAKOnP6y
// SIG // wQH9+pW9dmAKzsIZkhD3iLEdjP9HZwZqVhegpZIdMwxD
// SIG // /6vKAw3CTKFnoBdfvth0rI+2KAls3EmE69R+jMLWMFTg
// SIG // yqvpysvLGIlxtInDnhkXoNE65qpOAie1I1JassUyd2V9
// SIG // uK2I+PzFusW3vnz+gN09ObNhBJXuJeCn8EplF9swR7IG
// SIG // DGYHEEjY9GOaPU+Wq5brm0W/QFRoPDeeIw==
// SIG // End signature block

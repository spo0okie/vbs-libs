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
// SIG // Begin signature block
// SIG // MIIH0QYJKoZIhvcNAQcCoIIHwjCCB74CAQExDzANBglg
// SIG // hkgBZQMEAgEFADB3BgorBgEEAYI3AgEEoGkwZzAyBgor
// SIG // BgEEAYI3AgEeMCQCAQEEEBDgyQbOONQRoqMAEEvTUJAC
// SIG // AQACAQACAQACAQACAQAwMTANBglghkgBZQMEAgEFAAQg
// SIG // 00uOrcBX+o81jbBM1uJ9QwZrejpkDWrKkQPA6bFBVrig
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
// SIG // BgkqhkiG9w0BCQQxIgQgUJpJRSreI9f7Nwbhgfycd5U8
// SIG // L+4oNuXG4sxdpoPLDK8wDQYJKoZIhvcNAQEBBQAEggEA
// SIG // QULu1F5oEk5gAl56f1nDbwV0c3K4lIziM1f2hGrKto1Z
// SIG // D2+OUr8rCneJZ3+y8mK43T8Prf9aXskOUS6eEcKD8u3o
// SIG // oAf21q/9OsaF70FV4at5hA1hGxiMIAlS96eP9F385Cb3
// SIG // pcFOYtXSfNQJ1SmSCxHrES7bXTmHYeyirdEuJXXK/nCK
// SIG // b/QFONbNRwZoxb8Mb9A9Wa0KAjxVB0LJuEzYS0XJXiiy
// SIG // XAMkr+XxSdOlTARsVkhS3PDEpCjobES8gHPb6v92S2Ih
// SIG // 9EeT7k99p3UMKT5PBLGizMska3DicWPk8JUitXxWti5w
// SIG // UG7vl+4JkPVRi59FRkCYbvv4SkfWZGXqIg==
// SIG // End signature block

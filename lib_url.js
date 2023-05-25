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
// SIG // MIIIXwYJKoZIhvcNAQcCoIIIUDCCCEwCAQExDzANBglg
// SIG // hkgBZQMEAgEFADB3BgorBgEEAYI3AgEEoGkwZzAyBgor
// SIG // BgEEAYI3AgEeMCQCAQEEEBDgyQbOONQRoqMAEEvTUJAC
// SIG // AQACAQACAQACAQACAQAwMTANBglghkgBZQMEAgEFAAQg
// SIG // 00uOrcBX+o81jbBM1uJ9QwZrejpkDWrKkQPA6bFBVrig
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
// SIG // AQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIFCaSUUq3iPX
// SIG // +zcG4YH8nHeVPC/uKDblxuLMXaaDywyvMA0GCSqGSIb3
// SIG // DQEBAQUABIIBAA89Saya6smv1ZM1xpFROnQjQJl15HQp
// SIG // gMzs0yylXMxSJyme1svH/r70v0kd2Ssu8wmmQ31g1GLl
// SIG // PxQuH+/yvSLsGLdZXA9cbZWiIuAT2hdBsx17J42jgslJ
// SIG // cfBQaHYnw70T2Yw7dM66scXez4+Oa/cfTHBhUNWONCGp
// SIG // Ji1jjScojFH8+uBg02lxeSU1FxSD2MfBE+akxV/nbGQB
// SIG // q3pag6EGkJlRAAkEe5RwOYE2WZMB/bcnoNKVm0pu8PgO
// SIG // /azpBRkQjDs1bprApYpMQ/gJTwTeBPndjOK2cSarjIdm
// SIG // v+muhkkHwg3wE5ElaGlAPwQNwX1HfrAgA3ETmwUUDs2oI4Y=
// SIG // End signature block

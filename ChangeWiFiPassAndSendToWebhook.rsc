# ----------------------------------
# Variables: Customize these
# ----------------------------------
:global webhookBaseUrl "https://endpointWebHookAddress"
:global webhookUser "Username"
:global webhookPass "Password"
:local wifiInterface "wlan1"
:local wifiSecurityProfile "default"
:local passwordLength 16
# ----------------------------------
# Get SSID from interface
# ----------------------------------
:local ssid [/interface wireless get $wifiInterface ssid]

# ----------------------------------
# Generate password using system time as seed
# ----------------------------------
:local chars "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
:local newPass ""
:local charsLen [:len $chars]

:for i from=1 to=$passwordLength do={
    # use seconds + loop index to generate weak pseudo-random index
    :local timeSeed [/system clock get time]
    :local second [:tonum [:pick $timeSeed 6 8]]
    :local index (($second + $i * 13) % $charsLen)
    :local c [:pick $chars $index]
    :set newPass ($newPass . $c)
    :delay 1ms
}

# ----------------------------------
# Update WiFi password
# ----------------------------------
/interface wireless security-profiles set \
    [find name=$wifiSecurityProfile] \
    wpa2-pre-shared-key=$newPass

# ----------------------------------
# Send new password + SSID via GET to webhook with basic auth
# ----------------------------------
:local url "$webhookBaseUrl?newpass=$newPass&ssid=$ssid"

/tool fetch \
    url=$url \
    user=$webhookUser \
    password=$webhookPass \
    mode=https \
    http-method=get \
    keep-result=no

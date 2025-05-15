# This script is tested on MikroTik RouterOS 7.18.2 
# It forwards the last SMS message to the telegram and removes it.
# Use on your own responsibility.
# The code is delivered with no warraties. Proceed with caution.
# --- CONFIG ---
:local botToken "your telegram token"
:local chatID "channel id"  ;

# --- FETCH LAST SMS ---
:local smsID ""
:foreach sms in=[/tool sms inbox print as-value] do={
    :set smsID ($sms->".id")
}

# --- IF NONE, EXIT CLEANLY ---
:if ($smsID = "") do={
    :log warning "SMS-TGM: No SMS found"
    :return ""
}

# --- GET FIELDS ---
:local sender [/tool sms inbox get $smsID phone]
:local timestamp [/tool sms inbox get $smsID timestamp]
:local message [/tool sms inbox get $smsID message]

:log info ("SMS-TGM: From " . $sender)
:log info ("SMS-TGM: Time " . $timestamp)
:log info ("SMS-TGM: Message " . $message)

# --- URL ENCODE (space and newline only) ---
:local safeMessage ""
:for i from=0 to=([:len $message] - 1) do={
    :local ch [:pick $message $i]
    :if ($ch = " ") do={
        :set safeMessage ($safeMessage . "%20")
    } else={
        :if ($ch = "\n") do={
            :set safeMessage ($safeMessage . "%0A")
        } else={
            :set safeMessage ($safeMessage . $ch)
        }
    }
}

# --- BUILD TELEGRAM URL ---
:local url "https://api.telegram.org/bot"
:set url ($url . $botToken)
:set url ($url . "/sendMessage?chat_id=" . $chatID)
:set url ($url . "&text=From:%20" . $sender . "%0AAt:%20" . $timestamp . "%0AMessage:%20" . $safeMessage)

:log info ("SMS-TGM: Telegram URL: " . $url)

# --- SEND TO TELEGRAM ---
/tool fetch url=$url keep-result=no
:log info "SMS-TGM: Sent message to Telegram"

# --- DELETE SMS (optional) ---
/tool sms inbox remove $smsID
:log info ("SMS-TGM: Deleted SMS ID " . $smsID)

# --- CLEAN EXIT ---
#:put ""
:return ""

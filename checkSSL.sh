#!/bin/bash
# https://serverfault.com/questions/661978/displaying-a-remote-ssl-certificate-details-using-cli-tools
echo | openssl s_client -showcerts -servername $1 -connect $1:443 2>/dev/null | openssl x509 -inform pem -noout -text|grep -i after | awk '{for(i=4;i<=NF;++i)printf $i""FS ; print ""}'

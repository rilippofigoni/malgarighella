    #!/bin/bash                                      
    # Get all domains
    _dom=$@
     
    # Die if no domains are given
    [ $# -eq 0 ] && { echo "Usage: $0 domain1.com domain2.com ..."; exit 1; }
    for d in $_dom
    do
    _ip=$(host $d | grep 'has add' | head -1 | awk '{ print $4}')
    [ "$_ip" == "" ] && { echo "Error: $d is not valid domain or dns error."; continue; }
    echo "Getting information for domain: $d [ $_ip ]..."
    #nmap --traceroute --script traceroute-geolocation.nse -p 80 "$_ip"
    #nmap -p 80 --script dns-brute.nse "$_ip"
    nmap -v -d --script ssl-heartbleed --script-args vulns.showall -sV "$_ip"
    #nmap -sS -O "$_ip"/24  
    #nmap -O -sV –version-intensity 9 -p1-100 -Pn –osscan-guess "$_ip"
    #nmap -sV -p 22,53,110,143,4564 "$_ip" 
    echo ""
    done
    
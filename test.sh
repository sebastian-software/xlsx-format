#!/bin/bash
set -euxo pipefail
TZONES=(America/New_York Europe/London Asia/Seoul America/Los_Angeles Europe/Berlin Asia/Kolkata Asia/Shanghai America/Cancun America/Anchorage America/Barbados Asia/Tokyo  America/Cayman Pacific/Honolulu America/Mexico_City Asia/Hong_Kong Europe/Paris Atlantic/Azores)

export NVM_DIR="$HOME/.nvm"
[ -s "$NVM_DIR/nvm.sh" ] && \. "$NVM_DIR/nvm.sh"  # This loads nvm

if [ -e datetest.js ]; then
	nvm use 20
	# sudo n 20;
	for TZ in ${TZONES[@]}; do
		echo "$TZ"
		env TZ="$TZ" mocha -R dot datetest.js
	done
fi

# min test 
for n in 20 10 0.8 0.10 0.12 4 6 8 12 14 16 18 22 24; do
	nvm use $n
	# sudo n $n
	env WTF=1 make testdot_misc
	for TZ in ${TZONES[@]}; do
		# sudo n $n
		nvm use $n
		env WTF=1 TZ="$TZ" make testdot_misc
	done
done

# full test
for n in 24 16 0.12; do
	for TZ in America/New_York Asia/Seoul Asia/Kolkata Europe/Paris; do
		# sudo n $n
		nvm use $n
		env WTF=1 TZ="$TZ" make testdot
	done
done

# bun
for TZ in ${TZONES[@]}; do
	echo "$TZ";
	env TZ="$TZ" WTF=1 make test-bun_misc;
done

# deno
for TZ in ${TZONES[@]}; do
	echo "$TZ";
	env TZ="$TZ" WTF=1 make test-deno_misc;
	env TZ="$TZ" WTF=1 make test-denocp_misc;
done

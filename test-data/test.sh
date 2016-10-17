#!/bin/bash

set -ex

mkdir data

# First test that we can save attachments
../ytnef/ytnef -f data winmail.dat
for s in *.tnef
do
  ../ytnef/ytnef -f data $s
done

diff results data


# Now do some basic "deep checking" of specific fields.

../ytnefprint/ytnefprint ./winmail.dat  | grep -A 1 PR_MESSAGE_DELIVERY_TIME | grep "Tuesday June 17, 2003 3:23:00 pm"
../ytnefprint/ytnefprint ./winmail.dat  | grep -A 1 PR_TNEF_CORRELATION_KEY  | grep "[00000000CFF233DD7623274EB78A4912C247A0E464532700.]"
../ytnefprint/ytnefprint ./winmail.dat  | grep -A 1 PR_RTF_SYNC_BODY_CRC     | grep 872404792
../ytnefprint/ytnefprint ./winmail.dat  | grep -A 1 PR_RTF_SYNC_BODY_COUNT   | grep 90
../ytnefprint/ytnefprint ./winmail.dat  | grep -A 18 PR_RTF_COMPRESSED       | grep '\pard Casdasdfasdfasd\\par'

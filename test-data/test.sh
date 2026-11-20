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

../ytnefprint/ytnefprint ./ytnef-oob-TNEFVersion-SwapWord | grep 'ERROR: Field with size of 0'

# Regression files: these must not crash, and nothing may be written
# outside the output directory.
rm -rf regress-out
mkdir -p regress-out/out
../ytnef/ytnef -f regress-out/out regress/path-traversal.tnef
test ! -e regress-out/ESC1
test -e regress-out/out/.._ESC1
../ytnef/ytnef -f regress-out/out regress/nested-tnef.tnef
../ytnef/ytnef -f regress-out/out regress/short-attachment.tnef
../ytnefprint/ytnefprint regress/empty-string-prop.tnef > /dev/null
../ytnefprint/ytnefprint regress/duplicate-attribute.tnef > /dev/null
../ytnefprint/ytnefprint regress/early-attachment-attr.tnef > /dev/null
../ytnefprint/ytnefprint regress/duplicate-mapi.tnef > /dev/null
! ../ytnef/ytnef -s 0 -f regress-out/out winmail.dat
! ../ytnef/ytnef -s 4096 -f regress-out/out winmail.dat
../ytnef/ytnef -s 2047 -f regress-out/out winmail.dat
rm -rf regress-out

#!/bin/bash

set -ex -o pipefail

# Inputs come from $srcdir (set by automake); outputs go to the build dir.
srcdir=${srcdir:-.}

rm -rf data
mkdir data

# First test that we can save attachments
../ytnef/ytnef -f data "$srcdir"/winmail.dat
for s in "$srcdir"/*.tnef
do
  ../ytnef/ytnef -f data $s
done

diff "$srcdir"/results data
rm -rf data


# Now do some basic "deep checking" of specific fields.

../ytnefprint/ytnefprint "$srcdir"/winmail.dat  | grep -A 1 PR_MESSAGE_DELIVERY_TIME | grep "Tuesday June 17, 2003 3:23:00 pm"
../ytnefprint/ytnefprint "$srcdir"/winmail.dat  | grep -A 1 PR_TNEF_CORRELATION_KEY  | grep "[00000000CFF233DD7623274EB78A4912C247A0E464532700.]"
../ytnefprint/ytnefprint "$srcdir"/winmail.dat  | grep -A 1 PR_RTF_SYNC_BODY_CRC     | grep 872404792
../ytnefprint/ytnefprint "$srcdir"/winmail.dat  | grep -A 1 PR_RTF_SYNC_BODY_COUNT   | grep 90
../ytnefprint/ytnefprint "$srcdir"/winmail.dat  | grep -A 18 PR_RTF_COMPRESSED       | grep '\pard Casdasdfasdfasd\\par'

../ytnefprint/ytnefprint "$srcdir"/ytnef-oob-TNEFVersion-SwapWord | grep 'ERROR: Field with size of 0'

# Regression files: these must not crash, and nothing may be written
# outside the output directory.
rm -rf regress-out
mkdir -p regress-out/out
../ytnef/ytnef -f regress-out/out "$srcdir"/regress/path-traversal.tnef
test ! -e regress-out/ESC1
test -e regress-out/out/.._ESC1
../ytnef/ytnef -f regress-out/out "$srcdir"/regress/nested-tnef.tnef
../ytnef/ytnef -f regress-out/out "$srcdir"/regress/short-attachment.tnef
../ytnefprint/ytnefprint "$srcdir"/regress/empty-string-prop.tnef > /dev/null
../ytnefprint/ytnefprint "$srcdir"/regress/duplicate-attribute.tnef > /dev/null
../ytnefprint/ytnefprint "$srcdir"/regress/early-attachment-attr.tnef > /dev/null
../ytnefprint/ytnefprint "$srcdir"/regress/duplicate-mapi.tnef > /dev/null
# "! cmd" does not trip set -e, so check the exit status explicitly.
if ../ytnef/ytnef -s 0 -f regress-out/out "$srcdir"/winmail.dat; then exit 1; fi
if ../ytnef/ytnef -s 4096 -f regress-out/out "$srcdir"/winmail.dat; then exit 1; fi
../ytnef/ytnef -s 2047 -f regress-out/out "$srcdir"/winmail.dat
# a parse error must give a nonzero exit status
if ../ytnef/ytnef -f regress-out/out "$srcdir"/regress/truncated-field.tnef; then exit 1; fi
rm -rf regress-out

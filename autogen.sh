#!/bin/sh
set -e
mkdir -p m4
# Homebrew installs GNU libtoolize as glibtoolize; tell autoreconf to use it.
# Running it again after autoreconf would remove the compile/missing helpers.
if [ -z "$LIBTOOLIZE" ] && ! command -v libtoolize >/dev/null 2>&1 &&
   command -v glibtoolize >/dev/null 2>&1; then
  LIBTOOLIZE=glibtoolize
  export LIBTOOLIZE
fi
autoreconf -vfi

#!/bin/bash
set -ev
mkdir -p m4
autoreconf -vfi
./configure --disable-dependency-tracking
if [ "$TRAVIS_OS_NAME" == "osx" ] || [ "$CC" == "clang" ]; then
  make CFLAGS="-fsanitize=address -g" LDFLAGS="-fsanitize=address"
else
  make
fi
sudo make install

export LD_LIBRARY_PATH=/usr/local/lib:${LD_LIBRARY_PATH}
cd test-data
./test.sh

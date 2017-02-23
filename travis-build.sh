#!/bin/bash
set -ev
mkdir -p m4
autoreconf -vfi
./configure --disable-dependency-tracking
make CFLAGS="-fsanitize=address -g"
sudo make install

export LD_LIBRARY_PATH=/usr/local/lib:${LD_LIBRARY_PATH}
cd test-data
./test.sh

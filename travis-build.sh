#!/bin/bash
set -ev
cd ytneflib
./autogen.sh
./configure --disable-dependency-tracking
make
sudo make install
cd ../ytnef
./autogen.sh
./configure --disable-dependency-tracking
make

export LD_LIBRARY_PATH=/usr/local/lib:${LD_LIBRARY_PATH}
cd ../test-data
./test.sh

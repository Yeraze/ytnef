#!/bin/bash
set -ev
export CC=$1
cd ytneflib
./autogen.sh
./configure
make
sudo make install
cd ../ytnef
./autogen.sh
./configure 
make

export LD_LIBRARY_PATH=/usr/local/lib:${LD_LIBRARY_PATH}
cd ../test-data
./test.sh

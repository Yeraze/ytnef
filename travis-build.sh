#!/bin/bash
export CC=$1
set -ev
cd ytneflib
./autogen.sh
./configure
make
sudo make install
cd ../ytnef
./autogen.sh
./configure 
make

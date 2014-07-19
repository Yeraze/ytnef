#!/bin/bash
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

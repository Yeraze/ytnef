#!/bin/bash
aclocal
autoheader
automake --add-missing 
autoconf
autoreconf -vfi

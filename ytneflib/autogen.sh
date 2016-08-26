#!/bin/bash
aclocal
autoheader
automake --add-missing
autoconf
autoreconf -vfi
case `uname` in Darwin*) glibtoolize --copy --force ;;
  *) libtoolize --copy --force ;; esac

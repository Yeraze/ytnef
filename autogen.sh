#!/bin/bash
mkdir -p m4
autoreconf -vfi
case `uname` in Darwin*) glibtoolize --copy --force ;;
  *) ;; esac

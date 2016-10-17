#!/bin/bash
autoreconf -vfi
case `uname` in Darwin*) glibtoolize --copy --force ;;
  *) ;; esac

aclocal
autoheader
automake --add-missing
autoconf
case `uname` in Darwin*) glibtoolize --copy --force ;;
  *) libtoolize --copy --force ;; esac

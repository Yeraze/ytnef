# Process this file with autoconf to produce a configure script.
AC_INIT(libytnef, 1.7, randall.hand@gmail.com)
AM_INIT_AUTOMAKE
AC_CONFIG_SRCDIR([ytnef.c])
AM_CONFIG_HEADER([config.h])

# Checks for programs.
AC_PROG_CC
AC_PROG_LIBTOOL

# Checks for libraries.

# Checks for header files.
AC_HEADER_STDC
AC_CHECK_HEADERS([stdlib.h string.h])

# Checks for typedefs, structures, and compiler characteristics.
AC_C_BIGENDIAN

# Checks for library functions.
AC_FUNC_MALLOC
AC_CHECK_FUNCS([memset])

AC_CONFIG_FILES([Makefile])
AC_OUTPUT
AC_OUTPUT(libytnef.spec)

# Process this file with autoconf to produce a configure script.
AC_INIT(ytnef, 2.0-cvs, yerase@yeraze.com)
AM_INIT_AUTOMAKE
AC_CONFIG_SRCDIR([src/ytnef/main.c])
AC_CONFIG_HEADER([config.h])

# Checks for programs.
AC_PROG_CC

# Checks for libraries.

# Checks for header files.
AC_HEADER_STDC
AC_CHECK_HEADERS([stdlib.h string.h])

# Checks for typedefs, structures, and compiler characteristics.
AC_C_BIGENDIAN

# Checks for library functions.
AC_FUNC_MALLOC
AC_CHECK_FUNCS([memset strstr])

AC_CONFIG_FILES([Makefile])
AC_OUTPUT

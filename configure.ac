# Increment the:
# MAJOR version when you make incompatible API changes,
# MINOR version when you add functionality in a backwards-compatible manner
# PATCH version when you make backwards-compatible bug fixes.
m4_define([version_major],1)
m4_define([version_minor],9)
m4_define([version_micro],0)
m4_define([version_triplet],version_major.version_minor.version_micro)

AC_PREREQ([2.63])
AC_INIT([ytnef],
        [version_triplet],
        [randall.hand@gmail.com],
        [ytnef])

AC_CONFIG_MACRO_DIR([m4])
AC_CONFIG_SRCDIR([ytnef/main.c])
AC_CONFIG_HEADERS([config.h])

AM_INIT_AUTOMAKE([1.9 foreign dist-bzip2])

m4_ifdef([AM_SILENT_RULES],[AM_SILENT_RULES([yes])])

if test -z "$enable_maintainer_mode"; then
  enable_maintainer_mode=yes
fi
AM_MAINTAINER_MODE([enable])

# Libtool versioning
m4_define([lt_revision], m4_if(m4_eval(version_minor%2),1,0,version_micro))
m4_define([lt_age], m4_eval(version_minor*100+version_micro-lt_revision))
m4_define([lt_current], m4_eval(version_major+lt_age))
m4_define([lt_triplet],lt_current:lt_revision:lt_age)

LT_VERSION_INFO="lt_triplet()"
AC_SUBST([LT_VERSION_INFO])

# Checks for programs.
AC_PROG_CC
AC_PROG_LIBTOOL
#AC_USE_SYSTEM_EXTENSIONS
# Checks for libraries.

# Checks for header files.
AC_HEADER_STDC
AC_CHECK_HEADERS([stdlib.h string.h])

# Checks for typedefs, structures, and compiler characteristics.
AC_C_BIGENDIAN

# Checks for library functions.
AC_FUNC_MALLOC
AC_CHECK_FUNCS([memset strstr])

AC_SYS_LARGEFILE

# Initialize libtool
LT_PREREQ([2.2])
LT_INIT
LT_LANG([C])

AC_SUBST([YTNEF_MAJOR_VERSION],[version_major])
AC_SUBST([YTNEF_MINOR_VERSION],[version_minor])
AC_SUBST([YTNEF_MICRO_VERSION],[version_micro])

AC_SUBST([AM_CPPFLAGS])
AC_SUBST([AM_CFLAGS])
AC_SUBST([AM_CXXFLAGS])
AC_SUBST([AM_LDFLAGS])

AC_CONFIG_FILES([
Makefile
ytnef/Makefile
ytnefprint/Makefile
lib/Makefile
test-data/Makefile
])

AC_OUTPUT
AC_OUTPUT([ytnef.spec])

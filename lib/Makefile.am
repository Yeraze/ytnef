lib_LTLIBRARIES = libytnef.la

libytnef_la_SOURCES = \
	ytnef.c

libytnef_la_LDFLAGS = \
	-no-undefined

nobase_include_HEADERS = \
	ytnef.h \
	mapi.h mapidefs.h mapitags.h \
	tnef-types.h tnef-errors.h

pkgconfigdir = $(libdir)/pkgconfig
pkgconfig_DATA = libytnef.pc

EXTRA_DIST = libytnef.spec

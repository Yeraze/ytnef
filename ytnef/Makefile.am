bin_PROGRAMS = ytnef
bin_SCRIPTS = ytnefprocess
dist_man_MANS = ytnef.1 ytnefprocess.1

ytnef_SOURCES = main.c
ytnef_CPPFLAGS = -I$(top_srcdir)/lib
ytnef_LDADD = $(top_builddir)/lib/libytnef.la
ytnef_CFLAGS = \
        $(YTNEF_CFLAGS) \
        $(AM_CFLAGS) \
        -I$(top_srcdir) \
        $(NULL)


EXTRA_DIST = vcal.c vcard.c vtask.c utility.c settings.c settings.h \
	ytnefprocess README.ytnefprocess

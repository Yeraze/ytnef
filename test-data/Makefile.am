EXTRA_DIST = \
	long-filename.tnef \
	missing-filenames.tnef \
	multi-value-attribute.tnef \
	one-file.tnef \
	rtf.tnef \
	regress/path-traversal.tnef \
	regress/nested-tnef.tnef \
	regress/short-attachment.tnef \
	regress/empty-string-prop.tnef \
	regress/truncated-field.tnef \
	regress/duplicate-attribute.tnef \
	regress/early-attachment-attr.tnef \
	test.sh \
	winmail.dat \
	results/208225_(5_seconds)_Voice_Mail.mp3 \
	results/allproductsmar2000.dat \
	results/bookmark.htm \
	results/generpts.src \
	results/TechlibDEC99-JAN00.doc \
	results/TechlibNOV99.doc \
	results/TechlibDEC99.doc \
	results/zappa_av1.jpg

check_PROGRAMS = test_decompress_rtf test_lib
test_decompress_rtf_SOURCES = test_decompress_rtf.c
test_decompress_rtf_CPPFLAGS = -I$(top_srcdir)/lib
test_decompress_rtf_LDADD = $(top_builddir)/lib/libytnef.la

test_lib_SOURCES = test_lib.c
test_lib_CPPFLAGS = -I$(top_srcdir)/lib
test_lib_LDADD = $(top_builddir)/lib/libytnef.la

TESTS = test_decompress_rtf test_lib

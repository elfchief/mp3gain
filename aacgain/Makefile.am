bin_PROGRAMS = aacgain

if do_static_build
AACGAIN_EXT_LIBS = $(top_builddir)/faad2/libfaad/libfaad.la $(top_builddir)/mpeg4ip/lib/mp4v2/libmp4v2.la
else
AACGAIN_EXT_LIBS = -lfaad -lmp4v2
endif

# set the include path found by configure
INCLUDES = -I$(top_srcdir)/mp3gain -I$(top_srcdir)/mpeg4ip/lib/mp4v2 \
	-I$(top_srcdir)/mpeg4ip/include -I$(top_srcdir)/faad2/libfaad -I$(top_srcdir)/faad2/include \
	$(all_includes)

# the library search path.
aacgain_LDFLAGS = $(all_libraries)
aacgain_SOURCES = aacgain.cpp decoder.c MP4MetaFile.cpp syntax.c
aacgain_LDADD = $(AACGAIN_EXT_LIBS)\
	$(top_builddir)/mp3gain/mpglibDBL/libmpglib.la \
	$(top_builddir)/mp3gain/libmp3gain.la

CC = cc
LINK = cc

DEBUG = -g
CFLAGS = $(DEBUG) $(DEFINES)

INVLIBS = -lm


OBJS = main.o tnef.o

default: all

all: ytnef

ytnef: $(OBJS)
	$(LINK) -o $@ $(OBJS) $(INVLIBS)


main.o: main.c *.h vcal.c vcard.c vtask.c
tnef.o: tnef.c *.h

.c.o:
	$(CC) $(DEFINES) $(CFLAGS) $(INC) -c $<

clean:
	'rm' -f *.o ytnef
# DO NOT DELETE

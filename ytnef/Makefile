## SGI CC Compiler
#CC = CC -n32
## SUN CC Compiler
#CC = g++
## Linux CC Compiler
CC = gcc
LINK = gcc

DEBUG = -g
CFLAGS = $(DEBUG) $(DEFINES)

## SGI Library line  
#INVLIBS = -lglut -lGL -lGLU -lXmu -lXext -lXt -lX11 -lm
## SUN Library line  
#INVLIBS = -L/opt/gfx/lib -lXi -lXmu -lXext -lXt -lX11 -lglut -lGL -lGLU -lglut -lm
## Linux Library line   
INVLIBS = -lm


OBJS = main.o tnef.o

default: all

all: ytnef

ytnef: $(OBJS)
	$(LINK) -o $@ $(OBJS) $(INVLIBS)


main.o: main.c *.h
tnef.o: tnef.c *.h

.c.o:
	$(CC) $(DEFINES) $(CFLAGS) $(INC) -c $<

clean:
	'rm' -f *.o ytnef
# DO NOT DELETE

# Purpose: makefile for toolbar example (Unix)
# Created: 2000-03-14

CXX = $(shell wx-config --cxx)

PROGRAM = mp3gain

OBJECTS = $(PROGRAM).o mainFrame.o mp3Info.o 

# implementation

.SUFFIXES:	.o .cpp

.cpp.o :
	$(CXX) -c `wx-config --cxxflags` -o $@ $<

all:    $(PROGRAM)

$(PROGRAM):	$(OBJECTS)
	$(CXX) -o $(PROGRAM) $(OBJECTS) `wx-config --libs`

clean:
	rm -f *.o $(PROGRAM)

SOURCES += $$PWD/xlsx_wrapper.cpp

HEADERS += $$PWD/xlsx_wrapper.h

INCLUDEPATH *= $$PWD/
DEPENDPATH  *= $$PWD/

include($$PWD/../QXlsx/QXlsx.pri)
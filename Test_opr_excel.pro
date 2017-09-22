#-------------------------------------------------
#
# Project created by QtCreator 2017-09-22T09:41:51
#
#-------------------------------------------------

QT       += core gui

CONFIG  += qaxcontainer

greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

TARGET = Test_opr_excel
TEMPLATE = app


SOURCES += main.cpp\
        mainwindow.cpp \
    excelengine.cpp

HEADERS  += mainwindow.h \
    excelengine.h

FORMS    += mainwindow.ui

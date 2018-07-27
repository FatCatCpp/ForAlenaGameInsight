#-------------------------------------------------
#
# Project created by QtCreator 2018-07-11T09:55:40
#
#-------------------------------------------------

QT       += core gui
QT       += axcontainer
QT       += sql

greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

TARGET = for_fat
TEMPLATE = app


SOURCES += main.cpp\
        mainwindow.cpp \
    calculate.cpp

HEADERS  += mainwindow.h \
    calculate.h

FORMS    += mainwindow.ui

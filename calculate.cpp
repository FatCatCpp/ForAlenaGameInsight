#include "calculate.h"

calculate::calculate() {}

void calculate::startProcess(QString addr)
{
    QString cmd = "cmd /Q /C \"start";
    QString screen = "\"";
    QString strName = cmd + " " + addr + screen;
    QProcess::startDetached(strName);
}

int calculate::CalcWordCount(QString s)
{
    return s.split(QRegExp("(\\s|\\n|\\r)+"), QString::SkipEmptyParts).count();
}

int calculate::CalcSymbCount(QString s)
{
    return s.count(/*QLatin1Char('#')*/) - 1;
}

int calculate::CalcDigitCount(QString s)
{
    s = s.remove(QRegExp("[\\D]"));
    return (s.count() - 1);
}

int calculate::CalcSpaceCount(QString s)
{
    return s.count(" ");
}

int calculate::CalcConcreteWord(QString s, QString cont)
{
    return s.count(cont);
}

int calculate::CalcFilesCount(QStringList s)
{
    return s.size();
}

int calculate::CalcSizeFiles(QStringList s)
{
    int sizeAll = 0, size;
    for (int i = 0; i < s.size(); i++) {
        QFile myFile(s[i]);
        if (myFile.open(QIODevice::ReadOnly)) {
            size = myFile.size();
            myFile.close();
        }
        sizeAll += size;
    }
    sizeAll /= 1024;
    return (sizeAll);
}

int calculate::CalcPuncMarks(QString s)
{
    QString cont = ",.:;!?&";
    int m_counter = 0;
    for (int i = 0; i < s.size(); i++) {
        for (int j = 0; j < cont.size(); j++) {
            if (s[i] == cont[j])
                m_counter++;
        }
    }
    return m_counter;
}



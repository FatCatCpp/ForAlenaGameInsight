#ifndef CALCULATE_H
#define CALCULATE_H

#include <QStringList>
#include <QProcess>
#include <QFile>

class calculate
{
public:
    calculate();
    int CalcSizeFiles(QStringList); // считать кол-во
    int CalcWordCount(QString);     // считать кол-во слов
    int CalcSymbCount(QString);     // считать кол-во символов
    int CalcDigitCount(QString);    // считать кол-во цифр
    int CalcPuncMarks(QString);     // считать кол-во знаков препинания
    int CalcSpaceCount(QString);    // считать кол-во пробелов
    int CalcFilesCount(QStringList);// считать кол-во загруженных файлов
    int CalcConcreteWord(QString,
                         QString);  // считать сколько раз слово (2й параметр) встретится в тексте
    void startProcess(QString);     // запустить excel-файл
};

#endif // CALCULATE_H
